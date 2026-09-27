from datetime import timedelta
import logging
import os
import io
from flask import Flask, request, render_template, jsonify
from werkzeug.middleware.proxy_fix import ProxyFix

from flask_compress import Compress
from flask_minify import minify
from apscheduler.schedulers.background import BackgroundScheduler
from openpyxl import load_workbook

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024

compress = Compress()
compress.init_app(app)

dev = "development"
# dev = "production"
ports = '8080'

# Ensure static and static/tmp directories exist
TMP_DIR = os.path.join(app.root_path, 'static', 'tmp')
os.makedirs(TMP_DIR, exist_ok=True)

ALLOWED_EXTENSIONS_EXCEL = {'xls', 'xlsx'}
ALLOWED_EXTENSIONS_IMAGE = {'png', 'jpg', 'jpeg', 'gif', 'webp', 'svg'}

def allowed_file(filename, allowed_extensions):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in allowed_extensions

def parse_excel_file(file_storage_or_path):
    workbook = load_workbook(filename=file_storage_or_path, data_only=True)
    sheet = workbook.active
    data = []
    for row in sheet.iter_rows(values_only=True):
        # normalize row values
        row_clean = ["" if cell is None else str(cell).strip() for cell in row]
        if any(row_clean): # ignore completely empty rows
            data.append(row_clean)
    return data

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == "POST":
        return voucher(request)
    else:
        return render_template('index.html')

@app.route('/api/parse-excel', methods=['POST'])
def api_parse_excel():
    if 'excel' not in request.files or request.files['excel'].filename == '':
        return jsonify({'error': 'Tidak ada file Excel yang diunggah'}), 400
    
    excel_file = request.files['excel']
    if not allowed_file(excel_file.filename, ALLOWED_EXTENSIONS_EXCEL):
        return jsonify({'error': 'Format file harus .xlsx atau .xls'}), 400
    
    try:
        raw_data = parse_excel_file(excel_file)
        if not raw_data:
            return jsonify({'error': 'File Excel kosong'}), 400
        
        headers = raw_data[0] if len(raw_data) > 0 else []
        rows = raw_data[1:] if len(raw_data) > 1 else []
        
        return jsonify({
            'success': True,
            'filename': excel_file.filename,
            'total_rows': len(rows),
            'headers': headers,
            'rows': rows
        })
    except Exception as e:
        return jsonify({'error': f'Gagal membaca Excel: {str(e)}'}), 500

def create_vouch(data, configs, templs, logo_path=None):
    if templs == 1:
        return render_template('A4.html', data=data, configs=configs, logo_path=logo_path)
    elif templs == 2:
        return render_template('A4_2.html', data=data, configs=configs, logo_path=logo_path)
    elif templs == 3:
        return render_template('thermal.html', data=data, configs=configs, logo_path=logo_path)
    elif templs == 4:
        return render_template('thermal_cafe.html', data=data, configs=configs, logo_path=logo_path)
    else:
        return 'invalid templates', 400

def voucher(request):
    if request.method == 'POST':
        if 'excel' not in request.files or request.files['excel'].filename == '':
            return "No Excel file uploaded", 400

        excel_file = request.files['excel']
        if allowed_file(excel_file.filename, ALLOWED_EXTENSIONS_EXCEL):
            excel_data = parse_excel_file(excel_file)
            excel_data_without_first_element = excel_data[1:] if len(excel_data) > 1 else []

            configs = request.form.getlist('config')
            templs = int(request.form.get('templates', 1))

            logo_set = None
            if 'logo' in request.files and request.files['logo'].filename != '':
                logo_file = request.files['logo']
                if allowed_file(logo_file.filename, ALLOWED_EXTENSIONS_IMAGE):
                    safe_filename = os.path.basename(logo_file.filename)
                    logo_path = os.path.join(TMP_DIR, safe_filename)
                    logo_file.save(logo_path)
                    logo_set = f'/static/tmp/{safe_filename}'

            if 'namevoucher' in request.form and request.form['namevoucher'].strip():
                configs.append(f"namevoucher: {request.form['namevoucher'].strip()}")
            if 'katakata' in request.form and request.form['katakata'].strip():
                configs.append(f"katakata: {request.form['katakata'].strip()}")

            return create_vouch(excel_data_without_first_element, configs, templs, logo_set)
        else:
            return "Invalid file type for Excel file", 400

def cron10():
    try:
        if os.path.exists(TMP_DIR):
            pics = [f for f in os.listdir(TMP_DIR) if os.path.isfile(os.path.join(TMP_DIR, f))]
            for entry in pics:
                try:
                    os.remove(os.path.join(TMP_DIR, entry))
                except Exception:
                    pass
    except Exception as e:
        logging.error(f"Error in cron10 cleanup: {e}")

if not app.debug or os.environ.get('WERKZEUG_RUN_MAIN') == 'true':
    try:
        scheduler = BackgroundScheduler(daemon=True)
        scheduler.add_job(cron10, 'interval', minutes=5)
        scheduler.start()
    except Exception as e:
        logging.warning(f"Scheduler error: {e}")

if __name__ == "__main__":
    with app.app_context():
        app.wsgi_app = ProxyFix(app.wsgi_app)
        if dev == 'production':
            logging.basicConfig(level=logging.ERROR, format='%(levelname)s - %(message)s')
            try:
                minify(app=app, html=True, js=False, cssless=False)
            except Exception:
                pass
            app.run(debug=False, threaded=True, use_reloader=False, host='0.0.0.0', port=int(ports))
        else:
            logging.basicConfig(level=logging.DEBUG, format='%(levelname)s - %(message)s')
            try:
                minify(app=app, html=False, js=False, cssless=False)
            except Exception:
                pass
            app.run(debug=True, threaded=True, port=int(ports))
