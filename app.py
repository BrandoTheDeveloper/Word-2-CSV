import os
import logging
import time
from datetime import datetime, timedelta
from flask import Flask, send_file, render_template, jsonify, after_this_request
from werkzeug.utils import secure_filename
from werkzeug.exceptions import BadRequest
from flask_wtf.csrf import CSRFProtect
from flask_limiter import Limiter
from flask_limiter.util import get_remote_address
from dotenv import load_dotenv
from word_to_csv import parse_data_to_csv
from flask_wtf import FlaskForm
from flask_wtf.file import FileField, FileRequired
from wtforms import SubmitField
from apscheduler.schedulers.background import BackgroundScheduler

load_dotenv()

app = Flask(__name__)
app.config['SECRET_KEY'] = os.getenv('SECRET_KEY', 'your-secret-key')
app.config['UPLOAD_FOLDER'] = os.getenv('UPLOAD_FOLDER', 'uploads')
app.config['ALLOWED_EXTENSIONS'] = set(os.getenv('ALLOWED_EXTENSIONS', 'docx').split(','))
app.config['FILE_RETENTION_PERIOD'] = int(os.getenv('FILE_RETENTION_PERIOD', 3600))  # 1 hour by default

csrf = CSRFProtect(app)
limiter = Limiter(app, key_func=get_remote_address)

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class UploadForm(FlaskForm):
    file = FileField('File', validators=[FileRequired()])
    submit = SubmitField('Upload')

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

def validate_file(file):
    if file.filename == '':
        raise BadRequest("No file selected")
    if not allowed_file(file.filename):
        raise BadRequest("File type not allowed")

def cleanup_old_files():
    current_time = datetime.now()
    retention_period = timedelta(seconds=app.config['FILE_RETENTION_PERIOD'])
    upload_folder = app.config['UPLOAD_FOLDER']

    for filename in os.listdir(upload_folder):
        file_path = os.path.join(upload_folder, filename)
        file_modified = datetime.fromtimestamp(os.path.getmtime(file_path))
        if current_time - file_modified > retention_period:
            try:
                os.remove(file_path)
                logger.info(f"Deleted old file: {file_path}")
            except Exception as e:
                logger.error(f"Failed to delete old file {file_path}: {str(e)}")

@app.route('/', methods=['GET', 'POST'])
@limiter.limit("10 per minute")
def upload_file():
    form = UploadForm()
    if form.validate_on_submit():
        file = form.file.data
        try:
            validate_file(file)
            filename = secure_filename(file.filename)
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(file_path)
            csv_file = parse_data_to_csv(file_path)
            return send_file(csv_file, as_attachment=True, download_name=f"{os.path.splitext(filename)[0]}.csv")
        except BadRequest as e:
            logger.error(f"Bad request: {str(e)}")
            return jsonify({"error": str(e)}), 400
        except Exception as e:
            logger.error(f"An error occurred: {str(e)}")
            return jsonify({"error": "Internal Server Error"}), 500
    return render_template('upload.html', form=form)

@app.route('/health')
def health_check():
    return jsonify({"status": "healthy"}), 200

@app.errorhandler(400)
def bad_request(error):
    return jsonify({"error": "Bad Request"}), 400

@app.errorhandler(500)
def server_error(error):
    return jsonify({"error": "Internal Server Error"}), 500

if __name__ == '__main__':
    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
    
    # Set up the scheduler
    scheduler = BackgroundScheduler()
    scheduler.add_job(func=cleanup_old_files, trigger="interval", minutes=15)
    scheduler.start()

    app.run(debug=False, host='0.0.0.0', port=int(os.getenv('PORT', 8080)))