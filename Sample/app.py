import os
import sys
import importlib.util
from flask import Flask, render_template, request, send_file, flash
from werkzeug.utils import secure_filename

# Import your editor.py module
spec = importlib.util.spec_from_file_location("editor", "editor.py")
editor = importlib.util.module_from_spec(spec)
sys.modules["editor"] = editor
spec.loader.exec_module(editor)

app = Flask(__name__)
app.secret_key = 'your-secret-key-here'  # Change this for production!

# Base dir of the app (absolute path)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
app.config['UPLOAD_FOLDER'] = os.path.join(BASE_DIR, 'uploads')
app.config['PROCESSED_FOLDER'] = os.path.join(BASE_DIR, 'processed')
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB limit

# Allowed file extensions
ALLOWED_EXTENSIONS = {'pdf'}

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/process', methods=['POST'])
def process_file():
    if 'file' not in request.files:
        flash('No file part')
        return render_template('index.html')

    file = request.files['file']
    if file.filename == '':
        flash('No file selected')
        return render_template('index.html')

    if file and allowed_file(file.filename):
        os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
        os.makedirs(app.config['PROCESSED_FOLDER'], exist_ok=True)

        filename = secure_filename(file.filename)
        input_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(input_path)

        name, ext = os.path.splitext(filename)
        output_filename = f"{name}_processed{ext}"
        output_path = os.path.join(app.config['PROCESSED_FOLDER'], output_filename)

        try:
            success = editor.process_pdf(input_path, output_path)

            if success:
                # keep processed file, but clean upload
                if os.path.exists(input_path):
                    os.remove(input_path)

                flash('File processed successfully!')
                return render_template('index.html', download_file=output_filename)
            else:
                flash('Error processing PDF')
                return render_template('index.html')
        except Exception as e:
            flash(f'Error: {str(e)}')
            return render_template('index.html')

    flash('Invalid file type. Please upload a PDF file.')
    return render_template('index.html')


@app.route('/download/<filename>')
def download_file(filename):
    file_path = os.path.join(app.config['PROCESSED_FOLDER'], filename)
    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True, download_name=filename)
    flash('File not found')
    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)