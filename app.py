import os
from flask import Flask, request, render_template, send_file, flash, redirect, url_for
from werkzeug.utils import secure_filename
from main import AmazonSalesAnalysis
import time

app = Flask(__name__)
# Use environment variable for secret key in production
app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'dev-key-please-change-in-production')
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

# Ensure upload folder exists and is outside web root
UPLOAD_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'secure_uploads')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def secure_clean_old_files():
    """Clean up old uploaded and generated files"""
    for filename in os.listdir(UPLOAD_FOLDER):
        filepath = os.path.join(UPLOAD_FOLDER, filename)
        # Remove files older than 1 hour
        if os.path.isfile(filepath) and (time.time() - os.path.getmtime(filepath)) > 3600:
            try:
                os.remove(filepath)
            except:
                pass

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        # Clean old files first
        secure_clean_old_files()
        
        if 'file' not in request.files:
            flash('No file part')
            return redirect(request.url)
        
        file = request.files['file']
        if file.filename == '':
            flash('No selected file')
            return redirect(request.url)
        
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            # Generate unique filename to prevent overwrites
            unique_filename = f"{int(time.time())}_{filename}"
            filepath = os.path.join(UPLOAD_FOLDER, unique_filename)
            file.save(filepath)
            
            try:
                # Process the file
                analyzer = AmazonSalesAnalysis(filepath)
                if analyzer.load_data():
                    output_filename = f"analysis_{unique_filename}"
                    output_path = os.path.join(UPLOAD_FOLDER, output_filename)
                    analyzer.generate_excel_report()
                    # Move the generated report to secure location
                    os.rename('sales_analysis_report.xlsx', output_path)
                    return send_file(output_path,
                                   as_attachment=True,
                                   download_name='sales_analysis_report.xlsx')
                else:
                    flash('Error processing the file. Please ensure it has the correct format.')
            except Exception as e:
                flash(f'An error occurred while processing the file: {str(e)}')
            finally:
                # Clean up uploaded file
                try:
                    os.remove(filepath)
                except:
                    pass
            return redirect(request.url)
        else:
            flash('Allowed file types are .xlsx and .xls')
            return redirect(request.url)
    
    return render_template('upload.html')

if __name__ == '__main__':
    # In production, use environment variables for host and port
    host = os.environ.get('HOST', '127.0.0.1')
    port = int(os.environ.get('PORT', 5000))
    debug = os.environ.get('FLASK_DEBUG', 'False').lower() == 'true'
    
    app.run(host=host, port=port, debug=debug) 