import os
from flask import Flask, request, render_template, send_file, flash, redirect, url_for
from werkzeug.utils import secure_filename
from main import AmazonSalesAnalysis

app = Flask(__name__)
app.config['SECRET_KEY'] = 'your-secret-key-here'  # Required for flashing messages
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

# Ensure upload folder exists
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        if 'file' not in request.files:
            flash('No file part')
            return redirect(request.url)
        
        file = request.files['file']
        if file.filename == '':
            flash('No selected file')
            return redirect(request.url)
        
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(filepath)
            
            # Process the file
            analyzer = AmazonSalesAnalysis(filepath)
            if analyzer.load_data():
                analyzer.generate_excel_report()
                return send_file('sales_analysis_report.xlsx',
                               as_attachment=True,
                               download_name='sales_analysis_report.xlsx')
            else:
                flash('Error processing the file. Please ensure it has the correct format.')
                return redirect(request.url)
        else:
            flash('Allowed file types are .xlsx and .xls')
            return redirect(request.url)
    
    return render_template('upload.html')

if __name__ == '__main__':
    app.run(debug=True) 