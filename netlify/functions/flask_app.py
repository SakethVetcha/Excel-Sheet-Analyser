from flask import Flask, request, jsonify, send_file, render_template
import os
import json
from werkzeug.utils import secure_filename
from main import AmazonSalesAnalysis

app = Flask(__name__)
app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'dev-key-please-change')
app.config['UPLOAD_FOLDER'] = '/tmp'  # Use /tmp for Netlify Functions

ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        if 'file' not in request.files:
            return jsonify({'error': 'No file part'}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': 'No selected file'}), 400
        
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(filepath)
            
            try:
                analyzer = AmazonSalesAnalysis(filepath)
                if analyzer.load_data():
                    output_path = os.path.join(app.config['UPLOAD_FOLDER'], 'analysis_report.xlsx')
                    analyzer.generate_excel_report()
                    return send_file(output_path,
                                   as_attachment=True,
                                   download_name='sales_analysis_report.xlsx')
                else:
                    return jsonify({'error': 'Error processing file'}), 400
            except Exception as e:
                return jsonify({'error': str(e)}), 500
            finally:
                # Clean up
                try:
                    os.remove(filepath)
                    os.remove(output_path)
                except:
                    pass
        return jsonify({'error': 'Invalid file type'}), 400
    
    return render_template('upload.html') 