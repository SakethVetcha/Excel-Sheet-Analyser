from flask import Flask
from flask.cli import ScriptInfo
import os
import json
from main import AmazonSalesAnalysis

app = Flask(__name__)
app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'dev-key-please-change')
app.config['UPLOAD_FOLDER'] = '/tmp'  # Use /tmp for Netlify Functions

def handler(event, context):
    """Netlify Function handler"""
    # Parse the incoming request
    http_method = event.get('httpMethod', 'GET')
    path = event.get('path', '/')
    headers = event.get('headers', {})
    body = event.get('body', '')
    
    # Create a Flask request context
    with app.test_request_context(
        path=path,
        method=http_method,
        headers=headers,
        data=body
    ):
        try:
            # Handle the request with Flask
            response = app.full_dispatch_request()
            return {
                'statusCode': response.status_code,
                'headers': dict(response.headers),
                'body': response.get_data(as_text=True)
            }
        except Exception as e:
            return {
                'statusCode': 500,
                'body': json.dumps({'error': str(e)})
            }

# Your Flask routes
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