import os
import uuid
import time
import threading
import traceback
from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
from pdf2docx import Converter
import fitz

# --- Configuration ---
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
ALLOWED_EXTENSIONS = {'pdf'}

# --- Flask App Initialization ---
app = Flask(__name__)
CORS(app)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

conversion_jobs = {}

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def perform_conversion(job_id, input_path, output_format, optimizer_setting, original_filename=None, password=None):
    """Performs the actual PDF to DOCX conversion."""
    start_time = conversion_jobs[job_id]['start_time']
    stats = {}
    health_report = {
        "encryption": "None",
        "scanned_pages": "Detection not run (OCR skipped)",
        "font_issues": "Basic check passed",
        "warnings": []
    }
    output_paths = {}
    
    # Use original filename if provided, otherwise fall back to base filename
    if original_filename:
        base_output_name = original_filename
    else:
        base_output_name = os.path.splitext(os.path.basename(input_path))[0]
    
    docx_filename = f"{job_id}_{base_output_name}.docx"
    docx_output_path = os.path.join(OUTPUT_FOLDER, docx_filename)

    try:
        conversion_jobs[job_id]['status'] = 'analyzing'
        conversion_jobs[job_id]['progress'] = 5

        # PDF analysis
        pdf_doc = None
        try:
            pdf_doc = fitz.open(input_path)
            if pdf_doc.is_encrypted:
                health_report["encryption"] = "Encrypted (Not Supported)"
                raise ValueError("Password-protected PDFs are not supported.")
            
            stats["page_count"] = pdf_doc.page_count
            fonts = pdf_doc.get_page_fonts(0)
            if not fonts:
                health_report["font_issues"] = "No fonts detected on first page (potential issue)."

            estimated_total_time = 10 + stats["page_count"] * 0.5
            conversion_jobs[job_id]['estimated_time'] = estimated_total_time

        except Exception as analysis_error:
            print(f"Job {job_id}: Error during PDF analysis: {analysis_error}")
            health_report["warnings"].append(f"Analysis Error: {analysis_error}")
            if isinstance(analysis_error, ValueError) and "password" in str(analysis_error).lower():
                raise analysis_error

        finally:
            if pdf_doc:
                pdf_doc.close()

        # Conversion
        conversion_jobs[job_id]['status'] = 'processing'
        conversion_jobs[job_id]['progress'] = 10

        if optimizer_setting == "compact":
            health_report["warnings"].append("Compact optimization setting ignored.")
        elif optimizer_setting == "quality":
            health_report["warnings"].append("Quality optimization setting ignored.")

        cv = Converter(pdf_file=input_path, password=password)
        cv.convert(docx_filename=docx_output_path, start=0, end=None)
        cv.close()

        conversion_jobs[job_id]['progress'] = 95
        output_paths['docx'] = docx_output_path
        conversion_jobs[job_id]['output_paths'] = output_paths
        conversion_jobs[job_id]['status'] = 'complete'
        conversion_jobs[job_id]['progress'] = 100
        conversion_jobs[job_id]['estimated_time'] = 0

        stats["processing_time_seconds"] = round(time.time() - start_time, 2)
        if not health_report["warnings"]:
            health_report["warnings"].append("None")

        conversion_jobs[job_id]['stats'] = stats
        conversion_jobs[job_id]['health_report'] = health_report

        print(f"Job {job_id}: Conversion successful.")

    except Exception as e:
        print(f"Job {job_id}: Error during conversion: {e}")
        print(traceback.format_exc())
        conversion_jobs[job_id]['status'] = 'error'
        conversion_jobs[job_id]['error'] = f"Conversion failed: {e}"
        conversion_jobs[job_id]['progress'] = 0
        conversion_jobs[job_id]['stats'] = stats
        conversion_jobs[job_id]['health_report'] = health_report

@app.route('/')
def index():
    return jsonify({"message": "PDF to Word Converter API is running!"})

@app.route('/api/convert', methods=['POST'])
def upload_and_convert():
    if 'pdf_file' not in request.files:
        return jsonify({"message": "No 'pdf_file' part in the request"}), 400

    file = request.files['pdf_file']
    output_format = 'docx'
    optimizer_setting = request.form.get('optimizer_setting', 'balanced')
    password = request.form.get('password')
    original_filename = request.form.get('original_filename')

    if file.filename == '':
        return jsonify({"message": "No selected file"}), 400

    if file and allowed_file(file.filename):
        job_id = str(uuid.uuid4())
        filename = f"{job_id}_{file.filename}"
        input_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)

        try:
            file.save(input_path)

            conversion_jobs[job_id] = {
                'status': 'queued',
                'progress': 0,
                'estimated_time': None,
                'start_time': time.time(),
                'input_path': input_path,
                'original_filename': original_filename,
                'output_paths': None,
                'stats': None,
                'health': None,
                'error': None
            }

            thread = threading.Thread(
                target=perform_conversion,
                args=(job_id, input_path, output_format, optimizer_setting, original_filename, password)
            )
            thread.start()

            return jsonify({
                "message": "File uploaded successfully, conversion process initiated.",
                "job_id": job_id
            }), 202

        except Exception as e:
            return jsonify({"message": f"Failed to save file or start conversion: {e}"}), 500

    else:
        return jsonify({"message": "Invalid file type. Only PDF files are allowed."}), 400

@app.route('/api/progress/<job_id>', methods=['GET'])
def get_progress(job_id):
    job = conversion_jobs.get(job_id)
    if not job:
        return jsonify({"message": "Job ID not found"}), 404

    response = {
        "job_id": job_id,
        "status": job['status'],
        "progress": job.get('progress', 0),
        "estimated_time": job.get('estimated_time')
    }
    
    if job['status'] == 'error':
        response['message'] = job.get('error', 'An unknown error occurred.')
    elif job['status'] == 'queued':
        response['message'] = 'Conversion is queued and will start shortly.'
    elif job['status'] == 'analyzing':
        response['message'] = 'Analyzing PDF structure...'
    elif job['status'] == 'processing':
        response['message'] = 'Converting PDF to DOCX...'

    return jsonify(response), 200

@app.route('/api/results/<job_id>', methods=['GET'])
def get_results(job_id):
    job = conversion_jobs.get(job_id)
    if not job:
        return jsonify({"message": "Job ID not found"}), 404

    if job['status'] == 'error':
        return jsonify({
            "job_id": job_id,
            "status": job['status'],
            "message": job.get('error', 'Conversion failed.'),
            "statistics": job.get('stats'),
            "health_report": job.get('health_report')
        }), 400
    elif job['status'] != 'complete':
        return jsonify({
            "message": f"Job is not yet complete. Current status: {job['status']}"
        }), 202

    download_urls = {}
    if job.get('output_paths'):
        if 'docx' in job['output_paths']:
            filename = os.path.basename(job['output_paths']['docx'])
            download_urls['docx'] = f"/api/download/{filename}"

    response = {
        "job_id": job_id,
        "status": job['status'],
        "download_urls": download_urls,
        "statistics": job.get('stats'),
        "health_report": job.get('health_report')
    }
    return jsonify(response), 200

@app.route('/api/download/<filename>', methods=['GET'])
def download_file(filename):
    if '..' in filename or filename.startswith('/'):
        return jsonify({"message": "Invalid filename"}), 400

    try:
        parts = filename.split('_', 1)
        if len(parts) != 2:
            return send_from_directory(app.config['OUTPUT_FOLDER'], filename, as_attachment=True)
        
        job_id, original_output_name = parts
        job = conversion_jobs.get(job_id)
        original_filename = job.get('original_filename') if job else None
        
        if original_filename:
            download_name = f"{original_filename}.docx"
            return send_from_directory(
                app.config['OUTPUT_FOLDER'],
                filename,
                as_attachment=True,
                download_name=download_name
            )
        
        return send_from_directory(app.config['OUTPUT_FOLDER'], filename, as_attachment=True)
    
    except FileNotFoundError:
        return jsonify({"message": "File not found"}), 404

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)