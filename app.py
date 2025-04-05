import os
import uuid
import time
import threading
import traceback # For detailed error logging
from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
from pdf2docx import Converter # Use pdf2docx for conversion
import fitz # PyMuPDF, often bundled with pdf2docx or installed alongside

# --- Configuration ---
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
ALLOWED_EXTENSIONS = {'pdf'}

# --- Flask App Initialization ---
app = Flask(__name__)
CORS(app) # Enable CORS for all routes, adjust origins in production

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER
# Optional: Limit upload size (e.g., 50MB)
# app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024

# --- Ensure upload and output directories exist ---
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# --- In-memory storage for job status (Replace with DB/Redis in production) ---
conversion_jobs = {}

# --- Helper Functions ---
def allowed_file(filename):
    """Checks if the filename has an allowed extension."""
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def perform_conversion(job_id, input_path, output_format, optimizer_setting, password=None):
    """
    Performs the actual PDF to DOCX conversion using pdf2docx.
    Updates the conversion_jobs dictionary with status, results, stats, and health report.
    """
    start_time = conversion_jobs[job_id]['start_time']
    stats = {}
    health_report = {
        "encryption": "None",
        "scanned_pages": "Detection not run (OCR skipped)", # Updated as OCR is skipped
        "font_issues": "Basic check passed",
        "warnings": []
    }
    output_paths = {}
    base_filename = os.path.splitext(os.path.basename(input_path))[0]
    # We are focusing on DOCX output as discussed
    docx_filename = f"{job_id}_{base_filename}.docx"
    docx_output_path = os.path.join(OUTPUT_FOLDER, docx_filename)

    try:
        conversion_jobs[job_id]['status'] = 'analyzing'
        conversion_jobs[job_id]['progress'] = 5 # Small progress for analysis step

        # --- Pre-analysis with PyMuPDF (fitz) ---
        pdf_doc = None
        try:
            pdf_doc = fitz.open(input_path)
            if pdf_doc.is_encrypted:
                # Option A: Reject password-protected PDFs
                health_report["encryption"] = "Encrypted (Not Supported)"
                raise ValueError("Password-protected PDFs are not supported in this version.")
                # The code below for password handling is now effectively disabled.
                # health_report["encryption"] = "Encrypted"
                # if not password:
                #     # Try authenticating with empty password (common for some PDFs)
                #     if not pdf_doc.authenticate(''):
                #         raise ValueError("PDF is password protected, but no password provided or incorrect.")
                # elif not pdf_doc.authenticate(password):
                #     raise ValueError("Incorrect password provided for protected PDF.")
                # # If authentication succeeded, encryption is handled
                # health_report["encryption"] = "Handled (Password Provided)"

            stats["page_count"] = pdf_doc.page_count
            # Simple font check (more complex analysis is possible)
            fonts = pdf_doc.get_page_fonts(0) # Check fonts on first page
            if not fonts:
                health_report["font_issues"] = "No fonts detected on first page (potential issue)."

            # Estimate time based on page count (very rough)
            estimated_total_time = 10 + stats["page_count"] * 0.5 # 10s base + 0.5s per page
            conversion_jobs[job_id]['estimated_time'] = estimated_total_time

        except Exception as analysis_error:
            print(f"Job {job_id}: Error during PDF analysis: {analysis_error}")
            health_report["warnings"].append(f"Analysis Error: {analysis_error}")
            # Decide if analysis error is fatal or just a warning
            if isinstance(analysis_error, ValueError) and "password" in str(analysis_error).lower():
                raise analysis_error # Fatal password error
            # Otherwise, continue to conversion attempt but log warning

        finally:
            if pdf_doc:
                pdf_doc.close()

        # --- Conversion with pdf2docx ---
        conversion_jobs[job_id]['status'] = 'processing'
        conversion_jobs[job_id]['progress'] = 10 # Progress before starting convert

        # Note: pdf2docx.Converter doesn't offer fine-grained optimization params like "Quality/Balanced/Compact".
        # We'll ignore the optimizer_setting for now, unless specific parameters are found later.
        # The library aims for good general conversion.
        if optimizer_setting == "compact":
            health_report["warnings"].append("Compact optimization setting ignored (using pdf2docx defaults).")
        elif optimizer_setting == "quality":
            health_report["warnings"].append("Quality optimization setting ignored (using pdf2docx defaults).")


        # Initialize Converter
        cv = Converter(pdf_file=input_path, password=password)
        # Perform conversion - This is a blocking call
        cv.convert(docx_filename=docx_output_path, start=0, end=None) # Convert all pages
        cv.close()

        # --- Post-conversion ---
        conversion_jobs[job_id]['progress'] = 95 # Progress after conversion before final updates

        output_paths['docx'] = docx_output_path
        conversion_jobs[job_id]['output_paths'] = output_paths
        conversion_jobs[job_id]['status'] = 'complete'
        conversion_jobs[job_id]['progress'] = 100
        conversion_jobs[job_id]['estimated_time'] = 0

        # Final stats
        stats["processing_time_seconds"] = round(time.time() - start_time, 2)
        if not health_report["warnings"]:
            health_report["warnings"].append("None") # Add None if no warnings occurred

        conversion_jobs[job_id]['stats'] = stats
        conversion_jobs[job_id]['health_report'] = health_report

        print(f"Job {job_id}: Conversion successful.")

    except Exception as e:
        print(f"Job {job_id}: Error during conversion: {e}")
        print(traceback.format_exc()) # Print full traceback for debugging
        conversion_jobs[job_id]['status'] = 'error'
        conversion_jobs[job_id]['error'] = f"Conversion failed: {e}"
        conversion_jobs[job_id]['progress'] = 0 # Reset progress on error
        conversion_jobs[job_id]['stats'] = stats # Keep any stats gathered before error
        conversion_jobs[job_id]['health_report'] = health_report # Keep health report info

    finally:
        # Clean up input file? Optional, depends on requirements.
        # Be careful if multiple threads might access it.
        # if os.path.exists(input_path):
        #     os.remove(input_path)
        pass


# --- API Routes ---

@app.route('/')
def index():
    """A basic route for the homepage."""
    return jsonify({"message": "PDF to Word Converter API is running!"})

@app.route('/api/convert', methods=['POST'])
def upload_and_convert():
    """
    Handles file upload, starts conversion process in background.
    Returns a job ID.
    """
    if 'pdf_file' not in request.files:
        return jsonify({"message": "No 'pdf_file' part in the request"}), 400

    file = request.files['pdf_file']
    # Only DOCX is supported by this implementation
    # output_format = request.form.get('output_format', 'docx')
    output_format = 'docx' # Hardcode to docx for now
    optimizer_setting = request.form.get('optimizer_setting', 'balanced')
    password = request.form.get('password') # Get password if provided by frontend

    if file.filename == '':
        return jsonify({"message": "No selected file"}), 400

    if file and allowed_file(file.filename):
        # Generate unique job ID
        job_id = str(uuid.uuid4())
        filename = f"{job_id}_{file.filename}" # Use job_id to avoid collisions
        input_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)

        try:
            file.save(input_path)

            # Store job info (basic)
            conversion_jobs[job_id] = {
                'status': 'queued',
                'progress': 0,
                'estimated_time': None,
                'start_time': time.time(),
                'input_path': input_path,
                'output_paths': None,
                'stats': None,
                'health': None,
                'error': None
            }

            # Start conversion in a background thread
            thread = threading.Thread(target=perform_conversion, args=(job_id, input_path, output_format, optimizer_setting, password))
            thread.start()

            return jsonify({"message": "File uploaded successfully, conversion process initiated.", "job_id": job_id}), 202

        except Exception as e:
            return jsonify({"message": f"Failed to save file or start conversion: {e}"}), 500

    else:
        return jsonify({"message": "Invalid file type. Only PDF files are allowed."}), 400


@app.route('/api/progress/<job_id>', methods=['GET'])
def get_progress(job_id):
    """
    Returns the current status and progress of a conversion job.
    """
    job = conversion_jobs.get(job_id)
    if not job:
        return jsonify({"message": "Job ID not found"}), 404

    # Provide a more detailed progress response
    response = {
        "job_id": job_id,
        "status": job['status'],
        "progress": job.get('progress', 0),
        "estimated_time": job.get('estimated_time') # May be None initially or during processing
    }
    if job['status'] == 'error':
        response['message'] = job.get('error', 'An unknown error occurred during conversion.')
    elif job['status'] == 'queued':
        response['message'] = 'Conversion is queued and will start shortly.'
    elif job['status'] == 'analyzing':
        response['message'] = 'Analyzing PDF structure...'
    elif job['status'] == 'processing':
        response['message'] = 'Converting PDF to DOCX...'

    return jsonify(response), 200


@app.route('/api/results/<job_id>', methods=['GET'])
def get_results(job_id):
    """
    Returns the results (download links, stats, health report) of a completed job.
    """
    job = conversion_jobs.get(job_id)
    if not job:
        return jsonify({"message": "Job ID not found"}), 404

    if job['status'] == 'error':
        return jsonify({
            "job_id": job_id,
            "status": job['status'],
            "message": job.get('error', 'Conversion failed.'),
            "statistics": job.get('stats'), # Include stats even on error if available
            "health_report": job.get('health_report') # Include health report even on error
            }), 400 # Use a 400 or 500 status for backend error
    elif job['status'] != 'complete':
        return jsonify({"message": f"Job is not yet complete. Current status: {job['status']}"}), 202 # Accepted, but not ready

    # Generate download URLs (relative to /api/download endpoint)
    download_urls = {}
    if job.get('output_paths'):
        # Only expect 'docx' key now
        if 'docx' in job['output_paths']:
            filename = os.path.basename(job['output_paths']['docx'])
            download_urls['docx'] = f"/api/download/{filename}" # Example download route

    response = {
        "job_id": job_id,
        "status": job['status'],
        "download_urls": download_urls,
        "statistics": job.get('stats'),
        "health_report": job.get('health_report')
    }
    return jsonify(response), 200

# --- Download Route (Example) ---
# This route serves the generated files. Secure appropriately in production.
@app.route('/api/download/<filename>', methods=['GET'])
def download_file(filename):
    """Serves files from the output directory."""
    # Basic security check: prevent directory traversal
    if '..' in filename or filename.startswith('/'):
        return jsonify({"message": "Invalid filename"}), 400

    try:
        # Ensure the file being requested actually belongs to a known job output? (More secure)
        # For simplicity now, just serve if it exists in the output folder.
        return send_from_directory(app.config['OUTPUT_FOLDER'], filename, as_attachment=True)
    except FileNotFoundError:
        return jsonify({"message": "File not found"}), 404


# --- Main Execution ---
if __name__ == '__main__':
    # Note: Flask's development server is not suitable for production.
    # Use a production-ready WSGI server like Gunicorn or uWSGI.
    app.run(debug=True, host='0.0.0.0', port=5000) # Runs on port 5000