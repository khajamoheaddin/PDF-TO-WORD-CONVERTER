# PDF to Word Converter API

A simple Flask-based backend API to convert uploaded PDF files into DOCX format. This backend is designed to work with a frontend interface (like the provided WordPress plugin) to handle file uploads and display results.

## Features

* Upload PDF files via a POST request.
* Converts PDFs to DOCX format using the `pdf2docx` library[cite: 1].
* Provides progress tracking for ongoing conversions.
* Offers download links for the converted files.
* Includes basic PDF analysis (page count, simple font check).
* Handles potential errors during conversion.
* PDF to DOCX conversion with original filename preservation
* Progress tracking during conversion
* Basic PDF health reporting
* Secure file downloads with proper filename headers

## Filename Preservation

The API now preserves the original filename during conversion:
1. Frontend sends the original filename (without extension)
2. Backend uses this name when saving converted files
3. Download endpoint sets proper `Content-Disposition` headers

Example:
- Uploaded file: `my_document.pdf`
- Converted file: `[job-id]_my_document.docx`
- Downloaded as: `my_document.docx`

## Setup & Installation

[Previous setup instructions remain the same...]

## API Changes

The `/api/download` endpoint now:
- Automatically sets the correct download filename
- Preserves the original filename from upload
- Maintains security checks against path traversal

## Requirements

* Python 3.x
* pip

## Setup & Installation

1.  **Clone the repository:**
    ```bash
    git clone <your-repository-url>
    cd python-pdf-converter-backend
    ```
2.  **(Optional but recommended) Create and activate a virtual environment:**
    ```bash
    python -m venv venv
    # On Windows: .\venv\Scripts\activate
    # On macOS/Linux: source venv/bin/activate
    ```
3.  **Install dependencies:**
    ```bash
    pip install -r requirements.txt
    ```
    This will install Flask, Flask-Cors, pdf2docx, and PyMuPDF (fitz)[cite: 1].

## Running the Server

* **For Development:**
    ```bash
    python app.py
    ```
    The API will be available at `http://127.0.0.1:5000` by default.

* **For Production (e.g., using Gunicorn, as on Render):**
    ```bash
    gunicorn app:app
    ```
    (Render typically handles this via the Start Command you configure).

## API Endpoints

The base URL will be your deployment URL (e.g., `https://your-app.onrender.com` or `http://127.0.0.1:5000` locally). All endpoints are prefixed with `/api/`.

---

### 1. Start Conversion

* **Endpoint:** `/api/convert`
* **Method:** `POST`
* **Description:** Uploads a PDF file to start the conversion process.
* **Request:** `multipart/form-data`
    * `pdf_file`: The PDF file to convert (Required).
    * `output_format`: (Optional, currently ignored - defaults to 'docx'). String, e.g., "docx", "doc", "both".
    * `optimizer_setting`: (Optional, currently ignored - defaults to 'balanced'). String, e.g., "balanced", "quality", "compact".
    * `password`: (Optional, currently **not supported**). String, password for encrypted PDFs.
* **Success Response (202 Accepted):**
    ```json
    {
      "message": "File uploaded successfully, conversion process initiated.",
      "job_id": "unique-job-identifier-string"
    }
    ```
* **Error Responses:** `400 Bad Request` (e.g., no file, invalid file type), `500 Internal Server Error`.

---

### 2. Check Progress

* **Endpoint:** `/api/progress/<job_id>`
* **Method:** `GET`
* **Description:** Poll this endpoint to check the status and progress of a conversion job.
* **URL Parameters:**
    * `job_id`: The unique identifier returned by `/api/convert`.
* **Success Response (200 OK):**
    ```json
    {
      "job_id": "unique-job-identifier-string",
      "status": "queued | analyzing | processing | complete | error",
      "progress": 0-100, // Percentage
      "estimated_time": 15, // Estimated seconds remaining (or null)
      "message": "Optional status message..." // Provided for some statuses
    }
    ```
* **Error Response:** `404 Not Found` (Invalid job ID).

---

### 3. Get Results

* **Endpoint:** `/api/results/<job_id>`
* **Method:** `GET`
* **Description:** Fetches the results (download link, stats) once a job is complete.
* **URL Parameters:**
    * `job_id`: The unique identifier.
* **Success Response (200 OK - Job Complete):**
    ```json
    {
      "job_id": "unique-job-identifier-string",
      "status": "complete",
      "download_urls": {
        "docx": "/api/download/job-id_original-filename.docx"
        // "doc": "/api/download/job-id_original-filename.doc" // If implemented
      },
      "statistics": {
        "page_count": 10,
        "processing_time_seconds": 5.2
      },
      "health_report": {
        "encryption": "None",
        "scanned_pages": "Detection not run (OCR skipped)",
        "font_issues": "Basic check passed",
        "warnings": ["None"] // or ["Warning message 1", ...]
      }
    }
    ```
* **Response (202 Accepted - Job Not Complete):** Message indicating current status.
* **Response (400 Bad Request - Job Errored):** Includes status 'error' and an error message.
* **Error Response:** `404 Not Found` (Invalid job ID).

---

### 4. Download File

* **Endpoint:** `/api/download/<filename>`
* **Method:** `GET`
* **Description:** Downloads the converted file. The `<filename>` comes from the `download_urls` in the `/api/results` response.
* **URL Parameters:**
    * `filename`: The name of the file to download (e.g., `job-id_original-filename.docx`).
* **Success Response (200 OK):** The DOCX file content with appropriate headers for download.
* **Error Response:** `404 Not Found` (File does not exist).

## Notes

* **Job Storage:** This implementation uses in-memory storage for job tracking. This means job status will be lost if the server restarts. For production, consider using a database or Redis.
* **CORS:** Cross-Origin Resource Sharing (CORS) is enabled for all routes, allowing requests from different domains (like your WordPress site). Adjust origins in `app.py` for tighter security in production if needed.
* **Password Protection:** Password-protected PDFs are currently **not supported**, although some related code exists.
* **Optimization Settings:** The `optimizer_setting` parameter is currently ignored by the backend.