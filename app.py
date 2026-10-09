import os
import io
import tempfile
import logging
from flask import Flask, render_template, request, send_file, jsonify
from flask_cors import CORS
from doc_generator import generate_docx, generate_pdf

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__, static_folder='static', template_folder='templates')
CORS(app)

# Required form fields
REQUIRED_FIELDS = [
    ('title', 'Assignment Title'),
    ('course', 'Course Code'),
    ('sname', 'Student Name'),
    ('sid', 'Student ID'),
    ('semester', 'Semester'),
    ('tname', 'Teacher Name'),
    ('tdesignation', 'Teacher Designation')
]

def sanitize_filename_part(text, fallback='assignment'):
    """Generate a clean, safe ASCII filename segment."""
    clean = "".join(c for c in text if c.isalnum() or c in ('-', '_')).strip()
    return clean or fallback

def extract_form_data(req):
    """Safely parse JSON or form data, returning cleaned strings and field errors."""
    if req.is_json:
        data = req.get_json(silent=True) or {}
    else:
        data = req.form.to_dict()

    cleaned = {}
    missing = []
    for field_key, field_label in REQUIRED_FIELDS:
        val = str(data.get(field_key, '')).strip()
        cleaned[field_key] = val
        if not val:
            missing.append(field_label)

    return cleaned, missing

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/health')
def health_check():
    return jsonify({'status': 'ok', 'app': 'Assignment Cover Page Generator'}), 200

@app.route('/api/generate-docx', methods=['POST'])
def api_generate_docx():
    cleaned, missing = extract_form_data(request)
    if missing:
        msg = f"Please provide: {', '.join(missing)}."
        return jsonify({'error': msg, 'missing_fields': missing}), 400

    tmp_dir = tempfile.mkdtemp(prefix='sec_docx_')
    tmp_path = os.path.join(tmp_dir, 'generated.docx')

    try:
        generate_docx(
            title=cleaned['title'],
            course=cleaned['course'],
            sname=cleaned['sname'],
            sid=cleaned['sid'],
            semester=cleaned['semester'],
            tname=cleaned['tname'],
            tdesignation=cleaned['tdesignation'],
            output_path=tmp_path
        )

        with open(tmp_path, 'rb') as f:
            file_bytes = io.BytesIO(f.read())

        safe_course = sanitize_filename_part(cleaned['course'], fallback='CSE')
        filename = f"{safe_course}_Assignment_Cover.docx"

        response = send_file(
            file_bytes,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
        response.headers['Access-Control-Expose-Headers'] = 'Content-Disposition'
        return response

    except Exception as e:
        logger.error("DOCX generation error: %s", str(e), exc_info=True)
        return jsonify({'error': 'Failed to generate Word (.docx) file. Please check your inputs and try again.'}), 500

    finally:
        if os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except OSError:
                pass
        if os.path.exists(tmp_dir):
            try:
                os.rmdir(tmp_dir)
            except OSError:
                pass

@app.route('/api/generate-pdf', methods=['POST'])
def api_generate_pdf():
    cleaned, missing = extract_form_data(request)
    if missing:
        msg = f"Please provide: {', '.join(missing)}."
        return jsonify({'error': msg, 'missing_fields': missing}), 400

    tmp_dir = tempfile.mkdtemp(prefix='sec_pdf_')
    tmp_path = os.path.join(tmp_dir, 'generated.pdf')

    try:
        generate_pdf(
            title=cleaned['title'],
            course=cleaned['course'],
            sname=cleaned['sname'],
            sid=cleaned['sid'],
            semester=cleaned['semester'],
            tname=cleaned['tname'],
            tdesignation=cleaned['tdesignation'],
            output_path=tmp_path
        )

        with open(tmp_path, 'rb') as f:
            file_bytes = io.BytesIO(f.read())

        safe_course = sanitize_filename_part(cleaned['course'], fallback='CSE')
        filename = f"{safe_course}_Assignment_Cover.pdf"

        # Check query parameter for inline view vs attachment download
        disposition = request.args.get('disposition', 'attachment')
        as_attachment = (disposition != 'inline')

        response = send_file(
            file_bytes,
            as_attachment=as_attachment,
            download_name=filename,
            mimetype='application/pdf'
        )
        response.headers['Access-Control-Expose-Headers'] = 'Content-Disposition'
        return response

    except Exception as e:
        logger.error("PDF generation error: %s", str(e), exc_info=True)
        return jsonify({'error': 'Failed to generate PDF file. Please check your inputs and try again.'}), 500

    finally:
        if os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except OSError:
                pass
        if os.path.exists(tmp_dir):
            try:
                os.rmdir(tmp_dir)
            except OSError:
                pass

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
