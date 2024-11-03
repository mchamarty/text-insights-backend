import traceback
import logging
from flask import Flask, request, send_file, jsonify, make_response
from flask_cors import CORS  # Add this import
import pypff
import csv
from io import StringIO
import os
import datetime
import json

# Enhanced logging
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

app = Flask(__name__)

# Simple CORS configuration that allows all origins
CORS(app, supports_credentials=True)

#adding a comment just to trigger commit

# Set maximum content length and file size limits
MAX_FILE_SIZE = 100 * 1024 * 1024  # 100MB
app.config['MAX_CONTENT_LENGTH'] = 105 * 1024 * 1024  # 105MB to allow for form data overhead

UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.before_request
def before_request():
    logging.debug(f"Incoming request: {request.method} {request.path}")
    logging.debug(f"Headers: {dict(request.headers)}")

@app.after_request
def add_cors_headers(response):
    logging.debug(f"Response headers: {dict(response.headers)}")
    return response

@app.errorhandler(413)
def request_entity_too_large(error):
    response = jsonify({
        'error': f'File too large. Maximum allowed size is {MAX_FILE_SIZE/1024/1024:.1f}MB'
    })
    return response, 413

def decode_if_bytes(value):
    """Helper function to decode bytes to string."""
    if isinstance(value, bytes):
        try:
            return value.decode('utf-8')
        except UnicodeDecodeError:
            try:
                return value.decode('latin-1')
            except:
                return str(value)
    return str(value) if value is not None else ''

def process_pst_folder(folder, emails, csv_writer):
    """Recursively process a PST folder and its subfolders."""
    try:
        message_count = folder.get_number_of_sub_messages()
        logger.debug(f"Processing folder with {message_count} messages")
        
        for i in range(message_count):
            try:
                message = folder.get_sub_message(i)
                
                sender = decode_if_bytes(message.get_sender_name())
                subject = decode_if_bytes(message.get_subject())
                body = decode_if_bytes(message.get_plain_text_body())
                delivery_time = message.get_delivery_time()
                
                headers = decode_if_bytes(message.get_transport_headers())
                to = ''
                if headers:
                    for line in headers.split('\n'):
                        if line.lower().startswith('to:'):
                            to = line[3:].strip()
                            break
                
                attachments = []
                try:
                    attachment_count = message.get_number_of_attachments()
                    for j in range(attachment_count):
                        try:
                            attachment = message.get_attachment(j)
                            name = attachment.get_long_filename()
                            if not name:
                                name = attachment.get_short_filename()
                            if name:
                                attachments.append(decode_if_bytes(name))
                        except Exception as e:
                            logger.error(f"Error processing attachment {j}: {str(e)}")
                except Exception as e:
                    logger.error(f"Error getting attachments: {str(e)}")
                    attachment_count = 0
                
                email_data = {
                    'from': sender,
                    'to': to,
                    'subject': subject,
                    'body': body,
                    'date': str(delivery_time),
                    'attachments': attachments
                }
                
                emails.append(email_data)
                csv_writer.writerow([
                    sender,
                    to,
                    subject,
                    body,
                    delivery_time,
                    ', '.join(attachments)
                ])
                
            except Exception as e:
                logger.error(f"Error processing message: {str(e)}")
                continue
        
        for i in range(folder.get_number_of_sub_folders()):
            sub_folder = folder.get_sub_folder(i)
            process_pst_folder(sub_folder, emails, csv_writer)
            
    except Exception as e:
        logger.error(f"Error processing folder: {str(e)}")

@app.route('/')
def home():
    logger.debug("Home route accessed")
    return "Flask server is running!"

@app.route('/analyze-pst', methods=['POST'])
def analyze_pst():
    logger.debug(f"analyze-pst route accessed")
    logger.debug(f"Request method: {request.method}")
    logger.debug(f"Request headers: {request.headers}")
    
    if 'file' not in request.files:
        logger.error("No file in request")
        return jsonify({'error': 'No file part'}), 400
    
    file = request.files['file']
    if not file.filename:
        logger.error("No selected file")
        return jsonify({'error': 'No selected file'}), 400

    if not file.filename.endswith('.pst'):
        logger.error("Invalid file type")
        return jsonify({'error': 'Invalid file type'}), 400

    # Check file size
    file.seek(0, os.SEEK_END)
    size = file.tell()
    file.seek(0)
    
    if size > MAX_FILE_SIZE:
        logger.error(f"File too large: {size} bytes (max {MAX_FILE_SIZE} bytes)")
        return jsonify({'error': f'File size exceeds maximum allowed size of {MAX_FILE_SIZE/1024/1024:.1f}MB'}), 413

    temp_path = None
    try:
        logger.debug(f"Processing file: {file.filename}")
        
        temp_path = os.path.join(UPLOAD_FOLDER, f"temp_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.pst")
        file.save(temp_path)
        logger.debug(f"File saved to {temp_path}")
        
        pst = pypff.file()
        pst.open(temp_path)
        root = pst.get_root_folder()

        emails = []
        csv_data = StringIO()
        csv_writer = csv.writer(csv_data)
        csv_writer.writerow(['From', 'To', 'Subject', 'Body', 'Date', 'Attachments'])

        process_pst_folder(root, emails, csv_writer)
        pst.close()

        logger.debug(f"Found {len(emails)} emails")
        if len(emails) == 0:
            return jsonify({'error': 'No emails found in PST file'}), 400

        csv_filename = f"extracted_emails_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        csv_path = os.path.join(UPLOAD_FOLDER, csv_filename)
        with open(csv_path, 'w', newline='', encoding='utf-8') as csvfile:
            csvfile.write(csv_data.getvalue())
        
        logger.debug(f"Processing complete. CSV saved as {csv_filename}")
        
        try:
            json_data = {
                'emails': emails,
                'csv_url': f"/download/{csv_filename}"
            }
            json.dumps(json_data)  # Test serialization
            return jsonify(json_data)
        except TypeError as e:
            logger.error(f"JSON serialization error: {str(e)}")
            return jsonify({'error': 'Data contains non-serializable values'}), 500

    except Exception as e:
        logger.error(f"Error processing PST file: {str(e)}")
        logger.error(traceback.format_exc())
        return jsonify({'error': str(e)}), 500

    finally:
        if temp_path and os.path.exists(temp_path):
            os.remove(temp_path)
            logger.debug(f"Cleaned up temporary file {temp_path}")

@app.route('/download/<filename>')
def download_file(filename):
    logger.debug(f"Download requested for {filename}")
    try:
        file_path = os.path.join(UPLOAD_FOLDER, filename)
        if not os.path.exists(file_path):
            logger.error(f"File not found: {file_path}")
            return jsonify({'error': 'File not found'}), 404
        return send_file(file_path, as_attachment=True)
    except Exception as e:
        logger.error(f"Error downloading file: {str(e)}")
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    print("Starting Flask server on port 5000...")
    app.run(host='0.0.0.0', port=5000, debug=True)