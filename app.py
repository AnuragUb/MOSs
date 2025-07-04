from flask import Flask, render_template, request, jsonify, send_file, session
from flask_cors import CORS
import json
import os
from datetime import datetime, timedelta
import requests
import tempfile
import subprocess
from werkzeug.utils import secure_filename
import pandas as pd
import ffmpeg
import logging
import io
import vlc
import sys
import platform
import openpyxl
from openpyxl.styles import PatternFill
import csv
from video_utils import convert_wmv_to_mp4, upload_to_gcs
from google.cloud import storage
from google.cloud import firestore
from dotenv import load_dotenv
import uuid
from google.cloud import secretmanager
from google.oauth2 import service_account
import shutil
from google.auth import compute_engine
import urllib.parse
import re
import mimetypes

# Load environment variables
load_dotenv()

app = Flask(__name__)
CORS(app)

# --- FINGERPRINT ---
# Add a unique log message to verify deployment
current_version = "v3.0.0 - " + datetime.now().strftime("%Y-%m-%d %H:%M:%S")
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)
logger.info(f"--- APPLICATION STARTING - VERSION {current_version} ---")
# --- END FINGERPRINT ---

app.config['MAX_CONTENT_LENGTH'] = 500 * 1024 * 1024  # 500MB max file size
app.config['UPLOAD_FOLDER'] = 'uploads'
app.secret_key = os.environ.get('FLASK_SECRET_KEY', 'your-secret-key')  # Set securely in production!

# Configure maximum file size (1000MB)
app.config['MAX_CONTENT_LENGTH'] = 1000 * 1024 * 1024  # 1000MB in bytes

# Initialize VLC instance with proper error handling
try:
    # Create a basic vlc instance
    vlc_instance = vlc.Instance()
    if not vlc_instance:
        raise Exception("Failed to create VLC instance")
        
    # Create an empty vlc media player
    player = vlc_instance.media_player_new()
    if not player:
        raise Exception("Failed to create media player")
        
    logger.info("VLC initialized successfully")
except Exception as e:
    logger.warning(f"VLC initialization failed (this is normal in Cloud Run): {str(e)}")
    vlc_instance = None
    player = None

# Store usage statistics
USAGE_STATS = {
    'BI': 0,
    'BV': 0,
    'VI': 0,
    'VV': 0,
    'SRC': 0
}

# Ensure the uploads directory exists
UPLOADS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads')
if not os.path.exists(UPLOADS_DIR):
    os.makedirs(UPLOADS_DIR)

AUDD_API_TOKEN = os.environ.get('AUDD_API_TOKEN')

# Initialize Firestore client with proper error handling
try:
    firestore_client = firestore.Client()
    MUSIC_CO_COLLECTION = 'music_companies'
    # Optionally test the connection (can be commented out if not needed)
    # test_doc = firestore_client.collection(MUSIC_CO_COLLECTION).limit(1).get()
    logger.info("Firestore initialized successfully")
except Exception as e:
    logger.error(f"Error initializing Firestore: {str(e)}")
    firestore_client = None

# --- Global Variables ---
storage_client = None

def initialize_gcs_client():
    """Initializes the GCS client"""
    global storage_client
    try:
        storage_client = storage.Client()
        logger.info(f"Successfully initialized GCS client .")
    except Exception as e:
        logger.error(f"FATAL: Failed to initialize GCS client : {e}")
        storage_client = None

# Initialize GCS client on application startup
initialize_gcs_client()

# --- SRT Profile Management Endpoints ---
PROFILE_COLLECTION = 'SRT_validator_profiles'
PROFILE_DOC = 'global'
PROFILE_SUBCOL = 'profiles'
PROFILE_ICON_BUCKET = os.getenv('GCS_PROFILE_ICON_BUCKET', os.getenv('GCS_BUCKET_NAME', 'mos-aat'))
PROFILE_ICON_FOLDER = 'profile_icons/'
PROFILE_ICON_MAX_SIZE = 100 * 1024  # 100KB
PROFILE_ICON_ALLOWED = {'image/png', 'image/jpeg', 'image/jpg', 'image/svg+xml'}

@app.route('/api/srt-profiles', methods=['GET'])
def list_srt_profiles():
    try:
        profiles_ref = firestore_client.collection(PROFILE_COLLECTION).document(PROFILE_DOC).collection(PROFILE_SUBCOL)
        docs = profiles_ref.stream()
        profiles = []
        for doc in docs:
            data = doc.to_dict()
            data['id'] = doc.id
            profiles.append(data)
        return jsonify({'profiles': profiles})
    except Exception as e:
        logger.error(f"Error listing SRT profiles: {str(e)}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/srt-profiles', methods=['POST'])
def create_srt_profile():
    try:
        data = request.json
        name = data.get('name', '').strip()
        icon = data.get('icon', '')  # URL or empty
        parameters = data.get('parameters', {})
        now = datetime.utcnow()
        if not name:
            return jsonify({'error': 'Profile name is required'}), 400
        profiles_ref = firestore_client.collection(PROFILE_COLLECTION).document(PROFILE_DOC).collection(PROFILE_SUBCOL)
        # Optionally, check for duplicate name
        existing = list(profiles_ref.where('name', '==', name).stream())
        if existing:
            return jsonify({'error': 'Profile with this name already exists'}), 409
        doc_ref = profiles_ref.document()
        doc_ref.set({
            'name': name,
            'icon': icon,
            'parameters': parameters,
            'created_at': now,
            'updated_at': now
        })
        return jsonify({'id': doc_ref.id, 'name': name, 'icon': icon, 'parameters': parameters})
    except Exception as e:
        logger.error(f"Error creating SRT profile: {str(e)}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/srt-profiles/<profile_id>', methods=['PUT'])
def update_srt_profile(profile_id):
    try:
        data = request.json
        name = data.get('name', '').strip()
        icon = data.get('icon', '')
        parameters = data.get('parameters', {})
        now = datetime.utcnow()
        profiles_ref = firestore_client.collection(PROFILE_COLLECTION).document(PROFILE_DOC).collection(PROFILE_SUBCOL)
        doc_ref = profiles_ref.document(profile_id)
        if not doc_ref.get().exists:
            return jsonify({'error': 'Profile not found'}), 404
        update_data = {'updated_at': now}
        if name:
            update_data['name'] = name
        if icon is not None:
            update_data['icon'] = icon
        if parameters:
            update_data['parameters'] = parameters
        doc_ref.update(update_data)
        return jsonify({'id': profile_id, 'name': name, 'icon': icon, 'parameters': parameters})
    except Exception as e:
        logger.error(f"Error updating SRT profile: {str(e)}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/srt-profiles/<profile_id>', methods=['DELETE'])
def delete_srt_profile(profile_id):
    try:
        profiles_ref = firestore_client.collection(PROFILE_COLLECTION).document(PROFILE_DOC).collection(PROFILE_SUBCOL)
        doc_ref = profiles_ref.document(profile_id)
        if not doc_ref.get().exists:
            return jsonify({'error': 'Profile not found'}), 404
        doc_ref.delete()
        return jsonify({'status': 'success', 'message': 'Profile deleted'})
    except Exception as e:
        logger.error(f"Error deleting SRT profile: {str(e)}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/srt-profile-icon-upload', methods=['POST'])
def upload_srt_profile_icon():
    try:
        if 'icon' not in request.files:
            return jsonify({'error': 'No icon file uploaded'}), 400
        file = request.files['icon']
        if file.filename == '':
            return jsonify({'error': 'No selected file'}), 400
        # Check file size
        file.seek(0, 2)
        size = file.tell()
        file.seek(0)
        if size > PROFILE_ICON_MAX_SIZE:
            return jsonify({'error': f'Icon file too large (max {PROFILE_ICON_MAX_SIZE//1024}KB)'}), 400
        # Check file type
        mime = mimetypes.guess_type(file.filename)[0]
        if mime not in PROFILE_ICON_ALLOWED:
            return jsonify({'error': f'Invalid file type: {mime}'}), 400
        # Save to GCS
        ext = os.path.splitext(file.filename)[1]
        unique_name = f"{uuid.uuid4()}{ext}"
        gcs_path = f"{PROFILE_ICON_FOLDER}{unique_name}"
        bucket = storage_client.bucket(PROFILE_ICON_BUCKET)
        blob = bucket.blob(gcs_path)
        blob.upload_from_file(file, content_type=mime)
        blob.make_public()
        url = blob.public_url
        return jsonify({'url': url})
    except Exception as e:
        logger.error(f"Error uploading profile icon: {str(e)}")
        return jsonify({'error': str(e)}), 500

@app.route('/')
def index():
    return render_template('splash.html')

@app.route('/health')
def health_check():
    """Health check endpoint for Cloud Run startup probe"""
    try:
        # Basic health checks - only check essential services
        if storage_client is None:
            return jsonify({'status': 'unhealthy', 'error': 'GCS client not initialized'}), 503
        
        # Firestore is optional for basic functionality
        # VLC is optional and not required for startup
        
        return jsonify({
            'status': 'healthy', 
            'timestamp': datetime.now().isoformat(),
            'gcs_ready': storage_client is not None,
            'firestore_ready': firestore_client is not None
        }), 200
    except Exception as e:
        logger.error(f"Health check failed: {str(e)}")
        return jsonify({'status': 'unhealthy', 'error': str(e)}), 503

@app.route('/main')
def main():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    """DEPRECATED: This route is for basic local uploads, GCS is preferred."""
    try:
        if 'video' not in request.files:
            return jsonify({'error': 'No video file provided'}), 400
        
        file = request.files['video']
        if file.filename == '':
            return jsonify({'error': 'No selected file'}), 400
        
        # This is a fallback and does not handle large files or GCS correctly.
        # It's kept for potential local testing but not for production use.
        os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)

        return jsonify({'filename': filename})
                
    except Exception as e:
        logger.error(f"Error in basic upload: {str(e)}")
        return jsonify({'error': str(e)}), 500

@app.route('/uploads/<filename>')
def serve_video(filename):
    return send_file(os.path.join(UPLOADS_DIR, filename))

@app.route('/api/markers', methods=['GET', 'POST'])
def handle_markers():
    if request.method == 'POST':
        markers = request.json
        # Save markers to a file
        with open('markers.json', 'w') as f:
            json.dump(markers, f)
        return jsonify({'status': 'success'})
    else:
        # Load markers from file
        try:
            with open('markers.json', 'r') as f:
                markers = json.load(f)
            return jsonify(markers)
        except FileNotFoundError:
            return jsonify([])

@app.route('/api/export/<format>', methods=['POST'])
def export_markers(format):
    try:
        payload = request.json
        header_rows = payload.get('headerRows', [])
        markers = payload.get('markers', [])
        blank_lines = payload.get('blankLines', 0)
        fields_to_export = payload.get('fieldsToExport')
        field_labels = payload.get('fieldLabels', {})
        # Correctly get time_format, default to 'timecode' (HH:MM:SS)
        time_format = payload.get('timeFormat', 'timecode')

        def convert_time_fields(row):
            new_row = row.copy()
            def is_timecode_string(val):
                import re
                return isinstance(val, str) and re.match(r"^\d{2}:\d{2}:\d{2}(:\d{2})?$", val)

            def timecode_to_seconds(tc_str):
                parts = list(map(int, tc_str.split(':')))
                seconds = parts[0] * 3600 + parts[1] * 60 + parts[2]
                if len(parts) == 4:
                    seconds += parts[3] / 25.0 # Assuming 25 fps
                return seconds

            def seconds_to_timecode(s, with_frames=False):
                h = int(s // 3600)
                m = int((s % 3600) // 60)
                sec = int(s % 60)
                if with_frames:
                    f = int(round((s % 1) * 25))
                    return f"{h:02d}:{m:02d}:{sec:02d}:{f:02d}"
                return f"{h:02d}:{m:02d}:{sec:02d}"

            # The fields to convert
            for field in ['tcrIn', 'tcrOut', 'duration']:
                if field in new_row and new_row[field] and is_timecode_string(new_row[field]):
                    seconds = timecode_to_seconds(new_row[field])
                    new_row[field] = seconds_to_timecode(seconds, with_frames=(time_format == 'timecode_frames'))

            return new_row

        if format == 'excel':
            temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
            temp_path = temp_file.name
            temp_file.close()

            wb = openpyxl.Workbook()
            ws = wb.active

            # --- BEGIN: Write metadata rows ---
            for row in header_rows:
                ws.append(row)
            # --- END: Write metadata rows ---

            # Add blank lines if specified
            for _ in range(blank_lines):
                ws.append([])

            # --- BEGIN: Write marker table ---
            if markers:
                if fields_to_export:
                    # Add SEQ# to the beginning of fields_to_export if not already present
                    if 'seq' not in fields_to_export:
                        fields_to_export.insert(0, 'seq')
                    
                    # Add SEQ# to field labels if not present
                    if 'seq' not in field_labels:
                        field_labels['seq'] = 'SEQ#'
                    
                    ws.append([field_labels.get(field, field) for field in fields_to_export])
                    for i, row in enumerate(markers):
                        # Convert time fields to the selected format
                        row = convert_time_fields(row)
                        # Add SEQ# to the beginning of the row data
                        row_data = [i + 1]  # SEQ# is index + 1
                        # Convert usage array or stringified array to comma-separated string if it exists
                        if 'usage' in row:
                            if isinstance(row['usage'], list):
                                row['usage'] = ','.join(row['usage'])
                            elif isinstance(row['usage'], str) and row['usage'].startswith('[') and row['usage'].endswith(']'):
                                try:
                                    arr = json.loads(row['usage'])
                                    if isinstance(arr, list):
                                        row['usage'] = ','.join(arr)
                                except Exception:
                                    pass
                        # Add the rest of the fields
                        row_data.extend([row.get(field, '') for field in fields_to_export[1:]])
                        ws.append(row_data)
                else:
                    # Add SEQ# to the beginning of the keys
                    keys = ['seq'] + list(markers[0].keys())
                    ws.append(keys)
                    for i, row in enumerate(markers):
                        # Convert time fields to the selected format
                        row = convert_time_fields(row)
                        # Add SEQ# to the beginning of the row data
                        row_data = [i + 1]  # SEQ# is index + 1
                        # Convert usage array or stringified array to comma-separated string if it exists
                        if 'usage' in row:
                            if isinstance(row['usage'], list):
                                row['usage'] = ','.join(row['usage'])
                            elif isinstance(row['usage'], str) and row['usage'].startswith('[') and row['usage'].endswith(']'):
                                try:
                                    arr = json.loads(row['usage'])
                                    if isinstance(arr, list):
                                        row['usage'] = ','.join(arr)
                                except Exception:
                                    pass
                        # Add the rest of the values
                        row_data.extend(list(row.values()))
                        ws.append(row_data)
            # --- END: Write marker table ---

            # --- BEGIN: Color marked rows ---
            yellow_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
            red_fill = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')
            # Metadata rows + header row + blank lines
            start_row = len(header_rows) + blank_lines + 2 if markers else len(header_rows) + blank_lines + 1
            for i, marker in enumerate(markers):
                color = marker.get('markColor', '')
                # Handle combined markColor values (e.g., 'yellow,exception', 'red,exception')
                if 'yellow' in color:
                    for cell in ws[start_row + i]:
                        cell.fill = yellow_fill
                elif 'red' in color:
                    for cell in ws[start_row + i]:
                        cell.fill = red_fill
            # --- END: Color marked rows ---

            wb.save(temp_path)
            with open(temp_path, 'rb') as f:
                file_data = f.read()

            # Use original file name for export if provided
            export_file_name = payload.get('originalFileName', 'exported_file')
            return send_file(
                io.BytesIO(file_data),
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                as_attachment=True,
                download_name=f'{export_file_name}.xlsx'
            )

        elif format == 'csv':
            temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.csv')
            temp_path = temp_file.name
            temp_file.close()

            with open(temp_path, 'w', newline='', encoding='utf-8') as csvfile:
                writer = csv.writer(csvfile)
                # --- BEGIN: Write metadata rows ---
                for row in header_rows:
                    writer.writerow(row)
                # --- END: Write metadata rows ---

                # Add blank lines if specified
                for _ in range(blank_lines):
                    writer.writerow([])

                # --- BEGIN: Write marker table ---
                if markers:
                    if fields_to_export:
                        # Add SEQ# to the beginning of fields_to_export if not already present
                        if 'seq' not in fields_to_export:
                            fields_to_export.insert(0, 'seq')
                        
                        # Add SEQ# to field labels if not present
                        if 'seq' not in field_labels:
                            field_labels['seq'] = 'SEQ#'
                        
                        writer.writerow([field_labels.get(field, field) for field in fields_to_export])
                        for i, row in enumerate(markers):
                            # Convert time fields to the selected format
                            row = convert_time_fields(row)
                            # Add SEQ# to the beginning of the row data
                            row_data = [i + 1]  # SEQ# is index + 1
                            # Convert usage array or stringified array to comma-separated string if it exists
                            if 'usage' in row:
                                if isinstance(row['usage'], list):
                                    row['usage'] = ','.join(row['usage'])
                                elif isinstance(row['usage'], str) and row['usage'].startswith('[') and row['usage'].endswith(']'):
                                    try:
                                        arr = json.loads(row['usage'])
                                        if isinstance(arr, list):
                                            row['usage'] = ','.join(arr)
                                    except Exception:
                                        pass
                            # Add the rest of the fields
                            row_data.extend([row.get(field, '') for field in fields_to_export[1:]])
                            writer.writerow(row_data)
                    else:
                        # Add SEQ# to the beginning of the keys
                        keys = ['seq'] + list(markers[0].keys())
                        writer.writerow(keys)
                        for i, row in enumerate(markers):
                            # Convert time fields to the selected format
                            row = convert_time_fields(row)
                            # Add SEQ# to the beginning of the row data
                            row_data = [i + 1]  # SEQ# is index + 1
                            # Convert usage array or stringified array to comma-separated string if it exists
                            if 'usage' in row:
                                if isinstance(row['usage'], list):
                                    row['usage'] = ','.join(row['usage'])
                                elif isinstance(row['usage'], str) and row['usage'].startswith('[') and row['usage'].endswith(']'):
                                    try:
                                        arr = json.loads(row['usage'])
                                        if isinstance(arr, list):
                                            row['usage'] = ','.join(arr)
                                    except Exception:
                                        pass
                            # Add the rest of the values
                            row_data.extend(list(row.values()))
                            writer.writerow(row_data)
                # --- END: Write marker table ---

            # Use original file name for export if provided
            export_file_name = payload.get('originalFileName', 'exported_file')
            return send_file(
                io.BytesIO(open(temp_path, 'rb').read()),
                mimetype='text/csv',
                as_attachment=True,
                download_name=f'{export_file_name}.csv'
            )

        else:
            return jsonify({'error': 'Invalid export format'}), 400

    except Exception as e:
        logger.error(f"Export failed: {e}")
        return jsonify({'error': f'Export failed: {str(e)}'}), 500
    finally:
        if 'temp_path' in locals() and os.path.exists(temp_path):
            try:
                os.unlink(temp_path)
            except Exception as cleanup_error:
                logger.error(f"Error cleaning up temporary file: {str(cleanup_error)}")

@app.route('/api/usage-stats', methods=['GET', 'POST'])
def handle_usage_stats():
    global USAGE_STATS
    if request.method == 'POST':
        usage = request.json.get('usage')
        if usage in USAGE_STATS:
            USAGE_STATS[usage] += 1
    return jsonify(USAGE_STATS)

@app.route('/api/recognize-audio', methods=['POST'])
def recognize_audio():
    try:
        tcr_in = request.form.get('tcrIn')
        tcr_out = request.form.get('tcrOut')
        video_src = request.form.get('videoSrc')
        file = request.files.get('file')

        logger.info(f"Received recognize request - TCR In: {tcr_in}, TCR Out: {tcr_out}")
        logger.info(f"Video source: {video_src}, File: {file.filename if file else 'None'}")

        if not tcr_in or not tcr_out:
            return jsonify({'status': 'error', 'message': 'Missing TCR In or Out'}), 400

        # Calculate start and duration in seconds
        def time_to_seconds(t):
            h, m, s = map(float, t.split(':'))
            return int(h) * 3600 + int(m) * 60 + s
        
        start = time_to_seconds(tcr_in)
        end = time_to_seconds(tcr_out)
        duration = end - start
        
        if duration <= 0:
            return jsonify({'status': 'error', 'message': 'Invalid time range'}), 400

        # Prepare input file path
        if file:
            # Save uploaded file to temp
            try:
                # First, read the file into memory to check its size
                file_content = file.read()
                file_size = len(file_content)
                logger.info(f"Uploaded file size: {file_size} bytes")
                
                if file_size == 0:
                    return jsonify({
                        'status': 'error',
                        'message': 'The uploaded file is empty'
                    }), 400
                
                # Reset file pointer for saving
                file.seek(0)
                
                # Create a unique temporary directory for this upload
                upload_dir = tempfile.mkdtemp()
                input_path = os.path.join(upload_dir, 'input.mp4')
                logger.info(f"Created temporary directory: {upload_dir}")
                
                # Write the file in chunks to handle large files
                chunk_size = 8192
                bytes_written = 0
                with open(input_path, 'wb') as temp_in:
                    while True:
                        chunk = file.read(chunk_size)
                        if not chunk:
                            break
                        temp_in.write(chunk)
                        bytes_written += len(chunk)
                        logger.info(f"Upload progress: {bytes_written}/{file_size} bytes ({(bytes_written/file_size)*100:.2f}%)")
                
                logger.info(f"File upload completed. Total bytes written: {bytes_written}")
                
                if bytes_written != file_size:
                    logger.error(f"File size mismatch. Expected {file_size} bytes, got {bytes_written} bytes")
                    return jsonify({
                        'status': 'error',
                        'message': 'File upload was incomplete. Please try again.'
                    }), 400
                
                # Validate the video file
                try:
                    # First check if file exists and has content
                    if not os.path.exists(input_path) or os.path.getsize(input_path) == 0:
                        return jsonify({
                            'status': 'error',
                            'message': 'The video file is empty or incomplete. Please try uploading again.'
                        }), 400
                    
                    # Use ffprobe to get detailed video information
                    probe_cmd = [
                        'ffprobe',
                        '-v', 'error',
                        '-print_format', 'json',
                        '-show_format',
                        '-show_streams',
                        input_path
                    ]
                    
                    logger.info(f"Running ffprobe validation: {' '.join(probe_cmd)}")
                    probe_result = subprocess.run(probe_cmd, capture_output=True, text=True)
                    
                    if probe_result.returncode != 0:
                        logger.error(f"Invalid video file: {probe_result.stderr}")
                        return jsonify({
                            'status': 'error',
                            'message': 'The video file appears to be invalid or corrupted. Please try uploading again.'
                        }), 400
                    
                    # Log the probe result for debugging
                    logger.info(f"FFprobe result: {probe_result.stdout}")
                    
                    # Try to fix the video file using a different approach
                    try:
                        logger.info("Attempting to fix video file format")
                        fixed_path = os.path.join(upload_dir, 'fixed.mp4')
                        
                        # First try: Use -movflags +faststart
                        fix_cmd1 = [
                            'ffmpeg',
                            '-y',
                            '-v', 'error',
                            '-i', input_path,
                            '-c', 'copy',
                            '-movflags', '+faststart',
                            fixed_path
                        ]
                        
                        logger.info(f"Running first fix attempt: {' '.join(fix_cmd1)}")
                        fix_result1 = subprocess.run(fix_cmd1, capture_output=True, text=True)
                        
                        if fix_result1.returncode != 0:
                            logger.warning(f"First fix attempt failed: {fix_result1.stderr}")
                            
                            # Second try: Re-encode the video
                            logger.info("Attempting re-encoding fix")
                            fix_cmd2 = [
                                'ffmpeg',
                                '-y',
                                '-v', 'error',
                                '-i', input_path,
                                '-c:v', 'libx264',
                                '-c:a', 'aac',
                                '-movflags', '+faststart',
                                fixed_path
                            ]
                            
                            logger.info(f"Running second fix attempt: {' '.join(fix_cmd2)}")
                            fix_result2 = subprocess.run(fix_cmd2, capture_output=True, text=True)
                            
                            if fix_result2.returncode != 0:
                                logger.error(f"Second fix attempt failed: {fix_result2.stderr}")
                                return jsonify({
                                    'status': 'error',
                                    'message': 'Could not process the video file. Please try uploading a different file.'
                                }), 400
                        
                        # Replace original input path with fixed file
                        os.remove(input_path)
                        input_path = fixed_path
                        logger.info(f"Successfully fixed video file: {input_path}")
                        
                    except Exception as e:
                        logger.error(f"Error fixing video file: {str(e)}")
                        return jsonify({
                            'status': 'error',
                            'message': 'Error processing video file. Please try uploading again.'
                        }), 500
                    
                except Exception as e:
                    logger.error(f"Error validating video file: {str(e)}")
                    return jsonify({
                        'status': 'error',
                        'message': 'Error validating video file. Please try uploading again.'
                    }), 400
                    
            except Exception as e:
                logger.error(f"Error handling file upload: {str(e)}")
                return jsonify({
                    'status': 'error',
                    'message': f'Error handling file upload: {str(e)}'
                }), 500
        elif video_src:
            # Download the video file
            try:
                logger.info(f"Starting video download from URL: {video_src}")
                
                # Create a unique temporary directory for this download
                upload_dir = tempfile.mkdtemp()
                input_path = os.path.join(upload_dir, 'input.mp4')
                logger.info(f"Created temporary directory: {upload_dir}")
                
                # Get file size from headers if available
                head_response = requests.head(video_src, allow_redirects=True)
                content_length = head_response.headers.get('content-length')
                if content_length:
                    logger.info(f"Expected file size: {content_length} bytes")
                
                # Download the file with progress tracking
                response = requests.get(video_src, stream=True)
                response.raise_for_status()
                
                total_size = int(response.headers.get('content-length', 0))
                block_size = 8192
                downloaded_size = 0
                
                with open(input_path, 'wb') as temp_in:
                    for chunk in response.iter_content(chunk_size=block_size):
                        if chunk:
                            temp_in.write(chunk)
                            downloaded_size += len(chunk)
                            if total_size > 0:
                                progress = (downloaded_size / total_size) * 100
                                logger.info(f"Download progress: {downloaded_size}/{total_size} bytes ({progress:.2f}%)")
                
                logger.info(f"Download completed. Total bytes downloaded: {downloaded_size}")
                
                # Verify download is complete
                if content_length and int(content_length) != downloaded_size:
                    logger.error(f"Download incomplete. Expected {content_length} bytes, got {downloaded_size} bytes")
                    return jsonify({
                        'status': 'error',
                        'message': 'Video download was incomplete. Please try again.'
                    }), 400
                
                # Validate the downloaded file
                try:
                    # Check if file exists and has content
                    if not os.path.exists(input_path) or os.path.getsize(input_path) == 0:
                        return jsonify({
                            'status': 'error',
                            'message': 'The downloaded video file is empty or incomplete.'
                        }), 400
                    
                    # Use ffprobe to validate the video
                    probe_cmd = [
                        'ffprobe',
                        '-v', 'error',
                        '-print_format', 'json',
                        '-show_format',
                        '-show_streams',
                        input_path
                    ]
                    
                    logger.info(f"Running ffprobe validation: {' '.join(probe_cmd)}")
                    probe_result = subprocess.run(probe_cmd, capture_output=True, text=True)
                    
                    if probe_result.returncode != 0:
                        logger.error(f"Invalid video file: {probe_result.stderr}")
                        return jsonify({
                            'status': 'error',
                            'message': 'The downloaded video file appears to be invalid or corrupted.'
                        }), 400
                    
                    logger.info(f"FFprobe result: {probe_result.stdout}")
                    
                    # Try to fix the video file
                    try:
                        logger.info("Attempting to fix video file format")
                        fixed_path = os.path.join(upload_dir, 'fixed.mp4')
                        
                        # First try: Use -movflags +faststart
                        fix_cmd1 = [
                            'ffmpeg',
                            '-y',
                            '-v', 'error',
                            '-i', input_path,
                            '-c', 'copy',
                            '-movflags', '+faststart',
                            fixed_path
                        ]
                        
                        logger.info(f"Running first fix attempt: {' '.join(fix_cmd1)}")
                        fix_result1 = subprocess.run(fix_cmd1, capture_output=True, text=True)
                        
                        if fix_result1.returncode != 0:
                            logger.warning(f"First fix attempt failed: {fix_result1.stderr}")
                            
                            # Second try: Re-encode the video
                            logger.info("Attempting re-encoding fix")
                            fix_cmd2 = [
                                'ffmpeg',
                                '-y',
                                '-v', 'error',
                                '-i', input_path,
                                '-c:v', 'libx264',
                                '-c:a', 'aac',
                                '-movflags', '+faststart',
                                fixed_path
                            ]
                            
                            logger.info(f"Running second fix attempt: {' '.join(fix_cmd2)}")
                            fix_result2 = subprocess.run(fix_cmd2, capture_output=True, text=True)
                            
                            if fix_result2.returncode != 0:
                                logger.error(f"Second fix attempt failed: {fix_result2.stderr}")
                                return jsonify({
                                    'status': 'error',
                                    'message': 'Could not process the downloaded video file.'
                                }), 400
                        
                        # Replace original input path with fixed file
                        os.remove(input_path)
                        input_path = fixed_path
                        logger.info(f"Successfully fixed video file: {input_path}")
                        
                    except Exception as e:
                        logger.error(f"Error fixing video file: {str(e)}")
                        return jsonify({
                            'status': 'error',
                            'message': 'Error processing downloaded video file.'
                        }), 500
                    
                except Exception as e:
                    logger.error(f"Error validating downloaded file: {str(e)}")
                    return jsonify({
                        'status': 'error',
                        'message': 'Error validating downloaded video file.'
                    }), 400
                    
            except requests.RequestException as e:
                logger.error(f"Error downloading video: {str(e)}")
                return jsonify({
                    'status': 'error',
                    'message': f'Error downloading video: {str(e)}'
                }), 500
        else:
            return jsonify({
                'status': 'error',
                'message': 'No file or videoSrc provided'
            }), 400

        # Extract segment and convert to mp3
        with tempfile.NamedTemporaryFile(delete=False, suffix='.mp3') as temp_out:
            output_path = temp_out.name
            logger.info(f"Created output file: {output_path}")

        try:
            # Use ffmpeg-python to extract audio
            logger.info(f"Extracting audio segment from {start}s to {end}s")
            
            # Process the audio with more detailed logging
            logger.info("Starting FFmpeg processing...")
            
            # Use direct FFmpeg command with explicit options
            ffmpeg_cmd = [
                'ffmpeg',
                '-y',  # Overwrite output file
                '-v', 'error',  # Only show errors
                '-i', input_path,  # Input file
                '-ss', str(start),  # Start time
                '-t', str(duration),  # Duration
                '-vn',  # No video
                '-acodec', 'libmp3lame',  # Audio codec
                '-ab', '128k',  # Audio bitrate
                '-ac', '1',  # Mono audio
                '-ar', '44100',  # Sample rate
                '-f', 'mp3',  # Force MP3 format
                '-movflags', '+faststart',  # Move metadata to start of file
                output_path  # Output file
            ]
            
            logger.info(f"Running FFmpeg command: {' '.join(ffmpeg_cmd)}")
            
            # Run FFmpeg
            result = subprocess.run(ffmpeg_cmd, capture_output=True, text=True)
            
            if result.stderr:
                logger.warning(f"FFmpeg stderr output: {result.stderr}")
                # Check for specific error conditions
                if "moov atom not found" in result.stderr:
                    logger.error("Input video file appears to be incomplete or corrupted")
                    return jsonify({
                        'status': 'error',
                        'message': 'The video file appears to be incomplete or corrupted. Please ensure the video is fully uploaded before processing.'
                    }), 400
                elif "Invalid data found" in result.stderr:
                    logger.error("Invalid video data detected")
                    return jsonify({
                        'status': 'error',
                        'message': 'The video file contains invalid data. Please check the file format and try again.'
                    }), 400
            
            if result.returncode != 0:
                logger.error(f"FFmpeg failed with return code {result.returncode}")
                logger.error(f"FFmpeg error output: {result.stderr}")
                raise Exception(f"FFmpeg failed with return code {result.returncode}")
            
            # Verify the output file exists and has content
            if not os.path.exists(output_path):
                raise Exception("FFmpeg output file was not created")
            
            file_size = os.path.getsize(output_path)
            logger.info(f"FFmpeg output file size: {file_size} bytes")
            
            if file_size == 0:
                raise Exception("FFmpeg output file is empty")
            
            logger.info("Audio extraction completed successfully")
            
        except subprocess.CalledProcessError as e:
            logger.error(f"FFmpeg error: {e.stderr}")
            return jsonify({'status': 'error', 'message': 'ffmpeg error', 'stderr': e.stderr}), 500
        except Exception as e:
            logger.error(f"Error processing audio: {str(e)}")
            return jsonify({'status': 'error', 'message': f'Error processing audio: {str(e)}'}), 500

        # Send to audd.io with enhanced error handling
        try:
            logger.info("Sending audio to audd.io")
            with open(output_path, 'rb') as f:
                files = {'file': f}
                data = {
                    'api_token': AUDD_API_TOKEN,
                    'return': 'apple_music,spotify'
                }
                r = requests.post('https://api.audd.io/', data=data, files=files)
                r.raise_for_status()
                result = r.json()
                
                # Log the complete response for debugging
                logger.info(f"Complete audd.io response: {json.dumps(result, indent=2)}")
                
                if result.get('status') == 'error':
                    error_message = result.get('error', {}).get('message', 'Unknown error from audd.io')
                    logger.error(f"audd.io API error: {error_message}")
                    return jsonify({
                        'status': 'error',
                        'message': f'audd.io API error: {error_message}',
                        'details': result
                    }), 500
                
                logger.info(f"Received response from audd.io: {result.get('status', 'unknown')}")
                
        except requests.RequestException as e:
            logger.error(f"Error calling audd.io: {str(e)}")
            result = {'status': 'error', 'message': f'Error calling audd.io: {str(e)}'}
        except Exception as e:
            logger.error(f"Error processing audd.io response: {str(e)}")
            result = {'status': 'error', 'message': 'Invalid response from audd.io'}

        # Clean up temp files
        try:
            os.remove(output_path)
            if file or video_src:
                # Remove the entire temporary directory
                if 'upload_dir' in locals():
                    import shutil
                    shutil.rmtree(upload_dir)
            logger.info("Cleaned up temporary files")
        except Exception as e:
            logger.error(f"Error cleaning up files: {str(e)}")

        return jsonify(result)

    except Exception as e:
        logger.error(f"Unexpected error in recognize_audio: {str(e)}")
        return jsonify({'status': 'error', 'message': f'Unexpected error: {str(e)}'}), 500

@app.route('/api/parse-cue-sheet', methods=['POST'])
def parse_cue_sheet():
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400
            
        filename = secure_filename(file.filename)
        ext = filename.split('.')[-1].lower()
        
        # Validate file type
        if ext not in ['xlsx', 'xls', 'csv']:
            return jsonify({'error': f'Unsupported file type: {ext}. Please upload CSV or Excel files.'}), 400
        
        # Read file into pandas DataFrame
        try:
            if ext in ['xlsx', 'xls']:
                df = pd.read_excel(file, header=None)
            elif ext == 'csv':
                df = pd.read_csv(file, header=None)
        except Exception as e:
            logger.error(f"Error reading file {filename}: {str(e)}")
            return jsonify({'error': f'Error reading file: {str(e)}'}), 400
        
        # Check if file has enough rows
        if len(df) < 8:
            return jsonify({'error': 'File must have at least 8 rows (6 metadata + 1 header + 1 data row)'}), 400
        
        df = df.fillna('')  # Replace NaN with blank
        rows = df.values.tolist()
        
        # Extract metadata, header, and data
        metadata = [[str(cell) if cell is not None else '' for cell in row] for row in rows[:6]]
        header = [str(cell) if cell is not None else '' for cell in rows[6]] if len(rows) > 6 else []
        data_rows = rows[7:] if len(rows) > 7 else []
        
        # Filter out empty rows (rows where all cells are empty or whitespace)
        filtered_data_rows = []
        for i, row in enumerate(data_rows):
            # Check if row has any non-empty content
            row_has_content = any(str(cell).strip() for cell in row if cell is not None)
            if row_has_content:
                filtered_data_rows.append(row)
            else:
                logger.debug(f"Filtering out empty row {i}: {row}")
        
        logger.info(f"Original data rows: {len(data_rows)}")
        logger.info(f"Filtered data rows: {len(filtered_data_rows)}")
        
        # Convert data rows to list of dicts, all values as strings
        data = [dict(zip(header, [str(cell) if cell is not None else '' for cell in row])) for row in filtered_data_rows]
        
        logger.info(f"Final data dicts: {len(data)}")
        if len(data) > 0:
            logger.info(f"Sample data dict: {data[0]}")
        
        return jsonify({
            'metadata': metadata,
            'header': header,
            'data': data,
            'originalRowCount': len(data_rows),
            'filteredRowCount': len(filtered_data_rows)
        })
        
    except Exception as e:
        logger.error(f"Error in parse_cue_sheet: {str(e)}")
        return jsonify({'error': f'Server error: {str(e)}'}), 500

@app.route('/test-ffmpeg')
def test_ffmpeg():
    try:
        result = subprocess.run(['ffmpeg', '-version'], capture_output=True, text=True)
        if result.returncode == 0:
            return jsonify({
                'status': 'success',
                'message': 'FFmpeg is installed',
                'version': result.stdout.split('\n')[0]
            })
        else:
            return jsonify({
                'status': 'error',
                'message': 'FFmpeg command failed',
                'error': result.stderr
            }), 500
    except Exception as e:
        return jsonify({
            'status': 'error',
            'message': f'Error checking FFmpeg: {str(e)}'
        }), 500

@app.route('/api/vlc/play', methods=['POST'])
def vlc_play():
    try:
        video_path = request.json.get('video_path')
        if not video_path:
            return jsonify({'error': 'No video path provided'}), 400

        # Create media from path
        media = vlc_instance.media_new(video_path)
        player.set_media(media)
        player.play()
        
        return jsonify({
            'status': 'success',
            'message': 'Video started playing',
            'duration': player.get_length() / 1000  # Convert to seconds
        })
    except Exception as e:
        logger.error(f"Error playing video with VLC: {str(e)}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/vlc/pause', methods=['POST'])
def vlc_pause():
    try:
        player.pause()
        return jsonify({'status': 'success', 'message': 'Video paused'})
    except Exception as e:
        logger.error(f"Error pausing video: {str(e)}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/vlc/stop', methods=['POST'])
def vlc_stop():
    try:
        player.stop()
        return jsonify({'status': 'success', 'message': 'Video stopped'})
    except Exception as e:
        logger.error(f"Error stopping video: {str(e)}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/vlc/seek', methods=['POST'])
def vlc_seek():
    try:
        position = request.json.get('position')  # Position in seconds
        if position is None:
            return jsonify({'error': 'No position provided'}), 400
            
        # Convert position to milliseconds
        position_ms = int(position * 1000)
        player.set_time(position_ms)
        
        return jsonify({'status': 'success', 'message': f'Seeked to {position} seconds'})
    except Exception as e:
        logger.error(f"Error seeking video: {str(e)}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/vlc/status', methods=['GET'])
def vlc_status():
    try:
        state = player.get_state()
        position = player.get_time() / 1000  # Convert to seconds
        duration = player.get_length() / 1000  # Convert to seconds
        
        return jsonify({
            'status': 'success',
            'state': str(state),
            'position': position,
            'duration': duration
        })
    except Exception as e:
        logger.error(f"Error getting VLC status: {str(e)}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/vlc/volume', methods=['POST'])
def vlc_volume():
    try:
        volume = request.json.get('volume')
        if volume is None:
            return jsonify({'error': 'No volume provided'}), 400
            
        # Set volume (0-100)
        player.audio_set_volume(int(volume * 100))
        
        return jsonify({'status': 'success', 'message': f'Volume set to {volume * 100}%'})
    except Exception as e:
        logger.error(f"Error setting volume: {str(e)}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/music-co', methods=['GET'])
def list_music_co():
    """Lists music companies from Firestore, with optional search."""
    try:
        if firestore_client is None:
            return jsonify({'error': 'Firestore not initialized'}), 503
        
        search = request.args.get('search', '').lower()
        docs = firestore_client.collection(MUSIC_CO_COLLECTION).stream()
        results = []
        for doc in docs:
            data = doc.to_dict()
            name = data.get('name', '')
            if not search or search in name.lower():
                results.append({'id': doc.id, 'name': name})
        return jsonify(results)
    except Exception as e:
        logger.error(f"Error listing music companies: {str(e)}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/music-co', methods=['POST'])
def add_music_co():
    """Adds a new music company to Firestore."""
    try:
        if firestore_client is None:
            return jsonify({'error': 'Firestore not initialized'}), 503
        
        data = request.get_json()
        name = data.get('name', '').strip()
        
        if not name:
            return jsonify({'error': 'Music company name is required'}), 400
        
        # Check if name already exists
        existing_docs = firestore_client.collection(MUSIC_CO_COLLECTION).where('name', '==', name).stream()
        if list(existing_docs):
            return jsonify({'error': 'Music company with this name already exists'}), 409
        
        # Add new music company
        doc_ref = firestore_client.collection(MUSIC_CO_COLLECTION).add({'name': name})
        
        return jsonify({'id': doc_ref[1].id, 'name': name, 'status': 'success'})
    except Exception as e:
        logger.error(f"Error adding music company: {str(e)}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/music-co/<doc_id>', methods=['DELETE'])
def delete_music_co(doc_id):
    """Deletes a music company from Firestore."""
    try:
        if firestore_client is None:
            return jsonify({'error': 'Firestore not initialized'}), 503
        
        # Check if document exists
        doc_ref = firestore_client.collection(MUSIC_CO_COLLECTION).document(doc_id)
        if not doc_ref.get().exists:
            return jsonify({'error': 'Music company not found'}), 404
        
        # Delete the document
        doc_ref.delete()
        
        return jsonify({'status': 'success', 'message': 'Music company deleted successfully'})
    except Exception as e:
        logger.error(f"Error deleting music company: {str(e)}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/usage', methods=['GET'])
def list_usage():
    """Lists usage options from Firestore."""
    try:
        if firestore_client is None:
            return jsonify({'error': 'Firestore not initialized'}), 503
            
        usage_ref = firestore_client.collection('usage')
        usage_docs = usage_ref.stream()
        usage_list = [{'id': doc.id, 'name': doc.to_dict()['name']} for doc in usage_docs]
        return jsonify(usage_list)
    except Exception as e:
        logger.error(f"Error listing usage options: {str(e)}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/usage', methods=['POST'])
def add_usage():
    """Adds a new usage option to Firestore."""
    try:
        if firestore_client is None:
            return jsonify({'error': 'Firestore not initialized'}), 503
        
        data = request.get_json()
        name = data.get('name', '').strip()
        
        if not name:
            return jsonify({'error': 'Usage name is required'}), 400
        
        # Check if name already exists
        existing_docs = firestore_client.collection('usage').where('name', '==', name).stream()
        if list(existing_docs):
            return jsonify({'error': 'Usage option with this name already exists'}), 409
        
        # Add new usage option
        doc_ref = firestore_client.collection('usage').add({'name': name})
        
        return jsonify({'id': doc_ref[1].id, 'name': name, 'status': 'success'})
    except Exception as e:
        logger.error(f"Error adding usage option: {str(e)}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/usage/<doc_id>', methods=['DELETE'])
def delete_usage(doc_id):
    """Deletes a usage option from Firestore."""
    try:
        if firestore_client is None:
            return jsonify({'error': 'Firestore not initialized'}), 503
        
        # Check if document exists
        doc_ref = firestore_client.collection('usage').document(doc_id)
        if not doc_ref.get().exists:
            return jsonify({'error': 'Usage option not found'}), 404
        
        # Delete the document
        doc_ref.delete()
        
        return jsonify({'status': 'success', 'message': 'Usage option deleted successfully'})
    except Exception as e:
        logger.error(f"Error deleting usage option: {str(e)}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/unknown-tags', methods=['GET'])
def list_unknown_tags():
    """Lists title tag presets from Firestore."""
    try:
        if firestore_client is None:
            return jsonify({'error': 'Firestore not initialized'}), 503
            
        title_tag_preset_ref = firestore_client.collection('title_tag_preset')
        title_tag_preset_docs = title_tag_preset_ref.stream()
        title_tag_preset_list = [{'id': doc.id, 'name': doc.to_dict()['name']} for doc in title_tag_preset_docs]
        return jsonify(title_tag_preset_list)
    except Exception as e:
        logger.error(f"Error listing title tag presets: {str(e)}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/unknown-tags', methods=['POST'])
def add_unknown_tag():
    """Adds a new title tag preset to Firestore."""
    try:
        if firestore_client is None:
            return jsonify({'error': 'Firestore not initialized'}), 503
        
        data = request.get_json()
        name = data.get('name', '').strip()
        
        if not name:
            return jsonify({'error': 'Title tag preset name is required'}), 400
        
        # Check if name already exists
        existing_docs = firestore_client.collection('title_tag_preset').where('name', '==', name).stream()
        if list(existing_docs):
            return jsonify({'error': 'Title tag preset with this name already exists'}), 409
        
        # Add new title tag preset
        doc_ref = firestore_client.collection('title_tag_preset').add({'name': name})
        
        return jsonify({'id': doc_ref[1].id, 'name': name, 'status': 'success'})
    except Exception as e:
        logger.error(f"Error adding title tag preset: {str(e)}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/unknown-tags/<doc_id>', methods=['DELETE'])
def delete_unknown_tag(doc_id):
    """Deletes a title tag preset from Firestore."""
    try:
        if firestore_client is None:
            return jsonify({'error': 'Firestore not initialized'}), 503
        
        # Check if document exists
        doc_ref = firestore_client.collection('title_tag_preset').document(doc_id)
        if not doc_ref.get().exists:
            return jsonify({'error': 'Title tag preset not found'}), 404
        
        # Delete the document
        doc_ref.delete()
        
        return jsonify({'status': 'success', 'message': 'Title tag preset deleted successfully'})
    except Exception as e:
        logger.error(f"Error deleting title tag preset: {str(e)}")
        return jsonify({'error': str(e)}), 500

@app.route('/music-co')
def music_co_management():
    """Renders the Music Co Management page."""
    return render_template('music_co.html')

@app.route('/usage')
def usage_management():
    """Renders the Usage Management page."""
    return render_template('usage.html')

@app.route('/unknown-tags')
def unknown_tags_management():
    """Renders the Unknown Tags Management page."""
    return render_template('unknown_tags.html')

@app.route('/export-settings')
def export_settings():
    return render_template('export_settings.html')

@app.route('/api/recognize-gcs-segment', methods=['POST'])
def recognize_gcs_segment():
    """Recognize audio from a segment of a video stored in GCS"""
    try:
        tcr_in = request.form.get('tcrIn')
        tcr_out = request.form.get('tcrOut')
        gcs_path = request.form.get('gcsPath')
        
        logger.info(f"Received GCS recognize request - TCR In: {tcr_in}, TCR Out: {tcr_out}, GCS Path: {gcs_path}")

        if not tcr_in or not tcr_out or not gcs_path:
            return jsonify({'status': 'error', 'message': 'Missing TCR In, Out, or GCS Path'}), 400

        # Calculate start and duration in seconds
        def time_to_seconds(t):
            parts = t.split(':')
            if len(parts) == 3:
                h, m, s = map(float, parts)
                return int(h) * 3600 + int(m) * 60 + s
            elif len(parts) == 4:
                h, m, s, f = map(float, parts)
                return int(h) * 3600 + int(m) * 60 + s + (f / 25)  # Assuming 25fps
            return 0
        
        start = time_to_seconds(tcr_in)
        end = time_to_seconds(tcr_out)
        duration = end - start
        
        if duration <= 0:
            return jsonify({'status': 'error', 'message': 'Invalid time range'}), 400

        # Download video segment from GCS
        try:
            # Use the globally initialized storage client
            bucket = storage_client.bucket(os.getenv('GCS_BUCKET_NAME', 'mos-aat'))
            blob = bucket.blob(gcs_path)
            
            if not blob.exists():
                return jsonify({'status': 'error', 'message': 'Video file not found in GCS'}), 404
            
            # Create temporary directory for processing
            temp_dir = tempfile.mkdtemp()
            input_path = os.path.join(temp_dir, 'input.mp4')
            
            # Download the video file
            logger.info(f"Downloading video from GCS: {gcs_path}")
            blob.download_to_filename(input_path)
            
        except Exception as e:
            logger.error(f"Error downloading from GCS: {str(e)}")
            return jsonify({'status': 'error', 'message': f'Error downloading video: {str(e)}'}), 500

        # Extract audio segment
        output_path = os.path.join(temp_dir, 'output.mp3')
        
        try:
            logger.info(f"Extracting audio segment from {start}s to {end}s")
            
            ffmpeg_cmd = [
                'ffmpeg',
                '-y',  # Overwrite output file
                '-v', 'error',  # Only show errors
                '-i', input_path,  # Input file
                '-ss', str(start),  # Start time
                '-t', str(duration),  # Duration
                '-vn',  # No video
                '-acodec', 'libmp3lame',  # Audio codec
                '-ab', '128k',  # Audio bitrate
                '-ac', '1',  # Mono audio
                '-ar', '44100',  # Sample rate
                '-f', 'mp3',  # Force MP3 format
                output_path  # Output file
            ]
            
            logger.info(f"Running FFmpeg command: {' '.join(ffmpeg_cmd)}")
            
            result = subprocess.run(ffmpeg_cmd, capture_output=True, text=True)
            
            if result.returncode != 0:
                logger.error(f"FFmpeg error: {result.stderr}")
                return jsonify({'status': 'error', 'message': f'Error processing audio: {result.stderr}'}), 500
            
            # Check if output file was created and has content
            if not os.path.exists(output_path) or os.path.getsize(output_path) == 0:
                return jsonify({'status': 'error', 'message': 'Failed to extract audio segment'}), 500
                
            logger.info(f"Audio extraction successful. File size: {os.path.getsize(output_path)} bytes")
            
        except Exception as e:
            logger.error(f"Error processing audio: {str(e)}")
            return jsonify({'status': 'error', 'message': f'Error processing audio: {str(e)}'}), 500

        # Send to audd.io
        try:
            logger.info("Sending audio to audd.io")
            with open(output_path, 'rb') as f:
                files = {'file': f}
                data = {
                    'api_token': AUDD_API_TOKEN,
                    'return': 'apple_music,spotify'
                }
                r = requests.post('https://api.audd.io/', data=data, files=files)
                r.raise_for_status()
                result = r.json()
                
                logger.info(f"Complete audd.io response: {json.dumps(result, indent=2)}")
                
                if result.get('status') == 'error':
                    error_message = result.get('error', {}).get('message', 'Unknown error from audd.io')
                    logger.error(f"audd.io API error: {error_message}")
                    return jsonify({
                        'status': 'error',
                        'message': f'audd.io API error: {error_message}',
                        'details': result
                    }), 500
                
                logger.info(f"Received response from audd.io: {result.get('status', 'unknown')}")
                
        except requests.RequestException as e:
            logger.error(f"Error calling audd.io: {str(e)}")
            result = {'status': 'error', 'message': f'Error calling audd.io: {str(e)}'}
        except Exception as e:
            logger.error(f"Error processing audd.io response: {str(e)}")
            result = {'status': 'error', 'message': 'Invalid response from audd.io'}

        # Clean up temp files
        try:
            import shutil
            shutil.rmtree(temp_dir)
            logger.info("Cleaned up temporary files")
        except Exception as e:
            logger.error(f"Error cleaning up files: {str(e)}")

        return jsonify(result)

    except Exception as e:
        logger.error(f"Unexpected error in recognize_gcs_segment: {str(e)}")
        return jsonify({'status': 'error', 'message': f'Unexpected error: {str(e)}'}), 500

@app.route('/api/generate-upload-url', methods=['POST'])
def generate_upload_url():
    """Generate a signed URL for uploading a file directly to GCS."""
    try:
        data = request.get_json()
        filename = data.get('filename')
        content_type = data.get('contentType')

        if not filename or not content_type:
            return jsonify({'error': 'Missing filename or content type'}), 400

        # Generate a unique name for the file in GCS
        file_extension = os.path.splitext(filename)[1]
        unique_filename = f"{uuid.uuid4()}{file_extension}"
        gcs_path = f"videos/{unique_filename}"
        
        bucket_name = os.getenv('GCS_BUCKET_NAME')
        if not bucket_name:
            raise Exception("GCS_BUCKET_NAME environment variable not set.")
        
        bucket = storage_client.bucket(bucket_name)
        blob = bucket.blob(gcs_path)

        # By removing the service_account_email and access_token, the library
        # will automatically use the Cloud Run service's identity, which has
        # the necessary "Service Account Token Creator" role.
        url = blob.generate_signed_url(
            version="v4",
            expiration=timedelta(minutes=15),
            method="PUT",
            content_type=content_type
        )

        logger.info(f"Generated signed URL for {gcs_path} using the service's default identity.")
        
        return jsonify({
            'status': 'success',
            'signedUrl': url,
            'gcsPath': gcs_path,
            'filename': unique_filename
        })

    except Exception as e:
        logger.error(f"Error generating signed URL: {str(e)}")
        return jsonify({'error': f'Failed to generate upload URL: {str(e)}'}), 500

@app.route('/view-recognition')
def view_recognition():
    """Renders a page to display recognition data."""
    return render_template('view_recognition.html')

@app.route('/api/loadsave')
def loadsave():
    try:
        if firestore_client is None:
            logger.error("Firestore not initialized")
            return jsonify({'error': 'Database not available'}), 503
            
        session_id = session.get('session_id')
        if not session_id:
            return jsonify({})
            
        try:
            doc = firestore_client.collection('autosaves').document(session_id).get()
            if doc.exists:
                data = doc.to_dict()
                # Validate data before sending
                if not isinstance(data.get('markers'), list):
                    return jsonify({})
                return jsonify(data)
            return jsonify({})
        except Exception as db_error:
            logger.error(f"Database operation failed: {str(db_error)}")
            return jsonify({'error': 'Failed to load data'}), 500
            
    except Exception as e:
        logger.error(f"Error in loadsave: {str(e)}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/autosave', methods=['POST'])
def autosave():
    """Save markers and other data to Firestore for session-based autosave"""
    try:
        if firestore_client is None:
            logger.error("Firestore not initialized")
            return jsonify({'error': 'Database not available'}), 503
        
        # Generate or get session ID
        if 'session_id' not in session:
            session['session_id'] = str(uuid.uuid4())
        
        session_id = session['session_id']
        
        # Get data from request
        data = request.get_json()
        if not data:
            return jsonify({'error': 'No data provided'}), 400
        
        # Validate markers data
        markers = data.get('markers', [])
        if not isinstance(markers, list):
            return jsonify({'error': 'Invalid markers data'}), 400
        
        # Prepare data for storage
        save_data = {
            'markers': markers,
            'videoState': data.get('videoState'),
            'exportSettings': data.get('exportSettings', {}),
            'markedRows': data.get('markedRows', {}),
            'exceptionSettings': data.get('exceptionSettings', {}),
            'timestamp': datetime.now().isoformat()
        }
        
        # Save to Firestore
        try:
            firestore_client.collection('autosaves').document(session_id).set(save_data)
            logger.info(f"Autosave successful for session {session_id}")
            return jsonify({'status': 'success', 'message': 'Data saved successfully'})
        except Exception as db_error:
            logger.error(f"Database operation failed: {str(db_error)}")
            return jsonify({'error': 'Failed to save data'}), 500
            
    except Exception as e:
        logger.error(f"Error in autosave: {str(e)}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/upload-video-to-gcs', methods=['POST'])
def upload_video_to_gcs():
    """Upload video file to GCS and return the GCS path"""
    try:
        if 'video' not in request.files:
            return jsonify({'error': 'No video file provided'}), 400
        
        file = request.files['video']
        if file.filename == '':
            return jsonify({'error': 'No selected file'}), 400

        # Generate unique filename
        file_extension = os.path.splitext(file.filename)[1]
        unique_filename = f"{uuid.uuid4()}{file_extension}"
        gcs_path = f"videos/{unique_filename}"
        
        # Create temporary file
        with tempfile.NamedTemporaryFile(delete=False, suffix=file_extension) as temp_file:
            file.save(temp_file.name)
            temp_path = temp_file.name
        
        try:
            # Use the globally initialized storage client
            bucket = storage_client.bucket(os.getenv('GCS_BUCKET_NAME', 'mos-aat'))
            blob = bucket.blob(gcs_path)
            
            # Upload with content type
            content_type = 'video/mp4' if file_extension.lower() == '.mp4' else 'video/x-ms-wmv'
            blob.upload_from_filename(temp_path, content_type=content_type)
            
            logger.info(f"Successfully uploaded video to GCS: {gcs_path}")
            
            return jsonify({
                'status': 'success',
                'gcs_path': gcs_path,
                'filename': unique_filename
            })
            
        finally:
            # Clean up temporary file
            if os.path.exists(temp_path):
                os.remove(temp_path)
                
    except Exception as e:
        logger.error(f"Error uploading video to GCS: {str(e)}")
        return jsonify({'error': f'Upload failed: {str(e)}'}), 500

@app.route('/api/list-cloud-videos', methods=['GET'])
def list_cloud_videos():
    """List all videos in the cloud storage videos folder"""
    try:
        if storage_client is None:
            return jsonify({'error': 'GCS client not initialized'}), 503
        
        bucket_name = os.getenv('GCS_BUCKET_NAME', 'mos-aat')
        bucket = storage_client.bucket(bucket_name)
        
        # List all blobs in the videos folder
        blobs = bucket.list_blobs(prefix='videos/')
        
        videos = []
        for blob in blobs:
            # Skip the videos/ folder itself
            if blob.name == 'videos/':
                continue
                
            # Extract filename from path
            filename = os.path.basename(blob.name)
            
            # Create proxy URL for serving the video
            proxy_url = f"/api/serve-video/{blob.name}"
            
            video_info = {
                'name': blob.name,
                'filename': filename,
                'size': blob.size,
                'size_mb': round(blob.size / (1024 * 1024), 2),
                'created': blob.time_created.isoformat() if blob.time_created else None,
                'updated': blob.updated.isoformat() if blob.updated else None,
                'content_type': blob.content_type,
                'proxyUrl': proxy_url
            }
            videos.append(video_info)
        
        # Sort by creation date (newest first)
        videos.sort(key=lambda x: x['created'] or '', reverse=True)
        
        logger.info(f"Listed {len(videos)} videos from cloud storage")
        
        return jsonify({
            'status': 'success',
            'videos': videos,
            'total_count': len(videos)
        })
        
    except Exception as e:
        logger.error(f"Error listing cloud videos: {str(e)}")
        return jsonify({'error': f'Failed to list videos: {str(e)}'}), 500

@app.route('/api/load-cloud-video', methods=['POST'])
def load_cloud_video():
    """Load a specific video from cloud storage"""
    try:
        data = request.get_json()
        gcs_path = data.get('gcsPath')
        
        if not gcs_path:
            return jsonify({'error': 'No GCS path provided'}), 400
        
        if storage_client is None:
            return jsonify({'error': 'GCS client not initialized'}), 503
        
        bucket_name = os.getenv('GCS_BUCKET_NAME', 'mos-aat')
        bucket = storage_client.bucket(bucket_name)
        blob = bucket.blob(gcs_path)
        
        if not blob.exists():
            return jsonify({'error': 'Video not found in cloud storage'}), 404
        
        # Use server-side proxy URL
        proxy_url = f"/api/serve-video/{gcs_path}"
        
        # Extract filename from path
        filename = os.path.basename(gcs_path)
        
        logger.info(f"Loaded cloud video: {gcs_path}")
        
        return jsonify({
            'status': 'success',
            'gcsPath': gcs_path,
            'proxyUrl': proxy_url,
            'filename': filename,
            'size': blob.size,
            'contentType': blob.content_type
        })
        
    except Exception as e:
        logger.error(f"Error loading cloud video: {str(e)}")
        return jsonify({'error': f'Failed to load video: {str(e)}'}), 500

@app.route('/api/delete-cloud-video', methods=['DELETE'])
def delete_cloud_video():
    """Delete a video from cloud storage"""
    try:
        data = request.get_json()
        gcs_path = data.get('gcsPath')
        
        if not gcs_path:
            return jsonify({'error': 'No GCS path provided'}), 400
        
        if storage_client is None:
            return jsonify({'error': 'GCS client not initialized'}), 503
        
        bucket_name = os.getenv('GCS_BUCKET_NAME', 'mos-aat')
        bucket = storage_client.bucket(bucket_name)
        blob = bucket.blob(gcs_path)
        
        if not blob.exists():
            return jsonify({'error': 'Video not found in cloud storage'}), 404
        
        # Delete the blob
        blob.delete()
        
        logger.info(f"Deleted cloud video: {gcs_path}")
        
        return jsonify({
            'status': 'success',
            'message': f'Video {gcs_path} deleted successfully'
        })
        
    except Exception as e:
        logger.error(f"Error deleting cloud video: {str(e)}")
        return jsonify({'error': f'Failed to delete video: {str(e)}'}), 500

@app.route('/api/serve-video/<path:gcs_path>')
def serve_video_proxy(gcs_path):
    """Serve video files from GCS through a server-side proxy"""
    try:
        logger.info(f"Attempting to serve video: {gcs_path}")
        
        if storage_client is None:
            logger.error("GCS client not initialized")
            return jsonify({'error': 'GCS client not initialized'}), 503
        
        # URL decode the path to handle special characters
        gcs_path = urllib.parse.unquote(gcs_path)
        logger.info(f"Decoded path: {gcs_path}")
        
        bucket_name = os.getenv('GCS_BUCKET_NAME', 'mos-aat')
        logger.info(f"Using bucket: {bucket_name}")
        
        bucket = storage_client.bucket(bucket_name)
        blob = bucket.blob(gcs_path)
        
        logger.info(f"Checking if blob exists: {gcs_path}")
        if not blob.exists():
            logger.error(f"Video not found: {gcs_path}")
            return jsonify({'error': 'Video not found'}), 404
        
        logger.info(f"Blob exists, size: {blob.size} bytes")
        
        # Check if this is a range request (for video seeking)
        range_header = request.headers.get('Range', None)
        logger.info(f"Range header: {range_header}")
        
        if range_header:
            # Handle range requests for video seeking
            try:
                # Parse range header (e.g., "bytes=0-1023")
                range_match = re.match(r'bytes=(\d+)-(\d*)', range_header)
                if range_match:
                    start = int(range_match.group(1))
                    end = int(range_match.group(2)) if range_match.group(2) else blob.size - 1
                    
                    logger.info(f"Downloading range: {start}-{end}")
                    
                    # Use streaming for range requests
                    def generate_range():
                        try:
                            # Download the range in chunks
                            chunk_size = 1024 * 1024  # 1MB chunks
                            current_pos = start
                            
                            while current_pos <= end:
                                chunk_end = min(current_pos + chunk_size - 1, end)
                                chunk_data = blob.download_as_bytes(start=current_pos, end=chunk_end)
                                yield chunk_data
                                current_pos = chunk_end + 1
                        except Exception as e:
                            logger.error(f"Error in range streaming: {str(e)}")
                            raise
                    
                    response = app.response_class(
                        generate_range(),
                        status=206,  # Partial Content
                        mimetype=blob.content_type or 'video/mp4'
                    )
                    
                    response.headers['Accept-Ranges'] = 'bytes'
                    response.headers['Content-Length'] = str(end - start + 1)
                    response.headers['Content-Range'] = f'bytes {start}-{end}/{blob.size}'
                    response.headers['Cache-Control'] = 'public, max-age=3600'
                    
                    logger.info(f"Served video range: {gcs_path} ({start}-{end}/{blob.size})")
                    return response
                    
            except Exception as e:
                logger.error(f"Error handling range request for {gcs_path}: {str(e)}")
                # Fall back to full download
        
        # Full video streaming (for non-range requests or fallback)
        try:
            logger.info(f"Streaming full video: {gcs_path}")
            
            # Use streaming for full video download
            def generate_stream():
                try:
                    # Download in chunks to avoid memory issues
                    chunk_size = 1024 * 1024  # 1MB chunks
                    current_pos = 0
                    
                    while current_pos < blob.size:
                        chunk_end = min(current_pos + chunk_size - 1, blob.size - 1)
                        chunk_data = blob.download_as_bytes(start=current_pos, end=chunk_end)
                        yield chunk_data
                        current_pos = chunk_end + 1
                        
                except Exception as e:
                    logger.error(f"Error in video streaming: {str(e)}")
                    raise
            
            response = app.response_class(
                generate_stream(),
                status=200,
                mimetype=blob.content_type or 'video/mp4'
            )
            
            # Add headers for video streaming
            response.headers['Accept-Ranges'] = 'bytes'
            response.headers['Content-Length'] = str(blob.size)
            response.headers['Cache-Control'] = 'public, max-age=3600'
            
            logger.info(f"Successfully streaming video: {gcs_path} ({blob.size} bytes)")
            return response
            
        except Exception as e:
            logger.error(f"Error streaming video {gcs_path}: {str(e)}")
            logger.error(f"Exception type: {type(e).__name__}")
            import traceback
            logger.error(f"Traceback: {traceback.format_exc()}")
            return jsonify({'error': f'Failed to stream video: {str(e)}'}), 500
        
    except Exception as e:
        logger.error(f"Error serving video {gcs_path}: {str(e)}")
        logger.error(f"Exception type: {type(e).__name__}")
        import traceback
        logger.error(f"Traceback: {traceback.format_exc()}")
        return jsonify({'error': f'Failed to serve video: {str(e)}'}), 500

@app.route('/cloud-videos')
def cloud_videos():
    return render_template('cloud_videos.html')

@app.route('/subtitle-edit')
def subtitle_edit():
    return render_template('subtitle_edit.html')

@app.route('/api/upload-subtitle', methods=['POST'])
def upload_subtitle():
    if 'file' not in request.files:
        return {'error': 'No file uploaded'}, 400
    file = request.files['file']
    ext = file.filename.split('.')[-1].lower()
    if ext not in ['srt', 'vtt', 'ass', 'sub']:
        return {'error': 'Unsupported file type'}, 400
    content = file.read().decode('utf-8', errors='replace')
    return {'filename': file.filename, 'content': content}

@app.route('/api/download-subtitle', methods=['POST'])
def download_subtitle():
    data = request.json
    content = data.get('content', '')
    filename = data.get('filename', 'subtitles.srt')
    with tempfile.NamedTemporaryFile(delete=False, suffix='.srt', mode='w', encoding='utf-8') as f:
        f.write(content)
        temp_path = f.name
    return send_file(temp_path, as_attachment=True, download_name=filename)

@app.route('/srt-check')
def srt_check_page():
    return render_template('srt_check.html')

@app.route('/api/srt-check', methods=['POST'])
def srt_check_api():
    try:
        data = request.get_json()
        srt_content = data.get('content', '')
        language = data.get('language', 'en')
        frame_rate = float(data.get('frame_rate', 25.0))
        max_cps = int(data.get('max_cps', 20))
        max_line_length = int(data.get('max_line_length', 42))
        max_lines = int(data.get('max_lines', 2))
        min_duration = int(data.get('min_duration', 833))
        max_duration = int(data.get('max_duration', 7000))
        min_gap_frames = int(data.get('min_gap_frames', 2))
        max_gap_frames = int(data.get('max_gap_frames', 12))
        speaker_style = data.get('speaker_style', 'dash_both_lines_with_space')
        number_spelling_max = int(data.get('number_spelling_max', 10))
        sound_effects = data.get('sound_effects', 'music,laughter,applause,sighs,gasps,whispers')
        checks = data.get('checks', {})

        # Parse SRT
        import re
        def parse_srt(srt):
            pattern = re.compile(r"(\d+)\s+([\d:,]+)\s+-->\s+([\d:,]+)\s+([\s\S]+?)(?=\n\d+\n|\Z)", re.MULTILINE)
            subs = []
            for match in pattern.finditer(srt):
                idx = int(match.group(1))
                start = match.group(2)
                end = match.group(3)
                text = match.group(4).strip().replace('\r', '')
                subs.append({'idx': idx, 'start': start, 'end': end, 'text': text})
            return subs

        def time_to_seconds(t):
            t = t.replace(',', '.')
            parts = re.split('[:.]', t)
            if len(parts) == 4:
                h, m, s, ms = map(int, parts)
                return h*3600 + m*60 + s + ms/1000
            elif len(parts) == 3:
                h, m, s = map(int, parts)
                return h*3600 + m*60 + s
            return 0

        def time_to_ms(t):
            return time_to_seconds(t) * 1000

        def ms_from_time(t):
            t = t.replace(',', '.')
            parts = re.split('[:.]', t)
            if len(parts) == 4:
                h, m, s, ms = map(int, parts)
                return (h*3600 + m*60 + s)*1000 + ms
            elif len(parts) == 3:
                h, m, s = map(int, parts)
                return (h*3600 + m*60 + s)*1000
            return 0

        def convert_number_to_string(num, language='en'):
            numbers = {
                'en': {'1': 'one', '2': 'two', '3': 'three', '4': 'four', '5': 'five',
                       '6': 'six', '7': 'seven', '8': 'eight', '9': 'nine', '10': 'ten'}
            }
            return numbers.get(language, numbers['en']).get(str(num), str(num))

        errors = []
        subs = parse_srt(srt_content)

        # 1. Max CPS and Line Length
        for sub in subs:
            start_sec = time_to_seconds(sub['start'])
            end_sec = time_to_seconds(sub['end'])
            duration = max(end_sec - start_sec, 0.001)
            
            # Max line length
            lines = sub['text'].split('\n')
            for i, line in enumerate(lines):
                if len(line) > max_line_length:
                    errors.append({
                        'idx': sub['idx'], 'start': sub['start'], 'end': sub['end'],
                        'type': 'Line Length', 'message': f'Line {i+1} exceeds max length ({len(line)}/{max_line_length})', 'text': line
                    })
            
            # Max CPS
            total_chars = len(sub['text'].replace('\n', ''))
            cps = total_chars / duration if duration > 0 else 0
            if cps > max_cps:
                errors.append({
                    'idx': sub['idx'], 'start': sub['start'], 'end': sub['end'],
                    'type': 'CPS', 'message': f'Characters per second ({cps:.2f}) exceeds max ({max_cps})', 'text': sub['text']
                })

        # 2. Min Duration
        for sub in subs:
            duration_ms = ms_from_time(sub['end']) - ms_from_time(sub['start'])
            min_duration_actual = 500 if language == 'ja' else min_duration
            if duration_ms < min_duration_actual:
                errors.append({
                    'idx': sub['idx'], 'start': sub['start'], 'end': sub['end'],
                    'type': 'Min Duration', 'message': f'Duration ({duration_ms}ms) below minimum ({min_duration_actual}ms)', 'text': sub['text']
                })

        # 3. Max Duration
        for sub in subs:
            duration_ms = ms_from_time(sub['end']) - ms_from_time(sub['start'])
            if duration_ms > max_duration:
                errors.append({
                    'idx': sub['idx'], 'start': sub['start'], 'end': sub['end'],
                    'type': 'Max Duration', 'message': f'Duration ({duration_ms}ms) exceeds maximum ({max_duration}ms)', 'text': sub['text']
                })

        # 4. Number of Lines
        for sub in subs:
            line_count = len(sub['text'].split('\n'))
            if line_count > max_lines:
                errors.append({
                    'idx': sub['idx'], 'start': sub['start'], 'end': sub['end'],
                    'type': 'Line Count', 'message': f'Too many lines ({line_count} > {max_lines})', 'text': sub['text']
                })

        # 5. Ellipses check (use … not ...)
        if checks.get('ellipses', True):
            for sub in subs:
                if '...' in sub['text']:
                    errors.append({
                        'idx': sub['idx'], 'start': sub['start'], 'end': sub['end'],
                        'type': 'Ellipses', 'message': 'Use single smart character (…) not three dots (...)', 'text': sub['text']
                    })

        # 6. Numbers 1-10 should be spelled out
        if checks.get('number_spelling', True) and language not in ['ja', 'ar']:
            for sub in subs:
                text = sub['text']
                # Check for single digits 1-9
                for i in range(1, min(10, number_spelling_max + 1)):
                    pattern = rf'\b{i}\b'
                    if re.search(pattern, text):
                        # Check if it's not part of time, URL, etc.
                        if not re.search(rf'\b{i}[:.]', text) and not re.search(rf'[:.]{i}\b', text):
                            errors.append({
                                'idx': sub['idx'], 'start': sub['start'], 'end': sub['end'],
                                'type': 'Number Spelling', 'message': f'Number {i} should be spelled out ({convert_number_to_string(i, language)})', 'text': text
                            })
                # Check for 10
                if number_spelling_max >= 10 and re.search(r'\b10\b', text) and not re.search(r'\b10[:.]', text):
                    errors.append({
                        'idx': sub['idx'], 'start': sub['start'], 'end': sub['end'],
                        'type': 'Number Spelling', 'message': 'Number 10 should be spelled out (ten)', 'text': text
                    })

        # 7. Dialog formatting (speaker style)
        if checks.get('dialog_format', True) and language != 'ja':
            for sub in subs:
                text = sub['text']
                # Check for inconsistent dash usage based on speaker style
                if '-' in text:
                    lines = text.split('\n')
                    for i, line in enumerate(lines):
                        if line.strip().startswith('-') and not line.strip().startswith('- '):
                            errors.append({
                                'idx': sub['idx'], 'start': sub['start'], 'end': sub['end'],
                                'type': 'Dialog Format', 'message': f'Line {i+1}: Use hyphen with space for speaker', 'text': text
                            })

        # 8. Whitespace checks
        for sub in subs:
            text = sub['text']
            # Start/end whitespace
            if checks.get('start_whitespace', True) and re.match(r'^( |\n|\r\n)[^\s]', text):
                errors.append({'idx': sub['idx'], 'start': sub['start'], 'end': sub['end'], 'type': 'Whitespace', 'message': 'Whitespace at start of subtitle', 'text': text})
            if checks.get('end_whitespace', True) and re.match(r'[^\s]( |\n|\r\n)$', text):
                errors.append({'idx': sub['idx'], 'start': sub['start'], 'end': sub['end'], 'type': 'Whitespace', 'message': 'Whitespace at end of subtitle', 'text': text})
            # Spaces before punctuation
            if checks.get('space_before_punctuation', True) and re.search(r'[^\s]( |\n|\r\n)[!?).,\u061f\u060c\u2026]', text):
                errors.append({'idx': sub['idx'], 'start': sub['start'], 'end': sub['end'], 'type': 'Whitespace', 'message': 'Space before punctuation', 'text': text})
            # 2+ consecutive spaces
            if checks.get('consecutive_spaces', True) and re.search(r'( |\n|\r\n){2,}', text):
                errors.append({'idx': sub['idx'], 'start': sub['start'], 'end': sub['end'], 'type': 'Whitespace', 'message': '2+ consecutive spaces', 'text': text})

        # 9. Italics check
        if checks.get('italics', True):
            for sub in subs:
                text = sub['text']
                if 'i>' in text.lower():
                    errors.append({'idx': sub['idx'], 'start': sub['start'], 'end': sub['end'], 'type': 'Italics', 'message': 'Italics tag detected', 'text': text})

        # 10. Gap checks (bridge gaps, two frames gap)
        for i in range(len(subs)-1):
            cur = subs[i]
            nxt = subs[i+1]
            end_ms = ms_from_time(cur['end'])
            start_ms = ms_from_time(nxt['start'])
            gap_ms = start_ms - end_ms
            gap_frames = (gap_ms) / (1000.0 / frame_rate)
            
            if gap_frames < min_gap_frames:
                errors.append({'idx': cur['idx'], 'start': cur['start'], 'end': cur['end'], 'type': 'Gap', 'message': f'Less than {min_gap_frames} frames gap ({gap_frames:.1f})', 'text': cur['text']})
            elif gap_frames > min_gap_frames and gap_frames < max_gap_frames:
                errors.append({'idx': cur['idx'], 'start': cur['start'], 'end': cur['end'], 'type': 'Gap', 'message': f'Gap of {gap_frames:.1f} frames (should be {min_gap_frames})', 'text': cur['text']})

        # 11. Hearing impaired formatting (brackets for sound effects)
        if checks.get('hearing_impaired', True):
            sound_effects_list = [s.strip() for s in sound_effects.split(',')]
            for sub in subs:
                text = sub['text']
                for effect in sound_effects_list:
                    if effect.lower() in text.lower() and not re.search(rf'\[{effect}\]', text, re.IGNORECASE):
                        errors.append({
                            'idx': sub['idx'], 'start': sub['start'], 'end': sub['end'],
                            'type': 'Hearing Impaired', 'message': f'Sound effect should be in brackets: [{effect}]', 'text': text
                        })

        # 12. Glyph checks (special characters)
        if checks.get('invalid_chars', True):
            for sub in subs:
                text = sub['text']
                # Check for invalid characters
                invalid_chars = ['\u0000', '\u0001', '\u0002', '\u0003', '\u0004', '\u0005', '\u0006', '\u0007']
                for char in invalid_chars:
                    if char in text:
                        errors.append({
                            'idx': sub['idx'], 'start': sub['start'], 'end': sub['end'],
                            'type': 'Glyph', 'message': f'Invalid control character detected', 'text': text
                        })

        # 13. Frame rate validation
        if checks.get('frame_rate_alignment', True):
            for sub in subs:
                start_ms = ms_from_time(sub['start'])
                end_ms = ms_from_time(sub['end'])
                # Check if timing aligns with frame rate
                frame_duration = 1000.0 / frame_rate
                if start_ms % frame_duration != 0 or end_ms % frame_duration != 0:
                    errors.append({
                        'idx': sub['idx'], 'start': sub['start'], 'end': sub['end'],
                        'type': 'Frame Rate', 'message': f'Timing does not align with {frame_rate} fps frame boundaries', 'text': sub['text']
                    })

        return jsonify({
            'errors': errors, 
            'count': len(errors),
            'total_subtitles': len(subs) if subs else 0
        })
    except Exception as e:
        import traceback
        print(traceback.format_exc())
        return jsonify({'error': str(e)}), 500

@app.route('/api/export-srt-errors-pdf', methods=['POST'])
def export_srt_errors_pdf():
    """Generate a PDF report of SRT errors"""
    try:
        data = request.get_json()
        errors = data.get('errors', [])
        parameters = data.get('parameters', {})
        summary = data.get('summary', {})
        
        if not errors:
            return jsonify({'error': 'No errors to export'}), 400
        
        # Create PDF using reportlab
        from reportlab.lib.pagesizes import letter, A4
        from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib.units import inch
        from reportlab.lib import colors
        from reportlab.lib.enums import TA_CENTER, TA_LEFT
        import io
        
        # Create buffer for PDF
        buffer = io.BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=72, leftMargin=72, topMargin=72, bottomMargin=72)
        
        # Get styles
        styles = getSampleStyleSheet()
        title_style = ParagraphStyle(
            'CustomTitle',
            parent=styles['Heading1'],
            fontSize=16,
            spaceAfter=30,
            alignment=TA_CENTER
        )
        heading_style = ParagraphStyle(
            'CustomHeading',
            parent=styles['Heading2'],
            fontSize=14,
            spaceAfter=12,
            spaceBefore=20
        )
        normal_style = styles['Normal']
        
        # Build PDF content
        story = []
        
        # Title
        story.append(Paragraph("SRT Error Check Report", title_style))
        story.append(Spacer(1, 20))
        
        # Summary section
        story.append(Paragraph("Summary", heading_style))
        summary_data = [
            ['Total Errors', str(summary.get('totalErrors', 0))],
            ['Total Subtitles', str(summary.get('totalSubtitles', 0))],
            ['Error Categories', str(summary.get('errorCategories', 0))],
            ['Check Date', datetime.now().strftime('%Y-%m-%d %H:%M:%S')]
        ]
        
        summary_table = Table(summary_data, colWidths=[2*inch, 3*inch])
        summary_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (0, -1), colors.grey),
            ('TEXTCOLOR', (0, 0), (0, -1), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (1, 0), (1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        story.append(summary_table)
        story.append(Spacer(1, 20))
        
        # Parameters section
        story.append(Paragraph("Check Parameters", heading_style))
        param_data = [
            ['Parameter', 'Value'],
            ['Language', parameters.get('language', 'N/A')],
            ['Frame Rate', str(parameters.get('frameRate', 'N/A'))],
            ['Max CPS', str(parameters.get('maxCps', 'N/A'))],
            ['Max Line Length', str(parameters.get('maxLineLength', 'N/A'))],
            ['Max Lines', str(parameters.get('maxLines', 'N/A'))],
            ['Min Duration (ms)', str(parameters.get('minDuration', 'N/A'))],
            ['Max Duration (ms)', str(parameters.get('maxDuration', 'N/A'))],
            ['Min Gap (frames)', str(parameters.get('minGapFrames', 'N/A'))],
            ['Max Gap (frames)', str(parameters.get('maxGapFrames', 'N/A'))],
            ['Speaker Style', parameters.get('speakerStyle', 'N/A')],
            ['Number Spelling Max', str(parameters.get('numberSpellingMax', 'N/A'))],
            ['Sound Effects', parameters.get('soundEffects', 'N/A')]
        ]
        
        param_table = Table(param_data, colWidths=[2*inch, 3*inch])
        param_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (0, -1), colors.grey),
            ('TEXTCOLOR', (0, 0), (0, -1), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 9),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (1, 0), (1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        story.append(param_table)
        story.append(Spacer(1, 20))
        
        # Error details section
        story.append(Paragraph("Error Details", heading_style))
        
        # Group errors by type
        error_groups = {}
        for error in errors:
            error_type = error.get('type', 'Unknown')
            if error_type not in error_groups:
                error_groups[error_type] = []
            error_groups[error_type].append(error)
        
        # Create error tables for each type
        for error_type, type_errors in error_groups.items():
            story.append(Paragraph(f"{error_type} ({len(type_errors)} errors)", heading_style))
            
            # Create table for this error type
            error_data = [['Subtitle #', 'Time Range', 'Message', 'Text']]
            
            for error in type_errors:
                time_range = f"{error.get('start', 'N/A')} - {error.get('end', 'N/A')}"
                message = error.get('message', 'N/A')
                text = error.get('text', 'N/A')
                
                # Truncate long text for PDF
                if len(text) > 50:
                    text = text[:47] + "..."
                
                error_data.append([
                    str(error.get('idx', 'N/A')),
                    time_range,
                    message,
                    text
                ])
            
            error_table = Table(error_data, colWidths=[0.8*inch, 1.5*inch, 2.5*inch, 2*inch])
            error_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.darkblue),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, -1), 8),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.lightgrey),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.lightgrey])
            ]))
            story.append(error_table)
            story.append(Spacer(1, 12))
        
        # Build PDF
        doc.build(story)
        buffer.seek(0)
        
        return send_file(
            buffer,
            mimetype='application/pdf',
            as_attachment=True,
            download_name=f'srt_error_report_{datetime.now().strftime("%Y%m%d_%H%M%S")}.pdf'
        )
        
    except Exception as e:
        logger.error(f"Error generating PDF: {str(e)}")
        return jsonify({'error': f'Failed to generate PDF: {str(e)}'}), 500

if __name__ == '__main__':
    app.run(debug=True) 