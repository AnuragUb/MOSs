from flask import Flask, render_template, request, jsonify, send_file
from flask_cors import CORS
import json
import os
from datetime import datetime
import requests
import tempfile
import subprocess
from werkzeug.utils import secure_filename
import pandas as pd
import ffmpeg
import logging
import io

app = Flask(__name__)
CORS(app)

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

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

AUDD_API_TOKEN = '04e48a84490a8a2f0bf327b274404905'  # Replace with your actual API token

@app.route('/')
def index():
    return render_template('splash.html')

@app.route('/main')
def main():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_video():
    if 'video' not in request.files:
        return jsonify({'error': 'No video file provided'}), 400
    
    video_file = request.files['video']
    if video_file.filename == '':
        return jsonify({'error': 'No selected file'}), 400
    
    if video_file:
        filename = f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{video_file.filename}"
        filepath = os.path.join(UPLOADS_DIR, filename)
        video_file.save(filepath)
        return jsonify({'filename': filename})

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
    markers = request.json
    
    if format == 'excel':
        # Handle Excel export
        return jsonify({'status': 'success', 'message': 'Excel export handled by frontend'})
    elif format == 'csv':
        # Handle CSV export
        return jsonify({'status': 'success', 'message': 'CSV export handled by frontend'})
    else:
        return jsonify({'error': 'Invalid export format'}), 400

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
            with tempfile.NamedTemporaryFile(delete=False, suffix='.mp4') as temp_in:
                file.save(temp_in)
                input_path = temp_in.name
                logger.info(f"Saved uploaded file to: {input_path}")
        elif video_src:
            # Download the video file
            try:
                response = requests.get(video_src, stream=True)
                response.raise_for_status()
                with tempfile.NamedTemporaryFile(delete=False, suffix='.mp4') as temp_in:
                    for chunk in response.iter_content(chunk_size=8192):
                        temp_in.write(chunk)
                    input_path = temp_in.name
                    logger.info(f"Downloaded video to: {input_path}")
            except requests.RequestException as e:
                logger.error(f"Error downloading video: {str(e)}")
                return jsonify({'status': 'error', 'message': f'Error downloading video: {str(e)}'}), 500
        else:
            return jsonify({'status': 'error', 'message': 'No file or videoSrc provided'}), 400

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
                '-ab', '192k',  # Audio bitrate
                '-ac', '1',  # Mono audio
                '-ar', '44100',  # Sample rate
                '-f', 'mp3',  # Force MP3 format
                output_path  # Output file
            ]
            
            logger.info(f"Running FFmpeg command: {' '.join(ffmpeg_cmd)}")
            
            # Run FFmpeg
            result = subprocess.run(ffmpeg_cmd, capture_output=True, text=True)
            
            if result.stderr:
                logger.warning(f"FFmpeg stderr output: {result.stderr}")
            
            if result.returncode != 0:
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
                os.remove(input_path)
            logger.info("Cleaned up temporary files")
        except Exception as e:
            logger.error(f"Error cleaning up files: {str(e)}")

        return jsonify(result)

    except Exception as e:
        logger.error(f"Unexpected error in recognize_audio: {str(e)}")
        return jsonify({'status': 'error', 'message': f'Unexpected error: {str(e)}'}), 500

@app.route('/api/parse-cue-sheet', methods=['POST'])
def parse_cue_sheet():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
    file = request.files['file']
    filename = secure_filename(file.filename)
    ext = filename.split('.')[-1].lower()
    # Read file into pandas DataFrame
    if ext in ['xlsx', 'xls']:
        df = pd.read_excel(file, header=None)
    elif ext == 'csv':
        df = pd.read_csv(file, header=None)
    else:
        return jsonify({'error': 'Unsupported file type'}), 400
    # Convert to list of lists
    rows = df.values.tolist()
    # Extract metadata, header, and data
    metadata = [[str(cell) if cell is not None else '' for cell in row] for row in rows[:6]]
    header = [str(cell) if cell is not None else '' for cell in rows[6]] if len(rows) > 6 else []
    data_rows = rows[7:] if len(rows) > 7 else []
    # Convert data rows to list of dicts, all values as strings
    data = [dict(zip(header, [str(cell) if cell is not None else '' for cell in row])) for row in data_rows]
    return jsonify({
        'metadata': metadata,
        'header': header,
        'data': data
    })

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

if __name__ == '__main__':
    app.run(debug=True) 