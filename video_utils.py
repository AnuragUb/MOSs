import os
import subprocess
import logging
from google.cloud import storage
import tempfile

def convert_wmv_to_mp4(input_path, output_path):
    """
    Convert WMV file to MP4 using FFmpeg with optimized settings
    """
    try:
        # FFmpeg command with optimized settings for WMV to MP4 conversion
        cmd = [
            'ffmpeg',
            '-y',  # Overwrite output file if it exists
            '-i', input_path,  # Input file
            '-c:v', 'libx264',  # Video codec
            '-preset', 'medium',  # Encoding preset (balance between speed and quality)
            '-crf', '23',  # Constant Rate Factor (18-28 is good, lower is better quality)
            '-c:a', 'aac',  # Audio codec
            '-b:a', '128k',  # Audio bitrate
            '-movflags', '+faststart',  # Enable fast start for web playback
            output_path
        ]
        
        # Run FFmpeg with progress tracking
        process = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            universal_newlines=True
        )
        
        # Wait for the process to complete
        stdout, stderr = process.communicate()
        
        if process.returncode != 0:
            raise Exception(f"FFmpeg conversion failed: {stderr}")
        
        return True
    except Exception as e:
        logging.error(f"Error in convert_wmv_to_mp4: {str(e)}")
        raise

def upload_to_gcs(file_path, filename, content_type):
    """
    Upload a file to Google Cloud Storage and return a signed URL
    """
    try:
        storage_client = storage.Client()
        bucket = storage_client.bucket(os.getenv('GCS_BUCKET_NAME'))
        blob = bucket.blob(f"videos/{filename}")
        
        # Upload with content type
        blob.upload_from_filename(
            file_path,
            content_type=content_type
        )
        
        # Generate signed URL for immediate access
        url = blob.generate_signed_url(
            version="v4",
            expiration=3600,  # 1 hour
            method="GET"
        )
        
        return url
    except Exception as e:
        logging.error(f"Error in upload_to_gcs: {str(e)}")
        raise 