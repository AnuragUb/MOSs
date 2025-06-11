import os
from google.cloud import storage, firestore
from google.cloud import logging as cloud_logging

# Google Cloud Configuration
PROJECT_ID = os.getenv('GOOGLE_CLOUD_PROJECT', 'your-project-id')
BUCKET_NAME = os.getenv('GCS_BUCKET_NAME', 'your-bucket-name')

# Free Tier Limits
STORAGE_FREE_TIER = 5 * 1024 * 1024 * 1024  # 5GB
FIRESTORE_FREE_TIER = {
    'storage': 1 * 1024 * 1024 * 1024,  # 1GB
    'reads': 50000,  # per day
    'writes': 20000,  # per day
    'deletes': 20000  # per day
}
CLOUD_RUN_FREE_TIER = {
    'requests': 2000000,  # 2M requests per month
    'memory': 360000,  # GB-seconds
    'cpu': 180000,  # vCPU-seconds
    'egress': 1 * 1024 * 1024 * 1024  # 1GB egress from North America
}

# Initialize Google Cloud clients
storage_client = storage.Client()
firestore_client = firestore.Client()
logging_client = cloud_logging.Client()

# Get bucket
bucket = storage_client.bucket(BUCKET_NAME)

# Firestore collection names
MARKERS_COLLECTION = 'markers'
USAGE_STATS_COLLECTION = 'usage_stats'

# Temporary file storage
TEMP_DIR = '/tmp'
UPLOADS_DIR = os.path.join(TEMP_DIR, 'uploads')

# Ensure uploads directory exists
os.makedirs(UPLOADS_DIR, exist_ok=True)

# Cost optimization settings
MAX_FILE_SIZE = 100 * 1024 * 1024  # 100MB max file size
VIDEO_COMPRESSION = True  # Enable video compression
CACHE_DURATION = 3600  # 1 hour cache duration 