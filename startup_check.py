#!/usr/bin/env python3
"""
Startup check script for Cloud Run debugging
"""
import os
import sys
import logging
from datetime import datetime

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def check_environment():
    """Check environment variables and system state"""
    logger.info("=== Environment Check ===")
    
    # Check essential environment variables
    essential_vars = ['PORT', 'GOOGLE_CLOUD_PROJECT']
    for var in essential_vars:
        value = os.environ.get(var)
        logger.info(f"{var}: {value}")
    
    # Check optional environment variables
    optional_vars = ['GCS_BUCKET_NAME', 'FLASK_SECRET_KEY']
    for var in optional_vars:
        value = os.environ.get(var)
        logger.info(f"{var}: {'SET' if value else 'NOT SET'}")
    
    # Check system info
    logger.info(f"Python version: {sys.version}")
    logger.info(f"Current working directory: {os.getcwd()}")
    logger.info(f"Files in current directory: {os.listdir('.')}")

def check_dependencies():
    """Check if required dependencies are available"""
    logger.info("=== Dependency Check ===")
    
    try:
        import flask
        logger.info(f"Flask version: {flask.__version__}")
    except ImportError as e:
        logger.error(f"Flask not available: {e}")
    
    try:
        import google.cloud.storage
        logger.info("Google Cloud Storage available")
    except ImportError as e:
        logger.error(f"Google Cloud Storage not available: {e}")
    
    try:
        import google.cloud.firestore
        logger.info("Google Cloud Firestore available")
    except ImportError as e:
        logger.error(f"Google Cloud Firestore not available: {e}")
    
    try:
        import vlc
        logger.info("VLC available")
    except ImportError as e:
        logger.warning(f"VLC not available (this is normal in Cloud Run): {e}")

def main():
    """Main startup check function"""
    logger.info(f"=== Startup Check Started at {datetime.now()} ===")
    
    check_environment()
    check_dependencies()
    
    logger.info("=== Startup Check Completed ===")

if __name__ == "__main__":
    main() 