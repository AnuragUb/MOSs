#!/usr/bin/env python3
"""
GCS Credentials Setup Script
This script helps you set up Google Cloud Storage credentials for the MOS application.
"""

import os
import json
import sys
from pathlib import Path

def check_environment():
    """Check current environment setup"""
    print("🔍 Checking current GCS environment setup...")
    
    # Check for GOOGLE_APPLICATION_CREDENTIALS
    creds_path = os.getenv('GOOGLE_APPLICATION_CREDENTIALS')
    if creds_path:
        if os.path.exists(creds_path):
            print(f"✅ GOOGLE_APPLICATION_CREDENTIALS is set to: {creds_path}")
            return True
        else:
            print(f"❌ GOOGLE_APPLICATION_CREDENTIALS points to non-existent file: {creds_path}")
    else:
        print("❌ GOOGLE_APPLICATION_CREDENTIALS environment variable is not set")
    
    # Check for GOOGLE_CLOUD_PROJECT
    project_id = os.getenv('GOOGLE_CLOUD_PROJECT')
    if project_id:
        print(f"✅ GOOGLE_CLOUD_PROJECT is set to: {project_id}")
    else:
        print("❌ GOOGLE_CLOUD_PROJECT environment variable is not set")
    
    # Check for GCS_BUCKET_NAME
    bucket_name = os.getenv('GCS_BUCKET_NAME')
    if bucket_name:
        print(f"✅ GCS_BUCKET_NAME is set to: {bucket_name}")
    else:
        print("❌ GCS_BUCKET_NAME environment variable is not set")
    
    return False

def create_service_account_instructions():
    """Provide instructions for creating a service account"""
    print("\n📋 Instructions for creating a Google Cloud Service Account:")
    print("1. Go to Google Cloud Console: https://console.cloud.google.com/")
    print("2. Navigate to IAM & Admin → Service Accounts")
    print("3. Click 'Create Service Account'")
    print("4. Give it a name (e.g., 'mos-gcs-upload')")
    print("5. Add these roles:")
    print("   - Storage Object Admin (for uploading files)")
    print("   - Storage Object Viewer (for reading files)")
    print("   - Secret Manager Secret Accessor (if using Secret Manager)")
    print("6. Create and download the JSON key file")
    print("7. Save the JSON file in a secure location")

def setup_credentials_file():
    """Help user set up credentials file"""
    print("\n🔧 Setting up service account credentials...")
    
    # Ask for the credentials file path
    while True:
        creds_path = input("Enter the path to your service account JSON key file: ").strip()
        
        if not creds_path:
            print("❌ Please provide a valid path")
            continue
            
        # Remove quotes if present
        creds_path = creds_path.strip('"\'')
        
        if not os.path.exists(creds_path):
            print(f"❌ File not found: {creds_path}")
            continue
            
        # Validate JSON format
        try:
            with open(creds_path, 'r') as f:
                json.load(f)
            print(f"✅ Valid JSON credentials file found: {creds_path}")
            break
        except json.JSONDecodeError:
            print("❌ Invalid JSON format in credentials file")
            continue
        except Exception as e:
            print(f"❌ Error reading file: {e}")
            continue
    
    # Set environment variable
    os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = creds_path
    print(f"✅ Set GOOGLE_APPLICATION_CREDENTIALS to: {creds_path}")
    
    return creds_path

def setup_environment_variables():
    """Help user set up other required environment variables"""
    print("\n🔧 Setting up other environment variables...")
    
    # Project ID
    project_id = os.getenv('GOOGLE_CLOUD_PROJECT')
    if not project_id:
        project_id = input("Enter your Google Cloud Project ID: ").strip()
        if project_id:
            os.environ['GOOGLE_CLOUD_PROJECT'] = project_id
            print(f"✅ Set GOOGLE_CLOUD_PROJECT to: {project_id}")
    
    # Bucket name
    bucket_name = os.getenv('GCS_BUCKET_NAME')
    if not bucket_name:
        bucket_name = input("Enter your GCS bucket name: ").strip()
        if bucket_name:
            os.environ['GCS_BUCKET_NAME'] = bucket_name
            print(f"✅ Set GCS_BUCKET_NAME to: {bucket_name}")

def test_gcs_connection():
    """Test GCS connection"""
    print("\n🧪 Testing GCS connection...")
    
    try:
        from google.cloud import storage
        
        # Initialize client
        client = storage.Client()
        
        # Test bucket access
        bucket_name = os.getenv('GCS_BUCKET_NAME')
        if bucket_name:
            bucket = client.bucket(bucket_name)
            if bucket.exists():
                print(f"✅ Successfully connected to GCS bucket: {bucket_name}")
                return True
            else:
                print(f"❌ Bucket '{bucket_name}' does not exist or is not accessible")
                return False
        else:
            print("❌ GCS_BUCKET_NAME not set")
            return False
            
    except Exception as e:
        print(f"❌ GCS connection failed: {e}")
        return False

def create_env_file():
    """Create a .env file with the current environment variables"""
    print("\n📝 Creating .env file...")
    
    env_content = []
    
    creds_path = os.getenv('GOOGLE_APPLICATION_CREDENTIALS')
    if creds_path:
        env_content.append(f"GOOGLE_APPLICATION_CREDENTIALS={creds_path}")
    
    project_id = os.getenv('GOOGLE_CLOUD_PROJECT')
    if project_id:
        env_content.append(f"GOOGLE_CLOUD_PROJECT={project_id}")
    
    bucket_name = os.getenv('GCS_BUCKET_NAME')
    if bucket_name:
        env_content.append(f"GCS_BUCKET_NAME={bucket_name}")
    
    if env_content:
        with open('.env', 'w') as f:
            f.write('\n'.join(env_content))
        print("✅ Created .env file with current environment variables")
        print("💡 Remember to add .env to your .gitignore file for security")
    else:
        print("❌ No environment variables to save")

def main():
    """Main setup function"""
    print("🚀 GCS Credentials Setup for MOS Application")
    print("=" * 50)
    
    # Check current setup
    has_creds = check_environment()
    
    if not has_creds:
        print("\n❌ GCS credentials not properly configured")
        
        # Provide instructions
        create_service_account_instructions()
        
        # Setup credentials
        setup_credentials_file()
    
    # Setup other environment variables
    setup_environment_variables()
    
    # Test connection
    if test_gcs_connection():
        print("\n🎉 GCS setup completed successfully!")
        
        # Create .env file
        create_env_file()
        
        print("\n📋 Next steps:")
        print("1. Restart your application")
        print("2. Try uploading a video file")
        print("3. Check the logs for any remaining issues")
        
    else:
        print("\n❌ GCS setup failed. Please check your configuration and try again.")
        sys.exit(1)

if __name__ == "__main__":
    main() 