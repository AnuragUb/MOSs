#!/usr/bin/env python3
"""
Test script for GCS integration
This script tests the new GCS-based video upload and recognition functionality
"""

import os
import tempfile
import requests
import json
from google.cloud import storage

def test_gcs_connection():
    """Test if GCS connection works"""
    try:
        storage_client = storage.Client()
        bucket_name = os.getenv('GCS_BUCKET_NAME', 'mos-aat')
        bucket = storage_client.bucket(bucket_name)
        
        if bucket.exists():
            print(f"✅ GCS connection successful. Bucket '{bucket_name}' exists.")
            return True
        else:
            print(f"❌ GCS bucket '{bucket_name}' does not exist.")
            return False
    except Exception as e:
        print(f"❌ GCS connection failed: {str(e)}")
        return False

def test_upload_endpoint():
    """Test the upload endpoint"""
    try:
        # Create a small test file
        with tempfile.NamedTemporaryFile(delete=False, suffix='.mp4') as temp_file:
            temp_file.write(b'fake video content')
            temp_path = temp_file.name
        
        try:
            # Test upload
            with open(temp_path, 'rb') as f:
                files = {'video': ('test.mp4', f, 'video/mp4')}
                response = requests.post('http://localhost:5000/api/upload-video-to-gcs', files=files)
            
            if response.status_code == 200:
                result = response.json()
                if result.get('status') == 'success':
                    print(f"✅ Upload endpoint working. GCS path: {result.get('gcs_path')}")
                    return result.get('gcs_path')
                else:
                    print(f"❌ Upload failed: {result.get('error')}")
                    return None
            else:
                print(f"❌ Upload endpoint returned status {response.status_code}")
                return None
                
        finally:
            os.unlink(temp_path)
            
    except Exception as e:
        print(f"❌ Upload test failed: {str(e)}")
        return None

def test_recognize_endpoint(gcs_path):
    """Test the recognize endpoint"""
    if not gcs_path:
        print("❌ No GCS path provided for recognition test")
        return False
    
    try:
        data = {
            'tcrIn': '00:00:00:00',
            'tcrOut': '00:00:05:00',
            'gcsPath': gcs_path
        }
        
        response = requests.post('http://localhost:5000/api/recognize-gcs-segment', data=data)
        
        if response.status_code == 200:
            result = response.json()
            print(f"✅ Recognition endpoint working. Response: {result.get('status', 'unknown')}")
            return True
        else:
            print(f"❌ Recognition endpoint returned status {response.status_code}")
            print(f"Response: {response.text}")
            return False
            
    except Exception as e:
        print(f"❌ Recognition test failed: {str(e)}")
        return False

def main():
    """Run all tests"""
    print("🧪 Testing GCS Integration...")
    print("=" * 50)
    
    # Test 1: GCS Connection
    print("\n1. Testing GCS Connection...")
    gcs_ok = test_gcs_connection()
    
    # Test 2: Upload Endpoint
    print("\n2. Testing Upload Endpoint...")
    gcs_path = test_upload_endpoint()
    
    # Test 3: Recognition Endpoint
    print("\n3. Testing Recognition Endpoint...")
    if gcs_path:
        recognize_ok = test_recognize_endpoint(gcs_path)
    else:
        recognize_ok = False
    
    # Summary
    print("\n" + "=" * 50)
    print("📊 Test Summary:")
    print(f"GCS Connection: {'✅ PASS' if gcs_ok else '❌ FAIL'}")
    print(f"Upload Endpoint: {'✅ PASS' if gcs_path else '❌ FAIL'}")
    print(f"Recognition Endpoint: {'✅ PASS' if recognize_ok else '❌ FAIL'}")
    
    if gcs_ok and gcs_path and recognize_ok:
        print("\n🎉 All tests passed! GCS integration is working correctly.")
    else:
        print("\n⚠️  Some tests failed. Please check the configuration.")

if __name__ == "__main__":
    main() 