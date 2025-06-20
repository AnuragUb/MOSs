#!/usr/bin/env python3
"""
Simple verification script to check if the GCS integration setup is correct
This script checks the code structure without requiring external dependencies
"""

import os
import re

def check_file_exists(filepath):
    """Check if a file exists"""
    if os.path.exists(filepath):
        print(f"✅ {filepath} - Found")
        return True
    else:
        print(f"❌ {filepath} - Missing")
        return False

def check_endpoint_in_file(filepath, endpoint):
    """Check if an endpoint exists in a file"""
    if not os.path.exists(filepath):
        print(f"❌ {filepath} - File not found")
        return False
    
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            content = f.read()
            if endpoint in content:
                print(f"✅ {endpoint} - Found in {filepath}")
                return True
            else:
                print(f"❌ {endpoint} - Not found in {filepath}")
                return False
    except Exception as e:
        print(f"❌ Error reading {filepath}: {str(e)}")
        return False

def check_function_in_file(filepath, function_name):
    """Check if a function exists in a file"""
    if not os.path.exists(filepath):
        print(f"❌ {filepath} - File not found")
        return False
    
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            content = f.read()
            if f"function {function_name}" in content:
                print(f"✅ {function_name}() - Found in {filepath}")
                return True
            else:
                print(f"❌ {function_name}() - Not found in {filepath}")
                return False
    except Exception as e:
        print(f"❌ Error reading {filepath}: {str(e)}")
        return False

def check_css_class_in_file(filepath, css_class):
    """Check if a CSS class exists in a file"""
    if not os.path.exists(filepath):
        print(f"❌ {filepath} - File not found")
        return False
    
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            content = f.read()
            if css_class in content:
                print(f"✅ .{css_class} - Found in {filepath}")
                return True
            else:
                print(f"❌ .{css_class} - Not found in {filepath}")
                return False
    except Exception as e:
        print(f"❌ Error reading {filepath}: {str(e)}")
        return False

def main():
    """Run all verification checks"""
    print("🔍 Verifying GCS Integration Setup...")
    print("=" * 60)
    
    checks_passed = 0
    total_checks = 0
    
    # Check if required files exist
    print("\n📁 File Structure Checks:")
    files_to_check = [
        'app.py',
        'static/js/main.js',
        'static/css/style.css',
        'requirements.txt',
        'test_gcs_integration.py',
        'GCS_INTEGRATION_README.md'
    ]
    
    for file in files_to_check:
        total_checks += 1
        if check_file_exists(file):
            checks_passed += 1
    
    # Check backend endpoints
    print("\n🔧 Backend API Endpoints:")
    total_checks += 2
    if check_endpoint_in_file('app.py', '/api/upload-video-to-gcs'):
        checks_passed += 1
    if check_endpoint_in_file('app.py', '/api/recognize-gcs-segment'):
        checks_passed += 1
    
    # Check if original recognize endpoint still exists
    print("\n🔄 Original Functionality Preservation:")
    total_checks += 1
    if check_endpoint_in_file('app.py', '/api/recognize-audio'):
        checks_passed += 1
        print("✅ Original /api/recognize-audio endpoint preserved")
    
    # Check frontend functions
    print("\n🎨 Frontend JavaScript Functions:")
    total_checks += 1
    if check_function_in_file('static/js/main.js', 'uploadVideoToGCS'):
        checks_passed += 1
    
    # Check CSS styles
    print("\n💅 CSS Styles:")
    total_checks += 1
    if check_css_class_in_file('static/css/style.css', 'upload-progress'):
        checks_passed += 1
    
    # Check requirements
    print("\n📦 Dependencies:")
    total_checks += 1
    if check_endpoint_in_file('requirements.txt', 'google-cloud-storage'):
        checks_passed += 1
    
    # Summary
    print("\n" + "=" * 60)
    print("📊 Verification Summary:")
    print(f"Checks Passed: {checks_passed}/{total_checks}")
    print(f"Success Rate: {(checks_passed/total_checks)*100:.1f}%")
    
    if checks_passed == total_checks:
        print("\n🎉 All checks passed! GCS integration setup is complete.")
        print("\n📋 Next Steps:")
        print("1. Deploy to Cloud Run")
        print("2. Test with real video files")
        print("3. Monitor logs for any issues")
        print("4. Run the full test script when Google Cloud is configured")
    else:
        print(f"\n⚠️  {total_checks - checks_passed} check(s) failed.")
        print("Please review the failed checks above and fix any issues.")

if __name__ == "__main__":
    main() 