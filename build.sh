#!/usr/bin/env bash
# exit on error
set -o errexit

# Print debug information
echo "Starting build process..."
echo "Current directory: $(pwd)"
echo "Listing files:"
ls -la

# Install Python dependencies
echo "Upgrading pip..."
python -m pip install --upgrade pip
echo "Installing requirements..."
pip install -r requirements.txt

# Create necessary directories
echo "Creating uploads directory..."
mkdir -p uploads

echo "Build completed successfully!" 