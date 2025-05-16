#!/usr/bin/env bash
# exit on error
set -o errexit

# Print debug information
echo "Starting build process..."
echo "Current directory: $(pwd)"
echo "Listing files:"
ls -la

# Install system dependencies
echo "Installing system dependencies..."
apt-get update
apt-get install -y ffmpeg

# Verify FFmpeg installation
echo "Verifying FFmpeg installation..."
ffmpeg -version

# Install Python dependencies
echo "Upgrading pip..."
python -m pip install --upgrade pip
echo "Installing requirements..."
pip install -r requirements.txt

# Create necessary directories
echo "Creating uploads directory..."
mkdir -p uploads

echo "Build completed successfully!" 