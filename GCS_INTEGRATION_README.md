# GCS-Based Video Recognition System

## Overview

This implementation solves the recognize functionality issues by implementing a background Google Cloud Storage (GCS) upload system that allows for efficient video segment processing without hitting Cloud Run's 32MB request size limits.

## How It Works

### 1. Background Video Upload
When a user loads a local video file:
- The video is immediately displayed locally for instant playback
- In the background, the entire video is uploaded to Google Cloud Storage
- Users see a progress indicator during upload
- The video can be used for recognition once upload completes

### 2. GCS-Based Segment Processing
When the "Recognize" button is clicked:
- The server downloads the specific video segment from GCS based on TCR In/Out times
- FFmpeg extracts the audio segment from the downloaded video
- The audio is sent to the audd.io API for music recognition
- Results are returned to the user

## Key Benefits

1. **No File Size Limits**: Videos of any size can be processed since they're stored in GCS
2. **Instant Playback**: Users can start working immediately while upload happens in background
3. **Efficient Processing**: Only the required video segments are downloaded and processed
4. **Reliable**: GCS provides persistent storage and high availability
5. **Scalable**: Can handle multiple concurrent users and large video files

## Implementation Details

### Backend Changes

#### New API Endpoints

1. **`/api/upload-video-to-gcs`** (POST)
   - Uploads video files to Google Cloud Storage
   - Returns GCS path for future reference
   - Handles file validation and error checking

2. **`/api/recognize-gcs-segment`** (POST)
   - Downloads video segments from GCS
   - Extracts audio using FFmpeg
   - Sends to audd.io API
   - Returns recognition results

#### Key Features
- Unique filename generation using UUID
- Proper content type handling
- Temporary file cleanup
- Comprehensive error handling and logging
- Progress tracking for uploads

### Frontend Changes

#### Video Player Updates
- Background upload functionality
- Progress indicators for upload status
- GCS path tracking for local videos
- Fallback handling for upload failures

#### Recognition Button Updates
- Uses GCS-based recognition for local files
- Maintains existing URL-based recognition for server files
- Progress indicators during processing
- Error handling and user feedback

## Configuration Requirements

### Environment Variables
```bash
GCS_BUCKET_NAME=your-bucket-name
GOOGLE_APPLICATION_CREDENTIALS=path/to/service-account.json
AUDD_API_TOKEN=your-audd-api-token
```

### Google Cloud Storage Setup
1. Create a GCS bucket for video storage
2. Set up proper IAM permissions for your service account
3. Configure CORS if needed for direct browser uploads

### Dependencies
The system requires these Python packages:
- `google-cloud-storage>=2.14.0`
- `ffmpeg-python==0.2.0`
- `requests==2.31.0`

## Usage Flow

### For Local Video Files
1. User selects a video file
2. Video loads immediately in the player
3. Background upload to GCS starts with progress indicator
4. Once upload completes, recognition becomes available
5. User clicks "Recognize" on any row
6. Server processes the specific video segment from GCS
7. Results are displayed in the table

### For URL-Based Videos
1. User enters a video URL
2. Video loads from the URL
3. Recognition works as before (no GCS upload needed)
4. Server downloads the video and processes segments

## Error Handling

### Upload Failures
- Users can still work with the video locally
- Warning message indicates recognition may not work
- Upload can be retried by reloading the video

### Recognition Failures
- Detailed error messages from audd.io API
- Graceful fallback for network issues
- Logging for debugging purposes

### GCS Issues
- Connection timeout handling
- Bucket access permission checks
- File existence validation

## Performance Considerations

### Upload Performance
- Large files are uploaded in chunks
- Progress tracking provides user feedback
- Upload happens asynchronously

### Processing Performance
- Only required video segments are downloaded
- FFmpeg optimizations for audio extraction
- Temporary file cleanup after processing

### Storage Management
- Files are stored with unique names to prevent conflicts
- Consider implementing cleanup policies for old files
- Monitor storage costs and usage

## Testing

Use the provided test script to verify the integration:

```bash
python test_gcs_integration.py
```

This script tests:
1. GCS connection and bucket access
2. Video upload functionality
3. Recognition endpoint functionality

## Troubleshooting

### Common Issues

1. **GCS Connection Failed**
   - Check service account credentials
   - Verify bucket name and permissions
   - Ensure Google Cloud SDK is configured

2. **Upload Timeout**
   - Check network connectivity
   - Verify file size and upload speed
   - Consider implementing chunked uploads for very large files

3. **Recognition Errors**
   - Check audd.io API token
   - Verify video format compatibility
   - Check FFmpeg installation and configuration

### Debugging
- Check application logs for detailed error messages
- Use the test script to isolate issues
- Monitor GCS bucket for uploaded files
- Verify temporary file cleanup

## Future Enhancements

1. **Chunked Uploads**: Implement multipart uploads for very large files
2. **Video Compression**: Add optional video compression before upload
3. **Caching**: Implement recognition result caching
4. **Batch Processing**: Support for processing multiple segments simultaneously
5. **Cleanup Policies**: Automatic cleanup of old video files from GCS

## Security Considerations

1. **File Validation**: All uploaded files are validated for type and size
2. **Access Control**: GCS bucket should have proper access controls
3. **Temporary Files**: All temporary files are cleaned up after processing
4. **API Security**: audd.io API token should be kept secure

## Cost Optimization

1. **Storage Lifecycle**: Implement policies to delete old videos
2. **Processing Optimization**: Only download required segments
3. **Caching**: Cache recognition results to avoid reprocessing
4. **Monitoring**: Track usage and costs in Google Cloud Console 