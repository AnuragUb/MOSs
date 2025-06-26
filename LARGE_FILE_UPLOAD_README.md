# Large File Upload & Cloud Video Management

## Overview

This implementation adds advanced upload capabilities for large files (1GB+) and comprehensive cloud video management to the MOS application.

## Features

### 1. Parallel Chunked Uploads for Large Files

**For files over 1GB:**
- Automatic chunking into 5MB pieces
- Parallel upload with up to 4 concurrent chunks
- Progress tracking per chunk
- Resume capability for failed chunks
- Optimized for very large video files

**Upload Methods by File Size:**
- **< 32MB**: Server proxy upload
- **32MB - 1GB**: Direct GCS upload with signed URLs
- **> 1GB**: Parallel chunked upload with Transfer Service

### 2. Cloud Video Management

**Video Listing:**
- List all videos in MOS-AAT cloud storage
- Display file size, creation date, and metadata
- Generate signed URLs for secure access
- Sort by creation date (newest first)

**Video Operations:**
- Load videos directly from cloud storage
- Delete videos from cloud storage
- Refresh video list
- Secure access via signed URLs

### 3. Bulk Upload System

**Features:**
- Upload multiple files simultaneously
- Automatic method selection based on file size
- Progress tracking for each file
- Concurrent uploads (up to 3 files at once)
- Comprehensive error handling

## API Endpoints

### Chunked Upload Endpoints

#### `POST /api/generate-chunked-upload-urls`
Generate signed URLs for parallel chunked uploads.

**Request:**
```json
{
    "filename": "large_video.mp4",
    "contentType": "video/mp4",
    "fileSize": 2147483648,
    "chunkSize": 5242880
}
```

**Response:**
```json
{
    "status": "success",
    "chunkUrls": [
        {
            "chunkIndex": 0,
            "startByte": 0,
            "endByte": 5242879,
            "signedUrl": "https://..."
        }
    ],
    "gcsPath": "videos/uuid.mp4",
    "filename": "uuid.mp4",
    "totalChunks": 410,
    "chunkSize": 5242880
}
```

### Cloud Video Management Endpoints

#### `GET /api/list-cloud-videos`
List all videos in cloud storage.

**Response:**
```json
{
    "status": "success",
    "videos": [
        {
            "name": "videos/uuid.mp4",
            "filename": "uuid.mp4",
            "size": 2147483648,
            "size_mb": 2048.0,
            "created": "2024-01-01T12:00:00Z",
            "updated": "2024-01-01T12:00:00Z",
            "signed_url": "https://...",
            "content_type": "video/mp4"
        }
    ],
    "total_count": 1
}
```

#### `POST /api/load-cloud-video`
Load a specific video from cloud storage.

**Request:**
```json
{
    "gcsPath": "videos/uuid.mp4"
}
```

**Response:**
```json
{
    "status": "success",
    "gcsPath": "videos/uuid.mp4",
    "signedUrl": "https://...",
    "filename": "uuid.mp4",
    "size": 2147483648,
    "contentType": "video/mp4"
}
```

#### `DELETE /api/delete-cloud-video`
Delete a video from cloud storage.

**Request:**
```json
{
    "gcsPath": "videos/uuid.mp4"
}
```

**Response:**
```json
{
    "status": "success",
    "message": "Video videos/uuid.mp4 deleted successfully"
}
```

## Frontend Implementation

### JavaScript Functions

#### Chunked Upload Functions
- `uploadVideoChunked(file, progressDiv)`: Handle chunked uploads for large files
- `uploadChunksParallel(file, chunkUrls, gcsPath, progressDiv)`: Upload chunks in parallel
- `uploadVideoDirect(file, progressDiv)`: Handle direct uploads for medium files

#### Cloud Video Management Functions
- `loadCloudVideos()`: Fetch and display cloud videos
- `displayCloudVideos()`: Render cloud video list
- `loadCloudVideo(gcsPath)`: Load specific video from cloud
- `deleteCloudVideo(gcsPath)`: Delete video from cloud

#### Bulk Upload Functions
- `handleBulkUpload(event)`: Process multiple file uploads
- `uploadFileProxy(file, progressElement)`: Upload small files via proxy
- `uploadFileDirect(file, progressElement)`: Upload medium files directly
- `uploadFileChunked(file, progressElement)`: Upload large files in chunks

### UI Components

#### Cloud Video Management UI
```html
<div class="cloud-videos-section">
    <div class="d-flex justify-content-between align-items-center">
        <h6>Cloud Storage Videos</h6>
        <button onclick="loadCloudVideos()">Refresh</button>
    </div>
    <div id="cloudVideosContainer" class="cloud-videos-list">
        <!-- Video items -->
    </div>
</div>
```

#### Bulk Upload UI
```html
<button id="bulkUploadBtn">Bulk Upload</button>
<input type="file" id="bulkFileInput" multiple style="display: none;">
```

## Configuration

### Environment Variables
```bash
GCS_BUCKET_NAME=mos-aat
GOOGLE_CLOUD_PROJECT=your-project-id
GOOGLE_APPLICATION_CREDENTIALS=path/to/service-account.json
```

### Dependencies
```txt
google-cloud-storage==2.14.0
google-cloud-storage-transfer==1.8.0
```

## Usage Examples

### Uploading Large Files
1. Select a video file over 1GB
2. System automatically detects file size
3. Chunks file into 5MB pieces
4. Uploads chunks in parallel (4 at a time)
5. Shows progress for each chunk
6. Completes upload when all chunks are done

### Managing Cloud Videos
1. Click "Refresh" to load cloud video list
2. View all videos with metadata
3. Click "Load Video" to load into player
4. Click "Delete" to remove from cloud storage

### Bulk Upload
1. Click "Bulk Upload" button
2. Select multiple video files
3. System processes each file with appropriate method
4. Shows overall progress and individual file status
5. Refreshes cloud video list when complete

## Performance Considerations

### Upload Performance
- **Chunked uploads**: 4 concurrent chunks for large files
- **Bulk uploads**: 3 concurrent files
- **Progress tracking**: Real-time updates for all operations
- **Error handling**: Graceful failure with retry options

### Storage Management
- **Unique filenames**: UUID-based naming prevents conflicts
- **Signed URLs**: Secure access with expiration
- **Cleanup**: Manual deletion capability
- **Metadata**: Full file information tracking

### Network Optimization
- **Chunk size**: 5MB optimal for most connections
- **Concurrency limits**: Prevents overwhelming connections
- **Resume capability**: Failed chunks can be retried
- **Progress feedback**: User always knows upload status

## Error Handling

### Upload Failures
- Individual chunk failures don't stop entire upload
- Failed chunks can be retried
- Clear error messages for debugging
- Graceful degradation for network issues

### Cloud Storage Issues
- Connection timeout handling
- Permission verification
- File existence checks
- Signed URL generation fallbacks

## Security Features

### Access Control
- Signed URLs with expiration
- Service account authentication
- Bucket-level permissions
- Secure file operations

### Data Protection
- Unique file naming
- Temporary URL generation
- Secure deletion
- Audit logging

## Monitoring and Logging

### Upload Monitoring
- Progress tracking for all operations
- Chunk-level status monitoring
- File-level completion tracking
- Error logging with context

### Cloud Storage Monitoring
- File listing operations
- Access pattern tracking
- Storage usage monitoring
- Performance metrics

## Future Enhancements

### Planned Features
1. **Resumable uploads**: Resume interrupted uploads
2. **Video compression**: Automatic compression before upload
3. **Batch operations**: Bulk delete, move operations
4. **Storage policies**: Automatic cleanup based on age/size
5. **Transfer acceleration**: CDN integration for faster uploads

### Performance Improvements
1. **Dynamic chunk sizing**: Adjust based on connection speed
2. **Adaptive concurrency**: Adjust based on network conditions
3. **Caching**: Cache frequently accessed videos
4. **Compression**: Reduce storage costs and upload times

## Troubleshooting

### Common Issues

1. **Chunked upload fails**
   - Check network stability
   - Verify file size calculation
   - Check GCS permissions

2. **Cloud video not loading**
   - Verify signed URL expiration
   - Check bucket permissions
   - Verify file exists in storage

3. **Bulk upload stuck**
   - Check concurrent upload limits
   - Verify individual file sizes
   - Check network bandwidth

### Debug Information
- Console logging for all operations
- Progress indicators for user feedback
- Error messages with context
- Network status monitoring

## Cost Optimization

### Storage Costs
- Monitor storage usage
- Implement cleanup policies
- Use appropriate storage classes
- Track access patterns

### Transfer Costs
- Optimize chunk sizes
- Limit concurrent operations
- Use appropriate regions
- Monitor bandwidth usage

This implementation provides a robust, scalable solution for handling large video files and managing cloud storage efficiently. 