# Cloud Video Management Feature

## Overview

The Cloud Video Management feature allows users to view, load, and manage videos stored in the MOS-AAT Google Cloud Storage bucket. This provides a centralized way to access all uploaded videos without needing to re-upload them.

## Features

### 1. Cloud Video Listing
- **Automatic Loading**: Videos are automatically loaded when the page loads
- **Real-time Refresh**: Click "Refresh" to update the video list
- **Sorted Display**: Videos are sorted by creation date (newest first)
- **File Information**: Shows filename, size, creation date, and content type

### 2. Video Operations
- **Load Video**: Click "Load" to load a video directly into the player
- **Delete Video**: Click the trash icon to permanently delete from cloud storage
- **Secure Access**: Videos are accessed via signed URLs for security

### 3. User Interface
- **Collapsible Section**: Located in the video controls area
- **Responsive Design**: Adapts to different screen sizes
- **Loading States**: Shows spinner during operations
- **Error Handling**: Displays clear error messages

## API Endpoints

### `GET /api/list-cloud-videos`
Lists all videos in the cloud storage videos folder.

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

### `POST /api/load-cloud-video`
Loads a specific video from cloud storage.

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

### `DELETE /api/delete-cloud-video`
Deletes a video from cloud storage.

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

#### `loadCloudVideos()`
Fetches and displays the list of cloud videos.

#### `displayCloudVideos(videos)`
Renders the cloud video list with proper formatting.

#### `loadCloudVideo(gcsPath)`
Loads a specific video from cloud storage into the player.

#### `deleteCloudVideo(gcsPath)`
Deletes a video from cloud storage with confirmation.

### UI Components

The cloud video management section is located in the video controls area:

```html
<div class="cloud-videos-section mt-3">
    <div class="d-flex justify-content-between align-items-center mb-2">
        <h6 class="mb-0">
            <i class="fas fa-cloud me-2"></i>
            Cloud Storage Videos
        </h6>
        <button class="btn btn-outline-primary btn-sm" onclick="loadCloudVideos()">
            <i class="fas fa-sync-alt"></i> Refresh
        </button>
    </div>
    <div id="cloudVideosContainer" class="cloud-videos-list">
        <!-- Video items -->
    </div>
</div>
```

## Usage

### Viewing Cloud Videos
1. The cloud video list automatically loads when the page opens
2. Click "Refresh" to update the list if new videos have been uploaded
3. Each video shows:
   - Filename
   - Creation date
   - File size in MB
   - Content type

### Loading a Video
1. Click the "Load" button next to any video
2. The video will be loaded into the player with a signed URL
3. The video can then be used for TCR marking and recognition

### Deleting a Video
1. Click the trash icon next to any video
2. Confirm the deletion in the popup dialog
3. The video will be permanently removed from cloud storage
4. If the deleted video was currently loaded, the player will be cleared

## Security Features

### Signed URLs
- Videos are accessed via signed URLs with 1-hour expiration
- URLs are generated on-demand for security
- No direct bucket access from the frontend

### Confirmation Dialogs
- Delete operations require user confirmation
- Clear warning messages about permanent deletion

### Error Handling
- Graceful handling of missing videos
- Clear error messages for failed operations
- Fallback behavior when cloud storage is unavailable

## Styling

The cloud video management section uses custom CSS that matches the application's dark theme:

- **Background**: Dark theme with proper contrast
- **Cards**: Each video is displayed in a card with hover effects
- **Buttons**: Styled to match the application's button design
- **Scrollbar**: Custom scrollbar styling for the video list

## Integration

### With Existing Features
- **Video Player**: Seamlessly integrates with the existing video player
- **TCR Marking**: Loaded videos work with all TCR marking features
- **Recognition**: Cloud videos support the recognition functionality
- **Upload System**: New uploads automatically appear in the list

### State Management
- **Current Video Tracking**: Maintains the currently loaded video state
- **GCS Path Storage**: Stores the GCS path for recognition operations
- **Player Integration**: Properly handles video switching

## Error Scenarios

### Network Issues
- Shows loading spinner during operations
- Displays error messages for failed requests
- Allows retry of failed operations

### Missing Videos
- Handles cases where videos no longer exist in storage
- Shows appropriate error messages
- Removes missing videos from the list

### Permission Issues
- Graceful handling of GCS permission errors
- Clear error messages for access denied scenarios

## Performance Considerations

### Loading Optimization
- Videos are loaded on-demand
- Signed URLs are generated only when needed
- Efficient sorting and filtering of video lists

### Memory Management
- Proper cleanup of video resources
- Efficient DOM manipulation for video lists
- Minimal memory footprint for the feature

## Future Enhancements

### Potential Improvements
1. **Search and Filter**: Add search functionality for large video collections
2. **Bulk Operations**: Support for bulk delete or move operations
3. **Video Thumbnails**: Generate and display video thumbnails
4. **Folder Organization**: Support for organizing videos in folders
5. **Sharing**: Add ability to share videos with other users

### Advanced Features
1. **Video Metadata**: Display additional video metadata (duration, resolution, etc.)
2. **Playlist Support**: Create and manage video playlists
3. **Auto-cleanup**: Automatic cleanup of old videos based on policies
4. **Usage Analytics**: Track video usage and access patterns

## Troubleshooting

### Common Issues

1. **Videos Not Loading**
   - Check GCS bucket permissions
   - Verify the videos folder exists in the bucket
   - Check network connectivity

2. **Delete Operations Failing**
   - Verify GCS delete permissions
   - Check if the video is currently in use
   - Ensure proper authentication

3. **Signed URL Issues**
   - Check GCS service account permissions
   - Verify URL expiration settings
   - Check for clock synchronization issues

### Debug Information
- Console logging for all operations
- Network request monitoring
- Error message details in the UI
- GCS operation logging

This feature provides a comprehensive solution for managing cloud-stored videos within the MOS application, making it easy for users to access and work with their uploaded content. 