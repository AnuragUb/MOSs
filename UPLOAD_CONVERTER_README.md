# Upload Folder WMV to MP4 Converter

This converter monitors your web app's `uploads` folder and automatically converts WMV files to MP4 while preserving the folder structure.

## How It Works

1. **Monitors the `uploads` folder** every 30 seconds
2. **Scans recursively** - finds WMV files in all subfolders
3. **Preserves folder structure** - maintains the same folder hierarchy in the `converted` folder
4. **Skips existing files** - won't reconvert files that are already converted
5. **Logs all activity** - keeps detailed logs in the `logs` folder

## Folder Structure

```
Your Project/
├── uploads/                    # Input folder (WMV files)
│   ├── video1.wmv
│   ├── folder1/
│   │   ├── video2.wmv
│   │   └── subfolder/
│   │       └── video3.wmv
│   └── folder2/
│       └── video4.wmv
├── converted/                  # Output folder (MP4 files)
│   ├── video1.mp4
│   ├── folder1/
│   │   ├── video2.mp4
│   │   └── subfolder/
│   │       └── video3.mp4
│   └── folder2/
│       └── video4.mp4
├── logs/                       # Log files
└── upload_folder_converter.bat # The converter script
```

## Usage

### 1. **Start the Converter**
```cmd
upload_folder_converter.bat
```

### 2. **Let It Run**
- The converter will run continuously
- It scans every 30 seconds for new WMV files
- Just leave it running in the background

### 3. **Upload Files**
- Place WMV files in the `uploads` folder
- Create subfolders as needed
- The converter will automatically detect and convert them

## Features

- ✅ **Based on your working converter** - uses the same reliable logic
- ✅ **Recursive scanning** - finds files in all subfolders
- ✅ **Folder structure preservation** - maintains your folder hierarchy
- ✅ **Skip existing files** - won't waste time on already converted files
- ✅ **Detailed logging** - tracks all conversions and errors
- ✅ **30-second scan interval** - same as your working converter
- ✅ **FFmpeg optimization** - uses the same high-quality conversion settings

## Requirements

- **FFmpeg** must be installed and in your system PATH
- **Windows** (batch file)
- **Write permissions** to create `converted` and `logs` folders

## Configuration

The script uses these default settings:
- **Input folder:** `uploads`
- **Output folder:** `converted`
- **Log folder:** `logs`
- **Scan interval:** 30 seconds
- **FFmpeg settings:** `-c:v libx264 -crf 23 -preset medium -c:a aac -b:a 128k -movflags +faststart`

## Example Workflow

1. **Start the converter:**
   ```cmd
   upload_folder_converter.bat
   ```

2. **Upload files to your web app** (they go to `uploads/`)

3. **The converter automatically:**
   - Detects new WMV files
   - Converts them to MP4
   - Preserves folder structure
   - Logs the process

4. **Find converted files in `converted/` folder**

## Logs

Check the `logs/` folder for detailed conversion logs:
- `upload_conversion_YYYYMMDD_HHMMSS.log`
- Shows all conversions, skips, and errors
- Timestamped entries for easy tracking

## Stopping the Converter

Press `Ctrl+C` to stop the converter when needed. 