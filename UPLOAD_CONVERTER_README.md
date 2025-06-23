# Upload Folder WMV to MP4 Converter with Priority Processing

This converter monitors your web app's `uploads` folder and automatically converts WMV files to MP4 while preserving the folder structure. **NEW: Priority folder support for urgent files!**

## How It Works

1. **Monitors the `uploads` folder** every 30 seconds
2. **Processes PRIORITY folder first** - urgent files get converted immediately
3. **Then processes regular uploads** - normal files are converted after priority files
4. **Scans recursively** - finds WMV files in all subfolders
5. **Preserves folder structure** - maintains the same folder hierarchy in the `converted` folder
6. **Skips existing files** - won't reconvert files that are already converted
7. **Logs all activity** - keeps detailed logs in the `logs` folder

## Folder Structure

```
Your Project/
├── uploads/                    # Input folder (WMV files)
│   ├── priority/               # 🚨 PRIORITY FOLDER - urgent files processed FIRST
│   │   ├── urgent_video1.wmv
│   │   └── emergency/
│   │       └── urgent_video2.wmv
│   ├── video1.wmv             # Regular files processed SECOND
│   ├── folder1/
│   │   ├── video2.wmv
│   │   └── subfolder/
│   │       └── video3.wmv
│   └── folder2/
│       └── video4.wmv
├── converted/                  # Output folder (MP4 files)
│   ├── priority/               # Priority files converted first
│   │   ├── urgent_video1.mp4
│   │   └── emergency/
│   │       └── urgent_video2.mp4
│   ├── video1.mp4             # Regular files converted second
│   ├── folder1/
│   │   ├── video2.mp4
│   │   └── subfolder/
│   │       └── video3.mp4
│   └── folder2/
│       └── video4.mp4
├── logs/                       # Log files
└── upload_folder_converter.bat # The converter script
```

## Priority Processing

### 🚨 **Priority Folder: `uploads\priority\`**
- **Processed FIRST** in every scan cycle
- **Urgent files** get immediate attention
- **Perfect for time-sensitive conversions**
- **Same folder structure preservation**

### 📁 **Regular Uploads: `uploads\`**
- **Processed SECOND** after priority files
- **Normal workflow** for regular files
- **Same reliable conversion** as before

## Usage

### 1. **Start the Converter**
```cmd
upload_folder_converter.bat
```

### 2. **Let It Run**
- The converter will run continuously
- It scans every 30 seconds for new WMV files
- **Priority files are processed first**, then regular files
- Just leave it running in the background

### 3. **Upload Files**

#### **For Urgent Files:**
- Place WMV files in `uploads\priority\`
- Create subfolders as needed: `uploads\priority\emergency\`
- **These files will be converted FIRST**

#### **For Regular Files:**
- Place WMV files in `uploads\`
- Create subfolders as needed: `uploads\folder1\subfolder\`
- **These files will be converted AFTER priority files**

## Processing Order

1. **Scan Priority Folder** (`uploads\priority\`)
   - Convert all WMV files found here FIRST
   - Preserve folder structure in `converted\priority\`

2. **Scan Regular Uploads** (`uploads\`)
   - Convert all WMV files found here SECOND
   - Preserve folder structure in `converted\`

3. **Wait 30 seconds** and repeat

## Features

- ✅ **Priority processing** - urgent files converted first
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
- **Priority folder:** `uploads\priority`
- **Input folder:** `uploads`
- **Output folder:** `converted`
- **Log folder:** `logs`
- **Scan interval:** 30 seconds
- **FFmpeg settings:** `-c:v libx264 -crf 23 -preset medium -c:a aac -b:a 128k -movflags +faststart`

## Example Workflow

### **Urgent Conversion:**
1. **Place urgent file:** `uploads\priority\urgent_video.wmv`
2. **Converter detects it** in next scan (within 30 seconds)
3. **Converts immediately** to `converted\priority\urgent_video.mp4`
4. **Available for use** right away

### **Regular Conversion:**
1. **Place regular file:** `uploads\folder1\video.wmv`
2. **Converter processes priority files first**
3. **Then converts regular file** to `converted\folder1\video.mp4`
4. **Available after priority processing**

## Logs

Check the `logs/` folder for detailed conversion logs:
- `upload_conversion_YYYYMMDD_HHMMSS.log`
- Shows priority vs regular processing
- Timestamped entries for easy tracking
- Clear indication of which files are priority vs regular

## Stopping the Converter

Press `Ctrl+C` to stop the converter when needed.

## Tips

- **Use priority folder sparingly** - only for truly urgent files
- **Regular uploads folder** is perfect for normal workflow
- **Both folders preserve structure** - organize files however you need
- **Logs show processing order** - easy to track what's happening 