# Folder Structure Preservation Algorithm
# For C:\FTPData\SPN\Upload → C:\FTPData\SPN\Converted

## 🎯 **Algorithm Overview**

The system monitors `C:\FTPData\SPN\Upload` and recreates the exact folder structure in `C:\FTPData\SPN\Converted` with converted MP4 files.

## 📁 **Input Structure (WMV Files)**
```
C:\FTPData\SPN\Upload\
├── Project1\
│   ├── video1.wmv
│   ├── video2.wmv
│   └── Subfolder\
│       ├── video3.wmv
│       └── DeepFolder\
│           └── video4.wmv
├── Project2\
│   └── presentation.wmv
└── Misc\
    ├── intro.wmv
    └── outro.wmv
```

## 📁 **Output Structure (MP4 Files)**
```
C:\FTPData\SPN\Converted\
├── Project1\
│   ├── video1.mp4
│   ├── video2.mp4
│   └── Subfolder\
│       ├── video3.mp4
│       └── DeepFolder\
│           └── video4.mp4
├── Project2\
│   └── presentation.mp4
└── Misc\
    ├── intro.mp4
    └── outro.mp4
```

## 🔄 **Algorithm Steps**

### **Step 1: Directory Scanning**
```batch
:: Scan C:\FTPData\SPN\Upload recursively
for /d %%D in ("C:\FTPData\SPN\Upload\*") do (
    :: Process each subdirectory
    call :process_directory "%%D" "%%~nxD"
)
```

### **Step 2: Relative Path Calculation**
```batch
:: For each WMV file found:
set "input_file=C:\FTPData\SPN\Upload\Project1\Subfolder\video3.wmv"
set "relative_path=Project1\Subfolder"
set "filename=video3"
```

### **Step 3: Output Path Construction**
```batch
:: Build output path preserving structure
set "output_file=C:\FTPData\SPN\Converted\%relative_path%\%filename%.mp4"
:: Result: C:\FTPData\SPN\Converted\Project1\Subfolder\video3.mp4
```

### **Step 4: Directory Creation**
```batch
:: Create output directory if it doesn't exist
for %%D in ("%output_file%") do set "output_dir=%%~dpD"
if not exist "!output_dir!" mkdir "!output_dir!"
```

### **Step 5: File Conversion**
```batch
:: Convert WMV to MP4
ffmpeg -i "!input_file!" -c:v libx264 -crf 23 -preset medium -c:a aac -b:a 128k -movflags +faststart -y "!output_file!"
```

## 📊 **Real-World Examples**

### **Example 1: Simple File**
- **Input:** `C:\FTPData\SPN\Upload\video1.wmv`
- **Output:** `C:\FTPData\SPN\Converted\video1.mp4`

### **Example 2: Single Level Folder**
- **Input:** `C:\FTPData\SPN\Upload\Project1\video1.wmv`
- **Output:** `C:\FTPData\SPN\Converted\Project1\video1.mp4`

### **Example 3: Deep Nested Structure**
- **Input:** `C:\FTPData\SPN\Upload\Client\2024\Q1\Meeting\recording.wmv`
- **Output:** `C:\FTPData\SPN\Converted\Client\2024\Q1\Meeting\recording.mp4`

## 🔧 **Configuration Settings**

### **Key Settings in `watchdog_config_server.ini`:**
```ini
# Watch folder - where WMV files are uploaded
watch_folder=C:\FTPData\SPN\Upload

# Output base folder - where converted MP4 files will be saved
output_base_folder=C:\FTPData\SPN\Converted

# Process subdirectories recursively (true/false)
recursive=true

# Skip files that already have MP4 equivalents (true/false)
skip_existing=true
```

## 📝 **Log Output Examples**

### **Successful Conversion:**
```
[2024-12-01 10:30:15] Scanning directory: C:\FTPData\SPN\Upload\Project1
[2024-12-01 10:30:15] Converting: C:\FTPData\SPN\Upload\Project1\video1.wmv
[2024-12-01 10:30:15] Output: C:\FTPData\SPN\Converted\Project1\video1.mp4
[2024-12-01 10:30:45] SUCCESS: Converted C:\FTPData\SPN\Upload\Project1\video1.wmv to C:\FTPData\SPN\Converted\Project1\video1.mp4
```

### **Skipped File:**
```
[2024-12-01 10:31:00] SKIP: C:\FTPData\SPN\Converted\Project1\video1.mp4 already exists
```

### **Directory Creation:**
```
[2024-12-01 10:31:15] Creating directory: C:\FTPData\SPN\Converted\Project1\Subfolder
[2024-12-01 10:31:15] Converting: C:\FTPData\SPN\Upload\Project1\Subfolder\video2.wmv
```

## 🚀 **Installation and Usage**

### **1. Install the Service:**
```cmd
# Run as Administrator
install_service_server.bat
```

### **2. Service Management:**
```cmd
# Check status
sc query WMVConverterServer

# Start service
sc start WMVConverterServer

# Stop service
sc stop WMVConverterServer
```

### **3. Monitor Logs:**
```cmd
# View recent logs
type C:\WMVConverter\logs\watchdog_*.log

# View service wrapper logs
type C:\WMVConverter\logs\server_wrapper_*.log
```

## ✅ **Benefits of This Approach**

1. **Exact Structure Preservation** - Maintains all folder levels and names
2. **Automatic Directory Creation** - Creates output folders as needed
3. **Recursive Processing** - Handles unlimited folder depth
4. **Skip Existing Files** - Avoids re-converting already processed files
5. **Comprehensive Logging** - Tracks all operations for debugging
6. **FTP-Friendly** - Works seamlessly with FTP uploads

## 🔍 **Troubleshooting**

### **Common Issues:**

1. **Missing Output Folders**
   - Check if `recursive=true` in config
   - Verify write permissions on `C:\FTPData\SPN\Converted`

2. **Files Not Converting**
   - Check if WMV files are in the correct input folder
   - Verify FFmpeg is installed and in PATH
   - Review log files for error messages

3. **Permission Issues**
   - Run service installer as Administrator
   - Check folder permissions on FTP directories 