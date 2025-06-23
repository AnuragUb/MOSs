# WMV Watchdog Converter for Windows Server 2012 R2

A comprehensive watchdog system that monitors folders for WMV files and automatically converts them to MP4 format while preserving the original folder structure.

## Features

- **Folder Structure Preservation**: Maintains the exact folder hierarchy from the watch folder to the output folder
- **Automatic Conversion**: Monitors specified folders and converts WMV files as they appear
- **Windows Service**: Can run as a background Windows service for 24/7 operation
- **Configurable**: Extensive configuration options via INI file
- **Logging**: Comprehensive logging with rotation and retention
- **Error Handling**: Robust error handling and recovery
- **Performance Optimized**: Designed for server environments

## System Requirements

- Windows Server 2012 R2 Standard (or later)
- PowerShell 4.0 or later
- FFmpeg installed and available in system PATH
- Administrator privileges for service installation

## Installation

### 1. Install FFmpeg

The system requires FFmpeg to be installed. You can install it using Chocolatey:

```powershell
# Install Chocolatey (if not already installed)
Set-ExecutionPolicy Bypass -Scope Process -Force
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))

# Install FFmpeg
choco install ffmpeg -y
```

Alternatively, download FFmpeg manually from https://ffmpeg.org/download.html and add it to your system PATH.

### 2. Configure the Watchdog

1. Edit `watchdog_config.ini` to set your desired paths and settings:

```ini
[Paths]
watch_folder=C:\WatchFolder
output_base_folder=C:\ConvertedOutput
log_folder=C:\WMVConverter\logs

[Service]
scan_interval=30
```

2. Create the watch folder and output folder:
```powershell
New-Item -ItemType Directory -Path "C:\WatchFolder" -Force
New-Item -ItemType Directory -Path "C:\ConvertedOutput" -Force
```

### 3. Install as Windows Service

Run the installation script as Administrator:

```powershell
# Navigate to the WMVConverter directory
cd C:\path\to\WMVConverter

# Run the service installer
.\install_watchdog_service.ps1
```

Or with custom parameters:

```powershell
.\install_watchdog_service.ps1 -WatchFolder "D:\Videos\Input" -OutputBaseFolder "D:\Videos\Converted" -ScanInterval 60
```

### 4. Start the Service

```powershell
# Start the service
Start-Service -Name "WMVWatchdogConverter"

# Check service status
Get-Service -Name "WMVWatchdogConverter"

# View recent logs
Get-Content "C:\WMVConverter\logs\watchdog_*.log" | Select-Object -Last 50
```

## Usage

### Running as a Service (Recommended)

Once installed as a service, the watchdog will:
- Start automatically on system boot
- Run continuously in the background
- Monitor the configured watch folder
- Convert WMV files as they appear
- Log all activities

### Running Manually

You can also run the watchdog manually for testing:

```powershell
# Basic usage
.\watchdog_converter.ps1

# With custom parameters
.\watchdog_converter.ps1 -WatchFolder "D:\Videos" -OutputBaseFolder "D:\Converted" -ScanInterval 60

# Using enhanced version with config file
.\watchdog_converter_enhanced.ps1 -ConfigFile "watchdog_config.ini"
```

### Folder Structure Example

**Input Structure:**
```
C:\WatchFolder\
├── Project1\
│   ├── video1.wmv
│   └── Subfolder\
│       └── video2.wmv
└── Project2\
    └── video3.wmv
```

**Output Structure:**
```
C:\ConvertedOutput\
├── Project1\
│   ├── video1.mp4
│   └── Subfolder\
│       └── video2.mp4
└── Project2\
    └── video3.mp4
```

## Configuration

### Main Configuration File (`watchdog_config.ini`)

The configuration file contains several sections:

#### [Paths]
- `watch_folder`: Folder to monitor for WMV files
- `output_base_folder`: Base folder for converted MP4 files
- `log_folder`: Folder for log files

#### [Service]
- `service_name`: Windows service name
- `scan_interval`: How often to scan for new files (seconds)
- `auto_start`: Whether to start service automatically on boot

#### [Conversion]
- `video_codec`: FFmpeg video codec (default: libx264)
- `crf`: Constant Rate Factor for quality (18-28, lower = better)
- `preset`: Encoding preset (ultrafast to veryslow)
- `audio_codec`: Audio codec (default: aac)
- `audio_bitrate`: Audio bitrate (default: 128k)

#### [Behavior]
- `delete_original`: Delete original WMV file after conversion
- `skip_existing`: Skip files that already have MP4 equivalents
- `recursive`: Process subdirectories recursively
- `max_concurrent`: Maximum concurrent conversions

#### [Logging]
- `detailed_logging`: Enable detailed logging
- `log_level`: Log level (debug, info, warning, error)
- `log_retention_days`: How long to keep log files
- `max_log_size_mb`: Maximum log file size before rotation

## Management Commands

### Service Management

```powershell
# Start the service
Start-Service -Name "WMVWatchdogConverter"

# Stop the service
Stop-Service -Name "WMVWatchdogConverter"

# Restart the service
Restart-Service -Name "WMVWatchdogConverter"

# Check service status
Get-Service -Name "WMVWatchdogConverter"

# Remove the service
Remove-Service -Name "WMVWatchdogConverter"
```

### Log Management

```powershell
# View recent logs
Get-Content "C:\WMVConverter\logs\watchdog_*.log" | Select-Object -Last 100

# View logs for a specific date
Get-Content "C:\WMVConverter\logs\watchdog_20241201_*.log"

# Clear old logs
Get-ChildItem "C:\WMVConverter\logs\*.log" | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-7) } | Remove-Item
```

## Troubleshooting

### Common Issues

1. **Service won't start**
   - Check if FFmpeg is installed and in PATH
   - Verify all configured folders exist
   - Check Windows Event Viewer for service errors

2. **Files not being converted**
   - Verify the watch folder path is correct
   - Check file permissions on watch and output folders
   - Review log files for error messages

3. **Poor conversion performance**
   - Adjust the `preset` setting in configuration
   - Increase `scan_interval` to reduce CPU usage
   - Consider using hardware acceleration if available

4. **Service stops unexpectedly**
   - Check log files for error messages
   - Verify sufficient disk space
   - Check Windows Event Viewer for system errors

### Log Analysis

The watchdog creates detailed logs with timestamps and log levels:

```
[2024-12-01 10:30:15] [INFO] === Enhanced WMV Watchdog Converter Started ===
[2024-12-01 10:30:15] [INFO] Watch Folder: C:\WatchFolder
[2024-12-01 10:30:15] [INFO] Found 2 WMV file(s)
[2024-12-01 10:30:16] [INFO] Converting: C:\WatchFolder\video1.wmv
[2024-12-01 10:30:45] [INFO] SUCCESS: Converted C:\WatchFolder\video1.wmv to C:\ConvertedOutput\video1.mp4
```

### Performance Tuning

For optimal performance on Windows Server 2012 R2:

1. **CPU Settings**: Use `preset=medium` for good balance of speed and quality
2. **Memory**: Set `memory_limit=0` for unlimited memory usage
3. **Scan Interval**: Use 30-60 seconds for most environments
4. **Concurrent Conversions**: Limit to 2-4 concurrent conversions based on CPU cores

## Security Considerations

- Run the service with minimal required privileges
- Restrict access to watch and output folders
- Regularly review log files for suspicious activity
- Consider enabling file integrity checking for sensitive environments

## Support

For issues or questions:
1. Check the log files in `C:\WMVConverter\logs\`
2. Review Windows Event Viewer for system errors
3. Verify FFmpeg installation and PATH configuration
4. Test with manual execution before running as service

## File Structure

```
WMVConverter/
├── watchdog_converter.ps1              # Basic watchdog script
├── watchdog_converter_enhanced.ps1     # Enhanced version with config file
├── watchdog_config.ini                 # Configuration file
├── install_watchdog_service.ps1        # Service installer
├── WATCHDOG_README.md                  # This file
├── convert_wmv_server.bat              # Original batch converter
├── server_config.ini                   # Original server config
└── install_wmv_service.ps1             # Original service installer
``` 