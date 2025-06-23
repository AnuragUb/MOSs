# Enhanced WMV Converter - Based on Working Server Integration

This enhanced version combines the reliability of your working WMV converter system with advanced features like folder structure preservation and configuration management.

## 🎯 **What's Enhanced**

### **Based on Your Working System:**
- ✅ **VBS Service Wrapper** - Uses the proven `wscript.exe` approach
- ✅ **Batch File Core** - Reliable batch processing like your `working_converter.bat`
- ✅ **Simple Service Installation** - Based on your `install_service_fixed.bat`
- ✅ **Windows Server 2012 R2 Compatible** - Tested and proven approach

### **New Features Added:**
- ✅ **Folder Structure Preservation** - Maintains exact folder hierarchy
- ✅ **Configuration File** - Easy settings management via INI file
- ✅ **Recursive Processing** - Handles subdirectories automatically
- ✅ **Enhanced Logging** - Better tracking and debugging
- ✅ **Flexible Paths** - Configurable input/output folders

## 📁 **File Structure**

```
WMVConverter/
├── watchdog_converter_enhanced.bat      # Enhanced batch converter (MAIN)
├── service_wrapper_enhanced.vbs         # VBS service wrapper
├── install_service_enhanced.bat         # Service installer
├── watchdog_config_simple.ini           # Configuration file
├── ENHANCED_README.md                   # This file
├── watchdog_converter.ps1               # PowerShell version (alternative)
├── watchdog_converter_enhanced.ps1      # Enhanced PowerShell version
├── install_watchdog_service.ps1         # PowerShell service installer
├── watchdog_config.ini                  # Full configuration file
└── setup_watchdog.ps1                   # Automated setup script
```

## 🚀 **Quick Installation**

### **Option 1: Enhanced Batch System (Recommended - Based on Your Working System)**

1. **Install the service:**
```cmd
# Run as Administrator
install_service_enhanced.bat
```

2. **Configure settings (optional):**
Edit `watchdog_config_simple.ini` to change paths or behavior.

3. **Drop WMV files in `C:\WMVConverter` and they'll be automatically converted!**

### **Option 2: PowerShell System (Alternative)**

1. **Run automated setup:**
```powershell
# Run as Administrator
.\setup_watchdog.ps1
```

2. **Start the service:**
```powershell
Start-Service -Name "WMVWatchdogConverter"
```

## 📋 **Configuration**

### **Simple Configuration (`watchdog_config_simple.ini`)**

```ini
# Watch folder - where to monitor for WMV files
watch_folder=C:\WMVConverter

# Output base folder - where converted MP4 files will be saved
output_base_folder=C:\WMVConverter\converted

# Log folder - where log files will be stored
log_folder=C:\WMVConverter\logs

# Scan interval in seconds
scan_interval=30

# Delete original WMV file after successful conversion (true/false)
delete_original=false

# Skip files that already have MP4 equivalents (true/false)
skip_existing=true

# Process subdirectories recursively (true/false)
recursive=true
```

## 🔄 **How It Works**

### **Folder Structure Preservation Example:**

**Input Structure:**
```
C:\WMVConverter\
├── Project1\
│   ├── video1.wmv
│   └── Subfolder\
│       └── video2.wmv
└── Project2\
    └── video3.wmv
```

**Output Structure:**
```
C:\WMVConverter\converted\
├── Project1\
│   ├── video1.mp4
│   └── Subfolder\
│       └── video2.mp4
└── Project2\
    └── video3.mp4
```

## 🛠 **Service Management**

### **Enhanced Batch System:**
```cmd
# Check service status
sc query WMVConverterEnhanced

# Start service
sc start WMVConverterEnhanced

# Stop service
sc stop WMVConverterEnhanced

# Remove service
sc delete WMVConverterEnhanced
```

### **PowerShell System:**
```powershell
# Check service status
Get-Service -Name "WMVWatchdogConverter"

# Start service
Start-Service -Name "WMVWatchdogConverter"

# Stop service
Stop-Service -Name "WMVWatchdogConverter"

# Remove service
Remove-Service -Name "WMVWatchdogConverter"
```

## 📊 **Logging**

### **Enhanced Batch System:**
- **Main logs:** `C:\WMVConverter\logs\watchdog_YYYYMMDD_HHMMSS.log`
- **Service wrapper logs:** `C:\WMVConverter\logs\service_wrapper_YYYYMMDD_HHMMSS.log`

### **PowerShell System:**
- **Main logs:** `C:\WMVConverter\logs\watchdog_YYYYMMDD_HHMMSS.log`

### **View Recent Logs:**
```cmd
# Enhanced Batch System
type C:\WMVConverter\logs\watchdog_*.log

# PowerShell System
Get-Content C:\WMVConverter\logs\watchdog_*.log | Select-Object -Last 50
```

## 🔧 **Troubleshooting**

### **Common Issues:**

1. **Service won't start**
   - Check if FFmpeg is installed: `ffmpeg -version`
   - Verify directories exist: `dir C:\WMVConverter`
   - Check Windows Event Viewer for service errors

2. **Files not being converted**
   - Verify watch folder path in config file
   - Check file permissions on folders
   - Review log files for error messages

3. **Folder structure not preserved**
   - Ensure `recursive=true` in config file
   - Check that subdirectories have proper permissions

### **Log Analysis:**

**Successful conversion:**
```
[2024-12-01 10:30:15] Converting: C:\WMVConverter\Project1\video1.wmv
[2024-12-01 10:30:45] SUCCESS: Converted C:\WMVConverter\Project1\video1.wmv to C:\WMVConverter\converted\Project1\video1.mp4
```

**Skipped file:**
```
[2024-12-01 10:31:00] SKIP: C:\WMVConverter\converted\Project1\video1.mp4 already exists
```

## 🎯 **Why This Approach Works**

### **Based on Your Proven System:**
1. **VBS Wrapper** - More reliable than PowerShell for Windows services
2. **Batch Processing** - Native Windows compatibility
3. **Simple Service Installation** - Uses proven `sc.exe` commands
4. **Fixed Base Path** - Uses `C:\WMVConverter` like your working system

### **Enhanced Features:**
1. **Configuration File** - Easy customization without editing code
2. **Recursive Processing** - Handles complex folder structures
3. **Better Logging** - More detailed tracking for troubleshooting
4. **Error Handling** - Robust error recovery and reporting

## 📞 **Support**

### **For Enhanced Batch System:**
1. Check logs in `C:\WMVConverter\logs\`
2. Verify FFmpeg installation: `ffmpeg -version`
3. Test manual execution: `watchdog_converter_enhanced.bat`

### **For PowerShell System:**
1. Check logs in `C:\WMVConverter\logs\`
2. Review Windows Event Viewer
3. Test manual execution: `.\watchdog_converter_enhanced.ps1`

## 🏆 **Recommendation**

**Use the Enhanced Batch System** (`install_service_enhanced.bat`) as it's based on your proven working approach with added folder structure preservation and configuration features.

The PowerShell system is available as an alternative with more advanced features, but the batch system should be more reliable on your Windows Server 2012 R2 environment. 