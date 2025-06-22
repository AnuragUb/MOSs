# WMV to MP4 Converter for Windows Server 2012 R2

This package provides a robust WMV to MP4 conversion solution specifically configured for Windows Server 2012 R2 Standard environments.

## Files Included

- `convert_wmv_server.bat` - Main conversion script with server optimizations
- `install_wmv_service.ps1` - PowerShell script to install as Windows Service
- `server_config.ini` - Configuration file for server settings
- `convert_wmv.bat` - Original desktop version (for reference)

## Prerequisites

### System Requirements
- Windows Server 2012 R2 Standard (or later)
- Administrator privileges
- Internet connection (for FFmpeg installation)
- Minimum 2GB RAM
- Sufficient disk space for video files

### PowerShell Execution Policy
Before running the installation script, ensure PowerShell execution policy allows script execution:

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope LocalMachine
```

## Installation Options

### Option 1: Manual Installation (Recommended for Testing)

1. **Extract Files**: Place all files in a dedicated folder (e.g., `C:\WMVConverter`)

2. **Run as Administrator**: Right-click `convert_wmv_server.bat` and select "Run as administrator"

3. **Monitor Logs**: Check the `logs` folder for conversion status

### Option 2: Windows Service Installation (Recommended for Production)

1. **Open PowerShell as Administrator**

2. **Navigate to Script Directory**:
   ```powershell
   cd C:\WMVConverter
   ```

3. **Install Service**:
   ```powershell
   .\install_wmv_service.ps1
   ```

4. **Verify Installation**:
   ```powershell
   Get-Service -Name "WMVConverterService"
   ```

## Configuration

### Basic Configuration
Edit `server_config.ini` to customize settings:

```ini
[Paths]
input_folder=C:\Videos\Input
output_folder=C:\Videos\Output
log_folder=C:\Videos\Logs

[Conversion]
crf=23                    # Quality setting (18-28)
preset=medium            # Speed vs quality balance
scan_interval=30         # Check for new files every 30 seconds
```

### Advanced Configuration

#### Performance Tuning
```ini
[Performance]
cpu_priority=high        # Use high priority for faster conversion
max_concurrent=4         # Allow 4 simultaneous conversions
hardware_acceleration=true  # Enable if GPU available
```

#### Monitoring Setup
```ini
[Monitoring]
email_notifications=true
smtp_server=smtp.company.com
notification_email=admin@company.com
notify_on_error=true
```

## Usage

### Adding Files for Conversion
1. Place WMV files in the input folder
2. The service will automatically detect and convert them
3. Converted MP4 files appear in the output folder

### Service Management

#### Start Service
```powershell
Start-Service -Name "WMVConverterService"
```

#### Stop Service
```powershell
Stop-Service -Name "WMVConverterService"
```

#### Check Status
```powershell
Get-Service -Name "WMVConverterService"
```

#### View Logs
```powershell
Get-Content "C:\WMVConverter\logs\conversion_YYYYMMDD_HHMMSS.log" -Tail 50
```

### Manual Conversion
For one-time conversions, run the batch file directly:
```cmd
convert_wmv_server.bat
```

## Troubleshooting

### Common Issues

#### 1. FFmpeg Installation Fails
**Symptoms**: "FFmpeg is not installed" error
**Solutions**:
- Ensure internet connection is available
- Check Windows Firewall settings
- Try manual installation from https://ffmpeg.org/download.html

#### 2. Service Won't Start
**Symptoms**: Service status shows "Stopped"
**Solutions**:
- Check event logs: `eventvwr.msc`
- Verify administrator privileges
- Check service dependencies

#### 3. Conversion Fails
**Symptoms**: Files not converting or errors in logs
**Solutions**:
- Check file permissions
- Verify sufficient disk space
- Review FFmpeg error messages in logs

#### 4. Performance Issues
**Symptoms**: Slow conversions or high CPU usage
**Solutions**:
- Adjust `preset` setting in config (use "faster" for speed)
- Reduce `max_concurrent` setting
- Check available system resources

### Log Analysis

#### Log File Locations
- Conversion logs: `logs\conversion_YYYYMMDD_HHMMSS.log`
- Service logs: `logs\service.log`
- Error logs: `logs\service_error.log`

#### Key Log Messages
- `[SUCCESS]` - Conversion completed successfully
- `[ERROR]` - Conversion failed (check details)
- `[SKIP]` - File already converted
- `[INFO]` - General information

## Security Considerations

### File Permissions
- Ensure input/output folders have appropriate permissions
- Consider using dedicated service account
- Restrict access to configuration files

### Network Security
- If using email notifications, configure SMTP securely
- Consider VPN for remote access to logs
- Monitor for unauthorized file access

### Resource Protection
- Set appropriate file size limits
- Monitor disk space usage
- Configure log rotation to prevent disk filling

## Performance Optimization

### For High-Volume Servers
1. **Increase scan interval** to reduce CPU usage
2. **Use faster preset** for quicker conversions
3. **Enable hardware acceleration** if available
4. **Distribute load** across multiple servers

### For Low-Resource Servers
1. **Reduce max_concurrent** to 1
2. **Use slower preset** for better compression
3. **Increase scan interval** to 60+ seconds
4. **Monitor memory usage**

## Maintenance

### Regular Tasks
- **Weekly**: Review log files for errors
- **Monthly**: Clean old log files
- **Quarterly**: Update FFmpeg to latest version
- **Annually**: Review and update configuration

### Backup Strategy
- Backup configuration files
- Archive important conversion logs
- Document custom settings

## Support

### Getting Help
1. Check the logs first for error details
2. Review this README for common solutions
3. Verify system requirements are met
4. Test with a simple WMV file

### Logging Levels
- **ERROR**: Critical issues requiring immediate attention
- **WARNING**: Potential problems to monitor
- **INFO**: General operational information
- **DEBUG**: Detailed troubleshooting information

## Version History

- **v2.0**: Server-optimized version with service installation
- **v1.0**: Original desktop version

## License

This software is provided as-is for internal use. Ensure compliance with FFmpeg licensing requirements for production deployment. 