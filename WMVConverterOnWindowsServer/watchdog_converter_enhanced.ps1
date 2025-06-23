# Enhanced WMV Watchdog Converter for Windows Server 2012 R2
# Monitors a folder for WMV files and converts them while preserving folder structure
# Reads configuration from watchdog_config.ini

param(
    [string]$ConfigFile = "watchdog_config.ini",
    [switch]$RunAsService = $false
)

# Function to read INI configuration
function Read-IniFile {
    param([string]$FilePath)
    
    $config = @{}
    
    if (!(Test-Path $FilePath)) {
        Write-Error "Configuration file not found: $FilePath"
        return $null
    }
    
    $currentSection = ""
    
    Get-Content $FilePath | ForEach-Object {
        $line = $_.Trim()
        
        # Skip comments and empty lines
        if ($line -eq "" -or $line.StartsWith("#")) {
            return
        }
        
        # Check if it's a section header
        if ($line.StartsWith("[") -and $line.EndsWith("]")) {
            $currentSection = $line.Substring(1, $line.Length - 2)
            $config[$currentSection] = @{}
            return
        }
        
        # Parse key-value pairs
        if ($line.Contains("=") -and $currentSection -ne "") {
            $parts = $line.Split("=", 2)
            $key = $parts[0].Trim()
            $value = $parts[1].Trim()
            
            # Convert boolean strings to actual booleans
            if ($value -eq "true") { $value = $true }
            elseif ($value -eq "false") { $value = $false }
            
            # Convert numeric values
            if ($value -match "^\d+$") { $value = [int]$value }
            elseif ($value -match "^\d+\.\d+$") { $value = [double]$value }
            
            $config[$currentSection][$key] = $value
        }
    }
    
    return $config
}

# Function to write to log with rotation
function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    
    Write-Host $logEntry
    
    if ($script:LogFile) {
        Add-Content -Path $script:LogFile -Value $logEntry
        
        # Check log file size and rotate if necessary
        $logFileInfo = Get-Item $script:LogFile -ErrorAction SilentlyContinue
        if ($logFileInfo -and $logFileInfo.Length -gt ($script:Config.Paths.max_log_size_mb * 1MB)) {
            $archivePath = $script:LogFile.Replace(".log", "_$(Get-Date -Format 'yyyyMMdd_HHmmss').log")
            Move-Item $script:LogFile $archivePath
            Write-Host "Log file rotated to: $archivePath"
        }
    }
}

# Function to clean old log files
function Remove-OldLogs {
    param([string]$LogFolder, [int]$RetentionDays)
    
    try {
        $cutoffDate = (Get-Date).AddDays(-$RetentionDays)
        $oldLogs = Get-ChildItem -Path $LogFolder -Filter "*.log" | Where-Object { $_.LastWriteTime -lt $cutoffDate }
        
        foreach ($log in $oldLogs) {
            Remove-Item $log.FullName -Force
            Write-Log "Removed old log file: $($log.Name)" "DEBUG"
        }
    }
    catch {
        Write-Log "Error cleaning old logs: $($_.Exception.Message)" "ERROR"
    }
}

# Function to check if FFmpeg is available
function Test-FFmpeg {
    try {
        $null = Get-Command ffmpeg -ErrorAction Stop
        return $true
    }
    catch {
        return $false
    }
}

# Function to convert WMV to MP4 with enhanced error handling
function Convert-WMVToMP4 {
    param(
        [string]$InputFile,
        [string]$OutputFile
    )
    
    try {
        # Create output directory if it doesn't exist
        $outputDir = Split-Path $OutputFile -Parent
        if (!(Test-Path $outputDir)) {
            New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
        }
        
        # Build FFmpeg arguments from configuration
        $ffmpegArgs = @(
            "-i", "`"$InputFile`"",
            "-c:v", $script:Config.Conversion.video_codec,
            "-crf", $script:Config.Conversion.crf,
            "-preset", $script:Config.Conversion.preset,
            "-c:a", $script:Config.Conversion.audio_codec,
            "-b:a", $script:Config.Conversion.audio_bitrate
        )
        
        if ($script:Config.Conversion.fast_start) {
            $ffmpegArgs += @("-movflags", "+faststart")
        }
        
        $ffmpegArgs += @("-y", "`"$OutputFile`"")
        
        Write-Log "Converting: $InputFile" "INFO"
        Write-Log "Output: $OutputFile" "DEBUG"
        Write-Log "FFmpeg args: $($ffmpegArgs -join ' ')" "DEBUG"
        
        # Run FFmpeg conversion
        $process = Start-Process -FilePath "ffmpeg" -ArgumentList $ffmpegArgs -Wait -PassThru -NoNewWindow
        
        if ($process.ExitCode -eq 0) {
            Write-Log "SUCCESS: Converted $InputFile to $OutputFile" "INFO"
            
            # Delete original file if configured
            if ($script:Config.Behavior.delete_original) {
                Remove-Item $InputFile -Force
                Write-Log "Deleted original file: $InputFile" "INFO"
            }
            
            return $true
        } else {
            Write-Log "ERROR: Failed to convert $InputFile (Exit code: $($process.ExitCode))" "ERROR"
            return $false
        }
    }
    catch {
        Write-Log "ERROR: Exception during conversion of $InputFile - $($_.Exception.Message)" "ERROR"
        return $false
    }
}

# Function to process a single WMV file
function Process-WMVFile {
    param([string]$FilePath)
    
    try {
        # Check file size limit
        if ($script:Config.Security.max_file_size_mb -gt 0) {
            $fileSize = (Get-Item $FilePath).Length / 1MB
            if ($fileSize -gt $script:Config.Security.max_file_size_mb) {
                Write-Log "SKIP: File too large ($([math]::Round($fileSize, 2)) MB): $FilePath" "WARNING"
                return
            }
        }
        
        # Get relative path from watch folder
        $relativePath = $FilePath.Substring($script:Config.Paths.watch_folder.Length).TrimStart('\')
        $relativeDir = Split-Path $relativePath -Parent
        
        # Create output path preserving folder structure
        if ($relativeDir) {
            $outputDir = Join-Path $script:Config.Paths.output_base_folder $relativeDir
        } else {
            $outputDir = $script:Config.Paths.output_base_folder
        }
        
        # Generate output filename
        $fileName = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
        $outputFile = Join-Path $outputDir "$fileName.mp4"
        
        # Check if output file already exists
        if ($script:Config.Behavior.skip_existing -and (Test-Path $outputFile)) {
            Write-Log "SKIP: $outputFile already exists" "INFO"
            return
        }
        
        # Convert the file
        $success = Convert-WMVToMP4 -InputFile $FilePath -OutputFile $outputFile
        
        if ($success) {
            Write-Log "Conversion completed successfully for: $FilePath" "INFO"
        }
    }
    catch {
        Write-Log "ERROR: Failed to process $FilePath - $($_.Exception.Message)" "ERROR"
    }
}

# Function to scan for WMV files
function Scan-ForWMVFiles {
    try {
        Write-Log "Scanning for WMV files in: $($script:Config.Paths.watch_folder)" "DEBUG"
        
        # Find all WMV files
        $searchParams = @{
            Path = $script:Config.Paths.watch_folder
            Filter = "*.wmv"
            File = $true
        }
        
        if ($script:Config.Behavior.recursive) {
            $searchParams.Recurse = $true
        }
        
        $wmvFiles = Get-ChildItem @searchParams
        
        if ($wmvFiles.Count -eq 0) {
            Write-Log "No WMV files found" "DEBUG"
            return
        }
        
        Write-Log "Found $($wmvFiles.Count) WMV file(s)" "INFO"
        
        # Process each WMV file
        foreach ($file in $wmvFiles) {
            Process-WMVFile -FilePath $file.FullName
        }
    }
    catch {
        Write-Log "ERROR: Failed to scan for WMV files - $($_.Exception.Message)" "ERROR"
    }
}

# Main execution
Write-Log "=== Enhanced WMV Watchdog Converter Started ===" "INFO"

# Read configuration
$script:Config = Read-IniFile -FilePath $ConfigFile
if (!$script:Config) {
    Write-Error "Failed to read configuration file: $ConfigFile"
    exit 1
}

# Set up logging
$script:LogFile = Join-Path $script:Config.Paths.log_folder "watchdog_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

# Create necessary directories
$directories = @($script:Config.Paths.watch_folder, $script:Config.Paths.output_base_folder, $script:Config.Paths.log_folder)
foreach ($dir in $directories) {
    if (!(Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
        Write-Log "Created directory: $dir" "INFO"
    }
}

# Clean old logs
Remove-OldLogs -LogFolder $script:Config.Paths.log_folder -RetentionDays $script:Config.Logging.log_retention_days

# Check if FFmpeg is available
if (!(Test-FFmpeg)) {
    Write-Log "ERROR: FFmpeg is not installed or not in PATH" "ERROR"
    Write-Log "Please install FFmpeg and ensure it's available in the system PATH" "ERROR"
    exit 1
}

Write-Log "FFmpeg is available" "INFO"
Write-Log "Watch Folder: $($script:Config.Paths.watch_folder)" "INFO"
Write-Log "Output Base Folder: $($script:Config.Paths.output_base_folder)" "INFO"
Write-Log "Scan Interval: $($script:Config.Service.scan_interval) seconds" "INFO"
Write-Log "Log File: $script:LogFile" "INFO"

# Main loop
Write-Log "Starting watchdog loop. Press Ctrl+C to stop." "INFO"
try {
    while ($true) {
        Scan-ForWMVFiles
        
        Write-Log "Waiting $($script:Config.Service.scan_interval) seconds before next scan..." "DEBUG"
        Start-Sleep -Seconds $script:Config.Service.scan_interval
    }
}
catch {
    Write-Log "ERROR: Watchdog loop interrupted - $($_.Exception.Message)" "ERROR"
}
finally {
    Write-Log "=== Enhanced WMV Watchdog Converter Stopped ===" "INFO"
} 