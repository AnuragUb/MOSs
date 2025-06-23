# WMV Watchdog Converter for Windows Server 2012 R2
# Monitors a folder for WMV files and converts them while preserving folder structure

param(
    [string]$WatchFolder = "C:\WatchFolder",
    [string]$OutputBaseFolder = "C:\ConvertedOutput",
    [int]$ScanInterval = 30,
    [string]$LogFolder = "C:\WMVConverter\logs",
    [switch]$RunAsService = $false
)

# Create log folder if it doesn't exist
if (!(Test-Path $LogFolder)) {
    New-Item -ItemType Directory -Path $LogFolder -Force | Out-Null
}

# Set log file with timestamp
$LogFile = Join-Path $LogFolder "watchdog_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

# Function to write to log
function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] $Message"
    Write-Host $logEntry
    Add-Content -Path $LogFile -Value $logEntry
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

# Function to convert WMV to MP4
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
        
        # FFmpeg conversion parameters optimized for server
        $ffmpegArgs = @(
            "-i", "`"$InputFile`"",
            "-c:v", "libx264",
            "-crf", "23",
            "-preset", "medium",
            "-c:a", "aac",
            "-b:a", "128k",
            "-movflags", "+faststart",
            "-y",
            "`"$OutputFile`""
        )
        
        Write-Log "Converting: $InputFile"
        Write-Log "Output: $OutputFile"
        
        # Run FFmpeg conversion
        $process = Start-Process -FilePath "ffmpeg" -ArgumentList $ffmpegArgs -Wait -PassThru -NoNewWindow
        
        if ($process.ExitCode -eq 0) {
            Write-Log "SUCCESS: Converted $InputFile to $OutputFile"
            return $true
        } else {
            Write-Log "ERROR: Failed to convert $InputFile (Exit code: $($process.ExitCode))"
            return $false
        }
    }
    catch {
        Write-Log "ERROR: Exception during conversion of $InputFile - $($_.Exception.Message)"
        return $false
    }
}

# Function to process a single WMV file
function Process-WMVFile {
    param([string]$FilePath)
    
    try {
        # Get relative path from watch folder
        $relativePath = $FilePath.Substring($WatchFolder.Length).TrimStart('\')
        $relativeDir = Split-Path $relativePath -Parent
        
        # Create output path preserving folder structure
        if ($relativeDir) {
            $outputDir = Join-Path $OutputBaseFolder $relativeDir
        } else {
            $outputDir = $OutputBaseFolder
        }
        
        # Generate output filename
        $fileName = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
        $outputFile = Join-Path $outputDir "$fileName.mp4"
        
        # Check if output file already exists
        if (Test-Path $outputFile) {
            Write-Log "SKIP: $outputFile already exists"
            return
        }
        
        # Convert the file
        $success = Convert-WMVToMP4 -InputFile $FilePath -OutputFile $outputFile
        
        if ($success) {
            # Optionally move or delete the original file after successful conversion
            # Uncomment the next line if you want to delete the original WMV file
            # Remove-Item $FilePath -Force
            Write-Log "Conversion completed successfully"
        }
    }
    catch {
        Write-Log "ERROR: Failed to process $FilePath - $($_.Exception.Message)"
    }
}

# Function to scan for WMV files
function Scan-ForWMVFiles {
    try {
        Write-Log "Scanning for WMV files in: $WatchFolder"
        
        # Find all WMV files recursively
        $wmvFiles = Get-ChildItem -Path $WatchFolder -Filter "*.wmv" -Recurse -File
        
        if ($wmvFiles.Count -eq 0) {
            Write-Log "No WMV files found"
            return
        }
        
        Write-Log "Found $($wmvFiles.Count) WMV file(s)"
        
        # Process each WMV file
        foreach ($file in $wmvFiles) {
            Process-WMVFile -FilePath $file.FullName
        }
    }
    catch {
        Write-Log "ERROR: Failed to scan for WMV files - $($_.Exception.Message)"
    }
}

# Main execution
Write-Log "=== WMV Watchdog Converter Started ==="
Write-Log "Watch Folder: $WatchFolder"
Write-Log "Output Base Folder: $OutputBaseFolder"
Write-Log "Scan Interval: $ScanInterval seconds"
Write-Log "Log File: $LogFile"

# Check if FFmpeg is available
if (!(Test-FFmpeg)) {
    Write-Log "ERROR: FFmpeg is not installed or not in PATH"
    Write-Log "Please install FFmpeg and ensure it's available in the system PATH"
    exit 1
}

Write-Log "FFmpeg is available"

# Create output base folder if it doesn't exist
if (!(Test-Path $OutputBaseFolder)) {
    New-Item -ItemType Directory -Path $OutputBaseFolder -Force | Out-Null
    Write-Log "Created output base folder: $OutputBaseFolder"
}

# Main loop
Write-Log "Starting watchdog loop. Press Ctrl+C to stop."
try {
    while ($true) {
        Scan-ForWMVFiles
        
        Write-Log "Waiting $ScanInterval seconds before next scan..."
        Start-Sleep -Seconds $ScanInterval
    }
}
catch {
    Write-Log "ERROR: Watchdog loop interrupted - $($_.Exception.Message)"
}
finally {
    Write-Log "=== WMV Watchdog Converter Stopped ==="
} 