# WMV Watchdog Converter Setup Script
# Automates the complete installation process for Windows Server 2012 R2

param(
    [string]$WatchFolder = "C:\WatchFolder",
    [string]$OutputBaseFolder = "C:\ConvertedOutput",
    [int]$ScanInterval = 30,
    [switch]$SkipFFmpegInstall = $false,
    [switch]$SkipServiceInstall = $false
)

# Check if running as Administrator
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Error "This script must be run as Administrator"
    Write-Host "Please right-click PowerShell and select 'Run as Administrator'" -ForegroundColor Red
    exit 1
}

Write-Host "=== WMV Watchdog Converter Setup ===" -ForegroundColor Green
Write-Host "This script will install and configure the WMV Watchdog Converter" -ForegroundColor Yellow
Write-Host ""

# Get script directory
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# Function to write colored output
function Write-Status {
    param([string]$Message, [string]$Status = "INFO")
    
    $timestamp = Get-Date -Format "HH:mm:ss"
    switch ($Status) {
        "SUCCESS" { Write-Host "[$timestamp] ✓ $Message" -ForegroundColor Green }
        "ERROR" { Write-Host "[$timestamp] ✗ $Message" -ForegroundColor Red }
        "WARNING" { Write-Host "[$timestamp] ⚠ $Message" -ForegroundColor Yellow }
        default { Write-Host "[$timestamp] ℹ $Message" -ForegroundColor Cyan }
    }
}

# Step 1: Check PowerShell version
Write-Status "Checking PowerShell version..."
$psVersion = $PSVersionTable.PSVersion.Major
if ($psVersion -lt 4) {
    Write-Status "ERROR: PowerShell 4.0 or later is required. Current version: $psVersion" "ERROR"
    exit 1
}
Write-Status "PowerShell version $psVersion is compatible" "SUCCESS"

# Step 2: Install FFmpeg (if not skipped)
if (-not $SkipFFmpegInstall) {
    Write-Status "Checking FFmpeg installation..."
    
    # Check if FFmpeg is already installed
    $ffmpegInstalled = $false
    try {
        $null = Get-Command ffmpeg -ErrorAction Stop
        $ffmpegInstalled = $true
        Write-Status "FFmpeg is already installed" "SUCCESS"
    }
    catch {
        Write-Status "FFmpeg not found, installing..." "WARNING"
        
        # Check if Chocolatey is installed
        $chocoInstalled = $false
        try {
            $null = Get-Command choco -ErrorAction Stop
            $chocoInstalled = $true
            Write-Status "Chocolatey is already installed" "SUCCESS"
        }
        catch {
            Write-Status "Installing Chocolatey package manager..." "INFO"
            
            # Install Chocolatey
            Set-ExecutionPolicy Bypass -Scope Process -Force
            [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
            iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))
            
            if ($LASTEXITCODE -eq 0) {
                Write-Status "Chocolatey installed successfully" "SUCCESS"
                $chocoInstalled = $true
                
                # Refresh environment variables
                refreshenv
            } else {
                Write-Status "Failed to install Chocolatey" "ERROR"
                exit 1
            }
        }
        
        if ($chocoInstalled) {
            Write-Status "Installing FFmpeg using Chocolatey..." "INFO"
            choco install ffmpeg -y
            
            if ($LASTEXITCODE -eq 0) {
                Write-Status "FFmpeg installed successfully" "SUCCESS"
                $ffmpegInstalled = $true
                
                # Refresh environment variables
                refreshenv
            } else {
                Write-Status "Failed to install FFmpeg via Chocolatey" "ERROR"
                Write-Status "Please install FFmpeg manually from https://ffmpeg.org/download.html" "WARNING"
                exit 1
            }
        }
    }
    
    # Verify FFmpeg installation
    if ($ffmpegInstalled) {
        try {
            $ffmpegVersion = ffmpeg -version | Select-String "ffmpeg version" | Select-Object -First 1
            Write-Status "FFmpeg version: $ffmpegVersion" "SUCCESS"
        }
        catch {
            Write-Status "FFmpeg installation verification failed" "ERROR"
            exit 1
        }
    }
} else {
    Write-Status "Skipping FFmpeg installation as requested" "WARNING"
}

# Step 3: Create directories
Write-Status "Creating necessary directories..." "INFO"

$directories = @($WatchFolder, $OutputBaseFolder, "C:\WMVConverter\logs")
foreach ($dir in $directories) {
    if (!(Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
        Write-Status "Created directory: $dir" "SUCCESS"
    } else {
        Write-Status "Directory already exists: $dir" "INFO"
    }
}

# Step 4: Update configuration file
Write-Status "Updating configuration file..." "INFO"

$configFile = Join-Path $ScriptDir "watchdog_config.ini"
if (Test-Path $configFile) {
    # Read current config
    $configContent = Get-Content $configFile
    
    # Update paths in config
    $updatedContent = $configContent | ForEach-Object {
        $line = $_
        if ($line -match "^watch_folder=") {
            "watch_folder=$WatchFolder"
        } elseif ($line -match "^output_base_folder=") {
            "output_base_folder=$OutputBaseFolder"
        } elseif ($line -match "^scan_interval=") {
            "scan_interval=$ScanInterval"
        } else {
            $line
        }
    }
    
    # Write updated config
    $updatedContent | Out-File -FilePath $configFile -Encoding UTF8
    Write-Status "Configuration updated with custom paths" "SUCCESS"
} else {
    Write-Status "Configuration file not found: $configFile" "ERROR"
    exit 1
}

# Step 5: Install Windows Service (if not skipped)
if (-not $SkipServiceInstall) {
    Write-Status "Installing Windows service..." "INFO"
    
    $serviceInstaller = Join-Path $ScriptDir "install_watchdog_service.ps1"
    if (Test-Path $serviceInstaller) {
        # Run the service installer with our parameters
        & $serviceInstaller -WatchFolder $WatchFolder -OutputBaseFolder $OutputBaseFolder -ScanInterval $ScanInterval
        
        if ($LASTEXITCODE -eq 0) {
            Write-Status "Windows service installed successfully" "SUCCESS"
        } else {
            Write-Status "Failed to install Windows service" "ERROR"
            exit 1
        }
    } else {
        Write-Status "Service installer not found: $serviceInstaller" "ERROR"
        exit 1
    }
} else {
    Write-Status "Skipping Windows service installation as requested" "WARNING"
}

# Step 6: Test the setup
Write-Status "Testing the setup..." "INFO"

# Test FFmpeg
try {
    $null = Get-Command ffmpeg -ErrorAction Stop
    Write-Status "FFmpeg test: PASSED" "SUCCESS"
} catch {
    Write-Status "FFmpeg test: FAILED" "ERROR"
}

# Test directory access
try {
    $testFile = Join-Path $WatchFolder "test.txt"
    "test" | Out-File -FilePath $testFile -Encoding ASCII
    Remove-Item $testFile -Force
    Write-Status "Directory access test: PASSED" "SUCCESS"
} catch {
    Write-Status "Directory access test: FAILED" "ERROR"
}

# Test configuration file
if (Test-Path $configFile) {
    Write-Status "Configuration file test: PASSED" "SUCCESS"
} else {
    Write-Status "Configuration file test: FAILED" "ERROR"
}

# Step 7: Final instructions
Write-Host ""
Write-Host "=== Setup Complete ===" -ForegroundColor Green
Write-Host ""
Write-Host "Watchdog Converter has been installed with the following configuration:" -ForegroundColor Yellow
Write-Host "  Watch Folder: $WatchFolder" -ForegroundColor White
Write-Host "  Output Folder: $OutputBaseFolder" -ForegroundColor White
Write-Host "  Scan Interval: $ScanInterval seconds" -ForegroundColor White
Write-Host "  Log Folder: C:\WMVConverter\logs" -ForegroundColor White
Write-Host ""

if (-not $SkipServiceInstall) {
    Write-Host "To start the service, run:" -ForegroundColor Cyan
    Write-Host "  Start-Service -Name 'WMVWatchdogConverter'" -ForegroundColor White
    Write-Host ""
    Write-Host "To check service status:" -ForegroundColor Cyan
    Write-Host "  Get-Service -Name 'WMVWatchdogConverter'" -ForegroundColor White
    Write-Host ""
    Write-Host "To view logs:" -ForegroundColor Cyan
    Write-Host "  Get-Content 'C:\WMVConverter\logs\watchdog_*.log' | Select-Object -Last 50" -ForegroundColor White
} else {
    Write-Host "To run the watchdog manually:" -ForegroundColor Cyan
    Write-Host "  .\watchdog_converter_enhanced.ps1 -ConfigFile 'watchdog_config.ini'" -ForegroundColor White
}

Write-Host ""
Write-Host "The watchdog will automatically convert any WMV files placed in the watch folder" -ForegroundColor Yellow
Write-Host "and preserve the folder structure in the output folder." -ForegroundColor Yellow
Write-Host ""
Write-Host "For more information, see: WATCHDOG_README.md" -ForegroundColor Cyan 