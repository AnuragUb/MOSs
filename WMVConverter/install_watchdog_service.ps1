# Install WMV Watchdog Converter as Windows Service
# Run this script as Administrator

param(
    [string]$ServiceName = "WMVWatchdogConverter",
    [string]$DisplayName = "WMV Watchdog Converter Service",
    [string]$Description = "Monitors folders for WMV files and converts them to MP4 while preserving folder structure",
    [string]$WatchFolder = "C:\WatchFolder",
    [string]$OutputBaseFolder = "C:\ConvertedOutput",
    [int]$ScanInterval = 30,
    [string]$LogFolder = "C:\WMVConverter\logs"
)

# Check if running as Administrator
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Error "This script must be run as Administrator"
    exit 1
}

Write-Host "=== WMV Watchdog Converter Service Installer ===" -ForegroundColor Green
Write-Host "Service Name: $ServiceName"
Write-Host "Watch Folder: $WatchFolder"
Write-Host "Output Folder: $OutputBaseFolder"
Write-Host "Scan Interval: $ScanInterval seconds"
Write-Host ""

# Get the script directory
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$WatchdogScript = Join-Path $ScriptDir "watchdog_converter.ps1"

# Check if watchdog script exists
if (!(Test-Path $WatchdogScript)) {
    Write-Error "Watchdog script not found: $WatchdogScript"
    exit 1
}

# Create directories if they don't exist
$directories = @($WatchFolder, $OutputBaseFolder, $LogFolder)
foreach ($dir in $directories) {
    if (!(Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
        Write-Host "Created directory: $dir" -ForegroundColor Yellow
    }
}

# Check if service already exists
$existingService = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
if ($existingService) {
    Write-Host "Service '$ServiceName' already exists. Stopping and removing..." -ForegroundColor Yellow
    
    # Stop the service if it's running
    if ($existingService.Status -eq "Running") {
        Stop-Service -Name $ServiceName -Force
        Start-Sleep -Seconds 5
    }
    
    # Remove the service
    $service = Get-WmiObject -Class Win32_Service -Filter "Name='$ServiceName'"
    $service.Delete()
    Write-Host "Existing service removed." -ForegroundColor Green
}

# Create the service executable wrapper
$ServiceWrapper = @"
@echo off
powershell.exe -ExecutionPolicy Bypass -File "$WatchdogScript" -WatchFolder "$WatchFolder" -OutputBaseFolder "$OutputBaseFolder" -ScanInterval $ScanInterval -LogFolder "$LogFolder" -RunAsService
"@

$WrapperPath = Join-Path $ScriptDir "watchdog_service_wrapper.bat"
$ServiceWrapper | Out-File -FilePath $WrapperPath -Encoding ASCII

# Install the service using sc.exe
Write-Host "Installing service..." -ForegroundColor Yellow
$scCommand = "sc.exe create `"$ServiceName`" binPath= `"$WrapperPath`" start= auto DisplayName= `"$DisplayName`""
$result = Invoke-Expression $scCommand

if ($LASTEXITCODE -eq 0) {
    Write-Host "Service created successfully!" -ForegroundColor Green
    
    # Set service description
    $descriptionCommand = "sc.exe description `"$ServiceName`" `"$Description`""
    Invoke-Expression $descriptionCommand | Out-Null
    
    # Configure service to restart on failure
    $failureCommand = "sc.exe failure `"$ServiceName`" reset= 86400 actions= restart/60000/restart/60000/restart/60000"
    Invoke-Expression $failureCommand | Out-Null
    
    Write-Host ""
    Write-Host "=== Service Installation Complete ===" -ForegroundColor Green
    Write-Host "Service Name: $ServiceName"
    Write-Host "Display Name: $DisplayName"
    Write-Host "Status: Installed (not started)"
    Write-Host ""
    Write-Host "To start the service, run:" -ForegroundColor Cyan
    Write-Host "Start-Service -Name '$ServiceName'" -ForegroundColor White
    Write-Host ""
    Write-Host "To stop the service, run:" -ForegroundColor Cyan
    Write-Host "Stop-Service -Name '$ServiceName'" -ForegroundColor White
    Write-Host ""
    Write-Host "To remove the service, run:" -ForegroundColor Cyan
    Write-Host "Remove-Service -Name '$ServiceName'" -ForegroundColor White
    Write-Host ""
    Write-Host "Log files will be created in: $LogFolder" -ForegroundColor Yellow
    
} else {
    Write-Error "Failed to create service. Exit code: $LASTEXITCODE"
    Write-Host "Command executed: $scCommand"
    exit 1
}

# Clean up wrapper file
Remove-Item $WrapperPath -Force

Write-Host ""
Write-Host "=== Installation Summary ===" -ForegroundColor Green
Write-Host "Watch Folder: $WatchFolder"
Write-Host "Output Folder: $OutputBaseFolder"
Write-Host "Scan Interval: $ScanInterval seconds"
Write-Host "Log Folder: $LogFolder"
Write-Host ""
Write-Host "The service will automatically convert any WMV files placed in the watch folder"
Write-Host "and preserve the folder structure in the output folder." -ForegroundColor Yellow 