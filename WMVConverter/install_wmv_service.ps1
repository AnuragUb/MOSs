# Windows Server 2012 R2 WMV Converter Service Installer
# Run this script as Administrator

param(
    [string]$ServiceName = "WMVConverterService",
    [string]$ScriptPath = ".\convert_wmv_server.bat",
    [string]$Description = "WMV to MP4 Conversion Service"
)

# Check if running as Administrator
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Error "This script requires Administrator privileges. Please run as Administrator."
    exit 1
}

Write-Host "Installing WMV Converter Service..." -ForegroundColor Green

# Check if NSSM is available (Non-Sucking Service Manager)
$nssmPath = "C:\nssm\nssm.exe"
if (-not (Test-Path $nssmPath)) {
    Write-Host "NSSM not found. Installing NSSM..." -ForegroundColor Yellow
    
    # Create NSSM directory
    New-Item -ItemType Directory -Path "C:\nssm" -Force | Out-Null
    
    # Download NSSM (adjust URL for latest version)
    $nssmUrl = "https://nssm.cc/release/nssm-2.24.zip"
    $nssmZip = "$env:TEMP\nssm.zip"
    
    try {
        Invoke-WebRequest -Uri $nssmUrl -OutFile $nssmZip
        Expand-Archive -Path $nssmZip -DestinationPath "$env:TEMP\nssm" -Force
        Copy-Item "$env:TEMP\nssm\nssm-2.24\win64\nssm.exe" -Destination $nssmPath -Force
        Remove-Item $nssmZip -Force
        Remove-Item "$env:TEMP\nssm" -Recurse -Force
        Write-Host "NSSM installed successfully" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to install NSSM. Please download manually from https://nssm.cc/"
        exit 1
    }
}

# Check if service already exists
if (Get-Service -Name $ServiceName -ErrorAction SilentlyContinue) {
    Write-Host "Service '$ServiceName' already exists. Stopping and removing..." -ForegroundColor Yellow
    Stop-Service -Name $ServiceName -Force -ErrorAction SilentlyContinue
    & $nssmPath remove $ServiceName confirm
}

# Get full path to the batch script
$fullScriptPath = (Resolve-Path $ScriptPath).Path

# Install the service using NSSM
Write-Host "Installing service..." -ForegroundColor Green
& $nssmPath install $ServiceName $fullScriptPath
& $nssmPath set $ServiceName DisplayName $ServiceName
& $nssmPath set $ServiceName Description $Description
& $nssmPath set $ServiceName Start SERVICE_AUTO_START
& $nssmPath set $ServiceName AppDirectory (Split-Path $fullScriptPath)
& $nssmPath set $ServiceName AppStdout (Join-Path (Split-Path $fullScriptPath) "logs\service.log")
& $nssmPath set $ServiceName AppStderr (Join-Path (Split-Path $fullScriptPath) "logs\service_error.log")

# Start the service
Write-Host "Starting service..." -ForegroundColor Green
Start-Service -Name $ServiceName

# Verify service is running
Start-Sleep -Seconds 3
$service = Get-Service -Name $ServiceName
if ($service.Status -eq "Running") {
    Write-Host "Service installed and started successfully!" -ForegroundColor Green
    Write-Host "Service Name: $ServiceName" -ForegroundColor Cyan
    Write-Host "Status: $($service.Status)" -ForegroundColor Cyan
    Write-Host "Startup Type: $($service.StartType)" -ForegroundColor Cyan
} else {
    Write-Host "Service installed but failed to start. Status: $($service.Status)" -ForegroundColor Yellow
    Write-Host "Check the logs in the logs folder for more information." -ForegroundColor Yellow
}

Write-Host "`nService Management Commands:" -ForegroundColor Yellow
Write-Host "  Start:   Start-Service -Name '$ServiceName'" -ForegroundColor White
Write-Host "  Stop:    Stop-Service -Name '$ServiceName'" -ForegroundColor White
Write-Host "  Status:  Get-Service -Name '$ServiceName'" -ForegroundColor White
Write-Host "  Remove:  & '$nssmPath' remove '$ServiceName' confirm" -ForegroundColor White 