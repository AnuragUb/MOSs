' Server WMV Converter Service Wrapper
' For C:\FTPData\SPN\Upload structure

Option Explicit

Dim WshShell, FSO, LogFile, LogPath
Dim ScriptDir, ConverterPath, ConfigPath

' Get script directory
ScriptDir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)

' Set paths
ConverterPath = ScriptDir & "\watchdog_converter_enhanced.bat"
ConfigPath = ScriptDir & "\watchdog_config_server.ini"
LogPath = "C:\WMVConverter\logs"

' Create log file for service wrapper
Set FSO = CreateObject("Scripting.FileSystemObject")
If Not FSO.FolderExists(LogPath) Then
    FSO.CreateFolder(LogPath)
End If

LogFile = LogPath & "\server_wrapper_" & Year(Now) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2) & "_" & Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2) & ".log"

' Function to write to log
Sub WriteLog(message)
    Dim timestamp
    timestamp = Now
    FSO.OpenTextFile(LogFile, 8, True).WriteLine "[" & timestamp & "] " & message
End Sub

' Start logging
WriteLog "=== Server WMV Converter Service Wrapper Started ==="
WriteLog "Script Directory: " & ScriptDir
WriteLog "Converter Path: " & ConverterPath
WriteLog "Config Path: " & ConfigPath
WriteLog "Watch Folder: C:\FTPData\SPN\Upload"
WriteLog "Output Folder: C:\FTPData\SPN\Converted"

' Check if converter exists
If Not FSO.FileExists(ConverterPath) Then
    WriteLog "ERROR: Converter batch file not found: " & ConverterPath
    WScript.Quit 1
End If

' Check if config exists
If Not FSO.FileExists(ConfigPath) Then
    WriteLog "WARNING: Server config file not found, using defaults: " & ConfigPath
End If

' Check if watch folder exists
If Not FSO.FolderExists("C:\FTPData\SPN\Upload") Then
    WriteLog "WARNING: Watch folder does not exist: C:\FTPData\SPN\Upload"
    WriteLog "Creating watch folder..."
    FSO.CreateFolder("C:\FTPData\SPN\Upload")
End If

' Check if output folder exists
If Not FSO.FolderExists("C:\FTPData\SPN\Converted") Then
    WriteLog "WARNING: Output folder does not exist: C:\FTPData\SPN\Converted"
    WriteLog "Creating output folder..."
    FSO.CreateFolder("C:\FTPData\SPN\Converted")
End If

' Set working directory
Set WshShell = CreateObject("WScript.Shell")
WshShell.CurrentDirectory = ScriptDir
WriteLog "Set working directory to: " & ScriptDir

' Run the converter
WriteLog "Starting server WMV converter..."
WshShell.Run "cmd /c " & ConverterPath, 0, False

WriteLog "Service wrapper completed" 