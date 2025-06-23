' Server WMV Converter Service Wrapper - Fixed Version
' For C:\FTPData\SPN\Upload structure

Option Explicit

Dim WshShell, FSO, LogFile, LogPath
Dim ScriptDir, ConverterPath, ConfigPath, ErrorFile

' Get script directory
ScriptDir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)

' Set paths
ConverterPath = ScriptDir & "\watchdog_converter_enhanced.bat"
ConfigPath = ScriptDir & "\watchdog_config_server.ini"
LogPath = "C:\WMVConverter\logs"
ErrorFile = LogPath & "\service_error.log"

' Create log file for service wrapper
Set FSO = CreateObject("Scripting.FileSystemObject")
If Not FSO.FolderExists(LogPath) Then
    On Error Resume Next
    FSO.CreateFolder(LogPath)
    If Err.Number <> 0 Then
        ' Try to create in system temp directory if C:\WMVConverter\logs fails
        LogPath = CreateObject("WScript.Shell").Environment("SYSTEM")("TEMP")
        ErrorFile = LogPath & "\wmvconverter_error.log"
    End If
    On Error Goto 0
End If

LogFile = LogPath & "\server_wrapper_" & Year(Now) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2) & "_" & Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2) & ".log"

' Function to write to log
Sub WriteLog(message)
    On Error Resume Next
    Dim timestamp
    timestamp = Now
    FSO.OpenTextFile(LogFile, 8, True).WriteLine "[" & timestamp & "] " & message
    ' Also write to error file for debugging
    FSO.OpenTextFile(ErrorFile, 8, True).WriteLine "[" & timestamp & "] " & message
    On Error Goto 0
End Sub

' Function to write error
Sub WriteError(message)
    On Error Resume Next
    Dim timestamp
    timestamp = Now
    FSO.OpenTextFile(ErrorFile, 8, True).WriteLine "[" & timestamp & "] ERROR: " & message
    On Error Goto 0
End Sub

' Start logging
WriteLog "=== Server WMV Converter Service Wrapper Started ==="
WriteLog "Script Directory: " & ScriptDir
WriteLog "Converter Path: " & ConverterPath
WriteLog "Config Path: " & ConfigPath
WriteLog "Watch Folder: C:\FTPData\SPN\Upload"
WriteLog "Output Folder: C:\FTPData\SPN\Converted"
WriteLog "Log Path: " & LogPath

' Check if converter exists
If Not FSO.FileExists(ConverterPath) Then
    WriteError "Converter batch file not found: " & ConverterPath
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
    On Error Resume Next
    FSO.CreateFolder("C:\FTPData\SPN\Upload")
    If Err.Number <> 0 Then
        WriteError "Failed to create watch folder: " & Err.Description
    End If
    On Error Goto 0
End If

' Check if output folder exists
If Not FSO.FolderExists("C:\FTPData\SPN\Converted") Then
    WriteLog "WARNING: Output folder does not exist: C:\FTPData\SPN\Converted"
    WriteLog "Creating output folder..."
    On Error Resume Next
    FSO.CreateFolder("C:\FTPData\SPN\Converted")
    If Err.Number <> 0 Then
        WriteError "Failed to create output folder: " & Err.Description
    End If
    On Error Goto 0
End If

' Set working directory
Set WshShell = CreateObject("WScript.Shell")
On Error Resume Next
WshShell.CurrentDirectory = ScriptDir
If Err.Number <> 0 Then
    WriteError "Failed to set working directory: " & Err.Description
    WScript.Quit 1
End If
On Error Goto 0

WriteLog "Set working directory to: " & ScriptDir

' Run the converter with error handling
WriteLog "Starting server WMV converter..."
On Error Resume Next
WshShell.Run "cmd /c " & ConverterPath, 0, False
If Err.Number <> 0 Then
    WriteError "Failed to start converter: " & Err.Description
    WScript.Quit 1
End If
On Error Goto 0

WriteLog "Service wrapper completed successfully" 