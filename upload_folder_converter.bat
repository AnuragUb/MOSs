@echo off
setlocal enabledelayedexpansion

echo Starting Upload Folder WMV to MP4 Converter...
echo Current directory: %CD%

:: Set folders for web app
set "INPUT_FOLDER=uploads"
set "OUTPUT_FOLDER=converted"
set "LOG_FOLDER=logs"

:: Create folders if they don't exist
if not exist "%OUTPUT_FOLDER%" mkdir "%OUTPUT_FOLDER%"
if not exist "%LOG_FOLDER%" mkdir "%LOG_FOLDER%"

:: Set log file
set "LOG_FILE=%LOG_FOLDER%\upload_conversion_%date:~-4,4%%date:~-10,2%%date:~-7,2%_%time:~0,2%%time:~3,2%%time:~6,2%.log"
set "LOG_FILE=%LOG_FILE: =0%"

echo Log file: %LOG_FILE%
echo Input folder: %INPUT_FOLDER%
echo Output folder: %OUTPUT_FOLDER%

:loop
echo.
echo Scanning for WMV files in %INPUT_FOLDER% and subfolders...
set "files_found=0"

:: Find WMV files recursively (including subfolders)
for /r "%INPUT_FOLDER%" %%F in (*.wmv) do (
    set /a files_found+=1
    set "filename=%%~nF"
    set "input_file=%%F"
    
    :: Get relative path from input folder to preserve folder structure
    set "relative_path=%%~dpF"
    set "relative_path=!relative_path:%INPUT_FOLDER%=!"
    set "relative_path=!relative_path:~1!"
    set "relative_path=!relative_path:~0,-1!"
    
    :: Build output path preserving folder structure
    if "!relative_path!"=="" (
        set "output_file=%OUTPUT_FOLDER%\!filename!.mp4"
    ) else (
        set "output_file=%OUTPUT_FOLDER%\!relative_path!\!filename!.mp4"
    )
    
    :: Create output directory if it doesn't exist
    for %%D in ("!output_file!") do set "output_dir=%%~dpD"
    if not exist "!output_dir!" mkdir "!output_dir!"

    echo Found file: !input_file!
    echo Output: !output_file!

    :: Check if output already exists
    if not exist "!output_file!" (
        echo Converting: !input_file! to !output_file!
        ffmpeg -i "!input_file!" -c:v libx264 -crf 23 -preset medium -c:a aac -b:a 128k -movflags +faststart -y "!output_file!"
        if !ERRORLEVEL! equ 0 (
            echo SUCCESS: Converted !input_file!
            echo [%date% %time%] SUCCESS: Converted !input_file! to !output_file! >> "%LOG_FILE%"
        ) else (
            echo ERROR: Failed to convert !input_file!
            echo [%date% %time%] ERROR: Failed to convert !input_file! >> "%LOG_FILE%"
        )
    ) else (
        echo SKIP: !output_file! already exists
        echo [%date% %time%] SKIP: !output_file! already exists >> "%LOG_FILE%"
    )
)

if %files_found% equ 0 (
    echo No WMV files found
    echo [%date% %time%] No WMV files found >> "%LOG_FILE%"
) else (
    echo Processed %files_found% WMV files
    echo [%date% %time%] Processed %files_found% WMV files >> "%LOG_FILE%"
)

echo Waiting 30 seconds before next scan...
echo [%date% %time%] Waiting 30 seconds before next scan... >> "%LOG_FILE%"
timeout /t 30 /nobreak >nul
goto loop 