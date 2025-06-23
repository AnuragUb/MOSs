@echo off
setlocal enabledelayedexpansion

echo Starting Upload Folder WMV to MP4 Converter with Priority Processing...
echo Current directory: %CD%

:: Set folders for web app
set "PRIORITY_FOLDER=uploads\priority"
set "INPUT_FOLDER=uploads"
set "OUTPUT_FOLDER=converted"
set "LOG_FOLDER=logs"

:: Create folders if they don't exist
if not exist "%OUTPUT_FOLDER%" mkdir "%OUTPUT_FOLDER%"
if not exist "%LOG_FOLDER%" mkdir "%LOG_FOLDER%"
if not exist "%PRIORITY_FOLDER%" mkdir "%PRIORITY_FOLDER%"

:: Set log file
set "LOG_FILE=%LOG_FOLDER%\upload_conversion_%date:~-4,4%%date:~-10,2%%date:~-7,2%_%time:~0,2%%time:~3,2%%time:~6,2%.log"
set "LOG_FILE=%LOG_FILE: =0%"

echo Log file: %LOG_FILE%
echo Priority folder: %PRIORITY_FOLDER%
echo Input folder: %INPUT_FOLDER%
echo Output folder: %OUTPUT_FOLDER%

:loop
echo.
echo === Starting scan cycle ===
echo [%date% %time%] === Starting scan cycle === >> "%LOG_FILE%"

:: First, process priority folder (urgent files)
echo Processing PRIORITY folder first...
echo [%date% %time%] Processing PRIORITY folder first... >> "%LOG_FILE%"
call :process_folder "%PRIORITY_FOLDER%" "priority" "PRIORITY"

:: Then, process regular uploads folder
echo Processing regular uploads folder...
echo [%date% %time%] Processing regular uploads folder... >> "%LOG_FILE%"
call :process_folder "%INPUT_FOLDER%" "" "REGULAR"

echo === Scan cycle completed ===
echo [%date% %time%] === Scan cycle completed === >> "%LOG_FILE%"
echo Waiting 30 seconds before next scan...
echo [%date% %time%] Waiting 30 seconds before next scan... >> "%LOG_FILE%"
timeout /t 30 /nobreak >nul
goto loop

:: Function to process a specific folder
:process_folder
set "folder_path=%~1"
set "relative_prefix=%~2"
set "folder_type=%~3"
set "files_found=0"

echo Scanning for WMV files in %folder_type% folder: %folder_path%

:: Check if folder exists
if not exist "%folder_path%" (
    echo %folder_type% folder does not exist: %folder_path%
    echo [%date% %time%] %folder_type% folder does not exist: %folder_path% >> "%LOG_FILE%"
    goto :eof
)

:: Find WMV files recursively (including subfolders)
for /r "%folder_path%" %%F in (*.wmv) do (
    set /a files_found+=1
    set "filename=%%~nF"
    set "input_file=%%F"
    
    :: Get relative path from the specific folder to preserve folder structure
    set "relative_path=%%~dpF"
    set "relative_path=!relative_path:%folder_path%=!"
    set "relative_path=!relative_path:~1!"
    set "relative_path=!relative_path:~0,-1!"
    
    :: Build output path preserving folder structure
    if "!relative_path!"=="" (
        if "!relative_prefix!"=="" (
            set "output_file=%OUTPUT_FOLDER%\!filename!.mp4"
        ) else (
            set "output_file=%OUTPUT_FOLDER%\!relative_prefix!\!filename!.mp4"
        )
    ) else (
        if "!relative_prefix!"=="" (
            set "output_file=%OUTPUT_FOLDER%\!relative_path!\!filename!.mp4"
        ) else (
            set "output_file=%OUTPUT_FOLDER%\!relative_prefix!\!relative_path!\!filename!.mp4"
        )
    )
    
    :: Create output directory if it doesn't exist
    for %%D in ("!output_file!") do set "output_dir=%%~dpD"
    if not exist "!output_dir!" mkdir "!output_dir!"

    echo Found %folder_type% file: !input_file!
    echo Output: !output_file!

    :: Check if output already exists
    if not exist "!output_file!" (
        echo Converting %folder_type% file: !input_file! to !output_file!
        ffmpeg -i "!input_file!" -c:v libx264 -crf 23 -preset medium -c:a aac -b:a 128k -movflags +faststart -y "!output_file!"
        if !ERRORLEVEL! equ 0 (
            echo SUCCESS: Converted %folder_type% file !input_file!
            echo [%date% %time%] SUCCESS: Converted %folder_type% file !input_file! to !output_file! >> "%LOG_FILE%"
        ) else (
            echo ERROR: Failed to convert %folder_type% file !input_file!
            echo [%date% %time%] ERROR: Failed to convert %folder_type% file !input_file! >> "%LOG_FILE%"
        )
    ) else (
        echo SKIP: !output_file! already exists
        echo [%date% %time%] SKIP: !output_file! already exists >> "%LOG_FILE%"
    )
)

if %files_found% equ 0 (
    echo No WMV files found in %folder_type% folder
    echo [%date% %time%] No WMV files found in %folder_type% folder >> "%LOG_FILE%"
) else (
    echo Processed %files_found% WMV files from %folder_type% folder
    echo [%date% %time%] Processed %files_found% WMV files from %folder_type% folder >> "%LOG_FILE%"
)
goto :eof 