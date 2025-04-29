@echo off
echo Checking Python installation...

python --version > nul 2>&1
if errorlevel 1 (
    echo Python is not installed or not in PATH
    echo Please install Python 3.8 or later from https://www.python.org/downloads/
    pause
    exit /b 1
)

echo Installing base packages...
python -m pip install --upgrade pip
python -m pip install openpyxl

echo Installing core dependencies...
python -m pip install Flask==2.3.3 Werkzeug==2.3.7 python-docx==0.8.11

echo Installing data processing packages...
python -m pip install pandas==2.0.3 opencv-python==4.8.0.76

echo Installing image processing packages...
python -m pip install pillow-simd imagehash==4.3.1

echo Installing Google Cloud packages...
python -m pip install google-cloud-speech==2.21.0 google-cloud-storage==2.10.0

echo Installing additional utilities...
python -m pip install gunicorn==21.2.0 ffmpeg-python==0.2.0

echo Installing live Excel creation dependencies...
python -m pip install keyboard==0.13.5 pyperclip==1.8.2 websockets==12.0

echo Creating required directories...
mkdir uploads 2>nul
mkdir "Video to PDF\uploads" 2>nul

echo Checking Google Cloud credentials...
if not exist "uploads\sapheb-b87c6918d4ef.json" (
    echo WARNING: Google Cloud credentials file not found at uploads\sapheb-b87c6918d4ef.json
    echo The Video to PDF functionality will not work without valid credentials
    echo Please place your Google Cloud credentials file at this location
    pause
)

echo Starting server...
python server.py

if errorlevel 1 (
    echo Server failed to start. Please check the error message above.
    pause
    exit /b 1
)

pause