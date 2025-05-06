# Video to PDF Converter

This web application converts video tutorials into PDF documents with screenshots and text. It extracts audio from videos, transcribes it using Google Cloud Speech-to-Text, aligns the transcription with a provided script, extracts frames from the video, and generates a PDF document.

## Features

- Drag and drop interface for uploading video and Excel files
- Audio extraction from video using FFmpeg
- Speech recognition using Google Cloud Speech-to-Text
- Frame extraction from videos at key points
- PDF generation with screenshots and text
- Automatic cleanup of temporary files

## Requirements

- Python 3.7+
- FFmpeg (for audio extraction)
- Google Cloud account with Speech-to-Text API enabled
- Google Cloud Storage bucket
- Google Cloud service account with appropriate permissions

## Installation

1. Clone this repository or upload the files to your PythonAnywhere account

2. Install the required Python packages:
   ```
   pip install -r requirements.txt
   ```

3. Make sure FFmpeg is installed. On PythonAnywhere, it should be available by default.

4. Set up your Google Cloud credentials:
   - Create a project in Google Cloud Console
   - Enable the Speech-to-Text API and Cloud Storage API
   - Create a service account with appropriate permissions
   - Download the service account key as JSON
   - Upload this JSON file through the settings page of the application

## Deploying to PythonAnywhere

1. Sign up for a PythonAnywhere account if you don't have one

2. Go to the "Web" tab and create a new web app:
   - Select "Flask" as the framework
   - Choose Python 3.8 or newer

3. Set up your virtual environment:
   ```
   mkvirtualenv --python=/usr/bin/python3.8 myenv
   pip install -r requirements.txt
   ```

4. Configure your web app:
   - Set the source code directory to where you uploaded the files
   - Set the WSGI configuration file to point to your app.py
   - Add the following to the WSGI configuration file:
     ```python
     import sys
     path = '/home/yourusername/your-app-directory'
     if path not in sys.path:
         sys.path.append(path)
     
     from app import app as application
     ```

5. Create the required directories:
   ```
   mkdir uploads
   mkdir temp
   mkdir frames
   mkdir results
   ```

6. Set appropriate permissions:
   ```
   chmod 755 uploads temp frames results
   ```

7. Reload your web app from the PythonAnywhere dashboard

## Usage

1. Open the web application in your browser

2. Go to the Settings page and configure your Google Cloud settings:
   - Enter your Google Cloud Storage bucket name
   - Upload your Google Cloud service account key file

3. On the main page:
   - Drag and drop your video file or click to select
   - Drag and drop your Excel file with the script or click to select
   - Click "Process Files" to start the conversion

4. Wait for the processing to complete (this may take some time depending on the video length)

5. Download the generated PDF file

6. Optionally, click "Delete Temporary Files" to clean up

## File Structure

- `app.py`: Main Flask application
- `processing.py`: Core processing functions
- `templates/`: HTML templates
  - `index.html`: Main page with drag and drop interface
  - `settings.html`: Settings page for Google Cloud configuration
- `static/`: Static assets (CSS, JavaScript, images)
- `uploads/`: Uploaded files
- `temp/`: Temporary files during processing
- `frames/`: Extracted video frames
- `results/`: Generated PDF and DOCX files

## Notes

- The application requires significant processing power and may take time to process large videos
- Make sure your PythonAnywhere account has enough storage space for the uploaded videos and generated files
- The free tier of PythonAnywhere may have limitations on CPU usage and execution time
- Consider using a paid account for processing large videos or frequent usage