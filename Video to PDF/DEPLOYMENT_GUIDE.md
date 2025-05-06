# Deployment Guide for PythonAnywhere

This guide will help you deploy the Video to PDF Converter application on PythonAnywhere.

## Step 1: Upload All Files

Ensure all the following files are uploaded to your PythonAnywhere directory:

- `final_app.py` - Main application file
- `processing.py` - Processing functions
- `video_utils.py` - New utilities for frame extraction and PDF conversion
- `wsgi.py` - WSGI configuration file
- `requirements.txt` - Package dependencies
- `deploy_setup.py` - Deployment setup script
- Asset files:
  - `logo.jpg`
  - `sideimage.png`
  - `Copyright.docx`
  - `Template.docx`
- Templates directory with:
  - `templates/index.html`
  - `templates/settings.html`
  - `templates/get_frames.html`
  - `templates/convert_pdf.html`

## Step 2: Set Up Virtual Environment

If you haven't already, create a virtual environment:

```bash
mkvirtualenv --python=/usr/bin/python3.8 myenv
```

Activate the virtual environment:

```bash
workon myenv
```

## Step 3: Install Dependencies

Install the required packages:

```bash
pip install -r requirements.txt
```

## Step 4: Run Deployment Setup Script

Run the deployment setup script to create necessary directories and check for required files:

```bash
python deploy_setup.py
```

## Step 5: Configure Web App

1. Go to the Web tab in your PythonAnywhere dashboard
2. Configure your web app to use the correct WSGI file:
   - Source code: `/path/to/your/app`
   - Working directory: `/path/to/your/app`
   - WSGI configuration file: `/path/to/your/app/wsgi.py`
   - Virtual environment: `/path/to/your/virtualenv`

## Step 6: Install System Dependencies

For frame extraction and PDF conversion, you need to ensure the following system dependencies are available:

1. **FFmpeg** - Required for video processing
   - PythonAnywhere has FFmpeg pre-installed, but you may need to specify the full path in your code

2. **LibreOffice or unoconv** - Required for PDF conversion
   - PythonAnywhere has LibreOffice pre-installed
   - You may need to use the full path: `/usr/bin/libreoffice`

## Step 7: Update File Paths

If you're getting "not found" errors, check the following:

1. Make sure all required directories exist:
   - `uploads`
   - `temp`
   - `frames`
   - `results`

2. Ensure file paths in your code are absolute paths, not relative paths:
   - Update any hardcoded paths to use `os.path.join(app_dir, 'directory_name')`

## Step 8: Check Permissions

Ensure all directories have the correct permissions:

```bash
chmod -R 755 /path/to/your/app
```

## Step 9: Reload Web App

After making all changes, reload your web app from the PythonAnywhere dashboard.

## Step 10: Check Logs

If you're still experiencing issues, check the error logs:

1. Go to the Web tab in your PythonAnywhere dashboard
2. Click on the "Error log" link
3. Look for any error messages that might help identify the issue

## Troubleshooting "Not Found" Errors

If you're getting a "Not Found" error:

1. **Check URL Configuration**: Make sure the URL you're accessing matches a route in your application
2. **Check WSGI Configuration**: Ensure your WSGI file is correctly importing the Flask application
3. **Check File Paths**: Make sure all template files are in the correct location
4. **Check Directory Structure**: Ensure your directory structure matches what your application expects
5. **Check Permissions**: Make sure all files and directories have the correct permissions

## Common PythonAnywhere Issues

1. **Path Issues**: PythonAnywhere uses absolute paths. Make sure all paths in your code are correctly set.
2. **Package Installation**: Some packages with C extensions might not install correctly. Check if all packages installed successfully.
3. **System Dependencies**: Some features might require system dependencies that need special configuration on PythonAnywhere.
4. **File Permissions**: Make sure all files and directories have the correct permissions.
5. **Memory Limits**: PythonAnywhere has memory limits that might affect video processing. Consider processing smaller chunks if you hit memory limits.

If you continue to experience issues, please check the PythonAnywhere forums or contact their support for assistance.