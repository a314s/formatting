# Quick Fix Guide for "Not Found" Errors

If you're experiencing "Not Found" errors when deploying the application, follow these steps to quickly resolve the issue:

## Step 1: Run the Deployment Setup Script

```bash
python deploy_setup.py
```

This script will:
- Create all necessary directories
- Check for required files
- Verify that templates are in the correct location

## Step 2: Check Directory Structure

Ensure your directory structure looks like this:

```
/your_app_directory/
├── final_app.py
├── processing.py
├── video_utils.py
├── wsgi.py
├── requirements.txt
├── deploy_setup.py
├── test_routes.py
├── logo.jpg
├── sideimage.png
├── Copyright.docx
├── Template.docx
├── uploads/           (directory)
├── temp/              (directory)
├── frames/            (directory)
├── results/           (directory)
└── templates/
    ├── index.html
    ├── settings.html
    ├── get_frames.html
    └── convert_pdf.html
```

## Step 3: Verify WSGI Configuration

Make sure your `wsgi.py` file correctly imports the Flask application:

```python
import sys
import os

# Add the application directory to the Python path
app_dir = os.path.dirname(os.path.abspath(__file__))
if app_dir not in sys.path:
    sys.path.append(app_dir)

# Import the Flask application
from final_app import app as application
```

## Step 4: Test Routes

Run the test script to check if all routes are working:

```bash
python test_routes.py YOUR_APP_URL
```

Replace `YOUR_APP_URL` with your actual application URL (e.g., `http://yourusername.pythonanywhere.com`).

## Step 5: Reload Web App

After making changes, reload your web app from the PythonAnywhere dashboard:

1. Go to the Web tab
2. Click the "Reload" button for your web app

## Step 6: Check Logs

If you're still experiencing issues, check the error logs:

1. Go to the Web tab in your PythonAnywhere dashboard
2. Click on the "Error log" link
3. Look for any error messages

## Common Issues and Solutions

### Issue: Templates Not Found

**Solution:** Make sure all template files are in the `templates` directory and that Flask can find them. We've updated `final_app.py` to use absolute paths for the template folder.

### Issue: Static Files Not Found

**Solution:** Make sure all static files (CSS, JS, images) are in the correct location and that Flask can find them.

### Issue: Directory Permissions

**Solution:** Make sure all directories have the correct permissions:

```bash
chmod -R 755 /path/to/your/app
```

### Issue: Package Dependencies

**Solution:** Make sure all required packages are installed:

```bash
pip install -r requirements.txt
```

### Issue: System Dependencies

**Solution:** Make sure system dependencies like FFmpeg and LibreOffice are available. We've updated the code to use the full path to LibreOffice on PythonAnywhere.

## Need More Help?

For more detailed deployment instructions, refer to the `DEPLOYMENT_GUIDE.md` file.