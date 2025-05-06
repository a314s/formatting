#!/usr/bin/env python
"""
Deployment setup script for Video to PDF Converter
This script creates necessary directories and sets permissions
"""

import os
import sys

def setup_deployment():
    """Create necessary directories and set permissions"""
    print("Setting up deployment environment...")
    
    # Get the absolute path of the application directory
    app_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Define required directories
    required_dirs = [
        'uploads',
        'temp',
        'frames',
        'results'
    ]
    
    # Create directories if they don't exist
    for directory in required_dirs:
        dir_path = os.path.join(app_dir, directory)
        if not os.path.exists(dir_path):
            try:
                os.makedirs(dir_path)
                # Set directory permissions to 755 (rwxr-xr-x)
                os.chmod(dir_path, 0o755)
                print(f"Created directory with permissions: {dir_path}")
            except Exception as e:
                print(f"Error creating directory {dir_path}: {str(e)}")
                sys.exit(1)
        else:
            # Ensure correct permissions even if directory exists
            os.chmod(dir_path, 0o755)
            print(f"Directory exists, updated permissions: {dir_path}")
    
    # Ensure templates directory exists
    templates_dir = os.path.join(app_dir, 'templates')
    if not os.path.exists(templates_dir):
        print("ERROR: Templates directory not found!")
        print(f"Make sure you have the 'templates' directory at: {templates_dir}")
        sys.exit(1)
    else:
        # Set directory permissions to 755 (rwxr-xr-x)
        os.chmod(templates_dir, 0o755)
    
    # Check for required template files
    required_templates = [
        'index.html',
        'settings.html',
        'get_frames.html',
        'convert_pdf.html'
    ]
    
    for template in required_templates:
        template_path = os.path.join(templates_dir, template)
        if not os.path.exists(template_path):
            print(f"ERROR: Required template file not found: {template}")
            print(f"Make sure {template} exists at: {template_path}")
            sys.exit(1)
        else:
            # Set file permissions to 644 (rw-r--r--)
            os.chmod(template_path, 0o644)
    
    # Check for required Python files
    required_files = [
        'final_app.py',
        'processing.py',
        'video_utils.py',
        'wsgi.py'
    ]
    
    for file in required_files:
        file_path = os.path.join(app_dir, file)
        if not os.path.exists(file_path):
            print(f"ERROR: Required file not found: {file_path}")
            sys.exit(1)
        else:
            # Set file permissions to 644 (rw-r--r--)
            os.chmod(file_path, 0o644)
            print(f"Updated permissions for: {file_path}")
    
    # Check for required asset files
    asset_files = [
        'logo.jpg',
        'sideimage.png',
        'Copyright.docx',
        'Template.docx'
    ]
    
    for asset in asset_files:
        asset_path = os.path.join(app_dir, asset)
        if not os.path.exists(asset_path):
            print(f"WARNING: Asset file not found: {asset_path}")
            print(f"Some functionality may not work without this file.")
        else:
            # Set file permissions to 644 (rw-r--r--)
            os.chmod(asset_path, 0o644)
            print(f"Updated permissions for: {asset_path}")
    
    print("\nDeployment setup completed successfully!")
    print("\nDeployment Checklist:")
    print("1. Make sure all required packages are installed:")
    print("   pip install -r requirements.txt")
    print("2. Ensure ffmpeg is installed on the server")
    print("3. For PDF conversion, ensure LibreOffice or unoconv is installed")
    print("4. Restart the web server/application")
    print("5. Check the application logs for any errors")

if __name__ == "__main__":
    setup_deployment()