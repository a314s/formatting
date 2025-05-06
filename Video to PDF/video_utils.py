import os
import cv2
import zipfile
import tempfile
import shutil
import subprocess
from flask import send_file
from werkzeug.utils import secure_filename

def extract_frames_per_second(video_file, output_folder):
    """
    Extract frames from a video at 1 frame per second rate
    
    Args:
        video_file (str): Path to the video file
        output_folder (str): Path to save the extracted frames
        
    Returns:
        int: Number of frames extracted
    """
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    # Open the video file
    cap = cv2.VideoCapture(video_file)
    
    # Get video properties
    fps = cap.get(cv2.CAP_PROP_FPS)
    frame_count = int(cap.get(cv2.CAP_PROP_FRAME_COUNT))
    duration = frame_count / fps
    
    # Calculate frames to extract (1 per second)
    frames_to_extract = int(duration)
    extracted_count = 0
    
    for second in range(frames_to_extract):
        # Set position to the exact second
        cap.set(cv2.CAP_PROP_POS_MSEC, second * 1000)
        
        # Read the frame
        ret, frame = cap.read()
        if not ret:
            break
        
        # Save the frame
        frame_filename = os.path.join(output_folder, f'frame_{second:04d}.jpg')
        cv2.imwrite(frame_filename, frame)
        extracted_count += 1
    
    # Release the video capture object
    cap.release()
    
    return extracted_count

def create_frames_zip(frames_folder, zip_filename):
    """
    Create a zip file containing all frames from the frames folder
    
    Args:
        frames_folder (str): Path to the folder containing frames
        zip_filename (str): Path to save the zip file
        
    Returns:
        str: Path to the created zip file
    """
    with zipfile.ZipFile(zip_filename, 'w') as zipf:
        for root, dirs, files in os.walk(frames_folder):
            for file in files:
                file_path = os.path.join(root, file)
                # Add file to zip with relative path
                arcname = os.path.relpath(file_path, os.path.dirname(frames_folder))
                zipf.write(file_path, arcname)
    
    return zip_filename

def convert_docx_to_pdf(docx_file, pdf_file=None):
    """
    Convert a Word document to PDF
    
    Args:
        docx_file (str): Path to the Word document
        pdf_file (str, optional): Path to save the PDF file. If None, will use the same name as docx_file but with .pdf extension
        
    Returns:
        str: Path to the created PDF file
    """
    if pdf_file is None:
        pdf_file = os.path.splitext(docx_file)[0] + '.pdf'
    
    try:
        # Try using LibreOffice (for Linux/PythonAnywhere)
        # Use full path to LibreOffice on PythonAnywhere
        libreoffice_path = '/usr/bin/libreoffice'
        if not os.path.exists(libreoffice_path):
            libreoffice_path = 'libreoffice'  # Fall back to PATH lookup
            
        subprocess.run([
            libreoffice_path, '--headless', '--convert-to', 'pdf',
            '--outdir', os.path.dirname(pdf_file), docx_file
        ], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        
        # LibreOffice creates the PDF with the same basename as the input file
        # If the output path is different, we need to rename it
        libreoffice_pdf = os.path.join(
            os.path.dirname(pdf_file),
            os.path.basename(os.path.splitext(docx_file)[0] + '.pdf')
        )
        
        if libreoffice_pdf != pdf_file and os.path.exists(libreoffice_pdf):
            shutil.move(libreoffice_pdf, pdf_file)
            
        return pdf_file
        
    except (subprocess.SubprocessError, FileNotFoundError):
        # Try using unoconv as fallback
        try:
            subprocess.run(['unoconv', '-f', 'pdf', '-o', pdf_file, docx_file], 
                          check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            return pdf_file
        except (subprocess.SubprocessError, FileNotFoundError):
            # If on Windows, try using win32com
            try:
                import win32com.client
                import time
                
                word = win32com.client.DispatchEx("Word.Application")
                word.Visible = False
                word.DisplayAlerts = False
                
                # Get absolute paths
                docx_path = os.path.abspath(docx_file)
                
                # Ensure the source file exists
                if not os.path.exists(docx_path):
                    raise FileNotFoundError(f"Word document not found: {docx_path}")
                    
                # Open document with read-only and no alert flags
                doc = word.Documents.Open(
                    docx_path,
                    ReadOnly=True,
                    Visible=False,
                    ConfirmConversions=False
                )
                
                # Wait a moment for document to fully load
                time.sleep(2)
                
                # Save as PDF
                try:
                    doc.ExportAsFixedFormat(
                        OutputFileName=pdf_file,
                        ExportFormat=17,  # wdExportFormatPDF
                        OpenAfterExport=False,
                        OptimizeFor=0,    # wdExportOptimizeForPrint
                        Range=0,          # wdExportAllDocument
                        IncludeDocProps=True,
                        KeepIRM=True, 
                        DocStructureTags=True,
                        BitmapMissingFonts=True,
                        UseISO19005_1=False
                    )
                except Exception as save_error:
                    # Alternative method if ExportAsFixedFormat fails
                    doc.SaveAs(pdf_file, FileFormat=17)
                
                # Properly close the document
                doc.Close(SaveChanges=False)
                
                # Clean up
                word.Quit()
                
                return pdf_file
                
            except ImportError:
                raise Exception("PDF conversion failed. Neither LibreOffice, unoconv, nor win32com are available.")
    
    return pdf_file

def process_video_to_frames(video_file, output_dir):
    """
    Process a video file to extract frames at 1 frame per second and create a zip file
    
    Args:
        video_file (str): Path to the video file
        output_dir (str): Directory to save temporary files and the final zip
        
    Returns:
        tuple: (zip_file_path, frame_count)
    """
    # Create a temporary directory for frames
    frames_dir = os.path.join(output_dir, 'frames')
    os.makedirs(frames_dir, exist_ok=True)
    
    # Extract frames
    frame_count = extract_frames_per_second(video_file, frames_dir)
    
    # Create zip file
    video_name = os.path.splitext(os.path.basename(video_file))[0]
    zip_filename = os.path.join(output_dir, f"{video_name}_frames.zip")
    
    create_frames_zip(frames_dir, zip_filename)
    
    # Clean up frames directory
    shutil.rmtree(frames_dir)
    
    return zip_filename, frame_count