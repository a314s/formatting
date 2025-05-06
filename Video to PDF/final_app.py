import os
import uuid
import json
import tempfile
import shutil
from datetime import datetime
from flask import Flask, request, render_template, redirect, url_for, send_file, jsonify, session
from werkzeug.utils import secure_filename

# Import processing functions
from processing import (
    extract_audio, get_audio_properties, upload_to_gcs, transcribe_audio_gcs,
    load_script, align_script_with_transcription, extract_frames,
    extract_additional_frames, assign_steps_to_frames, generate_word_document
)

# Import new video utilities
from video_utils import process_video_to_frames, convert_docx_to_pdf

# Add the missing extract_doc_name function if it's not in processing.py
def extract_doc_name(script_lines):
    if not script_lines:
        return "Document"
    first_line = script_lines[0].strip()
    # Remove "welcome to the " at the start (case-insensitive)
    prefix = "welcome to the "
    low_line = first_line.lower()
    if (low_line.startswith(prefix)):
        first_line = first_line[len(prefix):].lstrip()

    # Find first period
    first_period_idx = first_line.find('.')
    if (first_period_idx == -1):
        # No period, use entire line
        return first_line
    # Check character after the first period
    if (first_period_idx + 1 < len(first_line)):
        next_char = first_line[first_period_idx + 1]
        if (next_char.isdigit()):
            # Find second period
            second_period_idx = first_line.find('.', first_period_idx + 1)
            if (second_period_idx == -1):
                # no second period found, use entire line
                return first_line
            return first_line[:second_period_idx+1].strip()
    # If no digit after first period or no second period needed
    return first_line[:first_period_idx+1].strip()

def process_files(video_path, excel_path, credential_path, bucket_name, frames_folder, temp_folder, job_id):
    """Process the uploaded files and generate the Word document (no PDF)"""
    # Extract base name from Excel file (without extension)
    excel_base_name = os.path.splitext(os.path.basename(excel_path))[0]
    audio_file = os.path.join(temp_folder, f"{excel_base_name}.wav")
    
    # Step 1: Extract audio from the video
    extract_audio(video_path, audio_file)
    
    # Get audio properties
    properties = get_audio_properties(audio_file)
    sample_rate_hertz = properties.get('sample_rate', 16000)
    audio_channel_count = properties.get('channels', 1)
    
    # Step 2: Upload audio file to Google Cloud Storage
    gcs_uri = upload_to_gcs(audio_file, bucket_name, credential_path)
    
    # Step 3: Transcribe the audio using Google Cloud Speech-to-Text
    transcription_words = transcribe_audio_gcs(gcs_uri, credential_path, sample_rate_hertz, audio_channel_count)
    
    # Step 4: Load the original script from Excel
    script_lines, df = load_script(excel_path)
    
    # Use our local extract_doc_name function
    doc_name = extract_doc_name(script_lines)
    
    # Step 5: Align the script with the transcription to assign timestamps
    script_alignment = align_script_with_transcription(script_lines, transcription_words)
    
    # Adjust lengths to match
    min_length = min(len(df), len(script_alignment))
    df = df.iloc[:min_length].reset_index(drop=True)
    script_alignment = script_alignment[:min_length]
    
    # Step 6: Extract frames from the video
    frame_data = extract_frames(video_path, frames_folder)
    
    # Step 7: Extract additional frames at script end times
    frame_data = extract_additional_frames(video_path, script_alignment, frame_data, frames_folder)
    
    # Step 8: Assign steps to frames based on the script alignment
    frame_step_mapping = assign_steps_to_frames(frame_data, script_alignment)
    
    # Step 9: Generate Word document only (no PDF conversion)
    results_folder = os.path.join(os.path.dirname(frames_folder), 'results')
    os.makedirs(results_folder, exist_ok=True)
    
    # Copy required assets to the results folder
    # Get the directory containing this script (final_app.py)
    script_dir = os.path.dirname(os.path.abspath(__file__))
    # Assets are in the same directory as the script
    video_dir = script_dir
    assets = {
        'logo.jpg': 'logo.jpg',
        'sideimage.png': 'sideimage.png.png',  # Fix double extension
        'Copyright.docx': 'Copyright.docx',
        'Template.docx': 'Template.docx'
    }
    
    for dest_name, src_name in assets.items():
        src = os.path.join(video_dir, src_name)
        dst = os.path.join(results_folder, dest_name)
        if os.path.exists(src):
            shutil.copy(src, dst)
        else:
            raise ValueError(f"Required asset not found: {src}")
    
    # Generate output filenames - change the suffix from -PDF to -DOCX
    docx_name = f"{os.path.splitext(os.path.basename(video_path))[0]}-DOCX.docx"
    docx_path = os.path.join(results_folder, docx_name)
    
    # Generate the document (Word document only, no PDF conversion)
    docx_path = generate_word_document(frame_step_mapping, video_path, doc_name, docx_path, results_folder)
    
    return {
        'docx_filename': docx_name,
        'job_id': job_id
    }

# Flask app code removed since we're using this as a module