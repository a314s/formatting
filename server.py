import http.server
import socketserver
import os
import json
import sys
import tempfile
import shutil
from urllib.parse import parse_qs, urlparse
import mimetypes
import threading
from pathlib import Path
import io
import email
import websockets
import asyncio
import keyboard
import pyperclip
import openpyxl
from openpyxl import Workbook
import time
from datetime import datetime
import win32com.client
import pythoncom

# Import functions from the formatting script
from formatting import clean_text, remove_blank_rows, process_excel

# Import functions from Video to PDF project
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'Video to PDF'))
from final_app import process_files

def handle_error(e):
    """Simple error handler that returns a user-friendly error message"""
    error_msg = str(e)
    if "google.cloud" in error_msg.lower():
        return "Error with Google Cloud services. Please check your credentials."
    elif "ffmpeg" in error_msg.lower():
        return "Error processing video. Please check if the video file is valid."
    elif "excel" in error_msg.lower():
        return "Error processing Excel file. Please check if the file is valid."
    else:
        return f"An error occurred: {error_msg}"

# Define the ports
HTTP_PORT = 8001
WS_PORT = 8002

# Get the directory of the current script
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(SCRIPT_DIR, 'uploads')
os.makedirs(UPLOAD_DIR, exist_ok=True)

# Store active Excel sessions
excel_sessions = {}

class ExcelSession:
    def __init__(self, filename, save_location):
        self.filename = filename
        self.save_location = save_location
        self.workbook = Workbook()
        self.sheet = self.workbook.active
        self.row = 1
        self.websocket = None
        self.monitoring = False

    async def add_text(self, text):
        # Add the text to Excel
        self.sheet.cell(row=self.row, column=1).value = text
        
        # Add a blank row
        self.row += 1
        self.sheet.cell(row=self.row, column=1).value = ''
        
        # Move to next row
        self.row += 1

        # Send update to client
        if self.websocket:
            await self.websocket.send(json.dumps({
                'type': 'update',
                'text': text,
                'row': self.row - 2
            }))

    def save(self):
        save_path = os.path.join(self.save_location, f"{self.filename}.xlsx")
        self.workbook.save(save_path)
        return save_path

class MultipartFormParser:
    def __init__(self, content_type, content_length, rfile):
        self.content_type = content_type
        self.content_length = content_length
        self.rfile = rfile
        
    def parse(self):
        boundary = self.content_type.split('boundary=')[1].encode()
        remainbytes = self.content_length
        
        # Read until first boundary
        while True:
            line = self.rfile.readline()
            remainbytes -= len(line)
            if boundary in line:
                break
        
        fields = {}
        while remainbytes > 0:
            # Parse headers
            headers = {}
            while True:
                line = self.rfile.readline()
                remainbytes -= len(line)
                if line == b'\r\n':
                    break
                    
                # Parse header line
                line = line.decode('utf-8').strip()
                if ':' in line:
                    key, value = line.split(':', 1)
                    headers[key.strip().lower()] = value.strip()
            
            # Get field name from Content-Disposition
            if 'content-disposition' in headers:
                disposition = headers['content-disposition']
                items = disposition.split(';')
                for item in items:
                    if 'name=' in item:
                        name = item.split('=')[1].strip('"')
                        break
            else:
                continue
            
            # Read content until boundary
            content = io.BytesIO()
            while remainbytes > 0:
                line = self.rfile.readline()
                remainbytes -= len(line)
                if boundary in line:
                    break
                content.write(line)
            
            # Store field value
            value = content.getvalue().strip(b'\r\n')
            fields[name.encode()] = value
            
            if remainbytes <= 0:
                break
        
        return fields

class MultiToolHandler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=SCRIPT_DIR, **kwargs)
    
    def do_GET(self):
        parsed_path = urlparse(self.path)
        path = parsed_path.path
        
        if path == '/':
            self.path = '/index.html'
        elif path == '/shutdown':
            self.send_response(200)
            self.send_header('Content-type', 'text/html')
            self.end_headers()
            self.wfile.write(b'Server shutting down...')
            threading.Thread(target=self.server.shutdown).start()
            return
        elif path == '/get_frames':
            self.redirect_to('/templates/get_frames.html')
            return
        elif path == '/convert_to_pdf':
            self.redirect_to('/templates/convert_pdf.html')
            return
        elif path == '/settings':
            self.redirect_to('/templates/settings.html')
            return
        elif path.startswith('/ws-port'):
            self.send_response(200)
            self.send_header('Content-type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'port': WS_PORT}).encode())
            return
        elif path == '/api/history':
            self.send_response(200)
            self.send_header('Content-type', 'application/json')
            self.end_headers()
            history = self.get_upload_history()
            self.wfile.write(json.dumps(history).encode())
            return
        elif path.startswith('/download/'):
            # Extract job_id and filename from path
            parts = path.split('/')
            if len(parts) >= 4:
                job_id = parts[2]
                filename = parts[3]
                job_dir = os.path.join(UPLOAD_DIR, job_id)
                results_dir = os.path.join(job_dir, 'results')
                file_path = os.path.join(results_dir, filename)
                
                if os.path.exists(file_path):
                    with open(file_path, 'rb') as f:
                        self.send_response(200)
                        if filename.endswith('.pdf'):
                            self.send_header('Content-type', 'application/pdf')
                        elif filename.endswith('.docx'):
                            self.send_header('Content-type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
                        self.send_header('Content-Disposition', f'attachment; filename="{filename}"')
                        self.end_headers()
                        shutil.copyfileobj(f, self.wfile)
                    return
        elif os.path.exists(path[1:]):  # Remove leading slash
            with open(path[1:], 'rb') as f:
                self.send_response(200)
                if path.endswith('.xlsx'):
                    self.send_header('Content-type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                self.end_headers()
                shutil.copyfileobj(f, self.wfile)
            return
        
        return super().do_GET()
    
    def do_POST(self):
        if self.path == '/api/format-excel':
            self.handle_format_excel()
        elif self.path == '/upload':
            self.handle_video_upload()
        elif self.path.startswith('/process/'):
            self.handle_video_processing()
        elif self.path.startswith('/cleanup/'):
            self.handle_cleanup()
        elif self.path == '/api/start-excel-session':
            self.handle_start_excel_session()
        elif self.path == '/api/convert-word':
            self.handle_word_conversion()
        else:
            self.send_response(404)
            self.end_headers()
            self.wfile.write(b'Not Found')

    def handle_word_conversion(self):
        try:
            content_length = int(self.headers['Content-Length'])
            content_type = self.headers['Content-Type']
            
            if not content_type.startswith('multipart/form-data'):
                raise ValueError("Expected multipart/form-data")
            
            parser = MultipartFormParser(content_type, content_length, self.rfile)
            form = parser.parse()
            
            if not form or b'file' not in form:
                raise ValueError("No file uploaded")

            # Create a unique ID for this conversion
            conversion_id = os.urandom(16).hex()
            conversion_dir = os.path.join(UPLOAD_DIR, conversion_id)
            os.makedirs(conversion_dir, exist_ok=True)

            # Save the Word file
            word_path = os.path.join(conversion_dir, 'document.docx')
            with open(word_path, 'wb') as f:
                f.write(form[b'file'])

            # Convert to PDF
            pdf_path = os.path.join(conversion_dir, 'document.pdf')
            self.convert_word_to_pdf(word_path, pdf_path)

            self.send_response(200)
            self.send_header('Content-type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({
                'id': conversion_id,
                'success': True
            }).encode())

        except Exception as e:
            self.send_error(500, str(e))

    def convert_word_to_pdf(self, word_path, pdf_path):
        pythoncom.CoInitialize()
        try:
            word = win32com.client.Dispatch('Word.Application')
            doc = word.Documents.Open(word_path)
            doc.SaveAs(pdf_path, FileFormat=17)  # 17 = PDF format
            doc.Close()
            word.Quit()
        finally:
            pythoncom.CoUninitialize()

    def redirect_to(self, path):
        self.send_response(302)
        self.send_header('Location', path)
        self.end_headers()
    
    def handle_format_excel(self):
        content_length = int(self.headers['Content-Length'])
        content_type = self.headers['Content-Type']
        
        if not content_type.startswith('multipart/form-data'):
            self.send_error(400, "Expected multipart/form-data")
            return
        
        parser = MultipartFormParser(content_type, content_length, self.rfile)
        form = parser.parse()
        
        if not form or b'file' not in form:
            self.send_error(400, "No file uploaded")
            return
        
        file_content = form[b'file']
        
        if not file_content:
            self.send_error(400, "Not a file")
            return
        
        options = {
            'remove_blank_lines': form.get(b'removeBlankLines', b'false').decode() == 'true',
            'capitalize_sentences': form.get(b'capitalizeSentences', b'false').decode() == 'true',
            'add_periods': form.get(b'addPeriods', b'false').decode() == 'true',
            'remove_spaces_quotes': form.get(b'removeSpacesQuotes', b'false').decode() == 'true',
            'remove_spaces_unquoted': form.get(b'removeSpacesUnquoted', b'false').decode() == 'true',
            'remove_lone_quotes': form.get(b'removeLoneQuotes', b'false').decode() == 'true',
            'remove_ellipsis': form.get(b'removeEllipsis', b'false').decode() == 'true'
        }
        
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
            tmp_file.write(file_content)
            tmp_filename = tmp_file.name
        
        try:
            if options['remove_blank_lines']:
                remove_blank_rows(tmp_filename)
            
            if any([options['capitalize_sentences'], options['add_periods'],
                    options['remove_spaces_quotes'], options['remove_spaces_unquoted'],
                    options['remove_lone_quotes'], options['remove_ellipsis']]):
                custom_process_excel(tmp_filename, options)
            
            with open(tmp_filename, 'rb') as f:
                processed_data = f.read()
            
            self.send_response(200)
            self.send_header('Content-type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            self.send_header('Content-Disposition', 'attachment; filename="modified_file.xlsx"')
            self.send_header('Content-Length', str(len(processed_data)))
            self.end_headers()
            self.wfile.write(processed_data)
            
        except Exception as e:
            self.send_error(500, f"Error processing Excel: {str(e)}")
        
        finally:
            if os.path.exists(tmp_filename):
                os.unlink(tmp_filename)

    def handle_video_upload(self):
        try:
            content_length = int(self.headers['Content-Length'])
            content_type = self.headers['Content-Type']
            
            if not content_type.startswith('multipart/form-data'):
                raise ValueError("Expected multipart/form-data")
            
            parser = MultipartFormParser(content_type, content_length, self.rfile)
            form = parser.parse()
            
            if not form or b'video' not in form or b'excel' not in form:
                raise ValueError("Missing video or excel file")

            job_id = os.urandom(16).hex()
            job_dir = os.path.join(UPLOAD_DIR, job_id)
            os.makedirs(job_dir)

            video_path = os.path.join(job_dir, 'video.mp4')
            excel_path = os.path.join(job_dir, 'script.xlsx')
            
            with open(video_path, 'wb') as f:
                f.write(form[b'video'])
            with open(excel_path, 'wb') as f:
                f.write(form[b'excel'])

            self.send_response(200)
            self.send_header('Content-type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'job_id': job_id}).encode())

        except Exception as e:
            self.send_error(500, str(e))

    def handle_video_processing(self):
        try:
            job_id = self.path.split('/')[-1]
            job_dir = os.path.join(UPLOAD_DIR, job_id)
            
            if not os.path.exists(job_dir):
                raise ValueError("Invalid job ID")

            video_path = os.path.join(job_dir, 'video.mp4')
            excel_path = os.path.join(job_dir, 'script.xlsx')

            if not os.path.exists(video_path) or not os.path.exists(excel_path):
                raise ValueError("Missing required files")

            # Create frames folder
            frames_folder = os.path.join(job_dir, 'frames')
            os.makedirs(frames_folder, exist_ok=True)

            # Create required folders
            temp_folder = os.path.join(job_dir, 'temp')
            frames_folder = os.path.join(job_dir, 'frames')
            results_folder = os.path.join(job_dir, 'results')
            os.makedirs(temp_folder, exist_ok=True)
            os.makedirs(frames_folder, exist_ok=True)
            os.makedirs(results_folder, exist_ok=True)

            # Copy required assets to results directory
            video_to_pdf_dir = os.path.join(SCRIPT_DIR, 'Video to PDF')
            assets = {
                'logo.jpg': 'logo.jpg',
                'sideimage.png': 'sideimage.png.png',  # Fix double extension
                'Copyright.docx': 'Copyright.docx',
                'Template.docx': 'Template.docx'
            }
            for dest_name, src_name in assets.items():
                src = os.path.join(video_to_pdf_dir, src_name)
                dst = os.path.join(results_folder, dest_name)
                if os.path.exists(src):
                    shutil.copy(src, dst)
                else:
                    raise ValueError(f"Required asset not found: {src}")

            # Process files using the Video to PDF project's function
            result = process_files(
                video_path=video_path,
                excel_path=excel_path,
                credential_path=os.path.join(UPLOAD_DIR, 'sapheb-b87c6918d4ef.json'),
                bucket_name='sap_pdf',
                frames_folder=frames_folder,
                temp_folder=temp_folder,
                job_id=job_id
            )

            # Ensure the docx file exists
            results_dir = os.path.join(job_dir, 'results')
            docx_path = os.path.join(results_dir, result['docx_filename'])
            if not os.path.exists(docx_path):
                raise ValueError(f"Generated file not found: {docx_path}")

            self.send_response(200)
            self.send_header('Content-type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({
                'download_url': f'/download/{job_id}/{result["docx_filename"]}'
            }).encode())

        except Exception as e:
            error_message = handle_error(e)
            self.send_error(500, error_message)

    def handle_cleanup(self):
        try:
            job_id = self.path.split('/')[-1]
            job_dir = os.path.join(UPLOAD_DIR, job_id)
            
            if os.path.exists(job_dir):
                shutil.rmtree(job_dir)
            
            self.send_response(200)
            self.send_header('Content-type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'success': True}).encode())

        except Exception as e:
            self.send_error(500, str(e))

    def handle_start_excel_session(self):
        try:
            content_length = int(self.headers['Content-Length'])
            data = self.rfile.read(content_length)
            params = json.loads(data)
            
            session_id = os.urandom(16).hex()
            excel_sessions[session_id] = ExcelSession(
                params['filename'],
                params['saveLocation']
            )
            
            self.send_response(200)
            self.send_header('Content-type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'session_id': session_id}).encode())

        except Exception as e:
            self.send_error(500, str(e))

async def handle_websocket(websocket, path):
    try:
        # First message should be session ID
        session_id = await websocket.recv()
        session = excel_sessions.get(session_id)
        
        if not session:
            await websocket.close()
            return
        
        # Store websocket connection in session
        session.websocket = websocket
        session.monitoring = True
        
        # Start clipboard monitoring with delay
        def on_clipboard_change():
            if session.monitoring:
                # Wait a brief moment for the clipboard to update
                time.sleep(0.1)
                text = pyperclip.paste()
                asyncio.run(session.add_text(text))
        
        # Set up hotkey with delay
        keyboard.add_hotkey('ctrl+c', on_clipboard_change)
        
        try:
            while True:
                msg = await websocket.recv()
                data = json.loads(msg)
                
                if data['type'] == 'done':
                    session.monitoring = False
                    save_path = session.save()
                    await websocket.send(json.dumps({
                        'type': 'saved',
                        'path': save_path
                    }))
                    break
        
        finally:
            keyboard.remove_hotkey('ctrl+c')
            session.monitoring = False
            del excel_sessions[session_id]
    
    except websockets.exceptions.ConnectionClosed:
        pass

def get_upload_history(self):
    history = []
    if os.path.exists(UPLOAD_DIR):
        for job_id in os.listdir(UPLOAD_DIR):
            job_dir = os.path.join(UPLOAD_DIR, job_id)
            if os.path.isdir(job_dir):
                # Get directory creation time
                created_time = os.path.getctime(job_dir)
                created_date = datetime.fromtimestamp(created_time).strftime('%Y-%m-%d %H:%M:%S')
                
                # Determine job type and files
                files = []
                job_type = "Unknown"
                
                if os.path.exists(os.path.join(job_dir, 'video.mp4')):
                    job_type = "Video to PDF"
                    files.append("video.mp4")
                    if os.path.exists(os.path.join(job_dir, 'script.xlsx')):
                        files.append("script.xlsx")
                
                results_dir = os.path.join(job_dir, 'results')
                if os.path.exists(results_dir):
                    for file in os.listdir(results_dir):
                        if file.endswith(('.pdf', '.docx')):
                            files.append(file)
                
                history.append({
                    'id': job_id,
                    'date': created_date,
                    'type': job_type,
                    'files': files
                })
    
    # Sort by date, newest first
    history.sort(key=lambda x: x['date'], reverse=True)
    return history

def run_websocket_server():
    asyncio.set_event_loop(asyncio.new_event_loop())
    start_server = websockets.serve(handle_websocket, "localhost", WS_PORT)
    asyncio.get_event_loop().run_until_complete(start_server)
    asyncio.get_event_loop().run_forever()

def run_http_server():
    with socketserver.TCPServer(("", HTTP_PORT), MultiToolHandler) as httpd:
        print(f"HTTP server running at http://localhost:{HTTP_PORT}")
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("Server stopped by user")
            httpd.server_close()

if __name__ == "__main__":
    if not mimetypes.guess_type('file.js')[0]:
        mimetypes.add_type('application/javascript', '.js')
    
    # Start WebSocket server in a separate thread
    ws_thread = threading.Thread(target=run_websocket_server)
    ws_thread.daemon = True
    ws_thread.start()
    
    # Start HTTP server in main thread
    run_http_server()