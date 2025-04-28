import http.server
import socketserver
import os
import json
import cgi
import sys
import tempfile
import shutil
from urllib.parse import parse_qs, urlparse
import mimetypes
import threading
from pathlib import Path

# Import functions from the formatting script
from formatting import clean_text, remove_blank_rows, process_excel

# Define the port
PORT = 8001  # Changed from 8000 to avoid conflict with running server

# Get the directory of the current script
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

class ExcelFormatterHandler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=SCRIPT_DIR, **kwargs)
    
    def do_GET(self):
        # Parse URL path
        parsed_path = urlparse(self.path)
        path = parsed_path.path
        
        # Serve the index.html file for the root path
        if path == '/':
            self.path = '/index.html'
        # Handle shutdown request
        elif path == '/shutdown':
            self.send_response(200)
            self.send_header('Content-type', 'text/html')
            self.end_headers()
            self.wfile.write(b'Server shutting down...')
            # Schedule server shutdown after response is sent
            threading.Thread(target=self.server.shutdown).start()
            return
        
        return super().do_GET()
    
    def do_POST(self):
        # Handle API endpoints
        if self.path == '/api/format-excel':
            self.handle_format_excel()
        else:
            self.send_response(404)
            self.end_headers()
            self.wfile.write(b'Not Found')
    
    def handle_format_excel(self):
        # Get content length
        content_length = int(self.headers['Content-Length'])
        
        # Get content type and boundary
        content_type, pdict = cgi.parse_header(self.headers['Content-Type'])
        
        if content_type != 'multipart/form-data':
            self.send_error(400, "Expected multipart/form-data")
            return
        
        # Parse multipart form data
        pdict['boundary'] = pdict['boundary'].encode('utf-8')
        pdict['CONTENT-LENGTH'] = content_length
        form = cgi.FieldStorage(
            fp=self.rfile,
            headers=self.headers,
            environ={'REQUEST_METHOD': 'POST', 'CONTENT_TYPE': self.headers['Content-Type']}
        )
        
        # Check if file was uploaded
        if 'file' not in form:
            self.send_error(400, "No file uploaded")
            return
        
        # Get the file item
        fileitem = form['file']
        
        # Check if it's a file
        if not fileitem.file:
            self.send_error(400, "Not a file")
            return
        
        # Get options
        options = {
            'remove_blank_lines': form.getvalue('removeBlankLines') == 'true',
            'capitalize_sentences': form.getvalue('capitalizeSentences') == 'true',
            'add_periods': form.getvalue('addPeriods') == 'true',
            'remove_spaces_quotes': form.getvalue('removeSpacesQuotes') == 'true',
            'remove_spaces_unquoted': form.getvalue('removeSpacesUnquoted') == 'true',
            'remove_lone_quotes': form.getvalue('removeLoneQuotes') == 'true',
            'remove_ellipsis': form.getvalue('removeEllipsis') == 'true'
        }
        
        # Create a temporary file to save the uploaded Excel
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
            # Write the file content
            shutil.copyfileobj(fileitem.file, tmp_file)
            tmp_filename = tmp_file.name
        
        try:
            # Process the Excel file based on options
            if options['remove_blank_lines']:
                remove_blank_rows(tmp_filename)
            
            if any([options['capitalize_sentences'], options['add_periods'],
                    options['remove_spaces_quotes'], options['remove_spaces_unquoted'],
                    options['remove_lone_quotes'], options['remove_ellipsis']]):
                # We need to modify the process_excel function to respect our options
                custom_process_excel(tmp_filename, options)
            
            # Read the processed file
            with open(tmp_filename, 'rb') as f:
                processed_data = f.read()
            
            # Send response
            self.send_response(200)
            self.send_header('Content-type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            self.send_header('Content-Disposition', f'attachment; filename="modified_{os.path.basename(fileitem.filename)}"')
            self.send_header('Content-Length', str(len(processed_data)))
            self.end_headers()
            self.wfile.write(processed_data)
            
        except Exception as e:
            self.send_error(500, f"Error processing Excel: {str(e)}")
        
        finally:
            # Clean up the temporary file
            if os.path.exists(tmp_filename):
                os.unlink(tmp_filename)

def custom_process_excel(input_filename, options):
    """
    Modified version of process_excel that respects the provided options.
    """
    import openpyxl
    
    # Load the workbook
    wb = openpyxl.load_workbook(input_filename)
    ws = wb.active
    
    consecutive_blank_count = 0
    row = 1  # Start from the first row
    
    while True:
        cell = ws.cell(row=row, column=1)
        cell_value = cell.value
        
        # Check for blank (or None) cell
        if cell_value is None or str(cell_value).strip() == "":
            consecutive_blank_count += 1
            # If we've reached two consecutive blanks, stop
            if consecutive_blank_count == 2:
                break
        else:
            # Reset blank count
            consecutive_blank_count = 0
            
            # Convert cell_value to string and clean it with our options
            text = str(cell_value)
            new_text = custom_clean_text(text, options)
            
            # Update the cell if changed
            if new_text != text:
                cell.value = new_text
        
        row += 1  # Move to the next row
    
    # Save the workbook
    wb.save(input_filename)

def custom_clean_text(original_text, options):
    """
    Modified version of clean_text that respects the provided options.
    """
    # 1. Strip leading/trailing whitespace
    text = original_text.strip()
    
    # 2. Remove all occurrences of '...' if option is enabled
    if options['remove_ellipsis']:
        text = text.replace("...", "")
    
    # 3. Process unquoted text with single characters separated by spaces
    if options['remove_spaces_unquoted']:
        import re
        # Improved approach: Find any sequence of single characters separated by spaces
        # Look for patterns like "P B 0 0 3 7 2 0 1"
        text = re.sub(r'\b([A-Za-z0-9](?:\s+[A-Za-z0-9]){2,})\b', lambda m: process_unquoted_singles(m), text)
    
    # 4. Remove lone quotation marks if option is enabled
    if options['remove_lone_quotes']:
        import re
        # Find lines with only one quotation mark (either single or double)
        # Count occurrences of each type of quote
        single_quotes = text.count("'")
        double_quotes = text.count('"')
        
        # If there's an odd number of either type, remove all of that type
        if single_quotes % 2 != 0:
            text = text.replace("'", "")
        if double_quotes % 2 != 0:
            text = text.replace('"', "")
    
    # 5. Process quoted text if option is enabled
    if options['remove_spaces_quotes']:
        def process_quoted(match):
            quote_char = match.group(1)  # The quote symbol (single or double)
            content = match.group(2)     # The text inside the quotes
            
            # Trim leading/trailing spaces inside the quotes
            content_stripped = content.strip()
            
            # If the content is "BOM" (any case), remove quotes entirely => BOM
            if content_stripped.upper() == "BOM":
                return "BOM"
            
            # Else, check if all tokens are exactly one character (letter or digit)
            tokens = content_stripped.split()
            if all(len(t) == 1 for t in tokens):
                # Remove spaces by joining tokens, e.g. "P R 3" -> "PR3"
                content_stripped = "".join(tokens)
            
            # Return with the original quote characters preserved, unless it's BOM
            return f"{quote_char}{content_stripped}{quote_char}"
        
        # Apply the regex substitution
        import re
        text = re.sub(r'(["\'])(.*?)(\1)', process_quoted, text)
    
    # 6. Convert multiple spaces to single space
    text = re.sub(r'\s{2,}', ' ', text)
    
    # 7. Capitalize the first letter, if there's any text and option is enabled
    if options['capitalize_sentences'] and text:
        text = text[0].upper() + text[1:]
    
    # 8. Ensure it ends with a period if option is enabled
    if options['add_periods'] and text and not text.endswith("."):
        text += "."
    
    return text

def process_unquoted_singles(match):
    """
    Process a match of single characters separated by spaces.
    Only join them if they are all single characters.
    """
    # Get the full matched text
    full_match = match.group(0)
    
    # Split by spaces
    parts = full_match.split()
    
    # Check if all parts are single characters
    if all(len(part) == 1 for part in parts) and len(parts) >= 3:
        # Join all single characters without spaces
        return ''.join(parts)
    
    # If not all single characters or less than 3, return unchanged
    return full_match

def run_server():
    with socketserver.TCPServer(("", PORT), ExcelFormatterHandler) as httpd:
        print(f"Server running at http://localhost:{PORT}")
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("Server stopped by user")
            httpd.server_close()

if __name__ == "__main__":
    # Add MIME type for JavaScript if not already registered
    if not mimetypes.guess_type('file.js')[0]:
        mimetypes.add_type('application/javascript', '.js')
    
    run_server()