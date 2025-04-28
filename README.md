# Excel Formatter UI

A browser-based user interface for formatting Excel files. This application allows you to:

- Drag and drop Excel files or select them from a file explorer
- Preview Excel content in the browser
- Apply various formatting options:
  - Remove blank lines
  - Capitalize sentences
  - Add periods to the end of lines
  - Remove spaces from groups of single letters in quotes
  - Remove ellipsis (...)
- Save formatted files with "_modified" added to the filename

## Features

- Modern, responsive UI
- Drag-and-drop file handling
- Excel file preview
- Multiple formatting options
- Server-side processing using Python

## Requirements

- Python 3.6 or higher
- Required Python packages:
  - openpyxl
  - http.server (standard library)

## How to Run

### Easy Method
1. Make sure you have Python installed
2. Install required packages:
   ```
   pip install openpyxl
   ```
3. Double-click the `start_excel_formatter.bat` file
4. Your browser will automatically open to the application

### Manual Method
1. Make sure you have Python installed
2. Install required packages:
   ```
   pip install openpyxl
   ```
3. Start the server:
   ```
   python server.py
   ```
4. Open your browser and navigate to:
   ```
   http://localhost:8001
   ```

## How to Stop the Server

You can stop the server in two ways:
1. Click the "Turn Off Server" button in the web interface
2. Press Ctrl+C in the terminal window where the server is running

## How to Use

1. **Add Files**: Drag Excel files into the drop zone or click "Browse Files" to select them
2. **Select a File**: Click on a file in the queue to preview it
3. **Choose Formatting Options**: Select the formatting options you want to apply
4. **Format Excel**: Click the "Format Excel" button to process the file
5. **Save Result**: The formatted file will be automatically downloaded with "_modified" added to the filename

## Formatting Options

- **Remove blank lines**: Removes all empty rows from the Excel file
- **Capitalize sentences**: Capitalizes the first letter of each cell in column A
- **Add periods to end of lines**: Ensures each cell in column A ends with a period
- **Remove spaces from single letters in quotes**: Removes spaces between single letters in quoted text (e.g., "P R 3" becomes "PR3")
- **Remove spaces from 3+ single characters (unquoted)**: Detects and removes spaces between 3 or more single characters even when not in quotes (e.g., "P B 0 0 3 7 2 0 1" becomes "PB0037201")
- **Remove lone quotation marks**: Removes all quotation marks of a specific type if there's an odd number of them in the text
- **Remove ellipsis**: Removes all instances of "..." from the text

## Controls

- **Format Excel**: Processes the selected Excel file with the chosen formatting options
- **Turn Off Server**: Safely shuts down the server when you're done using the application

## Technical Details

This application consists of:

- Frontend: HTML, CSS, and JavaScript
- Backend: Python server using http.server
- Processing: Uses the original formatting.py script functionality

The server handles file uploads, processes Excel files using the Python script, and returns the formatted file for download.