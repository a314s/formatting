document.addEventListener('DOMContentLoaded', function() {
    // DOM Elements
    const dropZone = document.getElementById('drop-zone');
    const fileInput = document.getElementById('file-input');
    const fileQueue = document.getElementById('file-queue');
    const excelPreview = document.getElementById('excel-preview');
    const formatButton = document.getElementById('format-button');
    const shutdownButton = document.getElementById('shutdown-button');
    const statusBar = document.getElementById('status-bar');
    
    // Options checkboxes
    const removeBlankLines = document.getElementById('remove-blank-lines');
    const capitalizeSentences = document.getElementById('capitalize-sentences');
    const addPeriods = document.getElementById('add-periods');
    const removeSpacesQuotes = document.getElementById('remove-spaces-quotes');
    const removeSpacesUnquoted = document.getElementById('remove-spaces-unquoted');
    const removeLoneQuotes = document.getElementById('remove-lone-quotes');
    const removeEllipsis = document.getElementById('remove-ellipsis');
    
    // State variables
    let files = [];
    let selectedFileIndex = -1;
    let workbook = null;
    
    // Event Listeners
    dropZone.addEventListener('dragover', handleDragOver);
    dropZone.addEventListener('dragleave', handleDragLeave);
    dropZone.addEventListener('drop', handleDrop);
    fileInput.addEventListener('change', handleFileSelect);
    formatButton.addEventListener('click', formatExcel);
    shutdownButton.addEventListener('click', shutdownServer);
    
    // Drag and Drop Handlers
    function handleDragOver(e) {
        e.preventDefault();
        e.stopPropagation();
        dropZone.classList.add('active');
    }
    
    function handleDragLeave(e) {
        e.preventDefault();
        e.stopPropagation();
        dropZone.classList.remove('active');
    }
    
    function handleDrop(e) {
        e.preventDefault();
        e.stopPropagation();
        dropZone.classList.remove('active');
        
        const dt = e.dataTransfer;
        const newFiles = dt.files;
        
        addFilesToQueue(newFiles);
    }
    
    function handleFileSelect(e) {
        const newFiles = e.target.files;
        addFilesToQueue(newFiles);
        fileInput.value = ''; // Reset file input
    }
    
    // File Queue Management
    function addFilesToQueue(newFiles) {
        for (let i = 0; i < newFiles.length; i++) {
            const file = newFiles[i];
            // Check if it's an Excel file
            if (file.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' || 
                file.type === 'application/vnd.ms-excel' ||
                file.name.endsWith('.xlsx') || 
                file.name.endsWith('.xls') ||
                file.name.endsWith('.xlsm') ||
                file.name.endsWith('.xltx') ||
                file.name.endsWith('.xltm')) {
                
                files.push(file);
                addFileToQueueUI(file, files.length - 1);
            } else {
                updateStatus(`Skipped non-Excel file: ${file.name}`);
            }
        }
        
        if (files.length > 0 && selectedFileIndex === -1) {
            selectFile(0);
        }
        
        updateFormatButtonState();
    }
    
    function addFileToQueueUI(file, index) {
        const fileItem = document.createElement('div');
        fileItem.className = 'file-item';
        fileItem.dataset.index = index;
        
        const fileName = document.createElement('div');
        fileName.className = 'file-item-name';
        fileName.textContent = file.name;
        
        const removeButton = document.createElement('span');
        removeButton.className = 'file-item-remove';
        removeButton.textContent = '×';
        removeButton.addEventListener('click', function(e) {
            e.stopPropagation();
            removeFile(index);
        });
        
        fileItem.appendChild(fileName);
        fileItem.appendChild(removeButton);
        
        fileItem.addEventListener('click', function() {
            selectFile(parseInt(this.dataset.index));
        });
        
        fileQueue.appendChild(fileItem);
    }
    
    function removeFile(index) {
        files.splice(index, 1);
        refreshFileQueueUI();
        
        if (selectedFileIndex === index) {
            if (files.length > 0) {
                selectFile(0);
            } else {
                selectedFileIndex = -1;
                clearExcelPreview();
            }
        } else if (selectedFileIndex > index) {
            selectedFileIndex--;
        }
        
        updateFormatButtonState();
    }
    
    function refreshFileQueueUI() {
        fileQueue.innerHTML = '';
        files.forEach((file, index) => {
            addFileToQueueUI(file, index);
        });
        
        // Update selected file highlight
        if (selectedFileIndex >= 0) {
            const selectedItem = fileQueue.querySelector(`[data-index="${selectedFileIndex}"]`);
            if (selectedItem) {
                selectedItem.classList.add('selected');
            }
        }
    }
    
    function selectFile(index) {
        selectedFileIndex = index;
        
        // Update UI selection
        const fileItems = fileQueue.querySelectorAll('.file-item');
        fileItems.forEach(item => item.classList.remove('selected'));
        
        const selectedItem = fileQueue.querySelector(`[data-index="${index}"]`);
        if (selectedItem) {
            selectedItem.classList.add('selected');
        }
        
        // Load and preview the selected Excel file
        loadExcelPreview(files[index]);
        
        updateFormatButtonState();
    }
    
    // Excel Preview
    function loadExcelPreview(file) {
        updateStatus(`Loading preview for ${file.name}...`);
        
        const reader = new FileReader();
        
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                workbook = XLSX.read(data, { type: 'array' });
                displayExcelPreview(workbook);
                updateStatus(`Loaded preview for ${file.name}`);
            } catch (error) {
                console.error('Error reading Excel file:', error);
                updateStatus(`Error loading preview: ${error.message}`);
                clearExcelPreview();
            }
        };
        
        reader.onerror = function() {
            console.error('Error reading file');
            updateStatus('Error reading file');
            clearExcelPreview();
        };
        
        reader.readAsArrayBuffer(file);
    }
    
    function displayExcelPreview(workbook) {
        if (!workbook || !workbook.SheetNames || workbook.SheetNames.length === 0) {
            clearExcelPreview();
            return;
        }
        
        // Get the first sheet
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        
        // Convert to HTML table
        const html = XLSX.utils.sheet_to_html(worksheet);
        
        // Display in preview
        excelPreview.innerHTML = html;
    }
    
    function clearExcelPreview() {
        excelPreview.innerHTML = '<p class="no-file-selected">No file selected</p>';
        workbook = null;
    }
    
    // Excel Formatting
    function formatExcel() {
        if (selectedFileIndex === -1 || !files[selectedFileIndex]) {
            updateStatus('No file selected for formatting');
            return;
        }
        
        const file = files[selectedFileIndex];
        updateStatus(`Formatting ${file.name}...`);
        
        // Create FormData to send to server
        const formData = new FormData();
        formData.append('file', file);
        
        // Add formatting options
        formData.append('removeBlankLines', removeBlankLines.checked);
        formData.append('capitalizeSentences', capitalizeSentences.checked);
        formData.append('addPeriods', addPeriods.checked);
        formData.append('removeSpacesQuotes', removeSpacesQuotes.checked);
        formData.append('removeSpacesUnquoted', removeSpacesUnquoted.checked);
        formData.append('removeLoneQuotes', removeLoneQuotes.checked);
        formData.append('removeEllipsis', removeEllipsis.checked);
        
        // Disable format button during processing
        formatButton.disabled = true;
        
        // Send to server for processing
        fetch('/api/format-excel', {
            method: 'POST',
            body: formData
        })
        .then(response => {
            if (!response.ok) {
                throw new Error(`Server returned ${response.status}: ${response.statusText}`);
            }
            return response.blob();
        })
        .then(blob => {
            // Generate modified filename
            const lastDotIndex = file.name.lastIndexOf('.');
            const baseName = lastDotIndex !== -1 ? file.name.substring(0, lastDotIndex) : file.name;
            const extension = lastDotIndex !== -1 ? file.name.substring(lastDotIndex) : '';
            const modifiedFileName = `${baseName}_modified${extension}`;
            
            // Create download link
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = modifiedFileName;
            document.body.appendChild(a);
            a.click();
            
            // Cleanup
            setTimeout(function() {
                document.body.removeChild(a);
                URL.revokeObjectURL(url);
            }, 0);
            
            updateStatus(`Formatted and saved as ${modifiedFileName}`);
            
            // Reload the file to show the changes
            loadExcelPreview(file);
        })
        .catch(error => {
            console.error('Error formatting Excel:', error);
            updateStatus(`Error during formatting: ${error.message}`);
        })
        .finally(() => {
            // Re-enable format button
            updateFormatButtonState();
        });
    }
    
    // These functions are now handled by the server-side Python code
    
    // File saving is now handled as part of the formatExcel function
    
    // Utility Functions
    function updateStatus(message) {
        statusBar.textContent = message;
        console.log(message);
    }
    
    function updateFormatButtonState() {
        formatButton.disabled = selectedFileIndex === -1 || !workbook;
    }
    
    // Server shutdown function
    function shutdownServer() {
        if (confirm('Are you sure you want to shut down the server?')) {
            updateStatus('Shutting down server...');
            
            // Disable the shutdown button to prevent multiple clicks
            shutdownButton.disabled = true;
            
            // Send request to shutdown endpoint
            fetch('/shutdown')
                .then(response => {
                    if (response.ok) {
                        updateStatus('Server is shutting down. You can close this window.');
                    } else {
                        throw new Error(`Server returned ${response.status}: ${response.statusText}`);
                    }
                })
                .catch(error => {
                    console.error('Error shutting down server:', error);
                    updateStatus(`Error shutting down server: ${error.message}`);
                    // Re-enable the button if there was an error
                    shutdownButton.disabled = false;
                });
        }
    }
});