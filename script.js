// WebSocket connection
let ws = null;
let currentExcelSession = null;

// Word to PDF Converter Code
let wordFiles = new Map();  // Store Word files
let pdfFiles = new Map();   // Store converted PDFs

// Helper function for setting up drop zones
function setupDropZone(dropZoneElement, inputElement, handleFile) {
    dropZoneElement.addEventListener('click', function() {
        inputElement.click();
    });
    
    dropZoneElement.addEventListener('dragover', function(e) {
        e.preventDefault();
        this.classList.add('dragover');
    });
    
    ['dragleave', 'dragend'].forEach(type => {
        dropZoneElement.addEventListener(type, function() {
            this.classList.remove('dragover');
        });
    });
    
    dropZoneElement.addEventListener('drop', function(e) {
        e.preventDefault();
        this.classList.remove('dragover');
        
        if (e.dataTransfer.files.length) {
            inputElement.files = e.dataTransfer.files;
            handleFile(e.dataTransfer.files[0]);
        }
    });
}

function setupWordConverter() {
    const wordDropZone = document.getElementById('word-drop-zone');
    const wordInput = document.getElementById('word-input');
    const wordQueue = document.getElementById('word-queue');
    const pdfQueue = document.getElementById('pdf-queue');
    const convertAllBtn = document.getElementById('convert-all-btn');
    const downloadAllBtn = document.getElementById('download-all-btn');
    const progressBar = document.querySelector('#word-converter .progress');
    const progressBarInner = progressBar ? progressBar.querySelector('.progress-bar') : null;
    const errorContainer = document.getElementById('word-converter-error');

    // Setup drop zone
    setupDropZone(wordDropZone, wordInput, handleWordFile);

    // Handle file input change
    wordInput.addEventListener('change', function() {
        Array.from(this.files).forEach(handleWordFile);
    });

    // Handle file selection in word queue
    wordQueue.addEventListener('click', function(e) {
        const fileItem = e.target.closest('.file-item');
        if (fileItem) {
            // Toggle selection
            fileItem.classList.toggle('selected');
            const hasSelected = wordQueue.querySelector('.file-item.selected');
            convertAllBtn.disabled = !hasSelected;
        }
    });

    // Handle file selection in PDF queue
    pdfQueue.addEventListener('click', function(e) {
        const fileItem = e.target.closest('.file-item');
        if (fileItem) {
            // Toggle selection
            fileItem.classList.toggle('selected');
            const hasSelected = pdfQueue.querySelector('.file-item.selected');
            downloadAllBtn.disabled = !hasSelected;
        }
    });

    // Convert all selected files
    convertAllBtn.addEventListener('click', function() {
        const selectedFiles = Array.from(wordQueue.querySelectorAll('.file-item.selected'));
        if (selectedFiles.length === 0) return;

        progressBar.style.display = 'flex';
        progressBarInner.style.width = '0%';
        errorContainer.style.display = 'none';

        // Convert each selected file
        Promise.all(selectedFiles.map(fileItem => {
            const fileId = fileItem.dataset.fileId;
            const file = wordFiles.get(fileId);
            return convertWordToPDF(file);
        }))
        .then(results => {
            progressBarInner.style.width = '100%';
            setTimeout(() => {
                progressBar.style.display = 'none';
            }, 1000);
            updatePDFQueue();
        })
        .catch(error => {
            showWordConverterError(error.message);
        });
    });

    // Download all selected PDFs
    downloadAllBtn.addEventListener('click', function() {
        const selectedFiles = Array.from(pdfQueue.querySelectorAll('.file-item.selected'));
        selectedFiles.forEach(fileItem => {
            const fileId = fileItem.dataset.fileId;
            const pdfFile = pdfFiles.get(fileId);
            if (pdfFile) {
                downloadPDF(pdfFile.id, pdfFile.name);
            }
        });
    });
}

function handleWordFile(file) {
    if (!file.name.endsWith('.doc') && !file.name.endsWith('.docx')) {
        showWordConverterError('Please upload only Word files (.doc or .docx)');
        return;
    }

    const fileId = generateFileId();
    wordFiles.set(fileId, file);
    updateWordQueue();
}

function updateWordQueue() {
    const wordQueue = document.getElementById('word-queue');
    wordQueue.innerHTML = '';
    
    wordFiles.forEach((file, fileId) => {
        const fileItem = document.createElement('div');
        fileItem.className = 'file-item';
        fileItem.dataset.fileId = fileId;
        fileItem.innerHTML = `
            <span class="file-name">${file.name}</span>
            <div class="file-actions">
                <button class="btn btn-sm btn-danger" onclick="removeWordFile('${fileId}')">Remove</button>
            </div>
        `;
        wordQueue.appendChild(fileItem);
    });

    document.getElementById('convert-all-btn').disabled = wordFiles.size === 0;
}

function updatePDFQueue() {
    const pdfQueue = document.getElementById('pdf-queue');
    pdfQueue.innerHTML = '';
    
    pdfFiles.forEach((file, fileId) => {
        const fileItem = document.createElement('div');
        fileItem.className = 'file-item';
        fileItem.dataset.fileId = fileId;
        fileItem.innerHTML = `
            <span class="file-name">${file.name}</span>
            <div class="file-actions">
                <button class="btn btn-sm btn-primary" onclick="downloadPDF('${file.id}', '${file.name}')">Download</button>
                <button class="btn btn-sm btn-danger" onclick="removePDFFile('${fileId}')">Remove</button>
            </div>
        `;
        pdfQueue.appendChild(fileItem);
    });

    document.getElementById('download-all-btn').disabled = pdfFiles.size === 0;
}

function removeWordFile(fileId) {
    wordFiles.delete(fileId);
    updateWordQueue();
}

function removePDFFile(fileId) {
    pdfFiles.delete(fileId);
    updatePDFQueue();
}

function convertWordToPDF(file) {
    const formData = new FormData();
    formData.append('file', file);

    return fetch('/api/convert-word', {
        method: 'POST',
        body: formData
    })
    .then(response => {
        if (!response.ok) {
            return response.text().then(text => {
                throw new Error(text);
            });
        }
        return response.json();
    })
    .then(data => {
        if (data.success && data.pdf_filename) {
            const fileId = generateFileId();
            pdfFiles.set(fileId, {
                id: data.id, // This is the job/conversion ID
                name: data.pdf_filename // Use the filename returned by the server
            });
            updatePDFQueue();
        } else {
             throw new Error(data.error || 'Conversion failed on server');
        }
    });
}

function downloadPDF(fileId, fileName) {
    window.location.href = `/download/${fileId}/${fileName}`; // Removed '/pdf/' segment
}

function showWordConverterError(message) {
    const errorContainer = document.getElementById('word-converter-error');
    errorContainer.textContent = message;
    errorContainer.style.display = 'block';
}

function generateFileId() {
    return Math.random().toString(36).substring(2) + Date.now().toString(36);
}

// Function to load and display history
function loadHistory() {
    const historyTableBody = document.getElementById('history-table-body');
    const historyStatus = document.getElementById('history-status');
    
    fetch('/api/history')
        .then(response => response.json())
        .then(history => {
            historyTableBody.innerHTML = '';
            if (history.length === 0) {
                historyStatus.textContent = 'No processing history found.';
                return;
            }
            
            history.forEach(item => {
                const row = document.createElement('tr');
                row.innerHTML = `
                    <td>${item.date}</td>
                    <td>${item.type}</td>
                    <td>${item.files.join(', ')}</td>
                    <td class="action-buttons">
                        ${item.files.map(file => `
                            <button class="btn btn-sm btn-primary" onclick="downloadFile('${item.id}', '${file}')">
                                Download ${file}
                            </button>
                        `).join('')}
                        <button class="btn btn-sm btn-danger" onclick="deleteJob('${item.id}')">
                            Delete
                        </button>
                    </td>
                `;
                historyTableBody.appendChild(row);
            });
            historyStatus.textContent = '';
        })
        .catch(error => {
            historyStatus.textContent = 'Error loading history: ' + error.message;
        });
}

// Function to download a file from history
function downloadFile(jobId, filename) {
    window.location.href = `/download/${jobId}/${filename}`;
}

// Function to delete a job
function deleteJob(jobId) {
    if (confirm('Are you sure you want to delete this job and its files?')) {
        fetch(`/cleanup/${jobId}`, {
            method: 'POST'
        })
        .then(response => response.json())
        .then(data => {
            if (data.success) {
                loadHistory();  // Refresh the history list
            } else {
                alert('Failed to delete job: ' + data.error);
            }
        })
        .catch(error => {
            alert('Error deleting job: ' + error.message);
        });
    }
}

document.addEventListener('DOMContentLoaded', function() {
    // Initialize Word to PDF converter
    setupWordConverter();

    // Initialize Checklist functionality
    setupChecklistTab();

    // Initialize TTS Converter
    setupTTSTab();

    // Setup shutdown button
    const shutdownButton = document.getElementById('shutdown-button');
    if (shutdownButton) {
        shutdownButton.addEventListener('click', function() {
            if (confirm('Are you sure you want to shut down the server?')) {
                this.disabled = true;
                fetch('/shutdown')
                    .then(response => {
                        if (response.ok) {
                            alert('Server is shutting down. You can close this window.');
                            setTimeout(() => window.close(), 1000);
                        } else {
                            throw new Error(`Server returned ${response.status}: ${response.statusText}`);
                        }
                    })
                    .catch(error => {
                        console.error('Error shutting down server:', error);
                        alert(`Error shutting down server: ${error.message}`);
                        this.disabled = false;
                    });
            }
        });
    }

    // Setup tab navigation
    const tabButtons = document.querySelectorAll('.tab-button');
    const tabPanes = document.querySelectorAll('.tab-pane');

    tabButtons.forEach(button => {
        button.addEventListener('click', function() {
            const tabId = this.dataset.tab;
            console.log(`Tab button clicked: ${tabId}`); // DEBUG

            // Update active states
            tabButtons.forEach(btn => btn.classList.remove('active'));
            tabPanes.forEach(pane => pane.classList.remove('active'));
            
            this.classList.add('active');
            const targetPane = document.querySelector(`.tab-pane[data-tab="${tabId}"]`);
            if (targetPane) {
                console.log(`Found target pane for ${tabId}:`, targetPane); // DEBUG
                targetPane.classList.add('active');
                console.log(`Added 'active' class to pane for ${tabId}`); // DEBUG
            } else {
                console.error(`Could not find target pane for tabId: ${tabId}`); // DEBUG
            }

            // Load history if needed
            if (tabId === 'history') {
                loadHistory();
            }
        });
    });

    // Activate first tab by default
    if (tabButtons.length > 0) {
        tabButtons[0].click();
    }

    // Get WebSocket port from server
    fetch('/ws-port')
        .then(response => response.json())
        .then(data => {
            window.wsPort = data.port;
        })
        .catch(console.error);


    // Excel Formatter Code
    const dropZone = document.getElementById('drop-zone');
    const fileInput = document.getElementById('file-input');
    const fileQueue = document.getElementById('file-queue');
    const excelPreview = document.getElementById('excel-preview');
    const formatButton = document.getElementById('format-button');
    const statusBar = document.getElementById('status-bar');
    
    // Options checkboxes
    const removeBlankLines = document.getElementById('remove-blank-lines');
    const capitalizeSentences = document.getElementById('capitalize-sentences');
    const addPeriods = document.getElementById('add-periods');
    const removeSpacesQuotes = document.getElementById('remove-spaces-quotes');
    const removeSpacesUnquoted = document.getElementById('remove-spaces-unquoted');
    const removeLoneQuotes = document.getElementById('remove-lone-quotes');
    const removeEllipsis = document.getElementById('remove-ellipsis');
    
    // Excel Formatter State
    let files = [];
    let selectedFileIndex = -1;
    let workbook = null;
    
    // Excel Formatter Event Listeners
    dropZone.addEventListener('dragover', handleDragOver);
    dropZone.addEventListener('dragleave', handleDragLeave);
    dropZone.addEventListener('drop', handleDrop);
    fileInput.addEventListener('change', handleFileSelect);
    formatButton.addEventListener('click', formatExcel);
    
    // Excel Formatter Functions
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
        fileInput.value = '';
    }
    
    function addFilesToQueue(newFiles) {
        for (let i = 0; i < newFiles.length; i++) {
            const file = newFiles[i];
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
        
        if (selectedFileIndex >= 0) {
            const selectedItem = fileQueue.querySelector(`[data-index="${selectedFileIndex}"]`);
            if (selectedItem) {
                selectedItem.classList.add('selected');
            }
        }
    }
    
    function selectFile(index) {
        selectedFileIndex = index;
        
        const fileItems = fileQueue.querySelectorAll('.file-item');
        fileItems.forEach(item => item.classList.remove('selected'));
        
        const selectedItem = fileQueue.querySelector(`[data-index="${index}"]`);
        if (selectedItem) {
            selectedItem.classList.add('selected');
        }
        
        loadExcelPreview(files[index]);
        updateFormatButtonState();
    }
    
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
        
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const html = XLSX.utils.sheet_to_html(worksheet);
        excelPreview.innerHTML = html;
    }
    
    function clearExcelPreview() {
        excelPreview.innerHTML = '<p class="no-file-selected">No file selected</p>';
        workbook = null;
    }
    
    function formatExcel() {
        if (selectedFileIndex === -1 || !files[selectedFileIndex]) {
            updateStatus('No file selected for formatting');
            return;
        }
        
        const file = files[selectedFileIndex];
        updateStatus(`Formatting ${file.name}...`);
        
        const formData = new FormData();
        formData.append('file', file);
        
        formData.append('removeBlankLines', removeBlankLines.checked);
        formData.append('capitalizeSentences', capitalizeSentences.checked);
        formData.append('addPeriods', addPeriods.checked);
        formData.append('removeSpacesQuotes', removeSpacesQuotes.checked);
        formData.append('removeSpacesUnquoted', removeSpacesUnquoted.checked);
        formData.append('removeLoneQuotes', removeLoneQuotes.checked);
        formData.append('removeEllipsis', removeEllipsis.checked);
        
        formatButton.disabled = true;
        
        fetch('/api/format-excel', {
            method: 'POST',
            body: formData
        })
        .then(response => {
            console.log("Received response from /api/format-excel"); // DEBUG
            if (!response.ok) {
                 // Try to get error text, then throw
                 return response.text().then(text => {
                     throw new Error(text || `Server returned ${response.status}: ${response.statusText}`);
                 });
            }
            console.log("Response OK, processing blob..."); // DEBUG
            return response.blob();
        })
        .then(blob => {
            console.log("Blob received, creating download link..."); // DEBUG
            const lastDotIndex = file.name.lastIndexOf('.');
            const baseName = lastDotIndex !== -1 ? file.name.substring(0, lastDotIndex) : file.name;
            const extension = lastDotIndex !== -1 ? file.name.substring(lastDotIndex) : '';
            const modifiedFileName = `${baseName}_modified${extension}`;

            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = modifiedFileName;
            document.body.appendChild(a);
            console.log("Triggering download..."); // DEBUG
            a.click();

            // Cleanup link immediately after click simulation
            setTimeout(function() {
                console.log("Cleaning up download link..."); // DEBUG
                document.body.removeChild(a);
                URL.revokeObjectURL(url);
                 console.log("Download link cleanup complete."); // DEBUG
            }, 0);

            updateStatus(`Formatted and saved as ${modifiedFileName}`);
            console.log("Status updated. Skipping preview reload for now."); // DEBUG
            // loadExcelPreview(file); // Temporarily commented out
        })
        .catch(error => {
            console.error('Error formatting Excel:', error); // Log detailed error
            // Display the actual error message from the server or fetch failure
            updateStatus(`Error during formatting: ${error.message}`);
        })
        .finally(() => {
            console.log("Executing finally block..."); // DEBUG
            // Ensure the button is always re-enabled
            updateFormatButtonState(); // This should re-enable based on selectedFileIndex and workbook state
            console.log("Format button state updated in finally block."); // DEBUG
        });
    }
    
    function updateStatus(message) {
        statusBar.textContent = message;
        console.log(message);
    }
    
    function updateFormatButtonState() {
        formatButton.disabled = selectedFileIndex === -1 || !workbook;
    }
    

    // Video to PDF Code
    const videoDropZone = document.getElementById('video-drop-zone');
    const excelDropZone = document.getElementById('excel-drop-zone');
    const videoInput = document.getElementById('video-input');
    const excelInput = document.getElementById('excel-input');
    const videoInfo = document.getElementById('video-info');
    const excelInfo = document.getElementById('excel-info');
    const processBtn = document.getElementById('process-btn');
    const progressBar = document.querySelector('.progress');
    const progressBarInner = document.querySelector('.progress-bar');
    const resultContainer = document.getElementById('result-container');
    const resultMessage = document.getElementById('result-message');
    const downloadLink = document.getElementById('download-link');
    const cleanupBtn = document.getElementById('cleanup-btn');
    const errorContainer = document.getElementById('error-container');
    
    let videoFile = null;
    let scriptFile = null;
    let jobId = null;
    
    setupDropZone(videoDropZone, videoInput, handleVideoFile);
    setupDropZone(excelDropZone, excelInput, handleScriptFile);
    
    videoInput.addEventListener('change', function() {
        if (this.files.length) {
            handleVideoFile(this.files[0]);
        }
    });
    
    excelInput.addEventListener('change', function() {
        if (this.files.length) {
            handleScriptFile(this.files[0]);
        }
    });
    
    processBtn.addEventListener('click', function() {
        if (videoFile && scriptFile) {
            uploadFiles();
        }
    });
    
    cleanupBtn.addEventListener('click', function() {
        if (jobId) {
            cleanupFiles(jobId);
        }
    });
    
    function handleVideoFile(file) {
        if (!file.type.startsWith('video/')) {
            alert('Please select a valid video file.');
            return;
        }
        
        videoFile = file;
        videoInfo.textContent = `Selected: ${file.name} (${formatFileSize(file.size)})`;
        updateProcessButton();
    }
    
    function handleScriptFile(file) {
        const validExcelTypes = [
            'application/vnd.ms-excel',
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'application/vnd.ms-excel.sheet.macroEnabled.12'
        ];
        
        if (!validExcelTypes.includes(file.type) && 
            !file.name.endsWith('.xlsx') && 
            !file.name.endsWith('.xls')) {
            alert('Please select a valid Excel file.');
            return;
        }
        
        scriptFile = file;
        excelInfo.textContent = `Selected: ${file.name} (${formatFileSize(file.size)})`;
        updateProcessButton();
    }
    
    function updateProcessButton() {
        processBtn.disabled = !(videoFile && scriptFile);
    }
    
    function formatFileSize(bytes) {
        if (bytes === 0) return '0 Bytes';
        
        const k = 1024;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        
        return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
    }
    
    function uploadFiles() {
        resultContainer.style.display = 'none';
        errorContainer.style.display = 'none';
        progressBar.style.display = 'flex';
        progressBarInner.style.width = '10%';
        
        const formData = new FormData();
        formData.append('video', videoFile, videoFile.name);
        formData.append('excel', scriptFile, scriptFile.name);
        
        fetch('/upload', {
            method: 'POST',
            body: formData
        })
        .then(response => {
            if (!response.ok) {
                return response.text().then(text => {
                    throw new Error(text);
                });
            }
            return response.json();
        })
        .then(data => {
            if (data.error) {
                showError(data.error);
                return;
            }
            
            jobId = data.job_id;
            progressBarInner.style.width = '50%';
            processFiles(jobId);
        })
        .catch(error => {
            showError('Upload failed: ' + error.message);
        });
    }
    
    function processFiles(jobId) {
        fetch(`/process/${jobId}`, {
            method: 'POST',
            headers: {
                'Accept': 'application/json'
            }
        })
        .then(response => {
            if (!response.ok) {
                return response.text().then(text => {
                    throw new Error(text);
                });
            }
            return response.json();
        })
        .then(data => {
            progressBarInner.style.width = '100%';
            
            if (data.error) {
                showError(data.error);
                return;
            }
            
            resultMessage.textContent = `Your document has been generated successfully!`;
            downloadLink.href = data.download_url;
            resultContainer.style.display = 'block';
            
            setTimeout(() => {
                progressBar.style.display = 'none';
            }, 1000);
        })
        .catch(error => {
            showError('Processing failed: ' + error.message);
        });
    }
    
    function cleanupFiles(jobId) {
        fetch(`/cleanup/${jobId}`, {
            method: 'POST'
        })
        .then(response => response.json())
        .then(data => {
            if (data.success) {
                alert('Temporary files have been deleted.');
                cleanupBtn.disabled = true;
            } else {
                alert('Failed to delete temporary files: ' + data.error);
            }
        })
        .catch(error => {
            alert('Cleanup failed: ' + error.message);
        });
    }
    
    function showError(message) {
        errorContainer.textContent = message;
        errorContainer.style.display = 'block';
        progressBar.style.display = 'none';
    }

    // Live Excel Creator Code
    const startMonitoringBtn = document.getElementById('start-monitoring');
    const doneButton = document.getElementById('done-button');
    const sendToFormatterBtn = document.getElementById('send-to-formatter');
    const excelFilename = document.getElementById('excel-filename');
    const saveLocation = document.getElementById('save-location');
    const browseLocation = document.getElementById('browse-location');
    const previewSection = document.querySelector('#live-excel .preview-section');
    const liveExcelTable = document.getElementById('live-excel-table').querySelector('tbody');
    let savedExcelPath = null;

    // Removed hidden dirInput logic

    // Add click handler for browse button to call backend
    browseLocation.addEventListener('click', function() {
        fetch('/api/browse-directory')
            .then(response => {
                if (!response.ok) {
                    throw new Error('Failed to browse directory');
                }
                return response.json();
            })
            .then(data => {
                if (data.path) {
                    saveLocation.value = data.path;
                } else if (data.error) {
                    alert(`Error selecting directory: ${data.error}`);
                }
            })
            .catch(error => {
                console.error('Error browsing directory:', error);
                alert('Failed to open directory browser.');
            });
    });

    startMonitoringBtn.addEventListener('click', function() {
        if (!excelFilename.value || !saveLocation.value) {
            alert('Please enter a filename and choose a save location');
            return;
        }

        // Start new Excel session
        fetch('/api/start-excel-session', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                filename: excelFilename.value,
                saveLocation: saveLocation.value
            })
        })
        .then(response => response.json())
        .then(data => {
            currentExcelSession = data.session_id;
            connectWebSocket();
            previewSection.style.display = 'block';
            startMonitoringBtn.disabled = true;
        })
        .catch(error => {
            console.error('Error starting Excel session:', error);
            alert('Failed to start Excel session');
        });
    });

    doneButton.addEventListener('click', function() {
        if (ws && ws.readyState === WebSocket.OPEN) {
            ws.send(JSON.stringify({ type: 'done' }));
        }
    });

    sendToFormatterBtn.addEventListener('click', function() {
        if (savedExcelPath) {
            // Create a File object from the saved Excel file
            fetch(savedExcelPath)
                .then(response => response.blob())
                .then(blob => {
                    const file = new File([blob], excelFilename.value + '.xlsx', { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
                    addFilesToQueue([file]);
                    
                    // Switch to Excel Formatter tab
                    const formatterTab = document.querySelector('[data-tab="excel-formatter"]');
                    formatterTab.click();
                })
                .catch(error => {
                    console.error('Error loading Excel file:', error);
                    alert('Failed to load Excel file for formatting');
                });
        }
    });

    function connectWebSocket() {
        ws = new WebSocket(`ws://localhost:${window.wsPort}`);

        ws.onopen = function() {
            // Send session ID as first message
            ws.send(currentExcelSession);
        };

        ws.onmessage = function(event) {
            const data = JSON.parse(event.data);
            
            if (data.type === 'update') {
                // Add new row to preview table
                const row = document.createElement('tr');
                const cell = document.createElement('td');
                cell.textContent = data.text;
                row.appendChild(cell);
                liveExcelTable.appendChild(row);
                
                // Scroll to bottom
                const preview = document.getElementById('live-excel-preview');
                preview.scrollTop = preview.scrollHeight;
            } else if (data.type === 'saved') {
                savedExcelPath = data.path;
                sendToFormatterBtn.disabled = false;
                alert(`Excel file saved successfully at: ${data.path}`);
                // Reset UI
                previewSection.style.display = 'none';
                startMonitoringBtn.disabled = false;
                liveExcelTable.innerHTML = '';
                excelFilename.value = '';
                saveLocation.value = '';
                ws.close();
            }
        };

        ws.onerror = function(error) {
            console.error('WebSocket error:', error);
            alert('Error in Excel monitoring. Please try again.');
        };

        ws.onclose = function() {
            if (currentExcelSession) {
                currentExcelSession = null;
                previewSection.style.display = 'none';
                startMonitoringBtn.disabled = false;
            }
        };
    }
});
// TTS Converter Code
function setupTTSTab() {
    const ttsText = document.getElementById('tts-text');
    const ttsVoiceSelect = document.getElementById('tts-voice');
    const ttsVoiceSearch = document.getElementById('tts-voice-search');
    const ttsConvertBtn = document.getElementById('tts-convert-btn');
    const ttsResultContainer = document.getElementById('tts-result-container');
    const ttsResultsList = document.getElementById('tts-results-list');
    const ttsDownloadAll = document.getElementById('tts-download-all');
    const ttsErrorContainer = document.getElementById('tts-error-container');
    const ttsProgress = document.getElementById('tts-progress');
    const ttsProgressBar = ttsProgress.querySelector('.progress-bar');
    const ttsExcelDropZone = document.getElementById('tts-excel-drop-zone');
    const ttsExcelInput = document.getElementById('tts-excel-input');
    const ttsExcelPreview = document.getElementById('tts-excel-preview');
    const ttsExcelPreviewContent = document.getElementById('tts-excel-preview-content');

    let voices = [];
    let currentMode = 'text'; // 'text' or 'excel'
    let excelRows = [];

    // Load voices when the tab is set up
    loadTTSVoices();

    // Setup Excel drop zone
    setupDropZone(ttsExcelDropZone, ttsExcelInput, handleExcelFile);

    // Voice search functionality
    ttsVoiceSearch.addEventListener('input', function() {
        const searchTerm = this.value.toLowerCase();
        filterVoices(searchTerm);
    });

    // Convert button click handler
    ttsConvertBtn.addEventListener('click', async function() {
        const voice = ttsVoiceSelect.value;
        if (!voice || voice === 'loading') {
            showTTSError('Please select a voice.');
            return;
        }

        if (currentMode === 'text') {
            const text = ttsText.value.trim();
            if (!text) {
                showTTSError('Please enter some text to convert.');
                return;
            }
            await convertSingleText(text, voice);
        } else {
            if (excelRows.length === 0) {
                showTTSError('Please upload an Excel file first.');
                return;
            }
            await convertExcelRows(voice);
        }
    });

    // Download all button click handler
    ttsDownloadAll.addEventListener('click', function() {
        const audioElements = ttsResultsList.getElementsByTagName('audio');
        Array.from(audioElements).forEach((audio, index) => {
            const link = document.createElement('a');
            link.href = audio.src;
            link.download = `tts_output_${index + 1}.mp3`;
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        });
    });

    async function loadTTSVoices() {
        ttsVoiceSelect.innerHTML = '<option value="loading">Loading voices...</option>';
        ttsVoiceSelect.disabled = true;

        try {
            const response = await fetch('/api/tts-voices');
            if (!response.ok) {
                throw new Error(`Server returned ${response.status}`);
            }
            voices = await response.json();
            updateVoicesList();
            ttsVoiceSelect.disabled = false;
        } catch (error) {
            console.error('Error loading TTS voices:', error);
            ttsVoiceSelect.innerHTML = '<option value="" disabled selected>Error loading voices</option>';
            showTTSError(`Failed to load voices: ${error.message}`);
        }
    }

    function updateVoicesList(filteredVoices = null) {
        const voicesToShow = filteredVoices || voices;
        ttsVoiceSelect.innerHTML = '<option value="" disabled selected>Select a voice</option>';
        voicesToShow.forEach(voice => {
            const option = document.createElement('option');
            option.value = voice.id;
            option.textContent = voice.name;
            ttsVoiceSelect.appendChild(option);
        });
    }

    function filterVoices(searchTerm) {
        const filtered = voices.filter(voice =>
            voice.name.toLowerCase().includes(searchTerm)
        );
        updateVoicesList(filtered);
    }

    async function handleExcelFile(file) {
        if (!file.name.endsWith('.xlsx') && !file.name.endsWith('.xls')) {
            showTTSError('Please upload a valid Excel file (.xlsx or .xls)');
            return;
        }

        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                excelRows = XLSX.utils.sheet_to_json(firstSheet, { header: 1 })
                    .filter(row => row.length > 0 && row[0]); // Filter out empty rows

                // Show preview
                ttsExcelPreview.style.display = 'block';
                ttsExcelPreviewContent.innerHTML = excelRows
                    .slice(0, 5) // Show first 5 rows
                    .map(row => `<tr><td>${row[0]}</td></tr>`)
                    .join('') +
                    (excelRows.length > 5 ? '<tr><td>...</td></tr>' : '');

                currentMode = 'excel';
            } catch (error) {
                showTTSError('Error reading Excel file: ' + error.message);
            }
        };
        reader.readAsArrayBuffer(file);
    }

    async function convertSingleText(text, voice) {
        ttsProgress.style.display = 'block';
        ttsProgressBar.style.width = '50%';
        ttsConvertBtn.disabled = true;
        ttsErrorContainer.style.display = 'none';
        ttsResultContainer.style.display = 'none';

        try {
            const audioBlob = await convertTextToSpeech(text, voice);
            displayAudioResult([{ text, blob: audioBlob }]);
        } catch (error) {
            showTTSError(`Conversion failed: ${error.message}`);
        } finally {
            ttsConvertBtn.disabled = false;
            ttsProgress.style.display = 'none';
        }
    }

    async function convertExcelRows(voice) {
        ttsProgress.style.display = 'block';
        ttsConvertBtn.disabled = true;
        ttsErrorContainer.style.display = 'none';
        ttsResultContainer.style.display = 'none';

        const results = [];
        let completed = 0;

        try {
            for (const row of excelRows) {
                const text = row[0].toString().trim();
                if (text) {
                    const audioBlob = await convertTextToSpeech(text, voice);
                    results.push({ text, blob: audioBlob });
                    completed++;
                    ttsProgressBar.style.width = `${(completed / excelRows.length) * 100}%`;
                }
            }
            displayAudioResult(results);
        } catch (error) {
            showTTSError(`Conversion failed: ${error.message}`);
        } finally {
            ttsConvertBtn.disabled = false;
            ttsProgress.style.display = 'none';
        }
    }

    async function convertTextToSpeech(text, voiceId) {
        const response = await fetch('/api/tts', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ text, voice_id: voiceId })
        });

        if (!response.ok) {
            const error = await response.json().catch(() => ({ error: `Server error: ${response.status}` }));
            throw new Error(error.error || `Server error: ${response.status}`);
        }

        return response.blob();
    }

    function displayAudioResult(results) {
        ttsResultsList.innerHTML = '';
        
        results.forEach((result, index) => {
            const resultDiv = document.createElement('div');
            resultDiv.className = 'card mb-3';
            
            const audioUrl = URL.createObjectURL(result.blob);
            resultDiv.innerHTML = `
                <div class="card-body">
                    <p class="card-text">${result.text}</p>
                    <audio controls src="${audioUrl}" class="w-100"></audio>
                    <a href="${audioUrl}" class="btn btn-sm btn-success mt-2" download="tts_output_${index + 1}.mp3">
                        Download MP3
                    </a>
                </div>
            `;
            
            ttsResultsList.appendChild(resultDiv);
        });

        ttsResultContainer.style.display = 'block';
        ttsDownloadAll.style.display = results.length > 1 ? 'block' : 'none';
    }

    function showTTSError(message) {
        ttsErrorContainer.textContent = message;
        ttsErrorContainer.style.display = 'block';
    }
}

async function convertSingleText(text, voice) {
    ttsProgress.style.display = 'block';
    ttsProgressBar.style.width = '50%';
    ttsConvertBtn.disabled = true;
    ttsErrorContainer.style.display = 'none';
    ttsResultContainer.style.display = 'none';

    try {
        const audioBlob = await convertTextToSpeech(text, voice);
        displayAudioResult([{ text, blob: audioBlob }]);
    } catch (error) {
        showTTSError(`Conversion failed: ${error.message}`);
    } finally {
        ttsConvertBtn.disabled = false;
        ttsProgress.style.display = 'none';
    }
}

async function convertExcelRows(voice) {
    ttsProgress.style.display = 'block';
    ttsConvertBtn.disabled = true;
    ttsErrorContainer.style.display = 'none';
    ttsResultContainer.style.display = 'none';

    const results = [];
    let completed = 0;

    try {
        for (const row of excelRows) {
            const text = row[0].toString().trim();
            if (text) {
                const audioBlob = await convertTextToSpeech(text, voice);
                results.push({ text, blob: audioBlob });
                completed++;
                ttsProgressBar.style.width = `${(completed / excelRows.length) * 100}%`;
            }
        }
        displayAudioResult(results);
    } catch (error) {
        showTTSError(`Conversion failed: ${error.message}`);
    } finally {
        ttsConvertBtn.disabled = false;
        ttsProgress.style.display = 'none';
    }
}

async function convertTextToSpeech(text, voiceId) {
    const response = await fetch('/api/tts', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ text, voice_id: voiceId })
    });

    if (!response.ok) {
        const error = await response.json().catch(() => ({ error: `Server error: ${response.status}` }));
        throw new Error(error.error || `Server error: ${response.status}`);
    }

    return response.blob();
}

function displayAudioResult(results) {
    ttsResultsList.innerHTML = '';
    
    results.forEach((result, index) => {
        const resultDiv = document.createElement('div');
        resultDiv.className = 'card mb-3';
        
        const audioUrl = URL.createObjectURL(result.blob);
        resultDiv.innerHTML = `
            <div class="card-body">
                <p class="card-text">${result.text}</p>
                <audio controls src="${audioUrl}" class="w-100"></audio>
                <a href="${audioUrl}" class="btn btn-sm btn-success mt-2" download="tts_output_${index + 1}.mp3">
                    Download MP3
                </a>
            </div>
        `;
        
        ttsResultsList.appendChild(resultDiv);
    });

    ttsResultContainer.style.display = 'block';
    ttsDownloadAll.style.display = results.length > 1 ? 'block' : 'none';
}

function showTTSError(message) {
    ttsErrorContainer.textContent = message;
    ttsErrorContainer.style.display = 'block';
    ttsResultContainer.style.display = 'none';
}

// Checklist Tab Code
function setupChecklistTab() {
    console.log('Setting up checklist tab...'); // Debug log
    
    const checklistDropZone = document.getElementById('checklist-drop-zone');
    const checklistInput = document.getElementById('checklist-input');
    const checklistFileInfo = document.getElementById('checklist-file-info');
    const processBtn = document.getElementById('process-checklist-btn');
    const progressBar = document.getElementById('checklist-progress');
    
    // Helper function to format file size
    function formatFileSize(bytes) {
        if (bytes === 0) return '0 Bytes';
        
        const k = 1024;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        
        return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
    }
    const progressBarInner = progressBar.querySelector('.progress-bar');
    const resultContainer = document.getElementById('checklist-result-container');
    const resultMessage = document.getElementById('checklist-result-message');
    const downloadLink = document.getElementById('checklist-download-link');
    const errorContainer = document.getElementById('checklist-error-container');

    // Debug log for elements
    console.log('Elements found:', {
        dropZone: checklistDropZone,
        input: checklistInput,
        fileInfo: checklistFileInfo,
        processBtn: processBtn
    });

    let docxFile = null;

    // Setup drag and drop events
    checklistDropZone.addEventListener('dragenter', handleDragEvent);
    checklistDropZone.addEventListener('dragover', handleDragEvent);
    checklistDropZone.addEventListener('dragleave', handleDragLeave);
    checklistDropZone.addEventListener('drop', handleDrop);

    // Setup button click handler
    const browseButton = checklistDropZone.querySelector('button');
    if (browseButton) {
        browseButton.addEventListener('click', (e) => {
            e.stopPropagation(); // Prevent event from bubbling to drop zone
            checklistInput.click();
        });
    }

    // Setup drop zone click handler (excluding button)
    checklistDropZone.addEventListener('click', (e) => {
        // Only trigger file input if the click wasn't on the button
        if (e.target === checklistDropZone || e.target.classList.contains('drop-zone__content') || e.target.tagName === 'P') {
            checklistInput.click();
        }
    });

    function handleDragEvent(e) {
        e.preventDefault();
        e.stopPropagation();
        checklistDropZone.classList.add('dragover');
        console.log('Drag event:', e.type); // Debug log
    }

    function handleDragLeave(e) {
        e.preventDefault();
        e.stopPropagation();
        checklistDropZone.classList.remove('dragover');
        console.log('Drag leave'); // Debug log
    }

    function handleDrop(e) {
        console.log('Drop event triggered'); // Debug log
        e.preventDefault();
        e.stopPropagation();
        checklistDropZone.classList.remove('dragover');

        const dt = e.dataTransfer;
        const files = dt.files;

        if (files.length > 0) {
            console.log('File dropped:', files[0].name); // Debug log
            handleChecklistFile(files[0]);
        }
    }

    // Handle file input change
    checklistInput.addEventListener('change', function(e) {
        console.log('File input change event'); // Debug log
        if (this.files.length > 0) {
            console.log('File selected:', this.files[0].name); // Debug log
            handleChecklistFile(this.files[0]);
        }
    });

    // Process button click handler
    processBtn.addEventListener('click', function() {
        if (!docxFile) {
            showChecklistError('Please select a Word document first.');
            return;
        }
        processChecklistFile();
    });

    function handleChecklistFile(file) {
        console.log('Handling file:', file); // Debug log
        if (!file.name.endsWith('.docx')) {
            showChecklistError('Please select a valid Word document (.docx)');
            return;
        }

        docxFile = file;
        
        // Update file info display
        checklistFileInfo.innerHTML = `
            <div class="alert alert-info mb-0">
                <i class="fas fa-file-word"></i>
                <span class="ms-2">${file.name} (${formatFileSize(file.size)})</span>
            </div>
        `;
        checklistFileInfo.style.display = 'block';
        
        // Update drop zone content
        const dropZoneContent = checklistDropZone.querySelector('.drop-zone__content');
        if (dropZoneContent) {
            dropZoneContent.innerHTML = `
                <p>File ready for processing</p>
                <button type="button" class="btn btn-outline-primary btn-sm">
                    <i class="fas fa-exchange-alt"></i> Choose Different File
                </button>
            `;
        }
        
        processBtn.disabled = false;
        errorContainer.style.display = 'none';
        console.log('File processed successfully'); // Debug log
    }

    function processChecklistFile() {
        console.log('Processing file:', docxFile); // Debug log
        
        const formData = new FormData();
        formData.append('file', docxFile, docxFile.name);  // Include filename

        progressBar.style.display = 'flex';
        progressBarInner.style.width = '50%';
        resultContainer.style.display = 'none';
        errorContainer.style.display = 'none';
        processBtn.disabled = true;

        fetch('/api/process-checklist', {
            method: 'POST',
            body: formData
        })
        .then(response => {
            console.log('Server response:', response.status); // Debug log
            if (!response.ok) {
                return response.text().then(text => {
                    console.error('Server error:', text); // Debug log
                    throw new Error(text);
                });
            }
            return response.json();
        })
        .then(data => {
            console.log('Server response data:', data); // Debug log
            progressBarInner.style.width = '100%';
            setTimeout(() => {
                progressBar.style.display = 'none';
            }, 1000);

            resultMessage.textContent = 'Your checklist has been generated successfully!';
            downloadLink.href = `/download/${data.id}/${data.filename}`;
            resultContainer.style.display = 'block';
        })
        .catch(error => {
            showChecklistError(error.message);
        })
        .finally(() => {
            processBtn.disabled = false;
        });
    }

    function showChecklistError(message) {
        errorContainer.innerHTML = `
            <div class="alert alert-danger mb-0">
                <i class="fas fa-exclamation-circle"></i>
                <span class="ms-2">${message}</span>
            </div>
        `;
        errorContainer.style.display = 'block';
        progressBar.style.display = 'none';
        resultContainer.style.display = 'none';
        
        // Reset file info display
        checklistFileInfo.style.display = 'none';
        
        // Reset drop zone content
        const dropZoneContent = checklistDropZone.querySelector('.drop-zone__content');
        if (dropZoneContent) {
            dropZoneContent.innerHTML = `
                <p>Drag a .docx file here or</p>
                <button type="button" class="btn btn-primary">Browse Files</button>
            `;
        }
    }
}