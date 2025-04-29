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

    // Setup tab switching
    const tabButtons = document.querySelectorAll('.tab-button');
    const tabPanes = document.querySelectorAll('.tab-pane');

    tabButtons.forEach(button => {
        button.addEventListener('click', function() {
            const tabId = this.dataset.tab;
            
            // Update active states
            tabButtons.forEach(btn => btn.classList.remove('active'));
            tabPanes.forEach(pane => pane.classList.remove('active'));
            
            this.classList.add('active');
            const targetPane = document.querySelector(`.tab-pane[data-tab="${tabId}"]`);
            if (targetPane) {
                targetPane.classList.add('active');
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
            if (!response.ok) {
                throw new Error(`Server returned ${response.status}: ${response.statusText}`);
            }
            return response.blob();
        })
        .then(blob => {
            const lastDotIndex = file.name.lastIndexOf('.');
            const baseName = lastDotIndex !== -1 ? file.name.substring(0, lastDotIndex) : file.name;
            const extension = lastDotIndex !== -1 ? file.name.substring(lastDotIndex) : '';
            const modifiedFileName = `${baseName}_modified${extension}`;
            
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = modifiedFileName;
            document.body.appendChild(a);
            a.click();
            
            setTimeout(function() {
                document.body.removeChild(a);
                URL.revokeObjectURL(url);
            }, 0);
            
            updateStatus(`Formatted and saved as ${modifiedFileName}`);
            loadExcelPreview(file);
        })
        .catch(error => {
            console.error('Error formatting Excel:', error);
            updateStatus(`Error during formatting: ${error.message}`);
        })
        .finally(() => {
            updateFormatButtonState();
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