document.addEventListener('DOMContentLoaded', function() {
    // Tab Switching
    const tabButtons = document.querySelectorAll('.tab-button');
    const tabPanes = document.querySelectorAll('.tab-pane');

    tabButtons.forEach(button => {
        button.addEventListener('click', () => {
            const tabId = button.dataset.tab;
            
            // Update active states
            tabButtons.forEach(btn => btn.classList.remove('active'));
            tabPanes.forEach(pane => pane.classList.remove('active'));
            
            button.classList.add('active');
            document.getElementById(tabId).classList.add('active');
        });
    });

    // Excel Formatter Code
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
    shutdownButton.addEventListener('click', shutdownServer);
    
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
    
    function shutdownServer() {
        if (confirm('Are you sure you want to shut down the server?')) {
            updateStatus('Shutting down server...');
            shutdownButton.disabled = true;
            
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
                    shutdownButton.disabled = false;
                });
        }
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
        formData.append('video', videoFile);
        formData.append('excel', scriptFile);
        
        fetch('/upload', {
            method: 'POST',
            body: formData
        })
        .then(response => response.json())
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
            method: 'POST'
        })
        .then(response => response.json())
        .then(data => {
            progressBarInner.style.width = '100%';
            
            if (data.error) {
                showError(data.error);
                return;
            }
            
            resultMessage.textContent = `Your PDF has been generated successfully!`;
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
});