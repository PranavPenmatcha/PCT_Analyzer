// Global variables
let selectedFile = null;
let analysisResults = null;

// DOM elements
const uploadArea = document.getElementById('uploadArea');
const fileInput = document.getElementById('fileInput');
const fileInfo = document.getElementById('fileInfo');
const fileName = document.getElementById('fileName');
const fileSize = document.getElementById('fileSize');
const analyzeBtn = document.getElementById('analyzeBtn');

const uploadSection = document.getElementById('uploadSection');
const loadingSection = document.getElementById('loadingSection');
const resultsSection = document.getElementById('resultsSection');
const errorSection = document.getElementById('errorSection');

// Initialize event listeners
document.addEventListener('DOMContentLoaded', function() {
    setupEventListeners();
});

function setupEventListeners() {
    // File input change
    fileInput.addEventListener('change', handleFileSelect);

    // Drag and drop
    uploadArea.addEventListener('dragover', handleDragOver);
    uploadArea.addEventListener('dragleave', handleDragLeave);
    uploadArea.addEventListener('drop', handleDrop);
    uploadArea.addEventListener('click', (e) => {
        // Only trigger file input if not clicking on the browse button
        if (e.target.id !== 'browseBtn' && !e.target.closest('#browseBtn')) {
            fileInput.click();
        }
    });

    // Browse button specific handler
    const browseBtn = document.getElementById('browseBtn');
    if (browseBtn) {
        browseBtn.addEventListener('click', (e) => {
            e.preventDefault();
            e.stopPropagation();
            fileInput.click();
        });
    }

    // Analyze button
    analyzeBtn.addEventListener('click', analyzeFile);
}

function handleFileSelect(event) {
    const file = event.target.files[0];
    if (file) {
        validateAndSetFile(file);
    }
}

function handleDragOver(event) {
    event.preventDefault();
    uploadArea.classList.add('dragover');
}

function handleDragLeave(event) {
    event.preventDefault();
    uploadArea.classList.remove('dragover');
}

function handleDrop(event) {
    event.preventDefault();
    uploadArea.classList.remove('dragover');
    
    const files = event.dataTransfer.files;
    if (files.length > 0) {
        validateAndSetFile(files[0]);
    }
}

function validateAndSetFile(file) {
    // Check file extension
    const allowedExtensions = ['.wdq', '.wdh', '.wdc'];
    const fileExtension = '.' + file.name.split('.').pop().toLowerCase();

    if (!allowedExtensions.includes(fileExtension)) {
        alert('Please select a valid WinDaq file (.wdq, .wdh, or .wdc)');
        return;
    }

    // Check file size (100MB limit)
    const maxSize = 100 * 1024 * 1024;
    if (file.size > maxSize) {
        alert('File size must be less than 100MB');
        return;
    }

    selectedFile = file;
    displayFileInfo(file);
}

function displayFileInfo(file) {
    fileName.textContent = file.name;
    fileSize.textContent = formatFileSize(file.size);
    fileInfo.style.display = 'flex';
}

function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

async function analyzeFile() {
    if (!selectedFile) {
        alert('Please select a file first');
        return;
    }
    
    // Show loading section
    showSection('loading');
    
    // Create form data
    const formData = new FormData();
    formData.append('wdqFile', selectedFile);
    
    try {
        // Simulate loading steps
        updateLoadingStep(1);
        
        const response = await fetch('/upload', {
            method: 'POST',
            body: formData
        });
        
        updateLoadingStep(2);
        
        const result = await response.json();
        
        updateLoadingStep(3);
        
        if (result.success) {
            analysisResults = result;
            displayResults(result);
            showSection('results');
        } else {
            throw new Error(result.error || 'Analysis failed');
        }
        
    } catch (error) {
        console.error('Analysis error:', error);
        showError(error.message);
    }
}

function updateLoadingStep(step) {
    // Reset all steps
    document.querySelectorAll('.step').forEach(s => s.classList.remove('active'));
    
    // Activate current step
    document.getElementById(`step${step}`).classList.add('active');
}

function displayResults(results) {
    // Update file name
    document.getElementById('resultFileName').textContent = results.filename;
    
    // Update summary
    document.getElementById('totalPulses').textContent = results.analysis.totalPulses;
    document.getElementById('peakRange').textContent = results.analysis.peakCurrentRange;
    document.getElementById('highestPeak').textContent = results.analysis.highestPeak + ' A';
    
    // Update pulses list
    const pulsesList = document.getElementById('pulsesList');
    pulsesList.innerHTML = '';
    
    results.analysis.pulses.forEach(pulse => {
        const pulseItem = document.createElement('div');
        pulseItem.className = 'pulse-item';
        pulseItem.innerHTML = `
            <div>
                <strong>Pulse ${pulse.number}</strong>
                <small>at ${pulse.time.toFixed(3)}s</small>
            </div>
            <div class="pulse-current">
                <strong>${pulse.peakCurrent.toFixed(0)} A</strong>
            </div>
        `;
        pulsesList.appendChild(pulseItem);
    });
    
    // Setup download button
    const downloadBtn = document.getElementById('downloadBtn');
    if (results.downloadUrl) {
        downloadBtn.onclick = (e) => {
            e.preventDefault();
            const link = document.createElement('a');
            link.href = results.downloadUrl;
            link.download = '';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        };
        downloadBtn.disabled = false;
        downloadBtn.innerHTML = '<i class="fas fa-download"></i> Download Excel';
    } else {
        downloadBtn.disabled = true;
        downloadBtn.innerHTML = '<i class="fas fa-times"></i> Download Not Available';
    }
}

function showSection(section) {
    // Hide all sections
    uploadSection.style.display = 'none';
    loadingSection.style.display = 'none';
    resultsSection.style.display = 'none';
    errorSection.style.display = 'none';
    
    // Show selected section
    switch(section) {
        case 'upload':
            uploadSection.style.display = 'block';
            break;
        case 'loading':
            loadingSection.style.display = 'block';
            break;
        case 'results':
            resultsSection.style.display = 'block';
            break;
        case 'error':
            errorSection.style.display = 'block';
            break;
    }
}

function showError(message) {
    document.getElementById('errorMessage').textContent = message;
    showSection('error');
}

function resetAnalysis() {
    selectedFile = null;
    analysisResults = null;
    fileInput.value = '';
    fileInfo.style.display = 'none';
    showSection('upload');
}

// Utility functions
function downloadFile(url, filename) {
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}
