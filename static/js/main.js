let markers = [];
let currentVideo = null;
let currentVideoFile = null; // Store the File object for local videos
let isSingleButtonMode = false;
let activePasteColumns = {};
let activeCell = null;
let isNextMarkTcrIn = true; // For single button mode toggle
let usageOptions = [];
let musicCoOptions = [];
let usageCounts = { BI: 0, BV: 0, VI: 0, VV: 0, SRC: 0, 'BI,BV': 0, 'VI,VV': 0, 'BI,VV': 0,'VI,BV': 0 };
let extraColumns = [];
let seqClickState = { row: null, count: 0, timeout: null }; // For tracking triple clicks
let headerRows = [];

// Add undo/redo history tracking
let history = [];
let historyIndex = -1;
const MAX_HISTORY = 50; // Maximum number of states to keep in history

// Function to save current state to history
function saveToHistory() {
    // Remove any future states if we're not at the end of history
    if (historyIndex < history.length - 1) {
        history = history.slice(0, historyIndex + 1);
    }
    
    // Create deep copy of current state
    const currentState = {
        markers: JSON.parse(JSON.stringify(markers)),
        markedRows: JSON.parse(JSON.stringify(markedRows))
    };
    
    // Add to history
    history.push(currentState);
    historyIndex++;
    
    // Trim history if it gets too long
    if (history.length > MAX_HISTORY) {
        history.shift();
        historyIndex--;
    }
}

// Function to undo last action
function undo() {
    if (historyIndex > 0) {
        historyIndex--;
        const previousState = history[historyIndex];
        markers = JSON.parse(JSON.stringify(previousState.markers));
        markedRows = JSON.parse(JSON.stringify(previousState.markedRows));
        updateMarkerTable();
    }
}

// Function to redo last undone action
function redo() {
    if (historyIndex < history.length - 1) {
        historyIndex++;
        const nextState = history[historyIndex];
        markers = JSON.parse(JSON.stringify(nextState.markers));
        markedRows = JSON.parse(JSON.stringify(nextState.markedRows));
        updateMarkerTable();
    }
}

// Cue Sheet Upload & Integration
let cueSheetParsed = null;
let cueSheetFile = null;

// Add these variables at the top with other global variables
let frameTimer = null;
let frameRate = 30; // Default frame rate, will be updated when video loads
let isFrameByFrame = false;

// Add a global object to track manual edits per row/column during copy mode
let manualEdits = {};

// Add at the top with other global variables
let isSequenceReversed = false;

// --- Marked Rows as Object with Multiple Colors ---
// markedRows[rowIndex] = { yellow: true, red: true, exception: true }
let markedRows = {}; // { rowIndex: { yellow: true, red: true, exception: true } }
let exceptionSettings = {}; // { rowIndex: { filmTitle: boolean, titlePrefix: boolean } }

// Load markedRows from localStorage on page load
(function() {
    try {
        const savedMarkedRows = localStorage.getItem('markedRows');
        if (savedMarkedRows) markedRows = JSON.parse(savedMarkedRows);
        
        const savedExceptionSettings = localStorage.getItem('exceptionSettings');
        if (savedExceptionSettings) exceptionSettings = JSON.parse(savedExceptionSettings);
    } catch (e) { 
        markedRows = {}; 
        exceptionSettings = {};
    }
})();

// Save markedRows to localStorage
function saveMarkedRows() {
    localStorage.setItem('markedRows', JSON.stringify(markedRows));
    localStorage.setItem('exceptionSettings', JSON.stringify(exceptionSettings));
}

// Helper function to check if a field should be auto-filled for a specific row
function shouldAutoFillField(rowIndex, fieldName) {
    const rowException = exceptionSettings[rowIndex];
    if (!rowException) return true; // No exception, allow auto-fill
    
    if (fieldName === 'filmTitle') {
        return !rowException.filmTitle;
    }
    if (fieldName === 'titlePrefix') {
        return !rowException.titlePrefix;
    }
    
    return true; // Default to allowing auto-fill for unknown fields
}

// Update default columns to include 'Title'
const defaultMarkerColumns = [
    { key: 'tcrIn', label: 'TCR In' },
    { key: 'tcrOut', label: 'TCR Out' },
    { key: 'duration', label: 'Duration' },
    { key: 'usage', label: 'Usage' },
    { key: 'title', label: 'Title' },
    { key: 'filmTitle', label: 'Film/Album Title' },
    { key: 'composer', label: 'Composer' },
    { key: 'lyricist', label: 'Lyricist' },
    { key: 'musicCo', label: 'Music Co' },
    { key: 'nocId', label: 'NOC ID' },
    { key: 'nocTitle', label: 'NOC Title' }
];

// Mapping for common column name variations
const columnNameMap = {
    'tcr in': 'tcrIn',
    'tcr out': 'tcrOut',
    'duration': 'duration',
    'usage': 'usage',
    'title': 'title',
    'film/album title': 'filmTitle',
    'film / album title': 'filmTitle',
    'film album title': 'filmTitle',
    'composer': 'composer',
    'lyricist': 'lyricist',
    'music co': 'musicCo',
    'noc id': 'nocId',
    'noc title': 'nocTitle'
};

// Add a new variable to track the pauseWithTCRMark toggle
let pauseWithTCRMark = false;

// Add a new variable to track checkbox clicks for row marking
let checkboxClickState = { row: null, count: 0, timeout: null, lastClickTime: 0 };

// Initialize the application
document.addEventListener('DOMContentLoaded', function() {
    console.log('Initializing main page components...');
    
    // Load markers from localStorage if available
    try {
        const savedMarkers = localStorage.getItem('markers');
        if (savedMarkers) {
            markers = JSON.parse(savedMarkers);
            console.log('Loaded markers from localStorage:', markers.length);
        }
        
        const savedHeaderRows = localStorage.getItem('headerRows');
        if (savedHeaderRows) {
            headerRows = JSON.parse(savedHeaderRows);
            console.log('Loaded headerRows from localStorage');
        }
    } catch (e) {
        console.error('Error loading markers from localStorage:', e);
    }
    
    initializeVideoPlayer();
    initializeMarkerTable();
    initializeResizeHandles();
    initializeExportButtons();
    setupKeyboardShortcuts();
    setupViewModeSwitcher();
    initializeAddColumn();
    initializeCueSheetUpload();
    initializeClearTableButton();
    initializeOffsetModal();
    initializeSequenceDirection();
    initializeTCRCellHandlers();
    initializeExportSettings();
    initializeRowMarking();
    initializeColumnResize();
    setupSeqHeaderDoubleClick();
    
    // Load options from Firestore
    loadUsageOptions();
    loadMusicCoOptions();
    
    // Load usage counts from localStorage
    loadUsageCounts();
    
    // Update usage counts based on current markers
    updateUsageCounts();
    
    // Initialize history with current state
    saveToHistory();
    
    console.log('Main page components initialized');
});

function initializeVideoPlayer() {
    const videoPlayer = document.getElementById('videoPlayer');
    const tcrInBtn = document.getElementById('tcrInBtn');
    const tcrOutBtn = document.getElementById('tcrOutBtn');
    const singleButtonMode = document.getElementById('singleButtonMode');
    const pauseWithTCRMarkToggle = document.getElementById('pauseWithTCRMark');
    const loadVideoBtn = document.getElementById('loadVideoBtn');
    const videoFileInput = document.getElementById('videoFileInput');
    const videoUrlInput = document.getElementById('videoUrlInput');
    const loadUrlBtn = document.getElementById('loadUrlBtn');

    // Initialize frame rate
    window.frameRate = 25; // Default frame rate

    // Restore pauseWithTCRMark from localStorage if available
    const savedPauseWithTCRMark = localStorage.getItem('pauseWithTCRMark');
    if (savedPauseWithTCRMark !== null) {
        pauseWithTCRMark = savedPauseWithTCRMark === 'true';
        if (pauseWithTCRMarkToggle) pauseWithTCRMarkToggle.checked = pauseWithTCRMark;
    }

    if (pauseWithTCRMarkToggle) {
        pauseWithTCRMarkToggle.addEventListener('change', function() {
            pauseWithTCRMark = this.checked;
            localStorage.setItem('pauseWithTCRMark', pauseWithTCRMark);
        });
    }

    videoPlayer.addEventListener('loadedmetadata', function() {
        currentVideo = videoPlayer.src;
        // Try to get the actual frame rate from the video
        try {
            const videoTrack = videoPlayer.videoTracks?.[0];
            if (videoTrack) {
                const settings = videoTrack.getSettings();
                if (settings.frameRate) {
                    window.frameRate = settings.frameRate;
                }
            }
        } catch (e) {
            console.warn('Could not detect frame rate, using default:', window.frameRate);
        }
    });

    tcrInBtn.addEventListener('click', () => markTCR('in'));
    tcrOutBtn.addEventListener('click', () => markTCR('out'));

    singleButtonMode.addEventListener('change', function() {
        isSingleButtonMode = this.checked;
        isNextMarkTcrIn = true; // Reset the toggle state when switching modes
        updateButtonMode();
    });

    loadVideoBtn.addEventListener('click', () => {
        videoFileInput.click();
    });

    // Handle video file upload
    document.getElementById('videoFileInput').addEventListener('change', function(e) {
        const file = e.target.files[0];
        const videoPlayer = document.getElementById('videoPlayer');
        const playerStatus = document.getElementById('playerStatus');
        if (file && file.type === 'video/mp4') {
            // Set exportSettings.fileName to the base name of the file if not set
            const baseName = file.name.replace(/\.[^/.]+$/, "");
            let exportSettings = JSON.parse(localStorage.getItem('exportSettings')) || {};
            if (!exportSettings.fileName || exportSettings.fileName === 'exported_file') {
                exportSettings.fileName = baseName;
                localStorage.setItem('exportSettings', JSON.stringify(exportSettings));
            }
            const url = URL.createObjectURL(file);
            videoPlayer.src = url;
            videoPlayer.load();
            playerStatus.style.display = 'block';
            playerStatus.className = 'alert alert-success';
            playerStatus.textContent = 'Local MP4 loaded!';
            setTimeout(() => { playerStatus.style.display = 'none'; }, 3000);
        } else {
            playerStatus.style.display = 'block';
            playerStatus.className = 'alert alert-danger';
            playerStatus.textContent = 'Please select a valid MP4 file.';
        }
    });

    loadUrlBtn.addEventListener('click', async function() {
        const url = videoUrlInput.value.trim();
        if (!url) {
            alert('Please enter a valid video URL');
            return;
        }

        try {
            // Show loading state
            loadUrlBtn.disabled = true;
            loadUrlBtn.textContent = 'Loading...';

            // Test if the URL is accessible
            const response = await fetch(url, { method: 'HEAD' });
            if (!response.ok) {
                throw new Error('Video URL is not accessible');
            }

            // Set the video source
            videoPlayer.src = url;
            videoPlayer.load();
            currentVideo = url;
            currentVideoFile = null; // Clear local file reference

            // Wait for video to load
            await new Promise((resolve, reject) => {
                videoPlayer.onloadeddata = resolve;
                videoPlayer.onerror = reject;
            });

            // Reset button state
            loadUrlBtn.disabled = false;
            loadUrlBtn.textContent = 'Load URL';
        } catch (error) {
            alert('Error loading video: ' + error.message);
            loadUrlBtn.disabled = false;
            loadUrlBtn.textContent = 'Load URL';
        }
    });

    // Handle Enter key in URL input
    videoUrlInput.addEventListener('keypress', function(e) {
        if (e.key === 'Enter') {
            loadUrlBtn.click();
        }
    });
}

function initializeMarkerTable() {
    const tableBody = document.getElementById('markerTableBody');
    const tableHeader = document.querySelector('.table thead tr:first-child');
    
    // Add checkbox column to header
    const checkboxHeader = document.createElement('th');
    checkboxHeader.innerHTML = '<input type="checkbox" id="selectAllRows">';
    tableHeader.insertBefore(checkboxHeader, tableHeader.firstChild);
    
    // Add select all functionality
    document.getElementById('selectAllRows').addEventListener('change', function(e) {
        const checkboxes = document.querySelectorAll('.row-checkbox');
        checkboxes.forEach(checkbox => {
            checkbox.checked = e.target.checked;
        });
    });
    
    // Add delete selected button functionality
    const deleteSelectedBtn = document.getElementById('deleteSelectedBtn');
    if (deleteSelectedBtn) {
        deleteSelectedBtn.addEventListener('click', deleteSelectedRows);
    }
    
    // Add double-click handler for special column headers
    tableHeader.addEventListener('dblclick', function(e) {
        if (e.target.tagName === 'TH' && e.target.classList.contains('special')) {
            const headerIndex = Array.from(e.target.parentElement.children).indexOf(e.target);
            const columnName = e.target.dataset.field;
            
            if (columnName) {
                // Toggle active paste mode for this column
                activePasteColumns[columnName] = !activePasteColumns[columnName];
                // Reset manualEdits for this column if activating
                if (activePasteColumns[columnName]) {
                    manualEdits[columnName] = {};
                } else {
                    delete manualEdits[columnName];
                }
                // Update header color to indicate active state
                updateHeaderColors();
                // Only show copy dropdown if activating
                if (activePasteColumns[columnName]) {
                    showCopyDropdown(e, columnName);
                }
            }
        }
    });
    
    // Add input change handler to copy first row value when in paste mode
    tableBody.addEventListener('change', function(e) {
        if (e.target.classList.contains('table-input')) {
            // Save state before making changes
            saveToHistory();
            
            if (activePasteColumns[e.target.dataset.field]) {
                const columnIndex = Array.from(e.target.parentElement.children).indexOf(e.target);
                const columnName = getColumnName(columnIndex);
                
                if (columnName === e.target.dataset.field && markers.length > 0) {
                    const firstRowValue = markers[0][columnName] || '';
                    applyValueToAllRows(columnName, firstRowValue);
                }
            }
        }
    });

    // Add event listener for recognize buttons
    tableBody.addEventListener('click', async function(e) {
        if (e.target.classList.contains('recognize-btn')) {
            const idx = +e.target.dataset.index;
            const marker = markers[idx];
            const tcrIn = marker.tcrIn;
            const tcrOut = marker.tcrOut;
            let videoSrc = currentVideo;
            let isLocal = false;
            if (videoSrc && videoSrc.startsWith('blob:')) {
                isLocal = true;
            }

            // If local, slice the file and send only the chunk
            if (isLocal && currentVideoFile) {
                // Get video duration in seconds
                const videoPlayer = document.getElementById('videoPlayer');
                const duration = videoPlayer.duration;
                const fileSize = currentVideoFile.size;
                // Convert TCR In/Out to seconds
                const tcrInSec = timeToSeconds(tcrIn);
                const tcrOutSec = timeToSeconds(tcrOut);
                
                // Create progress indicator
                const progressDiv = document.createElement('div');
                progressDiv.className = 'upload-progress';
                progressDiv.innerHTML = `
                    <div class="progress-bar">
                        <div class="progress-fill"></div>
                    </div>
                    <div class="progress-text">Processing: 0%</div>
                `;
                e.target.parentElement.appendChild(progressDiv);
                
                try {
                    // Create a temporary video element
                    const tempVideo = document.createElement('video');
                    tempVideo.src = URL.createObjectURL(currentVideoFile);
                    
                    // Wait for video to be loaded
                    await new Promise((resolve, reject) => {
                        tempVideo.onloadedmetadata = resolve;
                        tempVideo.onerror = reject;
                    });
                    
                    // Create MediaRecorder with specific audio settings
                    const stream = tempVideo.captureStream();
                    
                    // List of preferred MIME types in order of preference
                    const mimeTypes = [
                        'audio/webm;codecs=opus',
                        'audio/webm',
                        'audio/mp4',
                        'audio/mpeg'
                    ];
                    
                    // Find the first supported MIME type
                    let selectedMimeType = null;
                    for (const mimeType of mimeTypes) {
                        if (MediaRecorder.isTypeSupported(mimeType)) {
                            selectedMimeType = mimeType;
                            break;
                        }
                    }
                    
                    if (!selectedMimeType) {
                        throw new Error('No supported audio MIME types found');
                    }
                    
                    // Create MediaRecorder with the first supported MIME type
                    const mediaRecorder = new MediaRecorder(stream, {
                        mimeType: selectedMimeType,
                        audioBitsPerSecond: 256000
                    });
                    
                    const chunks = [];
                    mediaRecorder.ondataavailable = (e) => {
                        if (e.data.size > 0) {
                            chunks.push(e.data);
                        }
                    };
                    
                    mediaRecorder.onstop = async () => {
                        // Combine chunks into a single blob
                        const blob = new Blob(chunks, { type: selectedMimeType });
                        
                        // Verify the blob size
                        if (blob.size < 1000) { // Less than 1KB is suspicious
                            throw new Error('Generated audio segment is too small');
                        }
                        
                        // Update progress
                        progressDiv.querySelector('.progress-text').textContent = 'Uploading...';
                        
                        // Prepare form data
                        const formData = new FormData();
                        formData.append('file', blob, 'segment.' + selectedMimeType.split('/')[1].split(';')[0]);
                        formData.append('tcrIn', tcrIn);
                        formData.append('tcrOut', tcrOut);
                        formData.append('mimeType', selectedMimeType); // Send the MIME type to the server
                        
                        // Send to backend with progress tracking
                        const xhr = new XMLHttpRequest();
                        xhr.open('POST', '/api/recognize-audio', true);
                        
                        xhr.upload.onprogress = function(e) {
                            if (e.lengthComputable) {
                                const percentComplete = (e.loaded / e.total) * 100;
                                progressDiv.querySelector('.progress-fill').style.width = percentComplete + '%';
                                progressDiv.querySelector('.progress-text').textContent = `Uploading: ${Math.round(percentComplete)}%`;
                            }
                        };
                        
                        xhr.onload = function() {
                            if (xhr.status === 200) {
                                const result = JSON.parse(xhr.responseText);
                                marker.recognition = result;
                                updateMarkerTable();
                            } else {
                                alert('Error during recognition: ' + xhr.statusText);
                            }
                            progressDiv.remove();
                            // Cleanup
                            URL.revokeObjectURL(tempVideo.src);
                        };
                        
                        xhr.onerror = function() {
                            alert('Error during recognition');
                            progressDiv.remove();
                            // Cleanup
                            URL.revokeObjectURL(tempVideo.src);
                        };
                        
                        xhr.send(formData);
                    };
                    
                    // Start recording
                    try {
                        mediaRecorder.start();
                        
                        // Set the time range and play
                        tempVideo.currentTime = tcrInSec;
                        await tempVideo.play();
                        
                        // Stop recording after duration
                        setTimeout(() => {
                            tempVideo.pause();
                            mediaRecorder.stop();
                        }, (tcrOutSec - tcrInSec) * 1000);
                    } catch (error) {
                        throw new Error('Failed to start recording: ' + error.message);
                    }
                } catch (error) {
                    alert('Error processing video segment: ' + error.message);
                    progressDiv.remove();
                }
            } else if (!isLocal) {
                // Server file: send videoSrc and timecodes
                const formData = new FormData();
                formData.append('videoSrc', videoSrc);
                formData.append('tcrIn', tcrIn);
                formData.append('tcrOut', tcrOut);
                const resp = await fetch('/api/recognize-audio', {
                    method: 'POST',
                    body: formData
                });
                const result = await resp.json();
                marker.recognition = result;
                updateMarkerTable();
            }
        }
    });

    // Add datalists
    addDatalists();
}

function initializeExportButtons() {
    // This function is now obsolete since export is handled by exportBtn and settings
}

function getMostUsedUsage() {
    let max = 0;
    let mostUsed = usageOptions[0];
    for (let opt of usageOptions) {
        if (usageCounts[opt] > max) {
            max = usageCounts[opt];
            mostUsed = opt;
        }
    }
    return mostUsed;
}

// Update usage counts based on current markers
function updateUsageCounts() {
    // Reset counts
    usageCounts = {};
    
    // Initialize counts for all usage options
    usageOptions.forEach(option => {
        usageCounts[option] = 0;
    });
    
    // Count usage in current markers
    markers.forEach(marker => {
        if (marker.usage) {
            if (Array.isArray(marker.usage)) {
                marker.usage.forEach(usage => {
                    if (usageCounts.hasOwnProperty(usage)) {
                        usageCounts[usage]++;
                    }
                });
            } else if (typeof marker.usage === 'string') {
                // Handle comma-separated usage values
                const usages = marker.usage.split(',').map(u => u.trim());
                usages.forEach(usage => {
                    if (usageCounts.hasOwnProperty(usage)) {
                        usageCounts[usage]++;
                    }
                });
            }
        }
    });
    
    // Save usage counts to localStorage for persistence
    localStorage.setItem('usageCounts', JSON.stringify(usageCounts));
}

// Load usage counts from localStorage
function loadUsageCounts() {
    try {
        const savedCounts = localStorage.getItem('usageCounts');
        if (savedCounts) {
            usageCounts = JSON.parse(savedCounts);
        }
    } catch (e) {
        usageCounts = {};
    }
}

// Get sorted usage options (most used first)
function getSortedUsageOptions() {
    return [...usageOptions].sort((a, b) => {
        const countA = usageCounts[a] || 0;
        const countB = usageCounts[b] || 0;
        return countB - countA; // Sort in descending order (most used first)
    });
}

function markTCR(type) {
    const videoPlayer = document.getElementById('videoPlayer');
    const currentTime = videoPlayer.currentTime;
    
    // Save state before making changes
    saveToHistory();
    
    // Pause video if toggle is enabled
    if (pauseWithTCRMark && videoPlayer && !videoPlayer.paused) {
        videoPlayer.pause();
    }

    // Get selected rows
    const selectedCheckboxes = document.querySelectorAll('.row-checkbox:checked');
    const selectedRows = Array.from(selectedCheckboxes).map(checkbox => {
        const row = checkbox.closest('tr');
        return parseInt(row.querySelector('.seq-cell').textContent) - 1;
    });

    if (selectedRows.length > 0) {
        // Update TCR for selected rows
        selectedRows.forEach(rowIndex => {
            if (isSingleButtonMode) {
                if (isNextMarkTcrIn) {
                    markers[rowIndex].tcrIn = formatTime(currentTime);
                    markers[rowIndex].tcrOut = '';
                    markers[rowIndex].duration = '';
                } else {
                    markers[rowIndex].tcrOut = formatTime(currentTime);
                    markers[rowIndex].duration = calculateDuration(markers[rowIndex].tcrIn, markers[rowIndex].tcrOut);
                }
            } else {
                if (type === 'in') {
                    markers[rowIndex].tcrIn = formatTime(currentTime);
                    markers[rowIndex].tcrOut = '';
                    markers[rowIndex].duration = '';
                } else {
                    markers[rowIndex].tcrOut = formatTime(currentTime);
                    markers[rowIndex].duration = calculateDuration(markers[rowIndex].tcrIn, markers[rowIndex].tcrOut);
                }
            }
        });
    } else {
        // Original behavior for no selection
        if (isSingleButtonMode) {
            if (isNextMarkTcrIn) {
                addMarkerRow({
                    tcrIn: formatTime(currentTime),
                    tcrOut: '',
                    duration: '',
                    usage: getMostUsedUsage(),
                    filmTitle: '',
                    composer: '',
                    lyricist: '',
                    musicCo: '',
                    nocId: '',
                    nocTitle: ''
                });
            } else {
                if (markers.length > 0) {
                    const lastMarker = markers[markers.length - 1];
                    lastMarker.tcrOut = formatTime(currentTime);
                    lastMarker.duration = calculateDuration(lastMarker.tcrIn, lastMarker.tcrOut);
                }
            }
            isNextMarkTcrIn = !isNextMarkTcrIn;
        } else {
            if (type === 'in') {
                addMarkerRow({
                    tcrIn: formatTime(currentTime),
                    tcrOut: '',
                    duration: '',
                    usage: getMostUsedUsage(),
                    filmTitle: '',
                    composer: '',
                    lyricist: '',
                    musicCo: '',
                    nocId: '',
                    nocTitle: ''
                });
            } else {
                if (markers.length > 0) {
                    const lastMarker = markers[markers.length - 1];
                    lastMarker.tcrOut = formatTime(currentTime);
                    lastMarker.duration = calculateDuration(lastMarker.tcrIn, lastMarker.tcrOut);
                }
            }
        }
    }
    
    updateMarkerTable();
}

// Add new function for jumping to TCR
function jumpToTCR(type, rowIndex) {
    const videoPlayer = document.getElementById('videoPlayer');
    const marker = markers[rowIndex];
    
    if (!marker) return;
    
    const timecode = type === 'in' ? marker.tcrIn : marker.tcrOut;
    if (!timecode) return;
    
    const seconds = timeToSeconds(timecode);
    videoPlayer.currentTime = seconds;
}

// Add click handlers for TCR cells
function initializeTCRCellHandlers() {
    const tableBody = document.getElementById('markerTableBody');
    
    // Remove any existing click handlers
    tableBody.removeEventListener('click', handleTCRClick);
    
    // Add click handler to the table body
    tableBody.addEventListener('click', handleTCRClick);
}

function handleTCRClick(e) {
    // Check if we clicked on a TCR cell or its input
    const cell = e.target.closest('td');
    if (!cell || !cell.classList.contains('tcr-cell')) return;
    
    // Get the input field
    const input = cell.querySelector('input');
    if (!input || !input.dataset.field) return;
    
    // Check if it's a TCR field
    const field = input.dataset.field;
    if (field !== 'tcrIn' && field !== 'tcrOut') return;
    
    // Get the row index
    const row = cell.closest('tr');
    const rowIndex = parseInt(row.querySelector('.seq-cell').textContent) - 1;
    
    // Get the marker for this row
    const marker = markers[rowIndex];
    if (!marker) return;
    
    // Get the timecode
    const timecode = field === 'tcrIn' ? marker.tcrIn : marker.tcrOut;
    if (!timecode) return;
    
    // Convert timecode to seconds and jump
    const seconds = timeToSeconds(timecode);
    const videoPlayer = document.getElementById('videoPlayer');
    if (videoPlayer) {
        videoPlayer.currentTime = seconds;
    }
}

function updateButtonMode() {
    const tcrInBtn = document.getElementById('tcrInBtn');
    const tcrOutBtn = document.getElementById('tcrOutBtn');
    
    if (isSingleButtonMode) {
        tcrInBtn.textContent = 'TCR IN / OUT';
        tcrOutBtn.style.display = 'none';
    } else {
        tcrInBtn.textContent = 'TCR In';
        tcrOutBtn.style.display = 'inline-block';
    }
}

function applyValueToAllRows(column, value) {
    markers.forEach(marker => {
        marker[column] = value;
    });
    updateMarkerTable();
}

function getColumnName(index) {
    // Now includes 'title'
    const columns = ['seq', 'tcrIn', 'tcrOut', 'duration', 'usage', 'title',
        'filmTitle', 'composer', 'lyricist', 'musicCo', 'nocId', 'nocTitle'];
    return columns[index];
}

function makeInputResizable(input) {
    input.style.width = '100%';
    input.style.minHeight = '20px';
    input.style.padding = '4px 8px';
    input.style.boxSizing = 'border-box';
    
    // Add event listener for input changes to adjust height
    input.addEventListener('input', function() {
        this.style.height = 'auto';
        this.style.height = (this.scrollHeight) + 'px';
    });
}

function updateHeaderColors() {
    const headers = document.querySelectorAll('.table thead tr:first-child th.special');
    headers.forEach((header) => {
        const columnName = header.dataset.field;
        if (columnName) {
            if (activePasteColumns[columnName]) {
                header.classList.add('active-paste');
            } else {
                header.classList.remove('active-paste');
            }
        }
    });
}

function updateMarkerTable() {
    const tableBody = document.getElementById('markerTableBody');
    tableBody.innerHTML = '';
    
    // Create a copy of markers array and reverse if needed
    let displayMarkers = [...markers];
    if (isSequenceReversed) {
        displayMarkers.reverse();
    }
    
    displayMarkers.forEach((marker, displayIndex) => {
        const actualIndex = isSequenceReversed ? markers.length - 1 - displayIndex : displayIndex;
        const row = document.createElement('tr');
        
        // --- Apply color classes if marked ---
        const mark = markedRows[actualIndex] || {};
        if (mark.yellow) row.classList.add('marked-yellow');
        if (mark.red) row.classList.add('marked-red');
        if (mark.exception) row.classList.add('marked-exception');
        
        // Add checkbox cell with click tracking for row marking
        const checkboxCell = document.createElement('td');
        checkboxCell.style.position = 'relative';
        checkboxCell.className = 'checkbox-cell';
        checkboxCell.dataset.rowIndex = actualIndex;
        
        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.className = 'row-checkbox';
        checkboxCell.appendChild(checkbox);
        
        // Add click event listener for row marking (double-click = yellow, triple-click = red)
        checkboxCell.addEventListener('click', (e) => {
            // Don't trigger if clicking on the checkbox itself or remove buttons
            if (e.target === checkbox || e.target.classList.contains('remove-mark-btn')) return;
            
            handleCheckboxCellClick(actualIndex, e);
        });
        
        // Prevent checkbox clicks from bubbling to the cell
        checkbox.addEventListener('click', (e) => {
            e.stopPropagation();
        });
        
        // --- Add mark indicators and remove buttons if marked ---
        let markCount = 0;
        if (mark.yellow) {
            const markDot = document.createElement('span');
            markDot.className = 'mark-dot';
            markDot.style.backgroundColor = '#ffc107';
            markDot.title = 'Yellow mark';
            markDot.style.left = (markCount * 20) + 'px';
            checkboxCell.appendChild(markDot);
            
            const removeMarkBtn = document.createElement('button');
            removeMarkBtn.className = 'remove-mark-btn';
            removeMarkBtn.title = 'Remove yellow mark';
            removeMarkBtn.innerHTML = '&times;';
            removeMarkBtn.style.left = (markCount * 20) + 'px';
            removeMarkBtn.onclick = (e) => {
                e.stopPropagation();
                delete markedRows[actualIndex].yellow;
                if (Object.keys(markedRows[actualIndex]).length === 0) delete markedRows[actualIndex];
                saveMarkedRows();
                updateMarkerTable();
            };
            checkboxCell.appendChild(removeMarkBtn);
            markCount++;
        }
        if (mark.red) {
            const markDot = document.createElement('span');
            markDot.className = 'mark-dot';
            markDot.style.backgroundColor = '#dc3545';
            markDot.title = 'Red mark';
            markDot.style.left = (markCount * 20) + 'px';
            checkboxCell.appendChild(markDot);
            
            const removeMarkBtn = document.createElement('button');
            removeMarkBtn.className = 'remove-mark-btn';
            removeMarkBtn.title = 'Remove red mark';
            removeMarkBtn.innerHTML = '&times;';
            removeMarkBtn.style.left = (markCount * 20) + 'px';
            removeMarkBtn.onclick = (e) => {
                e.stopPropagation();
                delete markedRows[actualIndex].red;
                if (Object.keys(markedRows[actualIndex]).length === 0) delete markedRows[actualIndex];
                saveMarkedRows();
                updateMarkerTable();
            };
            checkboxCell.appendChild(removeMarkBtn);
            markCount++;
        }
        if (mark.exception) {
            const markDot = document.createElement('span');
            markDot.className = 'mark-dot';
            markDot.style.backgroundColor = '#6c757d';
            markDot.title = 'Exception mark';
            markDot.style.left = (markCount * 20) + 'px';
            checkboxCell.appendChild(markDot);
            
            const removeMarkBtn = document.createElement('button');
            removeMarkBtn.className = 'remove-mark-btn';
            removeMarkBtn.title = 'Remove exception mark';
            removeMarkBtn.innerHTML = '&times;';
            removeMarkBtn.style.left = (markCount * 20) + 'px';
            removeMarkBtn.onclick = (e) => {
                e.stopPropagation();
                delete markedRows[actualIndex].exception;
                delete exceptionSettings[actualIndex];
                if (Object.keys(markedRows[actualIndex]).length === 0) delete markedRows[actualIndex];
                saveMarkedRows();
                updateMarkerTable();
            };
            checkboxCell.appendChild(removeMarkBtn);
            markCount++;
        }
        row.appendChild(checkboxCell);
        
        // Update sequence number display
        const seqCell = document.createElement('td');
        seqCell.textContent = actualIndex + 1;
        seqCell.className = 'seq-cell';
        seqCell.dataset.row = actualIndex;
        seqCell.addEventListener('click', () => handleSeqClick(actualIndex, seqCell));
        row.appendChild(seqCell);
        
        // Add TCR In cell
        const tcrInCell = document.createElement('td');
        tcrInCell.className = 'tcr-cell'; // Add tcr-cell class
        const tcrInInput = document.createElement('input');
        tcrInInput.type = 'text';
        tcrInInput.className = 'table-input';
        tcrInInput.value = marker.tcrIn || '';
        tcrInInput.dataset.field = 'tcrIn';
        makeInputResizable(tcrInInput);
        tcrInInput.addEventListener('change', (e) => {
            marker.tcrIn = e.target.value;
            updateDuration(actualIndex);
        });
        // Highlight on blur
        tcrInInput.addEventListener('blur', function(e) {
            e.target.classList.add('tcr-blur');
        });
        tcrInInput.addEventListener('focus', function(e) {
            e.target.classList.remove('tcr-blur');
        });
        tcrInCell.appendChild(tcrInInput);
        row.appendChild(tcrInCell);
        
        // Add TCR Out cell
        const tcrOutCell = document.createElement('td');
        tcrOutCell.className = 'tcr-cell'; // Add tcr-cell class
        const tcrOutInput = document.createElement('input');
        tcrOutInput.type = 'text';
        tcrOutInput.className = 'table-input';
        tcrOutInput.value = marker.tcrOut || '';
        tcrOutInput.dataset.field = 'tcrOut';
        makeInputResizable(tcrOutInput);
        tcrOutInput.addEventListener('change', (e) => {
            marker.tcrOut = e.target.value;
            updateDuration(actualIndex);
        });
        // Highlight on blur
        tcrOutInput.addEventListener('blur', function(e) {
            e.target.classList.add('tcr-blur');
        });
        tcrOutInput.addEventListener('focus', function(e) {
            e.target.classList.remove('tcr-blur');
        });
        tcrOutCell.appendChild(tcrOutInput);
        row.appendChild(tcrOutCell);
        
        // Add Duration cell (read-only)
        const durationCell = document.createElement('td');
        durationCell.textContent = marker.duration || '';
        durationCell.dataset.field = 'duration';
        durationCell.addEventListener('dblclick', () => {
            row.style.backgroundColor = row.style.backgroundColor === 'yellow' ? '' : 'yellow';
        });
        row.appendChild(durationCell);
        
        // Add Usage cell with dropdown
        const usageCell = document.createElement('td');
        const usageSelect = document.createElement('select');
        usageSelect.className = 'table-input';
        usageSelect.dataset.field = 'usage';
        makeInputResizable(usageSelect);
        
        // Get sorted usage options (most used first)
        const sortedOptions = getSortedUsageOptions();
        
        sortedOptions.forEach(option => {
            const opt = document.createElement('option');
            opt.value = option;
            opt.textContent = option;
            if (marker.usage && marker.usage.includes(option)) opt.selected = true;
            usageSelect.appendChild(opt);
        });
        
        usageSelect.addEventListener('change', (e) => {
            marker.usage = Array.from(e.target.selectedOptions, option => option.value);
            updateUsageCounts();
            if (activePasteColumns['usage']) {
                if (!manualEdits['usage']) manualEdits['usage'] = {};
                manualEdits['usage'][actualIndex] = true;
            }
        });
        usageCell.appendChild(usageSelect);
        row.appendChild(usageCell);
        
        // Add other cells with resizable inputs
        ['title', 'filmTitle', 'composer', 'lyricist', 'musicCo', 'nocId', 'nocTitle'].forEach(field => {
            const cell = document.createElement('td');
            if (field === 'musicCo') {
                const select = document.createElement('select');
                select.className = 'table-input';
                select.dataset.field = field;
                makeInputResizable(select);
                
                // Add empty option
                const emptyOpt = document.createElement('option');
                emptyOpt.value = '';
                emptyOpt.textContent = '';
                select.appendChild(emptyOpt);
                
                // Fetch and add music co options
                fetch('/api/music-co')
                    .then(res => res.json())
                    .then(data => {
                        data.forEach(item => {
                            const opt = document.createElement('option');
                            opt.value = item.name;
                            opt.textContent = item.name;
                            if (marker[field] === item.name) opt.selected = true;
                            select.appendChild(opt);
                        });
                    });
                
                select.addEventListener('change', (e) => {
                    marker[field] = e.target.value;
                    if (activePasteColumns[field]) {
                        if (!manualEdits[field]) manualEdits[field] = {};
                        manualEdits[field][actualIndex] = true;
                    }
                });
                
                cell.appendChild(select);
            } else {
                const input = document.createElement('input');
                input.type = 'text';
                input.className = 'table-input';
                input.value = marker[field] || '';
                input.dataset.field = field;
                makeInputResizable(input);
                input.addEventListener('change', (e) => {
                    const newValue = e.target.value;
                    marker[field] = newValue;
                    if (activePasteColumns[field]) {
                        if (!manualEdits[field]) manualEdits[field] = {};
                        manualEdits[field][actualIndex] = true;
                    }
                });
                cell.appendChild(input);
            }
            row.appendChild(cell);
        });
        
        // Add Recognize button
        const recognizeCell = document.createElement('td');
        const recognizeBtn = document.createElement('button');
        recognizeBtn.className = 'btn btn-primary recognize-btn';
        recognizeBtn.textContent = 'Recognize';
        recognizeBtn.dataset.index = actualIndex;
        recognizeCell.appendChild(recognizeBtn);
        row.appendChild(recognizeCell);
        
        tableBody.appendChild(row);
    });
    
    updateHeaderColors();
    tableBody.parentElement.parentElement.scrollTop = 0;
}

function formatTime(seconds) {
    const hours = Math.floor(seconds / 3600);
    const minutes = Math.floor((seconds % 3600) / 60);
    const secs = Math.floor(seconds % 60);
    const frames = Math.floor((seconds % 1) * 25); // Convert decimal seconds to frames (25fps)
    
    return `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}:${secs.toString().padStart(2, '0')}:${frames.toString().padStart(2, '0')}`;
}

function timeToSeconds(timeStr) {
    if (!timeStr) return 0;
    
    const parts = timeStr.split(':').map(Number);
    if (parts.length === 3) {
        // HH:MM:SS format
        return parts[0] * 3600 + parts[1] * 60 + parts[2];
    } else if (parts.length === 4) {
        // HH:MM:SS:FF format
        return parts[0] * 3600 + parts[1] * 60 + parts[2] + (parts[3] / 25);
    }
    return 0;
}

function calculateDuration(inTime, outTime) {
    if (!inTime || !outTime) return '';
    
    const inSeconds = timeToSeconds(inTime);
    const outSeconds = timeToSeconds(outTime);
    const duration = outSeconds - inSeconds;
    
    // Get the current time format from export settings
    const settings = JSON.parse(localStorage.getItem('exportSettings')) || {};
    const timeFormat = settings.tcrFormat || 'timecode';
    
    // Format based on selected time format
    const hours = Math.floor(duration / 3600);
    const minutes = Math.floor((duration % 3600) / 60);
    const secs = Math.floor(duration % 60);
    
    if (timeFormat === 'time') {
        return `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}:${secs.toString().padStart(2, '0')}`;
    } else {
        const frames = Math.floor((duration % 1) * 25); // Convert decimal seconds to frames (25fps)
        return `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}:${secs.toString().padStart(2, '0')}:${frames.toString().padStart(2, '0')}`;
    }
}

function initializeResizeHandles() {
    const videoSection = document.getElementById('videoSection');
    const markerSection = document.getElementById('markerSection');
    const videoSectionResize = document.getElementById('videoSectionResize');
    const markerSectionResize = document.getElementById('markerSectionResize');
    
    let isResizing = false;
    let startX;
    let startWidth;
    let activeSection;
    
    function startResize(e) {
        isResizing = true;
        startX = e.clientX;
        activeSection = e.target === videoSectionResize ? videoSection : markerSection;
        startWidth = activeSection.offsetWidth;
        document.documentElement.style.cursor = 'col-resize';
        document.addEventListener('mousemove', resize);
        document.addEventListener('mouseup', stopResize);
    }
    
    function resize(e) {
        if (!isResizing) return;
        
        const deltaX = e.clientX - startX;
        let newWidth;
        
        if (activeSection === videoSection) {
            newWidth = startWidth + deltaX;
            // Set minimum and maximum widths for video section
            const minWidth = 400;
            const maxWidth = window.innerWidth * 0.8;
            
            if (newWidth >= minWidth && newWidth <= maxWidth) {
                videoSection.style.width = `${newWidth}px`;
                videoSection.style.flex = 'none';
                // Adjust marker section to fill remaining space
                markerSection.style.width = 'auto';
                markerSection.style.flex = '1';
            }
        } else {
            newWidth = startWidth - deltaX;
            // Set minimum and maximum widths for marker section
            const minWidth = 300;
            const maxWidth = window.innerWidth * 0.8;
            
            if (newWidth >= minWidth && newWidth <= maxWidth) {
                markerSection.style.width = `${newWidth}px`;
                markerSection.style.flex = 'none';
                // Adjust video section to fill remaining space
                videoSection.style.width = 'auto';
                videoSection.style.flex = '1';
            }
        }
    }
    
    function stopResize() {
        isResizing = false;
        document.documentElement.style.cursor = '';
        document.removeEventListener('mousemove', resize);
        document.removeEventListener('mouseup', stopResize);
    }
    
    videoSectionResize.addEventListener('mousedown', startResize);
    markerSectionResize.addEventListener('mousedown', startResize);
}

function setupKeyboardShortcuts() {
    let arrowInterval = null;
    let arrowTimeout = null;
    document.addEventListener('keydown', function(e) {
        const videoPlayer = document.getElementById('videoPlayer');
        const tableBody = document.getElementById('markerTableBody');
        const activeElement = document.activeElement;
        
        // If in input/select, let default behavior for arrows
        if (activeElement.tagName === 'INPUT' || activeElement.tagName === 'SELECT') {
            return;
        }
        
        // Undo (Ctrl + Z)
        if (e.ctrlKey && e.key === 'z' && !e.shiftKey) {
            e.preventDefault();
            undo();
            return;
        }
        
        // Redo (Ctrl + Shift + Z)
        if (e.ctrlKey && e.key === 'z' && e.shiftKey) {
            e.preventDefault();
            redo();
            return;
        }
        
        // Alt + 1 for TCR In (or both in single button mode)
        if (e.altKey && e.key === '1') {
            e.preventDefault();
            if (isSingleButtonMode) {
                markTCR(isNextMarkTcrIn ? 'in' : 'out');
            } else {
                markTCR('in');
            }
            return;
        }
        
        // Alt + 2 for TCR Out (only in two button mode)
        if (e.altKey && e.key === '2' && !isSingleButtonMode) {
            e.preventDefault();
            markTCR('out');
            return;
        }
        
        // Arrow keys for video seek
        if (e.key === 'ArrowRight' || e.key === 'ArrowLeft') {
            e.preventDefault();
            let seekAmount = 5;
            let direction = e.key === 'ArrowRight' ? 1 : -1;
            if (!arrowInterval) {
                videoPlayer.currentTime += direction * seekAmount;
                arrowInterval = setInterval(() => {
                    videoPlayer.currentTime += direction * seekAmount;
                }, 150);
            }
            clearTimeout(arrowTimeout);
            arrowTimeout = setTimeout(() => {
                clearInterval(arrowInterval);
                arrowInterval = null;
            }, 400);
            return;
        }
        
        // Frame-by-frame controls
        if (e.key === 'Period' || e.key === '>' || (e.shiftKey && e.key === '.')) {
            e.preventDefault();
            if (videoPlayer) {
                // Ensure video is paused
                videoPlayer.pause();
                // Calculate frame duration
                const frameDuration = 1 / window.frameRate;
                // Round to nearest frame
                const currentFrame = Math.round(videoPlayer.currentTime * window.frameRate);
                videoPlayer.currentTime = (currentFrame + 1) * frameDuration;
            }
            return;
        }
        
        if (e.key === 'Comma' || e.key === '<' || (e.shiftKey && e.key === ',')) {
            e.preventDefault();
            if (videoPlayer) {
                // Ensure video is paused
                videoPlayer.pause();
                // Calculate frame duration
                const frameDuration = 1 / window.frameRate;
                // Round to nearest frame
                const currentFrame = Math.round(videoPlayer.currentTime * window.frameRate);
                videoPlayer.currentTime = Math.max(0, (currentFrame - 1) * frameDuration);
            }
            return;
        }
        
        // Delete key
        if (e.key === 'Delete' || e.key === 'Backspace') {
            deleteSelectedRows();
            return;
        }
    });
    
    document.addEventListener('keyup', function(e) {
        if (e.key === 'ArrowRight' || e.key === 'ArrowLeft') {
            clearInterval(arrowInterval);
            arrowInterval = null;
        }
    });
}

function exportToExcel() {
    // This function is now obsolete, replaced by exportWithSettings
}

function exportToCSV() {
    // This function is now obsolete, replaced by exportWithSettings
}

function setupViewModeSwitcher() {
    const appContainer = document.querySelector('.app-container');
    const viewMode = document.getElementById('viewMode');
    const stackedResizeHandle = document.getElementById('stackedResizeHandle');
    const videoSection = document.getElementById('videoSection');
    const markerSection = document.getElementById('markerSection');
    if (!appContainer || !viewMode) return;
    viewMode.addEventListener('change', function() {
        if (this.value === 'stacked') {
            appContainer.classList.add('stacked');
            if (stackedResizeHandle) stackedResizeHandle.style.display = '';
            videoSection.style.height = '50vh';
            markerSection.style.height = '50vh';
        } else {
            appContainer.classList.remove('stacked');
            if (stackedResizeHandle) stackedResizeHandle.style.display = 'none';
            videoSection.style.height = '';
            markerSection.style.height = '';
        }
    });
    // Initialize stacked resize handle
    if (stackedResizeHandle) {
        let isResizing = false;
        let startY, startVideoHeight, startMarkerHeight;
        stackedResizeHandle.addEventListener('mousedown', function(e) {
            if (!appContainer.classList.contains('stacked')) return;
            isResizing = true;
            startY = e.clientY;
            startVideoHeight = videoSection.offsetHeight;
            startMarkerHeight = markerSection.offsetHeight;
            document.body.style.cursor = 'row-resize';
            document.addEventListener('mousemove', resizeStackedPanels);
            document.addEventListener('mouseup', stopResizeStackedPanels);
        });
        function resizeStackedPanels(e) {
            if (!isResizing) return;
            const deltaY = e.clientY - startY;
            let newVideoHeight = startVideoHeight + deltaY;
            let newMarkerHeight = startMarkerHeight - deltaY;
            const minHeight = 120;
            if (newVideoHeight < minHeight || newMarkerHeight < minHeight) return;
            videoSection.style.height = newVideoHeight + 'px';
            markerSection.style.height = newMarkerHeight + 'px';
        }
        function stopResizeStackedPanels() {
            isResizing = false;
            document.body.style.cursor = '';
            document.removeEventListener('mousemove', resizeStackedPanels);
            document.removeEventListener('mouseup', stopResizeStackedPanels);
        }
    }
}

function initializeAddColumn() {
    const addColumnBtn = document.getElementById('addColumnBtn');
    if (!addColumnBtn) return;
    addColumnBtn.addEventListener('click', function() {
        const colName = prompt('Enter column name:');
        if (!colName) return;
        const colType = prompt('Enter column type (text, dropdown, button):', 'text');
        if (!colType) return;
        let options = [];
        if (colType === 'dropdown') {
            const opts = prompt('Enter dropdown options (comma separated):');
            if (!opts) return;
            options = opts.split(',').map(s => s.trim()).filter(Boolean);
        }
        extraColumns.push({ name: colName, type: colType, options });
        // Add to all markers
        markers.forEach(m => { m[colName] = ''; });
        updateMarkerTable();
        updateMarkerTableHeader();
    });
}

function updateMarkerTableHeader() {
    const headerRow = document.querySelector('.table thead tr:first-child');
    if (!headerRow) return;
    // Remove any extra columns first
    while (headerRow.children.length > 13) headerRow.removeChild(headerRow.lastChild);
    // Add extra columns
    extraColumns.forEach(col => {
        const th = document.createElement('th');
        th.textContent = col.name;
        th.className = 'special';
        th.dataset.field = col.name.toLowerCase().replace(/[^a-z0-9]/g, '');
        headerRow.appendChild(th);
    });
}

// When adding a new row, auto-fill active columns only if they haven't been manually edited
function addMarkerRow(newMarker) {
    // For each special column, if active, copy from first row only if the field is empty
    const specialColumns = ['filmTitle', 'composer', 'lyricist', 'musicCo', 'nocId', 'nocTitle', 'title'];
    specialColumns.forEach(col => {
        if (activePasteColumns[col] && markers.length > 0 && !newMarker[col]) {
            newMarker[col] = markers[0][col] || '';
        }
    });
    // Allow up to 5 usage markers per row
    if (newMarker.usage && !Array.isArray(newMarker.usage)) {
        newMarker.usage = [newMarker.usage];
    }
    markers.push(newMarker);
    updateMarkerTable();
}

function initializeCueSheetUpload() {
    const loadCueBtn = document.getElementById('loadCueBtn');
    const cueFileInput = document.getElementById('cueFileInput');
    const cueActions = document.getElementById('cueActions');
    const loadStructureBtn = document.getElementById('loadStructureBtn');
    const loadDataBtn = document.getElementById('loadDataBtn');

    loadCueBtn.addEventListener('click', () => cueFileInput.click());
    cueFileInput.addEventListener('change', function(e) {
        if (e.target.files.length > 0) {
            cueSheetFile = e.target.files[0];
            cueActions.style.display = '';
        } else {
            cueActions.style.display = 'none';
        }
    });

    function resetCueSheetUpload() {
        cueFileInput.value = '';
        cueSheetFile = null;
        cueActions.style.display = 'none';
    }

    loadStructureBtn.addEventListener('click', function() {
        if (!cueSheetFile) return;
        parseCueSheetFile('structure');
        resetCueSheetUpload();
    });
    loadDataBtn.addEventListener('click', function() {
        if (!cueSheetFile) return;
        parseCueSheetFile('data');
        resetCueSheetUpload();
    });
}

// Add this function near the top with other utility functions
function clearAppData() {
    // Clear markers and related data
    markers = [];
    headerRows = [];
    extraColumns = [];
    
    // Clear marked rows and exception settings
    markedRows = {};
    exceptionSettings = {};
    localStorage.removeItem('markedRows');
    localStorage.removeItem('exceptionSettings');
    
    // Clear manual edits and paste mode data
    manualEdits = {};
    activePasteColumns = {};
    
    // Clear sequence state
    isSequenceReversed = false;
    
    // Clear cue sheet data
    cueSheetParsed = null;
    cueSheetFile = null;
    
    // Clear localStorage data
    localStorage.removeItem('headerRows');
    localStorage.removeItem('markers');
    localStorage.removeItem('showInfo');
    
    // Reset usage counts
    usageCounts = { BI: 0, BV: 0, VI: 0, VV: 0, SRC: 0, 'BI,BV': 0, 'VI,VV': 0, 'BI,VV': 0, 'VI,BV': 0 };
    
    console.log('Cleared all application data');
}

function parseCueSheetFile(mode) {
    if (!cueSheetFile) return;
    const formData = new FormData();
    formData.append('file', cueSheetFile);
    fetch('/api/parse-cue-sheet', {
        method: 'POST',
        body: formData
    })
    .then(resp => resp.json())
    .then(result => {
        if (result.error) {
            alert('Error parsing file: ' + result.error);
            return;
        }
        headerRows = result.metadata || [];
        // --- Extract Series Title and Episode Number for export file name ---
        let seriesTitle = '';
        let episodeNumber = '';
        for (const row of headerRows) {
            for (let i = 0; i < row.length; i++) {
                if (row[i] && row[i].toLowerCase().replace(/[^a-z0-9]/g, '') === 'seriestitle') {
                    for (let j = i + 1; j < row.length; j++) {
                        if (row[j] && row[j].trim() !== '') {
                            seriesTitle = row[j].trim();
                            break;
                        }
                    }
                }
                if (row[i] && row[i].toLowerCase().replace(/[^a-z0-9]/g, '') === 'episodenumber') {
                    for (let j = i + 1; j < row.length; j++) {
                        if (row[j] && row[j].trim() !== '') {
                            episodeNumber = row[j].trim();
                            break;
                        }
                    }
                }
            }
        }
        if (seriesTitle && episodeNumber) {
            const paddedEp = episodeNumber.padStart(4, '0');
            const exportFileName = `${seriesTitle}_${paddedEp}_Unmix HD_MusicCueSheet`;
            let exportSettings = JSON.parse(localStorage.getItem('exportSettings')) || {};
            exportSettings.fileName = exportFileName;
            localStorage.setItem('exportSettings', JSON.stringify(exportSettings));
            // Update export settings UI if present
            const fileNameInput = document.getElementById('fileName');
            if (fileNameInput) fileNameInput.value = exportFileName;
        }
        // --- End file name extraction ---
        if (mode === 'structure') {
            loadCueSheetStructure(result.header);
        } else if (mode === 'data') {
            loadCueSheetData(result.header, result.data);
        }
    })
    .catch(err => {
        alert('Failed to parse file.');
    });
}

function loadCueSheetStructure(header) {
    // Map file columns to app columns using mapping
    extraColumns = [];
    let mappedCols = [];
    let mappedKeys = [];
    let lowerHeader = header.map(h => h.trim().toLowerCase());
    defaultMarkerColumns.forEach(col => {
        let idx = lowerHeader.findIndex(h => columnNameMap[h] === col.key || h === col.label.toLowerCase());
        mappedCols.push(idx !== -1 ? header[idx] : col.label);
        mappedKeys.push(col.key);
    });
    // Add extra columns from file
    header.forEach((col, i) => {
        let key = columnNameMap[col.trim().toLowerCase()] || col.trim();
        if (!mappedKeys.includes(key) && col.trim().toLowerCase() !== 'recognize') {
            extraColumns.push({ 
                name: col, 
                type: 'text', 
                options: [],
                field: col.toLowerCase().replace(/[^a-z0-9]/g, '')
            });
        }
    });
    updateMarkerTableHeader();
    updateMarkerTable();
}

function extractFieldValue(headerRows, fieldName) {
    for (const row of headerRows) {
        if (row[0] && row[0].toLowerCase().includes(fieldName)) {
            // Search for the first non-empty cell after the field name
            for (let i = 1; i < row.length; i++) {
                if (row[i] && row[i].trim() !== '') {
                    return row[i].trim();
                }
            }
        }
    }
    return '';
}

function loadCueSheetData(header, data) {
    // Map file columns to app columns using mapping
    let lowerHeader = header.map(h => h.trim().toLowerCase());
    console.log('Lowercase header:', lowerHeader);
    
    // Extract show name (series title) from headerRows
    const showName = extractFieldValue(headerRows, 'series title');
    console.log('Show name:', showName);
    
    // Log the first row of data for debugging
    if (data.length > 0) {
        console.log('First row of data:', data[0]);
    }
    
    markers = data.map(row => {
        let marker = {};
        defaultMarkerColumns.forEach(col => {
            let idx = lowerHeader.findIndex(h => columnNameMap[h] === col.key || h === col.label.toLowerCase());
            console.log(`Mapping column ${col.key}:`, {
                'header': header[idx],
                'lowerHeader': lowerHeader[idx],
                'found': idx !== -1,
                'value': idx !== -1 ? row[header[idx]] || '' : ''
            });
            marker[col.key] = idx !== -1 ? row[header[idx]] || '' : '';
        });
        // Add extra columns
        extraColumns.forEach(col => {
            marker[col.name] = row[col.name] || '';
        });
        // Normalize TCR fields
        ['tcrIn', 'tcrOut'].forEach(field => {
            if (marker[field] && marker[field].match(/^[0-9]{2}:[0-9]{2}:[0-9]{2}$/)) {
                marker[field] += ':00';
            }
        });
        // Set filmTitle to showName if available
        if (showName) {
            marker.filmTitle = showName;
        }
        return marker;
    });

    // Extract and save show information
    const showInfo = {
        showName: showName,
        season: extractFieldValue(headerRows, 'season'),
        episodeNumber: extractFieldValue(headerRows, 'episode number')
    };

    // --- BEGIN: Logging for show info extraction ---
    if (!showInfo.showName || showInfo.showName === 'Unknown_Show') {
        console.warn('Show name not found in imported file. Using default.');
    } else {
        console.log('Show name found:', showInfo.showName);
    }
    if (!showInfo.season) {
        console.warn('Season not found in imported file.');
    } else {
        console.log('Season found:', showInfo.season);
    }
    if (!showInfo.episodeNumber) {
        console.warn('Episode number not found in imported file.');
    } else {
        console.log('Episode number found:', showInfo.episodeNumber);
    }
    // --- END: Logging for show info extraction ---

    // Save show information to localStorage
    localStorage.setItem('showInfo', JSON.stringify(showInfo));

    updateMarkerTable();

    // --- BEGIN: Save metadata and marker data to localStorage for export ---
    try {
        localStorage.setItem('headerRows', JSON.stringify(headerRows));
        localStorage.setItem('markers', JSON.stringify(markers));
        console.log('Saved headerRows and markers to localStorage for export.');
    } catch (e) {
        console.error('Failed to save headerRows/markers to localStorage:', e);
    }
    // --- END: Save metadata and marker data to localStorage for export ---
}

function initializeClearTableButton() {
    const clearTableBtn = document.getElementById('clearTableBtn');
    if (clearTableBtn) {
        clearTableBtn.addEventListener('click', () => {
            if (confirm('Are you sure you want to clear the table? This will remove all markers and settings.')) {
                clearAppData();
                updateMarkerTable();
            }
        });
    }
}

// Add this function to handle post-export cleanup
function handlePostExport() {
    if (confirm('Export completed. Would you like to clear the current data and start fresh?')) {
        clearAppData();
        updateMarkerTable();
    }
}

// Modify the exportToExcelWorkbook function
function exportToExcelWorkbook() {
    // ... existing export code ...
    
    // After successful export
    handlePostExport();
}

// Modify the exportToCSV function
function exportToCSV() {
    // ... existing export code ...
    
    // After successful export
    handlePostExport();
}

// Modify the exportToPlainExcel function
function exportToPlainExcel() {
    // ... existing export code ...
    
    // After successful export
    handlePostExport();
}

// Add this function to handle new file loading
function handleNewFileLoad() {
    if (markers.length > 0 || Object.keys(markedRows).length > 0) {
        if (confirm('Loading a new file. Would you like to clear the current data first?')) {
            clearAppData();
        }
    }
}

// Restore the original handleSeqClick function for the triple-click delete functionality
function handleSeqClick(rowIdx, cell) {
    if (seqClickState.row !== rowIdx) {
        // Reset state if a different row is clicked
        resetSeqClickState();
        seqClickState.row = rowIdx;
        seqClickState.count = 0;
    }
    
    seqClickState.count++;
    
    if (seqClickState.count === 1) {
        cell.style.backgroundColor = 'yellow';
    } else if (seqClickState.count === 2) {
        cell.style.backgroundColor = 'red';
    } else if (seqClickState.count === 3) {
        // Delete the row on third click
        markers.splice(rowIdx, 1);
        updateMarkerTable();
        resetSeqClickState();
        return;
    }
    
    // Reset after 2 seconds if no further click
    clearTimeout(seqClickState.timeout);
    seqClickState.timeout = setTimeout(() => {
        resetSeqClickState();
    }, 2000);
}

function resetSeqClickState() {
    // Remove highlight from all seq cells
    document.querySelectorAll('.seq-cell').forEach(cell => {
        cell.style.backgroundColor = '';
    });
    seqClickState = { row: null, count: 0, timeout: null };
}

// Simplified applyTimeOffset function to apply to all rows
function applyTimeOffset() {
    // Get the offset value
    const offsetValue = document.getElementById('offsetTime').value;
    const offsetDirection = document.querySelector('input[name="offsetDirection"]:checked').value;
    
    // Convert offset to seconds
    const offsetSeconds = timeToSeconds(offsetValue);
    
    // Determine the actual offset based on direction
    const actualOffset = offsetDirection === 'add' ? offsetSeconds : -offsetSeconds;
    
    // Apply offset to each marker
    markers.forEach(marker => {
        if (marker.tcrIn) {
            const inSeconds = timeToSeconds(marker.tcrIn) + actualOffset;
            marker.tcrIn = formatTime(Math.max(0, inSeconds)); // Ensure we don't go negative
        }
        if (marker.tcrOut) {
            const outSeconds = timeToSeconds(marker.tcrOut) + actualOffset;
            marker.tcrOut = formatTime(Math.max(0, outSeconds)); // Ensure we don't go negative
        }
        // Update duration
        marker.duration = calculateDuration(marker.tcrIn, marker.tcrOut);
    });
    
    // Update the table
    updateMarkerTable();
}

function initializeOffsetModal() {
    const offsetBtn = document.getElementById('offsetBtn');
    const offsetModal = document.getElementById('offsetModal');
    const closeBtn = document.querySelector('.offset-modal-close');
    const applyOffsetBtn = document.getElementById('applyOffset');

    // Open modal when offset button is clicked
    offsetBtn.addEventListener('click', function() {
        offsetModal.style.display = 'flex';
    });

    // Close modal when close button is clicked
    closeBtn.addEventListener('click', function() {
        offsetModal.style.display = 'none';
    });

    // Close modal when clicking outside the modal content
    offsetModal.addEventListener('click', function(event) {
        if (event.target === offsetModal) {
            offsetModal.style.display = 'none';
        }
    });

    // Apply the offset when apply button is clicked
    applyOffsetBtn.addEventListener('click', function() {
        applyTimeOffset();
        offsetModal.style.display = 'none';
    });
}

function showCopyDropdown(e, columnName) {
    if (!activePasteColumns[columnName]) return; // Only allow if in copy mode
    const header = e.target;
    const dropdown = document.createElement('div');
    dropdown.className = 'copy-dropdown';
    dropdown.innerHTML = `
        <div class="copy-title">Copy from row:</div>
        <div class="copy-options">
            ${markers.map((marker, index) => `
                <div class="copy-option" data-index="${index}">
                    Row ${index + 1}
                </div>
            `).join('')}
        </div>
        <button class="close-dropdown">Close</button>
    `;

    // Position the dropdown below the header
    const rect = header.getBoundingClientRect();
    dropdown.style.position = 'absolute';
    dropdown.style.top = `${rect.bottom}px`;
    dropdown.style.left = `${rect.left}px`;

    // Add click event to options
    dropdown.querySelectorAll('.copy-option').forEach(option => {
        option.addEventListener('click', () => {
            const sourceIndex = parseInt(option.dataset.index);
            const sourceValue = markers[sourceIndex][columnName];
            
            // Copy to all other rows in the same column, respecting manual edits and blank values
            markers.forEach((marker, index) => {
                if (index !== sourceIndex) {
                    // Only update if:
                    // 1. The field hasn't been manually edited
                    // 2. The field is empty (not blanked by user)
                    // 3. The field hasn't been previously copied to
                    if (!manualEdits[columnName] || !manualEdits[columnName][index]) {
                        // If the current value is empty (not blanked by user)
                        if (!marker[columnName] || marker[columnName].trim() === '') {
                            marker[columnName] = sourceValue;
                        }
                    }
                }
            });
            updateMarkerTable();
            dropdown.remove();
        });
    });

    // Close dropdown when clicking outside
    document.addEventListener('click', function closeDropdown(e) {
        if (!dropdown.contains(e.target)) {
            dropdown.remove();
            document.removeEventListener('click', closeDropdown);
        }
    });

    // Close button
    dropdown.querySelector('.close-dropdown').addEventListener('click', () => {
        dropdown.remove();
    });

    document.body.appendChild(dropdown);
}

function deleteSelectedRows() {
    const checkboxes = document.querySelectorAll('.row-checkbox:checked');
    if (checkboxes.length === 0) {
        alert('Please select rows to delete');
        return;
    }
    
    if (!confirm(`Are you sure you want to delete ${checkboxes.length} selected row(s)?`)) {
        return;
    }
    
    // Save state before making changes
    saveToHistory();
    
    // Get indices of selected rows
    const selectedIndices = Array.from(checkboxes).map(checkbox => 
        parseInt(checkbox.closest('tr').querySelector('.seq-cell').textContent) - 1
    );
    
    // Remove selected markers
    markers = markers.filter((_, index) => !selectedIndices.includes(index));
    
    // Update the table
    updateMarkerTable();
}

function initializeSequenceDirection() {
    // This function is now obsolete, sequence direction is toggled by double-clicking Seq header
}

function initializeExportSettings() {
    console.log('Initializing export settings...');
    const exportSettingsBtn = document.getElementById('exportSettingsBtn');
    const exportBtn = document.getElementById('exportBtn');
    
    if (exportSettingsBtn) {
        console.log('Found export settings button, adding click handler');
        exportSettingsBtn.addEventListener('click', () => {
            console.log('Opening export settings window...');
            try {
                // Save current state to localStorage
                localStorage.setItem('markers', JSON.stringify(markers));
                localStorage.setItem('headerRows', JSON.stringify(headerRows || []));
                
                // Open the settings window
                const settingsWindow = window.open('/export-settings', 'Export Settings', 'width=800,height=600');
                
                if (!settingsWindow) {
                    alert('Please allow popups for this site to use the export settings.');
                    return;
                }
                
                // Add event listener for when the settings window closes
                const checkWindow = setInterval(() => {
                    if (settingsWindow.closed) {
                        clearInterval(checkWindow);
                        console.log('Export settings window closed');
                        // Reload settings if needed
                        const savedSettings = localStorage.getItem('exportSettings');
                        if (savedSettings) {
                            console.log('Reloading saved export settings');
                        }
                    }
                }, 500);
            } catch (error) {
                console.error('Error opening export settings:', error);
                alert('Error opening export settings. Please try again.');
            }
        });
    } else {
        console.warn('Export settings button not found');
    }

    if (exportBtn) {
        console.log('Found export button, adding click handler');
        exportBtn.addEventListener('click', () => {
            console.log('Starting export process...');
            try {
                // Always save latest data before exporting
                localStorage.setItem('markers', JSON.stringify(markers));
                localStorage.setItem('headerRows', JSON.stringify(headerRows || []));
                exportWithSettings();
            } catch (error) {
                console.error('Error during export:', error);
                alert('Error during export. Please try again.');
            }
        });
    } else {
        console.warn('Export button not found');
    }
}

function initializeRowMarking() {
    // Keep only the exception modal functionality
    const rowMarkBtn = document.getElementById('rowMarkBtn');
    const exceptionModal = document.getElementById('exceptionModal');
    
    // Exception modal elements
    const exceptionCloseBtn = exceptionModal.querySelector('.close');
    const applyExceptionBtn = document.getElementById('applyException');
    const cancelExceptionBtn = document.getElementById('cancelException');

    // Update the row mark button to only handle exceptions
    if (rowMarkBtn) {
        rowMarkBtn.addEventListener('click', () => {
            const selectedRows = document.querySelectorAll('.row-checkbox:checked');
            if (selectedRows.length === 0) {
                alert('Please select rows to mark as exceptions');
                return;
            }
            exceptionModal.style.display = 'block';
        });
    }

    // Exception modal event handlers
    exceptionCloseBtn.addEventListener('click', () => {
        exceptionModal.style.display = 'none';
    });

    applyExceptionBtn.addEventListener('click', () => {
        const filmTitleException = document.getElementById('exceptionFilmTitle').checked;
        const titlePrefixException = document.getElementById('exceptionTitlePrefix').checked;
        
        markSelectedRowsWithException('exception', {
            filmTitle: filmTitleException,
            titlePrefix: titlePrefixException
        });
        exceptionModal.style.display = 'none';
    });

    cancelExceptionBtn.addEventListener('click', () => {
        exceptionModal.style.display = 'none';
    });

    // Close modal when clicking outside
    window.addEventListener('click', (e) => {
        if (e.target === exceptionModal) {
            exceptionModal.style.display = 'none';
        }
    });
}

function markSelectedRows(color) {
    const checkboxes = document.querySelectorAll('.row-checkbox:checked');
    checkboxes.forEach(checkbox => {
        const row = checkbox.closest('tr');
        const rowIndex = parseInt(row.querySelector('.seq-cell').textContent) - 1;
        if (!markedRows[rowIndex]) markedRows[rowIndex] = {};
        if (color) {
            markedRows[rowIndex][color] = true;
        } else {
            delete markedRows[rowIndex];
            delete exceptionSettings[rowIndex];
        }
    });
    saveMarkedRows();
    updateMarkerTable();
}

function markSelectedRowsWithException(color, exceptionConfig) {
    const checkboxes = document.querySelectorAll('.row-checkbox:checked');
    checkboxes.forEach(checkbox => {
        const row = checkbox.closest('tr');
        const rowIndex = parseInt(row.querySelector('.seq-cell').textContent) - 1;
        if (!markedRows[rowIndex]) markedRows[rowIndex] = {};
        if (color) {
            markedRows[rowIndex][color] = true;
            exceptionSettings[rowIndex] = exceptionConfig;
        } else {
            delete markedRows[rowIndex];
            delete exceptionSettings[rowIndex];
        }
    });
    saveMarkedRows();
    updateMarkerTable();
}

function initializeColumnResize() {
    const table = document.querySelector('.table');
    const headers = table.querySelectorAll('th');
    
    headers.forEach(header => {
        let startX, startWidth;
        let resizing = false;
        let handle = header.querySelector('.resize-handle');
        
        if (!handle) {
            handle = document.createElement('div');
            handle.className = 'resize-handle';
            header.appendChild(handle);
        }

        handle.addEventListener('mousedown', (e) => {
            e.preventDefault();
            resizing = true;
            startX = e.pageX;
            startWidth = header.offsetWidth;
            handle.classList.add('active');
            document.body.style.cursor = 'col-resize';
            document.addEventListener('mousemove', resize);
            document.addEventListener('mouseup', stopResize);
        });

        function resize(e) {
            if (!resizing) return;
            const width = Math.max(50, startWidth + (e.pageX - startX));
            const columnIndex = Array.from(header.parentNode.children).indexOf(header);
            
            // Set width for header
            header.style.width = `${width}px`;
            header.style.setProperty('--column-width', `${width}px`);
            
            // Set width for all cells in this column
            document.querySelectorAll(`.table tr`).forEach(row => {
                const cell = row.children[columnIndex];
                if (cell) {
                    cell.style.width = `${width}px`;
                    cell.style.setProperty('--column-width', `${width}px`);
                    
                    // Adjust input width if present
                    const input = cell.querySelector('.table-input');
                    if (input) {
                        input.style.width = '100%';
                    }
                }
            });
        }

        function stopResize() {
            resizing = false;
            handle.classList.remove('active');
            document.body.style.cursor = '';
            document.removeEventListener('mousemove', resize);
            document.removeEventListener('mouseup', stopResize);
        }
    });
}

function setupSeqHeaderDoubleClick() {
    const seqHeader = document.querySelector('th[data-field="seq"]');
    if (seqHeader) {
        seqHeader.addEventListener('dblclick', () => {
            isSequenceReversed = !isSequenceReversed;
            updateMarkerTable();
        });
    }
}

function exportWithSettings() {
    console.log('Starting export with settings...');
    
    // Check if required functions exist
    const requiredFunctions = [
        'exportToExcelWorkbook',
        'exportToCSV',
        'exportToPlainExcel',
        'prepareExportData'
    ];
    
    const missingFunctions = requiredFunctions.filter(func => typeof window[func] !== 'function');
    
    if (missingFunctions.length > 0) {
        console.error('Missing required functions:', missingFunctions);
        alert('Export functionality is not properly loaded. Please refresh the page.');
        return;
    }
    
    try {
        // Get settings from localStorage or default
        const settings = JSON.parse(localStorage.getItem('exportSettings')) || {
            fileName: '',
            fileType: 'excel',
            downloadLocation: 'ask',
            customLocation: '',
            includeHeader: true,
            fieldsToExport: [],
            tcrFormat: 'timecode'
        };
        
        console.log('Using export settings:', settings);
        
        // Make sure window.markers and window.headerRows are up to date
        window.markers = markers;
        window.headerRows = headerRows;
        
        // Prepare the data based on settings
        const data = prepareExportData(settings);
        console.log('Prepared data for export:', data);
        
        // Export based on file type
        switch (settings.fileType) {
            case 'excel':
                console.log('Exporting to Excel workbook...');
                exportToExcelWorkbook(data, settings);
                break;
            case 'csv':
                console.log('Exporting to CSV...');
                exportToCSV(data, settings);
                break;
            case 'plain':
                console.log('Exporting to plain Excel...');
                exportToPlainExcel(data, settings);
                break;
            default:
                console.error('Unknown export file type:', settings.fileType);
                alert('Unknown export file type. Please check your export settings.');
        }
    } catch (error) {
        console.error('Error during export:', error);
        alert('Error during export: ' + error.message);
    }
}

// --- Add markColor to marker data for export ---
function getMarkersWithMarkColor() {
    return markers.map((marker, idx) => {
        const mark = markedRows[idx] || {};
        let markColor = '';
        if (mark.yellow && mark.red && mark.exception) {
            markColor = 'yellow,red,exception';
        } else if (mark.yellow && mark.red) {
            markColor = 'yellow,red';
        } else if (mark.yellow && mark.exception) {
            markColor = 'yellow,exception';
        } else if (mark.red && mark.exception) {
            markColor = 'red,exception';
        } else if (mark.yellow) {
            markColor = 'yellow';
        } else if (mark.red) {
            markColor = 'red';
        } else if (mark.exception) {
            markColor = 'exception';
        }
        return { ...marker, markColor: markColor };
    });
}

// Add this function to load usage options from Firestore
function loadUsageOptions() {
    fetch('/api/usage')
        .then(response => response.json())
        .then(data => {
            usageOptions = data.map(item => item.name);
            // Update usage counts after loading options
            updateUsageCounts();
            // Refresh the table to show sorted options
            updateMarkerTable();
        })
        .catch(error => console.error('Error loading usage options:', error));
}

// Add this function to load music co options from Firestore
function loadMusicCoOptions() {
    fetch('/api/music-co')
        .then(response => response.json())
        .then(data => {
            musicCoOptions = data.map(item => item.name);
            updateMusicCoDropdown();
        })
        .catch(error => console.error('Error loading music co options:', error));
}

// Update the usage dropdown
function updateUsageDropdown() {
    const usageCells = document.querySelectorAll('.usage-cell');
    usageCells.forEach(cell => {
        const currentValue = cell.textContent;
        const input = document.createElement('input');
        input.type = 'text';
        input.className = 'form-control usage-input';
        input.value = currentValue;
        input.list = 'usageOptions';
        input.addEventListener('input', function(e) {
            const value = e.target.value;
            const datalist = document.getElementById('usageOptions');
            datalist.innerHTML = '';
            usageOptions
                .filter(option => option.toLowerCase().includes(value.toLowerCase()))
                .forEach(option => {
                    const optionElement = document.createElement('option');
                    optionElement.value = option;
                    datalist.appendChild(optionElement);
                });
        });
        cell.innerHTML = '';
        cell.appendChild(input);
    });
}

// Update the music co dropdown
function updateMusicCoDropdown() {
    const musicCoCells = document.querySelectorAll('.music-co-cell');
    musicCoCells.forEach(cell => {
        const currentValue = cell.textContent;
        const container = document.createElement('div');
        container.className = 'dropdown-container';
        
        const input = document.createElement('input');
        input.type = 'text';
        input.className = 'form-control music-co-input';
        input.value = currentValue;
        input.placeholder = 'Search Music Co...';
        
        const dropdown = document.createElement('div');
        dropdown.className = 'custom-dropdown';
        
        // Initially populate with all options
        updateDropdownOptions(dropdown, musicCoOptions, '');
        
        // Show dropdown on focus
        input.addEventListener('focus', () => {
            dropdown.style.display = 'block';
            updateDropdownOptions(dropdown, musicCoOptions, input.value);
        });
        
        // Filter options on input with improved search
        input.addEventListener('input', (e) => {
            const value = e.target.value.toLowerCase();
            const words = value.split(/\s+/).filter(word => word.length > 0);
            
            const filteredOptions = musicCoOptions.filter(option => {
                const optionLower = option.toLowerCase();
                // Match if all words are found in the option
                return words.every(word => optionLower.includes(word));
            });
            
            updateDropdownOptions(dropdown, filteredOptions, value);
            dropdown.style.display = 'block';
        });
        
        // Handle option selection
        dropdown.addEventListener('click', (e) => {
            if (e.target.classList.contains('dropdown-option')) {
                input.value = e.target.textContent;
                dropdown.style.display = 'none';
                // Update the marker data
                const rowIndex = findRowIndex(cell);
                if (rowIndex !== -1) {
                    markers[rowIndex].musicCo = input.value;
                }
            }
        });
        
        // Handle keyboard navigation
        input.addEventListener('keydown', (e) => {
            const options = dropdown.querySelectorAll('.dropdown-option');
            const currentIndex = Array.from(options).findIndex(opt => opt.classList.contains('selected'));
            
            switch(e.key) {
                case 'ArrowDown':
                    e.preventDefault();
                    if (currentIndex < options.length - 1) {
                        options[currentIndex]?.classList.remove('selected');
                        options[currentIndex + 1].classList.add('selected');
                        options[currentIndex + 1].scrollIntoView({ block: 'nearest' });
                    }
                    break;
                case 'ArrowUp':
                    e.preventDefault();
                    if (currentIndex > 0) {
                        options[currentIndex]?.classList.remove('selected');
                        options[currentIndex - 1].classList.add('selected');
                        options[currentIndex - 1].scrollIntoView({ block: 'nearest' });
                    }
                    break;
                case 'Enter':
                    e.preventDefault();
                    const selectedOption = dropdown.querySelector('.dropdown-option.selected');
                    if (selectedOption) {
                        input.value = selectedOption.textContent;
                        dropdown.style.display = 'none';
                        const rowIndex = findRowIndex(cell);
                        if (rowIndex !== -1) {
                            markers[rowIndex].musicCo = input.value;
                        }
                    }
                    break;
                case 'Escape':
                    e.preventDefault();
                    dropdown.style.display = 'none';
                    break;
            }
        });
        
        // Close dropdown when clicking outside
        document.addEventListener('click', (e) => {
            if (!container.contains(e.target)) {
                dropdown.style.display = 'none';
            }
        });
        
        container.appendChild(input);
        container.appendChild(dropdown);
        cell.innerHTML = '';
        cell.appendChild(container);
    });
}

// Helper function to update dropdown options
function updateDropdownOptions(dropdown, options, filterText) {
    dropdown.innerHTML = '';
    if (options.length === 0) {
        const noResults = document.createElement('div');
        noResults.className = 'dropdown-option no-results';
        noResults.textContent = 'No matches found';
        dropdown.appendChild(noResults);
        return;
    }
    
    options.forEach(option => {
        const optionElement = document.createElement('div');
        optionElement.className = 'dropdown-option';
        
        // Highlight matching text
        if (filterText) {
            const regex = new RegExp(`(${filterText})`, 'gi');
            optionElement.innerHTML = option.replace(regex, '<mark>$1</mark>');
        } else {
            optionElement.textContent = option;
        }
        
        dropdown.appendChild(optionElement);
    });
}

// Helper function to find row index from cell
function findRowIndex(cell) {
    const row = cell.closest('tr');
    const tbody = row.closest('tbody');
    return Array.from(tbody.children).indexOf(row);
}

// Add datalists to the document
function addDatalists() {
    const usageDatalist = document.createElement('datalist');
    usageDatalist.id = 'usageOptions';
    document.body.appendChild(usageDatalist);

    const musicCoDatalist = document.createElement('datalist');
    musicCoDatalist.id = 'musicCoOptions';
    document.body.appendChild(musicCoDatalist);
}

// --- Session-based autosave ---
function autoSave() {
    // Store video state
    let videoState = null;
    if (currentVideo) {
        videoState = {
            currentTime: currentVideo.currentTime,
            src: currentVideo.src,
            type: 'url'
        };
    } else if (currentVideoFile) {
        videoState = {
            currentTime: 0, // Reset to start for file-based videos
            name: currentVideoFile.name,
            size: currentVideoFile.size,
            type: 'file'
        };
    }

    fetch('/api/autosave', {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({
            markers,
            videoState,
            exportSettings: JSON.parse(localStorage.getItem('exportSettings') || '{}'),
            markedRows: JSON.parse(localStorage.getItem('markedRows') || '{}'),
            exceptionSettings: JSON.parse(localStorage.getItem('exceptionSettings') || '{}')
        })
    })
    .then(response => {
        if (!response.ok) {
            return response.json().then(err => {
                throw new Error(err.error || 'Autosave failed');
            });
        }
    })
    .catch(error => {
        console.error('Autosave failed:', error);
        // Don't show alert for autosave failures to avoid disrupting the user
    });
}

// Set up autosave interval
setInterval(autoSave, 30000); // Save every 30 seconds

// Restore from session autosave on page load
window.addEventListener('DOMContentLoaded', function() {
    fetch('/api/loadsave')
        .then(response => {
            if (!response.ok) {
                return response.json().then(err => {
                    throw new Error(err.error || 'Failed to load saved data');
                });
            }
            return response.json();
        })
        .then(data => {
            if (data && data.markers && data.markers.length > 0) {
                if (confirm('Restore your last saved work from this session?')) {
                    markers = data.markers;
                    
                    // Restore export settings
                    if (data.exportSettings) {
                        localStorage.setItem('exportSettings', JSON.stringify(data.exportSettings));
                    }
                    
                    // Restore marked rows and exception settings
                    if (data.markedRows) {
                        localStorage.setItem('markedRows', JSON.stringify(data.markedRows));
                        markedRows = data.markedRows;
                    }
                    if (data.exceptionSettings) {
                        localStorage.setItem('exceptionSettings', JSON.stringify(data.exceptionSettings));
                        exceptionSettings = data.exceptionSettings;
                    }
                    
                    // Restore video state if available
                    if (data.videoState) {
                        if (data.videoState.type === 'url' && data.videoState.src) {
                            loadVideoFromURL(data.videoState.src).then(() => {
                                if (currentVideo && data.videoState.currentTime) {
                                    currentVideo.currentTime = data.videoState.currentTime;
                                }
                            });
                        }
                        // Note: File-based videos can't be automatically restored due to security restrictions
                    }
                    
                    updateMarkerTable();
                }
            }
        })
        .catch(error => {
            console.error('Failed to load saved data:', error);
            // Don't show alert for load failures to avoid disrupting initial page load
        });
});

// Helper function to load video from URL
function loadVideoFromURL(url) {
    return new Promise((resolve, reject) => {
        const video = document.createElement('video');
        video.src = url;
        video.onloadedmetadata = () => {
            currentVideo = video;
            currentVideoFile = null;
            initializeVideoPlayer();
            resolve();
        };
        video.onerror = reject;
    });
}

// Add missing updateDuration function
function updateDuration(rowIndex) {
    if (rowIndex >= 0 && rowIndex < markers.length) {
        const marker = markers[rowIndex];
        if (marker.tcrIn && marker.tcrOut) {
            marker.duration = calculateDuration(marker.tcrIn, marker.tcrOut);
            // Update the duration cell in the table
            const durationCell = document.querySelector(`[data-row="${rowIndex}"]`);
            if (durationCell) {
                durationCell.textContent = marker.duration || '';
            }
        }
    }
}

// Add this function after the existing click handling functions
function handleCheckboxCellClick(rowIndex, event) {
    const currentTime = Date.now();
    const timeDiff = currentTime - checkboxClickState.lastClickTime;
    
    // Add visual feedback
    const checkboxCell = event.target.closest('.checkbox-cell');
    if (checkboxCell) {
        checkboxCell.classList.add('clicked');
        setTimeout(() => {
            checkboxCell.classList.remove('clicked');
        }, 300);
    }
    
    // Reset state if clicking on a different row or if too much time has passed
    if (checkboxClickState.row !== rowIndex || timeDiff > 500) {
        resetCheckboxClickState();
        checkboxClickState.row = rowIndex;
        checkboxClickState.count = 0;
    }
    
    checkboxClickState.count++;
    checkboxClickState.lastClickTime = currentTime;
    
    // Save state before making changes
    saveToHistory();
    
    if (checkboxClickState.count === 2) {
        // Double-click: Toggle yellow mark
        if (!markedRows[rowIndex]) markedRows[rowIndex] = {};
        if (markedRows[rowIndex].yellow) {
            delete markedRows[rowIndex].yellow;
        } else {
            markedRows[rowIndex].yellow = true;
            // Remove red mark if yellow is being added
            delete markedRows[rowIndex].red;
        }
        if (Object.keys(markedRows[rowIndex]).length === 0) delete markedRows[rowIndex];
        saveMarkedRows();
        updateMarkerTable();
        resetCheckboxClickState();
    } else if (checkboxClickState.count === 3) {
        // Triple-click: Toggle red mark
        if (!markedRows[rowIndex]) markedRows[rowIndex] = {};
        if (markedRows[rowIndex].red) {
            delete markedRows[rowIndex].red;
        } else {
            markedRows[rowIndex].red = true;
            // Remove yellow mark if red is being added
            delete markedRows[rowIndex].yellow;
        }
        if (Object.keys(markedRows[rowIndex]).length === 0) delete markedRows[rowIndex];
        saveMarkedRows();
        updateMarkerTable();
        resetCheckboxClickState();
    }
    
    // Reset after 1 second if no further click
    clearTimeout(checkboxClickState.timeout);
    checkboxClickState.timeout = setTimeout(() => {
        resetCheckboxClickState();
    }, 1000);
}

function resetCheckboxClickState() {
    checkboxClickState = { row: null, count: 0, timeout: null, lastClickTime: 0 };
}
