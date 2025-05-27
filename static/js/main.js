let markers = [];
let currentVideo = null;
let currentVideoFile = null; // Store the File object for local videos
let isSingleButtonMode = false;
let activePasteColumns = {};
let activeCell = null;
let isNextMarkTcrIn = true; // For single button mode toggle
let usageOptions = ['BI', 'BV', 'VI', 'VV', 'SRC'];
let usageCounts = { BI: 0, BV: 0, VI: 0, VV: 0, SRC: 0 };
let extraColumns = [];
let seqClickState = { row: null, count: 0, timeout: null }; // For tracking triple clicks
let headerRows = [];

// Cue Sheet Upload & Integration
let cueSheetParsed = null;
let cueSheetFile = null;

// Add these variables at the top with other global variables
let frameTimer = null;
let frameRate = 30; // Default frame rate, will be updated when video loads
let isFrameByFrame = false;

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

// Initialize the application
document.addEventListener('DOMContentLoaded', function() {
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
});

function initializeVideoPlayer() {
    const videoPlayer = document.getElementById('videoPlayer');
    const tcrInBtn = document.getElementById('tcrInBtn');
    const tcrOutBtn = document.getElementById('tcrOutBtn');
    const singleButtonMode = document.getElementById('singleButtonMode');
    const loadVideoBtn = document.getElementById('loadVideoBtn');
    const videoFileInput = document.getElementById('videoFileInput');
    const videoUrlInput = document.getElementById('videoUrlInput');
    const loadUrlBtn = document.getElementById('loadUrlBtn');

    videoPlayer.addEventListener('loadedmetadata', function() {
        currentVideo = videoPlayer.src;
        // Get the video's frame rate
        frameRate = videoPlayer.getVideoPlaybackQuality()?.totalVideoFrames / videoPlayer.duration || 30;
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

    videoFileInput.addEventListener('change', function(e) {
        const file = e.target.files[0];
        if (file) {
            const videoURL = URL.createObjectURL(file);
            videoPlayer.src = videoURL;
            videoPlayer.load();
            currentVideoFile = file; // Store the File object
            currentVideo = videoURL;
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
            const columnIndex = Array.from(e.target.parentElement.children).indexOf(e.target);
            const columnName = getColumnName(columnIndex);
            
            // Toggle active paste mode for this column
            activePasteColumns[columnName] = !activePasteColumns[columnName];
            
            // Update header color to indicate active state
            e.target.style.backgroundColor = activePasteColumns[columnName] ? '#90EE90' : '';
            
            // Show copy dropdown
            showCopyDropdown(e, columnName);
        }
    });
    
    // Add input change handler to copy first row value when in paste mode
    tableBody.addEventListener('change', function(e) {
        if (activePasteColumns[e.target.dataset.field] && e.target.classList.contains('table-input')) {
            const columnIndex = Array.from(e.target.parentElement.children).indexOf(e.target);
            const columnName = getColumnName(columnIndex);
            
            if (columnName === e.target.dataset.field && markers.length > 0) {
                const firstRowValue = markers[0][columnName] || '';
                applyValueToAllRows(columnName, firstRowValue);
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
}

function initializeExportButtons() {
    document.getElementById('exportExcel').addEventListener('click', exportToExcel);
    document.getElementById('exportCSV').addEventListener('click', exportToCSV);
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

function markTCR(type) {
    const videoPlayer = document.getElementById('videoPlayer');
    const currentTime = videoPlayer.currentTime;
    
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
        isNextMarkTcrIn = !isNextMarkTcrIn; // Toggle for next click
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
    
    updateMarkerTable();
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
    input.style.resize = 'both';
    input.style.overflow = 'auto';
    input.style.minWidth = '100px';
    input.style.minHeight = '20px';
}

function updateMarkerTable() {
    const tableBody = document.getElementById('markerTableBody');
    tableBody.innerHTML = '';
    
    markers.forEach((marker, index) => {
        const row = document.createElement('tr');
        
        // Add checkbox cell
        const checkboxCell = document.createElement('td');
        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.className = 'row-checkbox';
        checkboxCell.appendChild(checkbox);
        row.appendChild(checkboxCell);
        
        // Add sequence number cell
        const seqCell = document.createElement('td');
        seqCell.textContent = index + 1;
        seqCell.className = 'seq-cell';
        seqCell.dataset.row = index;
        seqCell.addEventListener('click', () => handleSeqClick(index, seqCell));
        row.appendChild(seqCell);
        
        // Add TCR In cell
        const tcrInCell = document.createElement('td');
        const tcrInInput = document.createElement('input');
        tcrInInput.type = 'text';
        tcrInInput.className = 'table-input';
        tcrInInput.value = marker.tcrIn || '';
        tcrInInput.dataset.field = 'tcrIn';
        makeInputResizable(tcrInInput);
        tcrInInput.addEventListener('change', (e) => {
            marker.tcrIn = e.target.value;
            updateDuration(index);
        });
        tcrInCell.appendChild(tcrInInput);
        row.appendChild(tcrInCell);
        
        // Add TCR Out cell
        const tcrOutCell = document.createElement('td');
        const tcrOutInput = document.createElement('input');
        tcrOutInput.type = 'text';
        tcrOutInput.className = 'table-input';
        tcrOutInput.value = marker.tcrOut || '';
        tcrOutInput.dataset.field = 'tcrOut';
        makeInputResizable(tcrOutInput);
        tcrOutInput.addEventListener('change', (e) => {
            marker.tcrOut = e.target.value;
            updateDuration(index);
        });
        tcrOutCell.appendChild(tcrOutInput);
        row.appendChild(tcrOutCell);
        
        // Add Duration cell (read-only)
        const durationCell = document.createElement('td');
        durationCell.textContent = marker.duration || '';
        row.appendChild(durationCell);
        
        // Add Usage cell with dropdown
        const usageCell = document.createElement('td');
        const usageSelect = document.createElement('select');
        usageSelect.className = 'table-input';
        usageSelect.dataset.field = 'usage';
        makeInputResizable(usageSelect);
        usageOptions.forEach(option => {
            const opt = document.createElement('option');
            opt.value = option;
            opt.textContent = option;
            if (marker.usage && marker.usage.includes(option)) opt.selected = true;
            usageSelect.appendChild(opt);
        });
        usageSelect.addEventListener('change', (e) => {
            marker.usage = Array.from(e.target.selectedOptions, option => option.value);
            updateUsageCounts();
        });
        usageCell.appendChild(usageSelect);
        row.appendChild(usageCell);
        
        // Add other cells with resizable inputs
        ['title', 'filmTitle', 'composer', 'lyricist', 'musicCo', 'nocId', 'nocTitle'].forEach(field => {
            const cell = document.createElement('td');
            const input = document.createElement('input');
            input.type = 'text';
            input.className = 'table-input';
            input.value = marker[field] || '';
            input.dataset.field = field;
            makeInputResizable(input);
            input.addEventListener('change', (e) => {
                marker[field] = e.target.value;
            });
            cell.appendChild(input);
            row.appendChild(cell);
        });
        
        // Add Recognize button
        const recognizeCell = document.createElement('td');
        const recognizeBtn = document.createElement('button');
        recognizeBtn.className = 'btn btn-primary recognize-btn';
        recognizeBtn.textContent = 'Recognize';
        recognizeBtn.dataset.index = index;
        recognizeCell.appendChild(recognizeBtn);
        row.appendChild(recognizeCell);
        
        tableBody.appendChild(row);
    });
    
    // Update header colors based on active paste columns
    const headers = document.querySelectorAll('.table thead tr:first-child th.special');
    headers.forEach(header => {
        const columnName = getColumnName(Array.from(header.parentElement.children).indexOf(header));
        header.style.backgroundColor = activePasteColumns[columnName] ? '#90EE90' : '';
    });
    
    // Scroll to top when table is updated
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
    const [hours, minutes, seconds, frames] = timeStr.split(':').map(Number);
    return hours * 3600 + minutes * 60 + seconds + (frames / 25); // Convert frames to seconds (25fps)
}

function calculateDuration(inTime, outTime) {
    if (!inTime || !outTime) return '';
    
    const inSeconds = timeToSeconds(inTime);
    const outSeconds = timeToSeconds(outTime);
    const duration = outSeconds - inSeconds;
    
    return formatTime(duration);
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
    document.addEventListener('keydown', function(e) {
        const videoPlayer = document.getElementById('videoPlayer');
        
        // Arrow key controls for video timeline
        if (e.key === 'ArrowLeft' || e.key === 'ArrowRight') {
            e.preventDefault();
            
            // Start frame-by-frame navigation after holding for 500ms
            if (!frameTimer) {
                frameTimer = setTimeout(() => {
                    isFrameByFrame = true;
                }, 500);
            }
            
            // Calculate the time increment (1 frame = 1/frameRate seconds)
            const frameIncrement = 1 / frameRate;
            const timeIncrement = isFrameByFrame ? frameIncrement : 1;
            
            if (e.key === 'ArrowLeft') {
                videoPlayer.currentTime = Math.max(0, videoPlayer.currentTime - timeIncrement);
            } else if (e.key === 'ArrowRight') {
                videoPlayer.currentTime = Math.min(videoPlayer.duration, videoPlayer.currentTime + timeIncrement);
            }
        }
        
        // Existing keyboard shortcuts
        if (e.altKey) {
            if (e.key === '1') {
                e.preventDefault();
                if (isSingleButtonMode) {
                    markTCR(isNextMarkTcrIn ? 'in' : 'out');
                } else {
                    markTCR('in');
                }
            } else if (e.key === '2' && !isSingleButtonMode) {
                e.preventDefault();
                markTCR('out');
            }
        }
    });

    // Reset frame-by-frame mode when key is released
    document.addEventListener('keyup', function(e) {
        if (e.key === 'ArrowLeft' || e.key === 'ArrowRight') {
            clearTimeout(frameTimer);
            frameTimer = null;
            isFrameByFrame = false;
        }
    });
}

function exportToExcel() {
    const includeHeader = document.getElementById('includeHeaderRows').checked;
    let ws;
    if (includeHeader && headerRows.length > 0) {
        // Prepare marker data as array of arrays
        const markerData = markers.map((marker, index) => [
            index + 1,
            marker.tcrIn,
            marker.tcrOut,
            marker.duration,
            marker.usage,
            marker.filmTitle,
            marker.composer,
            marker.lyricist,
            marker.musicCo,
            marker.nocId,
            marker.nocTitle,
            marker.titleColor || ''
        ]);
        // Combine headerRows and markerData
        const allData = headerRows.concat([['#Seq', 'TCR In', 'TCR Out', 'Duration', 'Usage', 'Film/Album Title', 'Composer', 'Lyricist', 'Music Co', 'NOC ID', 'NOC Title', 'Title Color']]).concat(markerData);
        ws = XLSX.utils.aoa_to_sheet(allData);
    } else {
        ws = XLSX.utils.json_to_sheet(markers.map((marker, index) => ({
            '#Seq': index + 1,
            'TCR In': marker.tcrIn,
            'TCR Out': marker.tcrOut,
            'Duration': marker.duration,
            'Usage': marker.usage,
            'Film/Album Title': marker.filmTitle,
            'Composer': marker.composer,
            'Lyricist': marker.lyricist,
            'Music Co': marker.musicCo,
            'NOC ID': marker.nocId,
            'NOC Title': marker.nocTitle,
            'Title Color': marker.titleColor || ''
        })));
    }
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Markers');
    XLSX.writeFile(wb, 'markers.xlsx');
}

function exportToCSV() {
    const includeHeader = document.getElementById('includeHeaderRows').checked;
    const headers = ['#Seq', 'TCR In', 'TCR Out', 'Duration', 'Usage', 'Film/Album Title', 'Composer', 'Lyricist', 'Music Co', 'NOC ID', 'NOC Title', 'Title Color'];
    let csvContent = '';
    if (includeHeader && headerRows.length > 0) {
        csvContent += headerRows.map(row => row.join(",")).join("\n") + "\n";
    }
    csvContent += headers.join(',') + '\n';
    csvContent += markers.map((marker, index) => [
        index + 1,
        marker.tcrIn,
        marker.tcrOut,
        marker.duration,
        marker.usage,
        marker.filmTitle,
        marker.composer,
        marker.lyricist,
        marker.musicCo,
        marker.nocId,
        marker.nocTitle,
        marker.titleColor || ''
    ].join(',')).join('\n');
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = 'markers.csv';
    link.click();
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
            extraColumns.push({ name: col, type: 'text', options: [] });
        }
    });
    updateMarkerTableHeader();
    updateMarkerTable();
}

function loadCueSheetData(header, data) {
    // Map file columns to app columns using mapping
    let lowerHeader = header.map(h => h.trim().toLowerCase());
    markers = data.map(row => {
        let marker = {};
        defaultMarkerColumns.forEach(col => {
            let idx = lowerHeader.findIndex(h => columnNameMap[h] === col.key || h === col.label.toLowerCase());
            marker[col.key] = idx !== -1 ? row[header[idx]] || '' : '';
        });
        // Add extra columns
        extraColumns.forEach(col => {
            marker[col.name] = row[col.name] || '';
        });
        return marker;
    });
    updateMarkerTable();
}

function initializeClearTableButton() {
    const clearTableBtn = document.getElementById('clearTableBtn');
    if (clearTableBtn) {
        clearTableBtn.addEventListener('click', function() {
            markers = [];
            updateMarkerTable();
        });
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
            
            // Copy to all other rows in the same column
            markers.forEach((marker, index) => {
                if (index !== sourceIndex) {
                    marker[columnName] = sourceValue;
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
    
    // Get indices of selected rows
    const selectedIndices = Array.from(checkboxes).map(checkbox => 
        parseInt(checkbox.closest('tr').querySelector('.seq-cell').textContent) - 1
    );
    
    // Remove selected markers
    markers = markers.filter((_, index) => !selectedIndices.includes(index));
    
    // Update the table
    updateMarkerTable();
} 