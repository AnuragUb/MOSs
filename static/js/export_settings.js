// Export settings management
let exportSettings = {
    fileName: '',
    fileType: 'excel',
    downloadLocation: 'ask',
    customLocation: '',
    includeHeader: true,
    blankLines: 0,
    fieldsToExport: [],
    tcrFormat: 'timecode',
    importFilmTitle: true,
    addSeriesTitlePrefix: true,
};

// Check if we're on the export settings page
function isExportSettingsPage() {
    return document.querySelector('.export-settings-container') !== null;
}

// Initialize the export settings page
document.addEventListener('DOMContentLoaded', function() {
    console.log('DOM Content Loaded in export settings page');
    
    // Only initialize if we're on the export settings page
    if (!isExportSettingsPage()) {
        console.log('Not on export settings page, skipping initialization');
        return;
    }

    // Wait for elements to be ready
    const initInterval = setInterval(() => {
        const requiredElements = [
            'closeSettings',
            'fileType',
            'fileName',
            'includeHeader',
            'browseLocation',
            'customLocation'
        ];

        const allElementsExist = requiredElements.every(id => document.getElementById(id) !== null);
        
        if (allElementsExist) {
            clearInterval(initInterval);
            try {
                console.log('Initializing export settings page...');
                initializeExportSettingsPage();
                loadSavedSettings();
                setupEventListeners();
                populateFieldsToExport();
                updateFileNamePreview();
                console.log('Export settings page initialized successfully');
            } catch (error) {
                console.error('Error initializing export settings page:', error);
                alert('Error initializing export settings. Please refresh the page.');
            }
        }
    }, 100);

    // Clear interval after 5 seconds to prevent infinite checking
    setTimeout(() => {
        clearInterval(initInterval);
        if (!document.getElementById('closeSettings')) {
            console.error('Required elements not found after timeout');
            alert('Error loading export settings. Please try again.');
        }
    }, 5000);
});

function initializeExportSettingsPage() {
    try {
        // Get fields from localStorage if available
        const savedFields = localStorage.getItem('exportFields');
        if (savedFields) {
            exportSettings.fieldsToExport = JSON.parse(savedFields);
        } else {
            // Default fields if none are saved
            exportSettings.fieldsToExport = [
                'seq',
                'tcrIn',
                'tcrOut',
                'duration',
                'usage',
                'title',
                'filmTitle',
                'composer',
                'lyricist',
                'musicCo',
                'nocId',
                'nocTitle'
            ];
        }

        // Get markers and header rows from localStorage
        const savedMarkers = localStorage.getItem('markers');
        const savedHeaderRows = localStorage.getItem('headerRows');
        
        if (savedMarkers) {
            window.markers = JSON.parse(savedMarkers);
        }
        if (savedHeaderRows) {
            window.headerRows = JSON.parse(savedHeaderRows);
        }

        // Initialize blank lines input
        const blankLinesInput = document.getElementById('blankLines');
        if (blankLinesInput) {
            blankLinesInput.value = exportSettings.blankLines;
            blankLinesInput.addEventListener('change', (e) => {
                exportSettings.blankLines = parseInt(e.target.value) || 0;
            });
        }

        // Add checkbox for addSeriesTitlePrefix only if it doesn't exist
        const existingPrefixCheckbox = document.getElementById('addSeriesTitlePrefix');
        if (!existingPrefixCheckbox) {
            // Find the section by heading text
            const headings = document.querySelectorAll('.settings-section h3');
            let contentSettingsSection = null;
            headings.forEach(h => {
                if (h.textContent.trim() === 'Content Settings') {
                    contentSettingsSection = h.parentElement;
                }
            });
            if (contentSettingsSection) {
                const prefixGroup = document.createElement('div');
                prefixGroup.className = 'setting-group';
                prefixGroup.innerHTML = `
                    <label>
                        <input type="checkbox" id="addSeriesTitlePrefix">
                        Add Series Title as Prefix to Title ("Series Title - (User Title)")
                    </label>
                    <span class="help-text">If enabled, exported title will be: Series Title - (User Title)</span>
                `;
                contentSettingsSection.appendChild(prefixGroup);
                const prefixCheckbox = document.getElementById('addSeriesTitlePrefix');
                prefixCheckbox.checked = exportSettings.addSeriesTitlePrefix;
                prefixCheckbox.addEventListener('change', (e) => {
                    exportSettings.addSeriesTitlePrefix = e.target.checked;
                    localStorage.setItem('exportSettings', JSON.stringify(exportSettings));
                });
            }
        }
    } catch (error) {
        console.error('Error in initializeExportSettingsPage:', error);
        throw error;
    }
}

function loadSavedSettings() {
    const savedSettings = localStorage.getItem('exportSettings');
    if (savedSettings) {
        exportSettings = { ...exportSettings, ...JSON.parse(savedSettings) };
        applySettingsToUI();
    }
}

function setupEventListeners() {
    try {
        // Close button
        const closeBtn = document.getElementById('closeSettings');
        if (closeBtn) {
            closeBtn.addEventListener('click', () => {
                window.close();
            });
        }

        // File type change
        const fileTypeSelect = document.getElementById('fileType');
        if (fileTypeSelect) {
            fileTypeSelect.addEventListener('change', (e) => {
                exportSettings.fileType = e.target.value;
                updateFileNamePreview();
            });
        }

        // Download location radio buttons
        const downloadLocationRadios = document.querySelectorAll('input[name="downloadLocation"]');
        if (downloadLocationRadios.length > 0) {
            downloadLocationRadios.forEach(radio => {
                radio.addEventListener('change', (e) => {
                    exportSettings.downloadLocation = e.target.value;
                    const customLocationInput = document.getElementById('customLocation');
                    const browseButton = document.getElementById('browseLocation');
                    
                    if (customLocationInput && browseButton) {
                        if (e.target.value === 'custom') {
                            customLocationInput.disabled = false;
                            browseButton.disabled = false;
                        } else {
                            customLocationInput.disabled = true;
                            browseButton.disabled = true;
                        }
                    }
                });
            });
        }

        // Browse button
        const browseButton = document.getElementById('browseLocation');
        if (browseButton) {
            browseButton.addEventListener('click', async () => {
                try {
                    const dirHandle = await window.showDirectoryPicker();
                    exportSettings.customLocation = dirHandle.name;
                    const customLocationInput = document.getElementById('customLocation');
                    if (customLocationInput) {
                        customLocationInput.value = dirHandle.name;
                    }
                } catch (err) {
                    console.error('Error selecting directory:', err);
                }
            });
        }

        // Include header checkbox
        const includeHeaderCheckbox = document.getElementById('includeHeader');
        if (includeHeaderCheckbox) {
            includeHeaderCheckbox.addEventListener('change', (e) => {
                exportSettings.includeHeader = e.target.checked;
            });
        }

        // TCR format radio buttons
        const tcrFormatRadios = document.querySelectorAll('input[name="tcrFormat"]');
        if (tcrFormatRadios.length > 0) {
            tcrFormatRadios.forEach(radio => {
                radio.addEventListener('change', (e) => {
                    exportSettings.tcrFormat = e.target.value;
                });
            });
        }

        // Import Film Title checkbox
        const importFilmTitleCheckbox = document.getElementById('importFilmTitle');
        if (importFilmTitleCheckbox) {
            importFilmTitleCheckbox.addEventListener('change', (e) => {
                exportSettings.importFilmTitle = e.target.checked;
            });
        }

        // Save settings button
        const saveButton = document.getElementById('saveSettings');
        if (saveButton) {
            saveButton.addEventListener('click', () => {
                saveSettings();
                window.close();
            });
        }

        // Reset settings button
        const resetButton = document.getElementById('resetSettings');
        if (resetButton) {
            resetButton.addEventListener('click', () => {
                if (confirm('Are you sure you want to reset all export settings to default?')) {
                    resetSettings();
                }
            });
        }
    } catch (error) {
        console.error('Error setting up event listeners:', error);
    }
}

function populateFieldsToExport() {
    const fieldsContainer = document.querySelector('.fields-container');
    if (!fieldsContainer) return;

    fieldsContainer.innerHTML = '';

    // Always use all available fields
    const fields = ['tcrIn', 'tcrOut', 'duration', 'usage', 'title', 'filmTitle', 'composer', 'lyricist', 'musicCo', 'nocId', 'nocTitle'];

    const fieldLabels = {
        'tcrIn': 'TCR In',
        'tcrOut': 'TCR Out',
        'duration': 'Duration',
        'usage': 'Usage',
        'title': 'Title',
        'filmTitle': 'Film/Album Title',
        'composer': 'Composer',
        'lyricist': 'Lyricist',
        'musicCo': 'Music Co',
        'nocId': 'NOC ID',
        'nocTitle': 'NOC Title'
    };

    fields.forEach(fieldKey => {
        const label = document.createElement('label');
        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.value = fieldKey;
        checkbox.checked = exportSettings.fieldsToExport.includes(fieldKey);
        
        checkbox.addEventListener('change', (e) => {
            if (e.target.checked) {
                if (!exportSettings.fieldsToExport.includes(fieldKey)) {
                    exportSettings.fieldsToExport.push(fieldKey);
                }
            } else {
                exportSettings.fieldsToExport = exportSettings.fieldsToExport.filter(f => f !== fieldKey);
            }
            // Save fields to localStorage
            localStorage.setItem('exportFields', JSON.stringify(exportSettings.fieldsToExport));
        });

        label.appendChild(checkbox);
        label.appendChild(document.createTextNode(fieldLabels[fieldKey] || fieldKey));
        fieldsContainer.appendChild(label);
    });
}

function updateFileNamePreview() {
    try {
        const fileNameInput = document.getElementById('fileName');
        const preview = document.querySelector('.file-name-preview');
        if (!fileNameInput || !preview) {
            console.warn('File name input or preview element not found');
            return;
        }
        const showInfo = JSON.parse(localStorage.getItem('showInfo')) || {
            showName: 'Unknown_Show',
            season: '',
            episodeNumber: ''
        };
        if (!showInfo.showName || showInfo.showName === 'Unknown_Show') {
            console.warn('Show name not found in localStorage. File name will use default.');
        }
        if (!showInfo.season) {
            console.warn('Season not found in localStorage. File name will not include season.');
        }
        if (!showInfo.episodeNumber) {
            console.warn('Episode number not found in localStorage. File name will not include episode number.');
        }
        let fileName = showInfo.showName;
        if (showInfo.season) fileName += `_Season${showInfo.season}`;
        if (showInfo.episodeNumber) fileName += `_${showInfo.episodeNumber.padStart(4, '0')}`;
        fileName += '_Unmix HD_MusicCueSheet';
        exportSettings.fileName = fileName;
        fileNameInput.value = fileName;
        preview.textContent = `Preview: ${fileName}.${exportSettings.fileType === 'csv' ? 'csv' : 'xlsx'}`;
    } catch (error) {
        console.error('Error updating file name preview:', error);
    }
}

function applySettingsToUI() {
    try {
        // Apply file type
        const fileTypeSelect = document.getElementById('fileType');
        if (fileTypeSelect) {
            fileTypeSelect.value = exportSettings.fileType;
        }
        
        // Apply download location
        const downloadLocationRadios = document.querySelectorAll('input[name="downloadLocation"]');
        if (downloadLocationRadios.length > 0) {
            downloadLocationRadios.forEach(radio => {
                radio.checked = radio.value === exportSettings.downloadLocation;
            });
        }
        
        // Apply custom location
        const customLocationInput = document.getElementById('customLocation');
        const browseButton = document.getElementById('browseLocation');
        if (customLocationInput && browseButton) {
            customLocationInput.value = exportSettings.customLocation;
            customLocationInput.disabled = exportSettings.downloadLocation !== 'custom';
            browseButton.disabled = exportSettings.downloadLocation !== 'custom';
        }
        
        // Apply include header
        const includeHeaderCheckbox = document.getElementById('includeHeader');
        if (includeHeaderCheckbox) {
            includeHeaderCheckbox.checked = exportSettings.includeHeader;
        }

        // Apply blank lines
        const blankLinesInput = document.getElementById('blankLines');
        if (blankLinesInput) {
            blankLinesInput.value = exportSettings.blankLines;
        }
        
        // Apply TCR format
        const tcrFormatRadios = document.querySelectorAll('input[name="tcrFormat"]');
        if (tcrFormatRadios.length > 0) {
            tcrFormatRadios.forEach(radio => {
                radio.checked = radio.value === exportSettings.tcrFormat;
            });
        }
        
        // Apply import film title setting
        const importFilmTitleCheckbox = document.getElementById('importFilmTitle');
        if (importFilmTitleCheckbox) {
            importFilmTitleCheckbox.checked = exportSettings.importFilmTitle;
        }

        // Apply addSeriesTitlePrefix setting
        const addSeriesTitlePrefixCheckbox = document.getElementById('addSeriesTitlePrefix');
        if (addSeriesTitlePrefixCheckbox) {
            addSeriesTitlePrefixCheckbox.checked = exportSettings.addSeriesTitlePrefix;
        }
        
        // Update file name preview
        updateFileNamePreview();
    } catch (error) {
        console.error('Error applying settings to UI:', error);
    }
}

function saveSettings() {
    localStorage.setItem('exportSettings', JSON.stringify(exportSettings));
}

function resetSettings() {
    exportSettings = {
        fileName: '',
        fileType: 'excel',
        downloadLocation: 'ask',
        customLocation: '',
        includeHeader: true,
        blankLines: 0,
        fieldsToExport: [],
        tcrFormat: 'timecode',
        importFilmTitle: true,
        addSeriesTitlePrefix: true,
    };
    
    initializeExportSettingsPage();
    applySettingsToUI();
    saveSettings();
}

// Export function that will be called from the main page
function exportWithSettings() {
    try {
        console.log('Starting export with settings');
        
        // Get settings from localStorage
        const settings = JSON.parse(localStorage.getItem('exportSettings')) || exportSettings;
        console.log('Loaded settings from localStorage:', settings);

        // --- BEGIN: Prepare export payload with metadata and marker data ---
        const headerRows = JSON.parse(localStorage.getItem('headerRows')) || [];
        // Get markers and add markColor
        let markers = [];
        if (typeof window.getMarkersWithMarkColor === 'function') {
            markers = window.getMarkersWithMarkColor();
        } else {
            // Fallback: add markColor from localStorage.markedRows
            const rawMarkers = JSON.parse(localStorage.getItem('markers')) || [];
            const markedRows = JSON.parse(localStorage.getItem('markedRows')) || {};
            markers = rawMarkers.map((marker, idx) => ({ ...marker, markColor: markedRows[idx] || '' }));
        }

        // Apply film title import if enabled
        if (settings.importFilmTitle) {
            // Get show name from the first header row (assuming it's in the first cell)
            const showName = headerRows[0] && headerRows[0][0] ? headerRows[0][0] : '';
            if (showName) {
                markers = markers.map(marker => ({
                    ...marker,
                    filmTitle: showName
                }));
            }
        }

        // Apply addSeriesTitlePrefix if enabled
        if (settings.addSeriesTitlePrefix) {
            markers = markers.map(marker => ({
                ...marker,
                title: `Series Title - (${marker.title || ''})`
            }));
        }

        const exportPayload = { 
            headerRows, 
            markers,
            blankLines: settings.blankLines
        };
        // --- END: Prepare export payload with metadata and marker data ---

        // Get file type from settings
        const fileType = settings.fileType || 'excel';
        console.log('Exporting as:', fileType);

        // Export based on file type
        switch (fileType) {
            case 'excel':
                exportToExcelWorkbook(exportPayload);
                break;
            case 'csv':
                exportToCSV(exportPayload);
                break;
            case 'plain':
                exportToPlainExcel(exportPayload);
                break;
            default:
                throw new Error('Invalid file type selected');
        }

        console.log('Export completed successfully');
    } catch (error) {
        console.error('Error during export:', error);
        alert('Export failed: ' + error.message);
    }
}

function getExportFileName(extension) {
    const settings = JSON.parse(localStorage.getItem('exportSettings')) || {};
    return (settings.fileName || 'exported_file') + '.' + extension;
}

function exportToExcelWorkbook(payload) {
    console.log('Starting Excel export with payload:', payload);
    fetch('/api/export/excel', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
        },
        body: JSON.stringify(payload)
    })
    .then(response => {
        console.log('Server response status:', response.status);
        if (!response.ok) {
            return response.json().then(err => {
                throw new Error(err.error || 'Export failed');
            });
        }
        return response.blob();
    })
    .then(blob => {
        console.log('Received blob of size:', blob.size);
        if (blob.size === 0) {
            throw new Error('Received empty file from server');
        }
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = getExportFileName('xlsx');
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        a.remove();
        console.log('Excel file downloaded successfully');
    })
    .catch(error => {
        console.error('Excel export failed:', error);
        alert('Export failed: ' + error.message);
        throw error;
    });
}

function exportToCSV(payload) {
    console.log('Starting CSV export with payload:', payload);
    fetch('/api/export/csv', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
        },
        body: JSON.stringify(payload)
    })
    .then(response => {
        console.log('Server response status:', response.status);
        if (!response.ok) {
            return response.json().then(err => {
                throw new Error(err.error || 'Export failed');
            });
        }
        return response.blob();
    })
    .then(blob => {
        console.log('Received blob of size:', blob.size);
        if (blob.size === 0) {
            throw new Error('Received empty file from server');
        }
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = getExportFileName('csv');
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        a.remove();
        console.log('CSV file downloaded successfully');
    })
    .catch(error => {
        console.error('CSV export failed:', error);
        alert('Export failed: ' + error.message);
        throw error;
    });
}

function exportToPlainExcel(payload) {
    exportToExcelWorkbook(payload);
}

function prepareExportData(settings) {
    let showName = '';
    if (window.headerRows) {
        showName = extractFieldValue(window.headerRows, 'series title');
    }
    if (!showName) {
        const showInfo = JSON.parse(localStorage.getItem('showInfo') || '{}');
        showName = showInfo.showName || '';
    }
    let markers = window.getMarkersWithMarkColor ? window.getMarkersWithMarkColor() : (window.markers || []);
    if (settings.addSeriesTitlePrefix && showName) {
        markers = markers.map(marker => ({
            ...marker,
            title: `${showName} - (${marker.title || ''})`
        }));
    }
    return { headerRows: window.headerRows, markers, blankLines: settings.blankLines };
} 