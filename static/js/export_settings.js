// Export settings management
let exportSettings = {
    fileName: '',
    fileType: 'excel',
    downloadLocation: 'ask',
    customLocation: '',
    includeHeader: true,
    fieldsToExport: [],
    tcrFormat: 'timecode'
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

    // Use default fields if no saved fields
    const fields = exportSettings.fieldsToExport.length > 0 ? 
        exportSettings.fieldsToExport : 
        ['tcrIn', 'tcrOut', 'duration', 'usage', 'title', 'filmTitle', 'composer', 'lyricist', 'musicCo', 'nocId', 'nocTitle'];

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
                exportSettings.fieldsToExport.push(fieldKey);
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
        
        // Get show name and episode info from localStorage or use defaults
        const showInfo = JSON.parse(localStorage.getItem('showInfo')) || {
            showName: 'Unknown_Show',
            season: '',
            episodeNumber: ''
        };
        
        let fileName = showInfo.showName;
        if (showInfo.season) {
            fileName += `_Season${showInfo.season}`;
        }
        if (showInfo.episodeNumber) {
            fileName += `_${showInfo.episodeNumber.padStart(4, '0')}`;
        }
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
        
        // Apply TCR format
        const tcrFormatRadios = document.querySelectorAll('input[name="tcrFormat"]');
        if (tcrFormatRadios.length > 0) {
            tcrFormatRadios.forEach(radio => {
                radio.checked = radio.value === exportSettings.tcrFormat;
            });
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
        fieldsToExport: [],
        tcrFormat: 'timecode'
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
        
        // Validate settings
        if (!settings.fieldsToExport || !settings.fieldsToExport.length) {
            throw new Error('No fields selected for export');
        }
        
        // Prepare data
        const data = prepareExportData(settings);
        if (!data || !data.length) {
            throw new Error('No data to export');
        }
        
        // Get file type from settings
        const fileType = settings.fileType || 'excel';
        console.log('Exporting as:', fileType);
        
        // Export based on file type
        switch (fileType) {
            case 'excel':
                exportToExcelWorkbook(data);
                break;
            case 'csv':
                exportToCSV(data);
                break;
            case 'plain':
                exportToPlainExcel(data);
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

function prepareExportData(settings) {
    try {
        console.log('Preparing export data with settings:', settings);
        
        // Get all marker rows
        const rows = Array.from(document.querySelectorAll('.table tbody tr'));
        console.log('Found rows:', rows.length);
        
        // Prepare data array
        const data = [];
        
        // Add header if enabled
        if (settings.includeHeader) {
            console.log('Including header rows');
            const headerData = {};
            // Get header rows from localStorage
            const headerRows = JSON.parse(localStorage.getItem('headerRows')) || [];
            if (Array.isArray(headerRows)) {
                headerRows.forEach((row, index) => {
                    if (Array.isArray(row)) {
                        row.forEach((cell, cellIndex) => {
                            if (cellIndex % 2 === 0 && cell) {
                                const key = cell.trim();
                                const value = row[cellIndex + 1] || '';
                                headerData[key] = value;
                            }
                        });
                    }
                });
                data.push(headerData);
            } else {
                console.warn('Header rows is not an array:', headerRows);
            }
        }
        
        // Add marker data
        console.log('Processing marker data');
        rows.forEach(row => {
            const rowData = {};
            settings.fieldsToExport.forEach(field => {
                const cell = row.querySelector(`[data-field="${field}"]`);
                if (cell) {
                    let value = cell.textContent.trim();
                    
                    // Format TCR if needed
                    if (field.toLowerCase().includes('tcr')) {
                        value = formatTCR(value, settings.tcrFormat || 'timecode');
                    }
                    
                    rowData[field] = value;
                }
            });
            data.push(rowData);
        });
        
        console.log('Export data prepared:', data);
        return data;
    } catch (error) {
        console.error('Error preparing export data:', error);
        throw new Error('Failed to prepare export data: ' + error.message);
    }
}

function formatTCR(timecode, format) {
    if (!timecode) return '';
    
    // Parse the timecode
    const [hours, minutes, seconds, frames] = timecode.split(':').map(Number);
    
    if (format === 'time') {
        // Convert to HH:MM:SS
        return `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}:${seconds.toString().padStart(2, '0')}`;
    }
    
    // Return original timecode format
    return timecode;
}

function exportToExcelWorkbook(data) {
    console.log('Starting Excel export with data:', data.slice(0, 2));  // Log first two items
    
    // Send data to server for Excel export
    fetch('/api/export/excel', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
        },
        body: JSON.stringify(data)
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
        
        // Create download link
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `${exportSettings.fileName}.xlsx`;
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

function exportToCSV(data) {
    // Send data to server for CSV export
    fetch('/api/export/csv', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
        },
        body: JSON.stringify(data)
    })
    .then(response => {
        if (!response.ok) {
            throw new Error('Export failed');
        }
        return response.blob();
    })
    .then(blob => {
        // Create download link
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `${exportSettings.fileName}.csv`;
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        a.remove();
    })
    .catch(error => {
        console.error('CSV export failed:', error);
        throw error;
    });
}

function exportToPlainExcel(data) {
    // This is similar to Excel workbook but with simpler formatting
    exportToExcelWorkbook(data);
} 