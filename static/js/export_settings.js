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

// Initialize the export settings page
document.addEventListener('DOMContentLoaded', function() {
    initializeExportSettings();
    loadSavedSettings();
    setupEventListeners();
    populateFieldsToExport();
    updateFileNamePreview();
});

function initializeExportSettings() {
    // Get all available fields from the marker table
    const tableHeaders = document.querySelectorAll('.table thead th');
    const availableFields = Array.from(tableHeaders)
        .map(th => ({
            name: th.textContent.trim(),
            key: th.dataset.field || th.textContent.toLowerCase().replace(/[^a-z0-9]/g, '')
        }))
        .filter(field => field.key && field.key !== 'seq' && field.key !== 'recognize');

    // Set default fields to export
    exportSettings.fieldsToExport = availableFields.map(field => field.key);
}

function loadSavedSettings() {
    const savedSettings = localStorage.getItem('exportSettings');
    if (savedSettings) {
        exportSettings = { ...exportSettings, ...JSON.parse(savedSettings) };
        applySettingsToUI();
    }
}

function setupEventListeners() {
    // Close button
    document.getElementById('closeSettings').addEventListener('click', () => {
        window.close();
    });

    // File type change
    document.getElementById('fileType').addEventListener('change', (e) => {
        exportSettings.fileType = e.target.value;
        updateFileNamePreview();
    });

    // Download location radio buttons
    document.querySelectorAll('input[name="downloadLocation"]').forEach(radio => {
        radio.addEventListener('change', (e) => {
            exportSettings.downloadLocation = e.target.value;
            const customLocationInput = document.getElementById('customLocation');
            const browseButton = document.getElementById('browseLocation');
            
            if (e.target.value === 'custom') {
                customLocationInput.disabled = false;
                browseButton.disabled = false;
            } else {
                customLocationInput.disabled = true;
                browseButton.disabled = true;
            }
        });
    });

    // Browse button
    document.getElementById('browseLocation').addEventListener('click', async () => {
        try {
            const dirHandle = await window.showDirectoryPicker();
            exportSettings.customLocation = dirHandle.name;
            document.getElementById('customLocation').value = dirHandle.name;
        } catch (err) {
            console.error('Error selecting directory:', err);
        }
    });

    // Include header checkbox
    document.getElementById('includeHeader').addEventListener('change', (e) => {
        exportSettings.includeHeader = e.target.checked;
    });

    // TCR format radio buttons
    document.querySelectorAll('input[name="tcrFormat"]').forEach(radio => {
        radio.addEventListener('change', (e) => {
            exportSettings.tcrFormat = e.target.value;
        });
    });

    // Save settings button
    document.getElementById('saveSettings').addEventListener('click', () => {
        saveSettings();
        window.close();
    });

    // Reset settings button
    document.getElementById('resetSettings').addEventListener('click', () => {
        if (confirm('Are you sure you want to reset all export settings to default?')) {
            resetSettings();
        }
    });
}

function populateFieldsToExport() {
    const fieldsContainer = document.querySelector('.fields-container');
    fieldsContainer.innerHTML = '';

    const tableHeaders = document.querySelectorAll('.table thead th');
    Array.from(tableHeaders).forEach(th => {
        const fieldKey = th.dataset.field || th.textContent.toLowerCase().replace(/[^a-z0-9]/g, '');
        if (fieldKey && fieldKey !== 'seq' && fieldKey !== 'recognize') {
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
            });

            label.appendChild(checkbox);
            label.appendChild(document.createTextNode(th.textContent.trim()));
            fieldsContainer.appendChild(label);
        }
    });
}

function updateFileNamePreview() {
    const fileNameInput = document.getElementById('fileName');
    const preview = document.querySelector('.file-name-preview');
    
    // Extract show name and episode number from header rows
    const showName = extractShowName();
    const episodeNumber = extractEpisodeNumber();
    const season = extractSeason();
    
    let fileName = showName;
    if (season) {
        fileName += `_Season${season}`;
    }
    if (episodeNumber) {
        fileName += `_${episodeNumber.padStart(4, '0')}`;
    }
    fileName += '_Unmix HD_MusicCueSheet';
    
    exportSettings.fileName = fileName;
    fileNameInput.value = fileName;
    preview.textContent = `Preview: ${fileName}.${exportSettings.fileType === 'csv' ? 'csv' : 'xlsx'}`;
}

function extractShowName() {
    const headerRows = window.headerRows || [];
    for (const row of headerRows) {
        const seriesTitleIndex = row.findIndex(cell => cell.toLowerCase().includes('series title'));
        if (seriesTitleIndex !== -1 && row[seriesTitleIndex + 1]) {
            return row[seriesTitleIndex + 1].trim();
        }
    }
    return 'Unknown_Show';
}

function extractEpisodeNumber() {
    const headerRows = window.headerRows || [];
    for (const row of headerRows) {
        const episodeIndex = row.findIndex(cell => cell.toLowerCase().includes('episode number'));
        if (episodeIndex !== -1 && row[episodeIndex + 1]) {
            return row[episodeIndex + 1].trim();
        }
    }
    return '0000';
}

function extractSeason() {
    const headerRows = window.headerRows || [];
    for (const row of headerRows) {
        const seasonIndex = row.findIndex(cell => cell.toLowerCase().includes('season'));
        if (seasonIndex !== -1 && row[seasonIndex + 1]) {
            return row[seasonIndex + 1].trim();
        }
    }
    return null;
}

function applySettingsToUI() {
    // Apply file type
    document.getElementById('fileType').value = exportSettings.fileType;
    
    // Apply download location
    const downloadLocationRadios = document.querySelectorAll('input[name="downloadLocation"]');
    downloadLocationRadios.forEach(radio => {
        radio.checked = radio.value === exportSettings.downloadLocation;
    });
    
    // Apply custom location
    const customLocationInput = document.getElementById('customLocation');
    customLocationInput.value = exportSettings.customLocation;
    customLocationInput.disabled = exportSettings.downloadLocation !== 'custom';
    document.getElementById('browseLocation').disabled = exportSettings.downloadLocation !== 'custom';
    
    // Apply include header
    document.getElementById('includeHeader').checked = exportSettings.includeHeader;
    
    // Apply TCR format
    const tcrFormatRadios = document.querySelectorAll('input[name="tcrFormat"]');
    tcrFormatRadios.forEach(radio => {
        radio.checked = radio.value === exportSettings.tcrFormat;
    });
    
    // Update file name preview
    updateFileNamePreview();
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
    
    initializeExportSettings();
    applySettingsToUI();
    saveSettings();
}

// Export function that will be called from the main page
function exportWithSettings() {
    const settings = JSON.parse(localStorage.getItem('exportSettings')) || exportSettings;
    
    // Get the selected fields in the correct order
    const selectedFields = settings.fieldsToExport;
    
    // Prepare the data based on settings
    const exportData = prepareExportData(settings);
    
    // Show loading indicator
    const loadingIndicator = document.createElement('div');
    loadingIndicator.className = 'loading-indicator';
    loadingIndicator.textContent = 'Exporting...';
    document.body.appendChild(loadingIndicator);
    
    // Export based on file type
    try {
        switch(settings.fileType) {
            case 'excel':
                exportToExcelWorkbook(exportData, settings);
                break;
            case 'csv':
                exportToCSV(exportData, settings);
                break;
            case 'plain':
                exportToPlainExcel(exportData, settings);
                break;
            default:
                throw new Error('Invalid file type');
        }
    } catch (error) {
        console.error('Export failed:', error);
        alert('Export failed: ' + error.message);
    } finally {
        // Remove loading indicator
        loadingIndicator.remove();
    }
}

function prepareExportData(settings) {
    // Get all marker rows
    const rows = Array.from(document.querySelectorAll('.table tbody tr'));
    
    // Prepare data array
    const data = [];
    
    // Add header if enabled
    if (settings.includeHeader) {
        const headerData = {};
        window.headerRows.forEach((row, index) => {
            row.forEach((cell, cellIndex) => {
                if (cellIndex % 2 === 0 && cell) {
                    const key = cell.trim();
                    const value = row[cellIndex + 1] || '';
                    headerData[key] = value;
                }
            });
        });
        data.push(headerData);
    }
    
    // Add marker data
    rows.forEach(row => {
        const rowData = {};
        settings.fieldsToExport.forEach(field => {
            const cell = row.querySelector(`[data-field="${field}"]`);
            if (cell) {
                let value = cell.textContent.trim();
                
                // Format TCR if needed
                if (field.toLowerCase().includes('tcr')) {
                    value = formatTCR(value, settings.tcrFormat);
                }
                
                rowData[field] = value;
            }
        });
        data.push(rowData);
    });
    
    return data;
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

function exportToExcelWorkbook(data, settings) {
    // Send data to server for Excel export
    fetch('/api/export/excel', {
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
        a.download = `${settings.fileName}.xlsx`;
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        a.remove();
    })
    .catch(error => {
        console.error('Excel export failed:', error);
        throw error;
    });
}

function exportToCSV(data, settings) {
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
        a.download = `${settings.fileName}.csv`;
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

function exportToPlainExcel(data, settings) {
    // This is similar to Excel workbook but with simpler formatting
    exportToExcelWorkbook(data, settings);
} 