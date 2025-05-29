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
    const data = prepareExportData(settings);
    
    // Export based on file type
    switch (settings.fileType) {
        case 'excel':
            exportToExcelWorkbook(data, settings);
            break;
        case 'csv':
            exportToCSV(data, settings);
            break;
        case 'plain':
            exportToPlainExcel(data, settings);
            break;
    }
}

function prepareExportData(settings) {
    const data = [];
    
    // Add header rows if enabled
    if (settings.includeHeader && window.headerRows) {
        data.push(...window.headerRows);
    }
    
    // Add column headers
    const headers = selectedFields.map(field => {
        const th = document.querySelector(`th[data-field="${field}"]`);
        return th ? th.textContent.trim() : field;
    });
    data.push(headers);
    
    // Add marker data
    window.markers.forEach(marker => {
        const row = selectedFields.map(field => {
            let value = marker[field] || '';
            
            // Format TCR values if needed
            if ((field === 'tcrIn' || field === 'tcrOut') && value) {
                value = formatTCR(value, settings.tcrFormat);
            }
            
            return value;
        });
        data.push(row);
    });
    
    return data;
}

function formatTCR(timecode, format) {
    if (format === 'time') {
        // Convert HH:MM:SS:FF to HH:MM:SS
        return timecode.split(':').slice(0, 3).join(':');
    }
    return timecode;
}

function exportToExcelWorkbook(data, settings) {
    const ws = XLSX.utils.aoa_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Markers');
    
    // Apply styling
    const range = XLSX.utils.decode_range(ws['!ref']);
    for (let R = range.s.r; R <= range.e.r; R++) {
        for (let C = range.s.c; C <= range.e.c; C++) {
            const cell_address = { c: C, r: R };
            const cell_ref = XLSX.utils.encode_cell(cell_address);
            if (!ws[cell_ref]) continue;
            
            ws[cell_ref].s = {
                font: { sz: 11 },
                alignment: { vertical: 'center' }
            };
        }
    }
    
    // Write the file
    XLSX.writeFile(wb, `${settings.fileName}.xlsx`);
}

function exportToCSV(data, settings) {
    const csvContent = data.map(row => 
        row.map(field => {
            if (typeof field === 'string' && (field.includes(',') || field.includes('"'))) {
                return `"${field.replace(/"/g, '""')}"`;
            }
            return field;
        }).join(',')
    ).join('\n');
    
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = `${settings.fileName}.csv`;
    link.click();
}

function exportToPlainExcel(data, settings) {
    const ws = XLSX.utils.aoa_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Markers');
    XLSX.writeFile(wb, `${settings.fileName}.xlsx`);
} 