/**
 * Merge-o-matic 2000 - Main Application
 */

// Application State
const AppState = {
    files: {},          // fileName -> { file, data, columns, headerRow, ... }
    templateWorkbook: null,  // Loaded template
    processingFiles: false
};

// DOM Elements
const elements = {};

/**
 * Initialize the application
 */
function init() {
    // Get all DOM elements
    elements.fileInput = document.getElementById('file-input');
    elements.uploadZone = document.getElementById('upload-zone');
    elements.uploadStatus = document.getElementById('upload-status');
    elements.fileList = document.getElementById('file-list');
    elements.fileConfigs = document.getElementById('file-configs');
    elements.columnsSection = document.getElementById('columns-section');
    elements.graphSection = document.getElementById('graph-section');
    elements.timeSection = document.getElementById('time-section');
    elements.downloadSection = document.getElementById('download-section');
    elements.graphColumns = document.getElementById('graph-columns');
    elements.generateGraphBtn = document.getElementById('generate-graph-btn');
    elements.graphContainer = document.getElementById('graph-container');
    elements.startDate = document.getElementById('start-date');
    elements.startTime = document.getElementById('start-time');
    elements.endDate = document.getElementById('end-date');
    elements.endTime = document.getElementById('end-time');
    elements.durationDays = document.getElementById('duration-days');
    elements.interval = document.getElementById('interval');
    elements.alignmentOptions = document.getElementById('alignment-options');
    elements.createFileBtn = document.getElementById('create-file-btn');
    elements.progressContainer = document.getElementById('progress-container');
    elements.progressFill = document.getElementById('progress-fill');
    elements.progressText = document.getElementById('progress-text');
    elements.downloadContainer = document.getElementById('download-container');
    elements.downloadBtn = document.getElementById('download-btn');
    elements.templateInput = document.getElementById('template-input');
    elements.templateStatus = document.getElementById('template-status');
    elements.globalLoader = document.getElementById('global-loader');
    elements.globalLoaderFill = document.getElementById('global-loader-fill');
    elements.globalLoaderText = document.getElementById('global-loader-text');

    // Set up event listeners
    setupEventListeners();
    
    // Set default dates
    setDefaultDates();
}

/**
 * Set up all event listeners
 */
function setupEventListeners() {
    // File upload
    elements.fileInput.addEventListener('change', handleFileSelect);
    elements.uploadZone.addEventListener('click', (e) => {
        if (e.target.tagName !== 'BUTTON') {
            elements.fileInput.click();
        }
    });
    elements.uploadZone.addEventListener('dragover', handleDragOver);
    elements.uploadZone.addEventListener('dragleave', handleDragLeave);
    elements.uploadZone.addEventListener('drop', handleDrop);

    // Template upload
    elements.templateInput.addEventListener('change', handleTemplateSelect);

    // Range mode toggle
    document.querySelectorAll('input[name="range-mode"]').forEach(radio => {
        radio.addEventListener('change', handleRangeModeChange);
    });

    // Graphing
    elements.generateGraphBtn.addEventListener('click', generateGraph);

    // Create file button
    elements.createFileBtn.addEventListener('click', createCombinedFile);
}

/**
 * Set default start/end dates
 */
function setDefaultDates() {
    const today = new Date();
    const twoWeeksLater = new Date(today);
    twoWeeksLater.setDate(twoWeeksLater.getDate() + 14);

    elements.startDate.value = formatDateForInput(today);
    elements.endDate.value = formatDateForInput(twoWeeksLater);
}

/**
 * Format date for input[type="date"]
 */
function formatDateForInput(date) {
    return date.toISOString().split('T')[0];
}

// ===== GLOBAL LOADER =====

function showGlobalLoader(text = 'Processing...') {
    elements.globalLoader.classList.remove('hidden');
    elements.globalLoaderText.textContent = text;
    elements.globalLoaderFill.style.width = '0%';
}

function updateGlobalLoader(percent, text) {
    elements.globalLoaderFill.style.width = `${percent}%`;
    if (text) {
        elements.globalLoaderText.textContent = text;
    }
}

function hideGlobalLoader() {
    elements.globalLoader.classList.add('hidden');
}

// ===== TEMPLATE HANDLING =====

async function handleTemplateSelect(e) {
    const file = e.target.files[0];
    if (!file) return;

    try {
        showGlobalLoader('Loading template...');
        updateGlobalLoader(50, 'Reading template file...');
        
        const arrayBuffer = await file.arrayBuffer();
        // Store raw bytes to preserve all formatting
        AppState.templateBytes = arrayBuffer;
        // Also parse it to verify it's valid and get sheet info
        const testRead = XLSX.read(arrayBuffer, { type: 'array' });
        
        if (!testRead.Sheets['Analysis']) {
            throw new Error('Template must have a sheet named "Analysis"');
        }
        
        updateGlobalLoader(100, 'Template loaded!');
        elements.templateStatus.textContent = `✅ ${file.name}`;
        elements.templateStatus.classList.add('loaded');
        
        setTimeout(hideGlobalLoader, 500);
    } catch (error) {
        console.error('Error loading template:', error);
        elements.templateStatus.textContent = `❌ ${error.message || 'Error loading template'}`;
        elements.templateStatus.classList.remove('loaded');
        AppState.templateBytes = null;
        hideGlobalLoader();
    }
}

// ===== FILE UPLOAD HANDLERS =====

function handleDragOver(e) {
    e.preventDefault();
    elements.uploadZone.classList.add('drag-over');
}

function handleDragLeave(e) {
    e.preventDefault();
    elements.uploadZone.classList.remove('drag-over');
}

function handleDrop(e) {
    e.preventDefault();
    elements.uploadZone.classList.remove('drag-over');
    
    const files = e.dataTransfer.files;
    processFiles(files);
}

function handleFileSelect(e) {
    const files = e.target.files;
    processFiles(files);
}

/**
 * Process uploaded files
 */
async function processFiles(fileList) {
    const validExtensions = ['.csv', '.xls', '.xlsx'];
    const filesToProcess = [];

    for (const file of fileList) {
        const ext = '.' + file.name.split('.').pop().toLowerCase();
        if (validExtensions.includes(ext)) {
            filesToProcess.push(file);
        }
    }

    if (filesToProcess.length === 0) {
        showStatus('No valid files selected. Please upload CSV or Excel files.', 'error');
        return;
    }

    AppState.processingFiles = true;
    showGlobalLoader(`Processing 0 of ${filesToProcess.length} files...`);

    for (let i = 0; i < filesToProcess.length; i++) {
        const file = filesToProcess[i];
        const progress = ((i) / filesToProcess.length) * 100;
        updateGlobalLoader(progress, `Processing ${i + 1} of ${filesToProcess.length}: ${file.name}`);

        try {
            const result = await FileHandlers.readFile(file);
            
            // Detect datetime columns
            const dateTimeCols = FileHandlers.detectDateTimeColumns(result.columns, result.data);
            const dateTimeCol = dateTimeCols[0] || null;
            
            // Get selectable columns (excludes datetime and index-like columns)
            const selectableColumns = FileHandlers.getSelectableColumns(result.columns, dateTimeCols);
            
            // Check for duplicate timestamps
            const hasDuplicates = dateTimeCol ? 
                FileHandlers.hasDuplicateTimestamps(result.data, dateTimeCol) : false;
            
            // Get date range
            const dateRange = dateTimeCol ?
                FileHandlers.getDateRange(result.data, dateTimeCol) : { earliest: null, latest: null };

            // Store file data
            AppState.files[file.name] = {
                file: file,
                data: result.data,
                columns: result.columns,
                selectableColumns: selectableColumns,
                headerRow: result.headerRow,
                dateTimeCol: dateTimeCol,
                dateTimeCols: dateTimeCols,
                hasDuplicates: hasDuplicates,
                dateRange: dateRange,
                selectedCols: {},
                units: {},
                cleanup: {},
                dupeHandling: 'Average values'
            };

            if (result.headerRow > 0) {
                console.log(`Detected data starting on line ${result.headerRow + 1} in ${file.name}`);
            }

        } catch (error) {
            console.error(`Error processing ${file.name}:`, error);
            showStatus(`Error reading ${file.name}: ${error.message}`, 'error');
        }
    }

    updateGlobalLoader(100, `✅ Processed ${filesToProcess.length} files!`);
    
    setTimeout(() => {
        hideGlobalLoader();
        AppState.processingFiles = false;
    }, 800);

    // Update UI
    updateFileList();
    updateFileConfigs();
    updateSectionVisibility();
    updateAlignmentOptions();
    updateDefaultDatesFromData();
    updateGraphColumnOptions();

    showStatus(`✅ Uploaded ${Object.keys(AppState.files).length} file(s)!`, 'success');
}

/**
 * Update the file list display
 */
function updateFileList() {
    elements.fileList.innerHTML = '';

    for (const [fileName, fileInfo] of Object.entries(AppState.files)) {
        const fileItem = document.createElement('div');
        fileItem.className = 'file-item';
        fileItem.innerHTML = `
            <div>
                <span class="file-item-name">${escapeHtml(fileName)}</span>
                <span class="file-item-size">(${FileHandlers.formatFileSize(fileInfo.file.size)})</span>
            </div>
            <button class="file-item-remove" data-filename="${escapeHtml(fileName)}" title="Remove file">✕</button>
        `;
        elements.fileList.appendChild(fileItem);
    }

    // Add remove handlers
    elements.fileList.querySelectorAll('.file-item-remove').forEach(btn => {
        btn.addEventListener('click', (e) => {
            const fileName = e.target.dataset.filename;
            delete AppState.files[fileName];
            updateFileList();
            updateFileConfigs();
            updateSectionVisibility();
            updateAlignmentOptions();
            updateGraphColumnOptions();
        });
    });
}

/**
 * Update file configuration panels
 */
function updateFileConfigs() {
    elements.fileConfigs.innerHTML = '';

    for (const [fileName, fileInfo] of Object.entries(AppState.files)) {
        const panel = createFileConfigPanel(fileName, fileInfo);
        elements.fileConfigs.appendChild(panel);
    }
}

/**
 * Create a configuration panel for a file
 */
function createFileConfigPanel(fileName, fileInfo) {
    const panel = document.createElement('div');
    panel.className = 'file-config';
    panel.id = `config-${sanitizeId(fileName)}`;

    // Header
    const header = document.createElement('div');
    header.className = 'file-config-header';
    header.innerHTML = `
        <span class="file-config-title">📈 Data available from: ${escapeHtml(fileName)}</span>
        <span class="file-config-toggle">▼</span>
    `;
    header.addEventListener('click', () => panel.classList.toggle('open'));

    // Content
    const content = document.createElement('div');
    content.className = 'file-config-content';

    // Data preview table
    const preview = FileHandlers.getPreview(fileInfo.data, 2);
    const previewHtml = createDataTable(fileInfo.columns, preview);

    // Duplicate warning
    let dupeWarningHtml = '';
    if (fileInfo.hasDuplicates) {
        dupeWarningHtml = `
            <div class="duplicate-warning">
                <div class="duplicate-warning-title">⚠️ Duplicate timestamps detected</div>
                <div class="form-group">
                    <label>How should duplicates be handled?</label>
                    <select class="select dupe-handling" data-filename="${escapeHtml(fileName)}">
                        <option value="Average values">Average values</option>
                        <option value="Maximum value">Maximum value</option>
                        <option value="Minimum value">Minimum value</option>
                    </select>
                </div>
            </div>
        `;
    }

    // Column selection with checkboxes
    const columnsHtml = createColumnCheckboxes(fileName, fileInfo);

    content.innerHTML = `
        <div class="data-preview">${previewHtml}</div>
        ${dupeWarningHtml}
        <h4 style="color: var(--text-primary); margin-bottom: var(--spacing-md);">
            Select data columns to include in the combination
        </h4>
        ${columnsHtml}
        <div class="column-settings" id="column-settings-${sanitizeId(fileName)}"></div>
    `;

    panel.appendChild(header);
    panel.appendChild(content);

    // Set up event listeners after adding to DOM
    setTimeout(() => {
        setupFileConfigListeners(fileName, fileInfo);
    }, 0);

    return panel;
}

/**
 * Create HTML for data preview table
 */
function createDataTable(columns, rows) {
    let html = '<table class="data-table"><thead><tr>';
    for (const col of columns) {
        html += `<th>${escapeHtml(col)}</th>`;
    }
    html += '</tr></thead><tbody>';

    for (const row of rows) {
        html += '<tr>';
        for (const col of columns) {
            const value = row[col];
            html += `<td>${value != null ? escapeHtml(String(value)) : ''}</td>`;
        }
        html += '</tr>';
    }

    html += '</tbody></table>';
    return html;
}

/**
 * Create column selector with checkboxes (allows multiple selection easily)
 */
function createColumnCheckboxes(fileName, fileInfo) {
    const selectableCols = fileInfo.selectableColumns || [];
    
    if (selectableCols.length === 0) {
        return '<div class="column-checkbox-list"><p class="checkbox-grid-empty">No selectable columns found</p></div>';
    }
    
    let html = `<div class="column-checkbox-list" id="col-list-${sanitizeId(fileName)}">`;
    for (const col of selectableCols) {
        const checkId = `col-check-${sanitizeId(fileName)}-${sanitizeId(col)}`;
        html += `
            <div class="column-checkbox-item">
                <input type="checkbox" 
                    id="${checkId}" 
                    data-filename="${escapeHtml(fileName)}" 
                    data-column="${escapeHtml(col)}"
                    class="column-checkbox">
                <label for="${checkId}">${escapeHtml(col)}</label>
            </div>
        `;
    }
    html += '</div>';
    
    return html;
}

/**
 * Set up event listeners for file config panel
 */
function setupFileConfigListeners(fileName, fileInfo) {
    const panel = document.getElementById(`config-${sanitizeId(fileName)}`);
    if (!panel) return;

    // Column checkboxes
    panel.querySelectorAll('.column-checkbox').forEach(checkbox => {
        checkbox.addEventListener('change', () => {
            updateSelectedColumns(fileName);
        });
    });

    // Duplicate handling
    const dupeSelect = panel.querySelector('.dupe-handling');
    if (dupeSelect) {
        dupeSelect.addEventListener('change', (e) => {
            AppState.files[fileName].dupeHandling = e.target.value;
        });
    }
}

/**
 * Update selected columns based on checkbox state
 */
function updateSelectedColumns(fileName) {
    const panel = document.getElementById(`config-${sanitizeId(fileName)}`);
    if (!panel) return;

    const fileInfo = AppState.files[fileName];
    const checkedBoxes = panel.querySelectorAll('.column-checkbox:checked');
    
    const selectedColumns = Array.from(checkedBoxes).map(cb => cb.dataset.column);
    
    // Update selected columns in state
    const newSelectedCols = {};
    const newUnits = {};
    const newCleanup = {};

    for (const col of selectedColumns) {
        newSelectedCols[col] = fileInfo.selectedCols[col] || col;
        newUnits[col] = fileInfo.units[col] || '';
        newCleanup[col] = fileInfo.cleanup[col] || 'Fill with nearest available value';
    }

    fileInfo.selectedCols = newSelectedCols;
    fileInfo.units = newUnits;
    fileInfo.cleanup = newCleanup;

    // Update column settings UI
    updateColumnSettingsUI(fileName, selectedColumns);
    
    // Update graph column options
    updateGraphColumnOptions();
}

/**
 * Update column settings panel UI
 */
function updateColumnSettingsUI(fileName, selectedColumns) {
    const fileInfo = AppState.files[fileName];
    const settingsContainer = document.getElementById(`column-settings-${sanitizeId(fileName)}`);
    
    if (!settingsContainer) return;

    // Build settings HTML
    let html = '';
    for (const col of selectedColumns) {
        html += `
            <div class="column-setting-item" id="col-setting-${sanitizeId(fileName)}-${sanitizeId(col)}">
                <div class="column-setting-header" onclick="this.parentElement.classList.toggle('open')">
                    <span class="column-setting-name">⚙️ Settings for ${escapeHtml(col)}</span>
                    <span>▼</span>
                </div>
                <div class="column-setting-content">
                    <div class="column-setting-grid">
                        <div class="form-group">
                            <label>Data Column Title</label>
                            <input type="text" class="input col-title" 
                                data-filename="${escapeHtml(fileName)}" 
                                data-column="${escapeHtml(col)}"
                                value="${escapeHtml(fileInfo.selectedCols[col] || col)}">
                        </div>
                        <div class="form-group">
                            <label>Units</label>
                            <input type="text" class="input col-units"
                                data-filename="${escapeHtml(fileName)}"
                                data-column="${escapeHtml(col)}"
                                value="${escapeHtml(fileInfo.units[col] || '')}">
                        </div>
                        <div class="form-group">
                            <label>Missing data handling</label>
                            <select class="select col-cleanup"
                                data-filename="${escapeHtml(fileName)}"
                                data-column="${escapeHtml(col)}">
                                <option value="Fill with nearest available value" 
                                    ${fileInfo.cleanup[col] === 'Fill with nearest available value' ? 'selected' : ''}>
                                    Fill with nearest available value
                                </option>
                                <option value="Fill with a linear interpolation between the nearest values"
                                    ${fileInfo.cleanup[col] === 'Fill with a linear interpolation between the nearest values' ? 'selected' : ''}>
                                    Fill with linear interpolation
                                </option>
                                <option value="Delete the entire row of data"
                                    ${fileInfo.cleanup[col] === 'Delete the entire row of data' ? 'selected' : ''}>
                                    Delete the entire row
                                </option>
                                <option value="Fill with zero"
                                    ${fileInfo.cleanup[col] === 'Fill with zero' ? 'selected' : ''}>
                                    Fill with zero
                                </option>
                            </select>
                        </div>
                    </div>
                </div>
            </div>
        `;
    }

    settingsContainer.innerHTML = html;

    // Add event listeners for settings inputs
    settingsContainer.querySelectorAll('.col-title').forEach(input => {
        input.addEventListener('change', (e) => {
            const fn = e.target.dataset.filename;
            const col = e.target.dataset.column;
            AppState.files[fn].selectedCols[col] = e.target.value;
            updateGraphColumnOptions();
        });
    });

    settingsContainer.querySelectorAll('.col-units').forEach(input => {
        input.addEventListener('change', (e) => {
            const fn = e.target.dataset.filename;
            const col = e.target.dataset.column;
            AppState.files[fn].units[col] = e.target.value;
        });
    });

    settingsContainer.querySelectorAll('.col-cleanup').forEach(select => {
        select.addEventListener('change', (e) => {
            const fn = e.target.dataset.filename;
            const col = e.target.dataset.column;
            AppState.files[fn].cleanup[col] = e.target.value;
        });
    });
}

// ===== SECTION VISIBILITY =====

function updateSectionVisibility() {
    const hasFiles = Object.keys(AppState.files).length > 0;

    elements.columnsSection.classList.toggle('hidden', !hasFiles);
    elements.graphSection.classList.toggle('hidden', !hasFiles);
    elements.timeSection.classList.toggle('hidden', !hasFiles);
    elements.downloadSection.classList.toggle('hidden', !hasFiles);
}

// ===== ALIGNMENT OPTIONS =====

function updateAlignmentOptions() {
    elements.alignmentOptions.innerHTML = '';

    for (const fileName of Object.keys(AppState.files)) {
        const div = document.createElement('div');
        div.className = 'form-group';
        div.innerHTML = `
            <label>Data from '${escapeHtml(fileName)}':</label>
            <select class="select alignment-select" data-filename="${escapeHtml(fileName)}">
                <option value="Fill with the nearest value">Fill with the nearest value</option>
                <option value="Do a linear interpolation from the nearest values">Do a linear interpolation from the nearest values</option>
                <option value="Take an average of the available values within the interval">Take an average of the available values within the interval</option>
            </select>
        `;
        elements.alignmentOptions.appendChild(div);
    }
}

// ===== DATE DEFAULTS =====

function updateDefaultDatesFromData() {
    let latestStart = null;

    for (const fileInfo of Object.values(AppState.files)) {
        if (fileInfo.dateRange.earliest) {
            if (!latestStart || fileInfo.dateRange.earliest > latestStart) {
                latestStart = fileInfo.dateRange.earliest;
            }
        }
    }

    if (latestStart) {
        const startDate = new Date(latestStart);
        startDate.setDate(startDate.getDate() + 1);
        elements.startDate.value = formatDateForInput(startDate);

        const endDate = new Date(startDate);
        endDate.setDate(endDate.getDate() + 14);
        elements.endDate.value = formatDateForInput(endDate);
    }
}

// ===== RANGE MODE =====

function handleRangeModeChange(e) {
    const mode = e.target.value;
    document.getElementById('end-date-inputs').classList.toggle('hidden', mode === 'duration');
    document.getElementById('duration-input').classList.toggle('hidden', mode === 'end-date');
}

// ===== GRAPHING =====

/**
 * Update graph column options with checkboxes
 */
function updateGraphColumnOptions() {
    elements.graphColumns.innerHTML = '';

    let hasColumns = false;

    for (const [fileName, fileInfo] of Object.entries(AppState.files)) {
        for (const [col, title] of Object.entries(fileInfo.selectedCols)) {
            hasColumns = true;
            const checkId = `graph-check-${sanitizeId(fileName)}-${sanitizeId(col)}`;
            const div = document.createElement('div');
            div.className = 'checkbox-grid-item';
            div.innerHTML = `
                <input type="checkbox" 
                    id="${checkId}"
                    data-filename="${escapeHtml(fileName)}"
                    data-column="${escapeHtml(col)}"
                    class="graph-column-checkbox">
                <label for="${checkId}">${escapeHtml(fileName)} - ${escapeHtml(title)}</label>
            `;
            elements.graphColumns.appendChild(div);
            
            // Add click handler to toggle checked class on parent
            div.querySelector('input').addEventListener('change', (e) => {
                div.classList.toggle('checked', e.target.checked);
            });
        }
    }

    if (!hasColumns) {
        elements.graphColumns.innerHTML = '<p class="checkbox-grid-empty">Select columns from your files above to enable graphing</p>';
    }
}

/**
 * Generate graph with selected columns - optimized for performance
 */
async function generateGraph() {
    const checkedBoxes = elements.graphColumns.querySelectorAll('.graph-column-checkbox:checked');
    
    if (checkedBoxes.length === 0) {
        showStatus('Please select at least one column to graph.', 'warning');
        return;
    }

    showGlobalLoader('Generating graph...');
    updateGlobalLoader(10, 'Loading data...');

    const traces = [];
    const totalColumns = checkedBoxes.length;
    let processed = 0;
    let totalPoints = 0;

    for (const checkbox of checkedBoxes) {
        const fileName = checkbox.dataset.filename;
        const col = checkbox.dataset.column;
        const fileInfo = AppState.files[fileName];
        
        if (!fileInfo) continue;

        const dateTimeCol = fileInfo.dateTimeCol;
        const data = fileInfo.data;

        if (!dateTimeCol) {
            console.warn(`No datetime column found for ${fileName}`);
            continue;
        }

        // Sort by datetime
        const sorted = [...data].sort((a, b) => {
            return new Date(a[dateTimeCol]) - new Date(b[dateTimeCol]);
        });

        let x = [];
        let y = [];
        
        for (const row of sorted) {
            const xVal = new Date(row[dateTimeCol]);
            const yVal = parseFloat(row[col]);
            
            if (!isNaN(xVal.getTime()) && !isNaN(yVal)) {
                x.push(xVal);
                y.push(yVal);
            }
        }

        // Downsample if too many points (keep max 2000 points per series for smooth interaction)
        const maxPoints = 2000;
        if (x.length > maxPoints) {
            const step = Math.ceil(x.length / maxPoints);
            const downsampledX = [];
            const downsampledY = [];
            for (let i = 0; i < x.length; i += step) {
                downsampledX.push(x[i]);
                downsampledY.push(y[i]);
            }
            x = downsampledX;
            y = downsampledY;
        }

        if (x.length > 0) {
            totalPoints += x.length;
            traces.push({
                x: x,
                y: y,
                mode: 'lines',  // No markers for better performance
                name: fileInfo.selectedCols[col] || col,
                type: 'scattergl',  // WebGL for faster rendering
                line: { width: 2 }
            });
        }

        processed++;
        updateGlobalLoader(10 + (processed / totalColumns) * 80, `Processing ${processed} of ${totalColumns} columns...`);
    }

    if (traces.length === 0) {
        hideGlobalLoader();
        showStatus('No valid data found for the selected columns.', 'warning');
        return;
    }

    updateGlobalLoader(95, 'Rendering chart...');

    const layout = {
        title: {
            text: 'Combined Data Plot',
            font: { size: 14, color: '#00ffff' }
        },
        hovermode: 'closest',  // Faster than 'x unified' for large datasets
        legend: { 
            orientation: 'h', 
            y: -0.15,
            font: { color: '#e0e0e0', size: 10 }
        },
        paper_bgcolor: '#12121a',
        plot_bgcolor: '#0a0a0f',
        font: { color: '#e0e0e0' },
        margin: { t: 40, b: 60, l: 50, r: 20 },
        xaxis: {
            gridcolor: '#2a2a3a',
            title: { text: 'Date/Time', font: { size: 11, color: '#00ffff' } },
            tickfont: { color: '#e0e0e0', size: 10 },
            rangeslider: { visible: false }  // We'll use drag selection instead
        },
        yaxis: {
            gridcolor: '#2a2a3a',
            title: { text: 'Value', font: { size: 11, color: '#00ffff' } },
            tickfont: { color: '#e0e0e0', size: 10 }
        },
        // Enable drag-to-zoom on x-axis
        dragmode: 'zoom',
        selectdirection: 'h'
    };

    elements.graphContainer.classList.remove('hidden');
    
    // Create the plot
    Plotly.newPlot('plot', traces, layout, { 
        responsive: true,
        displayModeBar: true,
        modeBarButtonsToRemove: ['lasso2d', 'select2d'],
        scrollZoom: true
    });
    
    // Listen for zoom/range selection events to update time inputs
    const plotElement = document.getElementById('plot');
    plotElement.on('plotly_relayout', function(eventData) {
        // Check if x-axis range was changed (via zoom or drag)
        if (eventData['xaxis.range[0]'] && eventData['xaxis.range[1]']) {
            const startDate = new Date(eventData['xaxis.range[0]']);
            const endDate = new Date(eventData['xaxis.range[1]']);
            
            if (!isNaN(startDate.getTime()) && !isNaN(endDate.getTime())) {
                updateTimeInputsFromGraph(startDate, endDate);
            }
        }
        // Also handle autorange reset (double-click to reset)
        if (eventData['xaxis.autorange']) {
            // User reset the view - could optionally reset inputs here
        }
    });
    
    updateGlobalLoader(100, `✅ Graph generated!`);
    setTimeout(hideGlobalLoader, 800);
    
    showStatus(`✅ Graph generated with ${traces.length} series (${totalPoints.toLocaleString()} points) - Drag to zoom and set time range!`, 'success');
}

/**
 * Update time inputs based on graph selection
 */
function updateTimeInputsFromGraph(startDate, endDate) {
    // Update start date/time
    elements.startDate.value = formatDateForInput(startDate);
    elements.startTime.value = formatTimeForInput(startDate);
    
    // Update end date/time
    elements.endDate.value = formatDateForInput(endDate);
    elements.endTime.value = formatTimeForInput(endDate);
    
    // Make sure we're in "End date" mode
    document.querySelector('input[name="range-mode"][value="end-date"]').checked = true;
    document.getElementById('end-date-inputs').classList.remove('hidden');
    document.getElementById('duration-input').classList.add('hidden');
    
    // Visual feedback
    highlightTimeInputs();
    
    showStatus(`📅 Time range updated from graph selection!`, 'success');
}

/**
 * Format time for input[type="time"]
 */
function formatTimeForInput(date) {
    const hours = String(date.getHours()).padStart(2, '0');
    const minutes = String(date.getMinutes()).padStart(2, '0');
    return `${hours}:${minutes}`;
}

/**
 * Briefly highlight the time inputs to show they were updated
 */
function highlightTimeInputs() {
    const inputs = [elements.startDate, elements.startTime, elements.endDate, elements.endTime];
    inputs.forEach(input => {
        input.style.transition = 'box-shadow 0.3s, border-color 0.3s';
        input.style.boxShadow = '0 0 15px rgba(0, 255, 255, 0.5)';
        input.style.borderColor = 'var(--neon-cyan)';
        
        setTimeout(() => {
            input.style.boxShadow = '';
            input.style.borderColor = '';
        }, 1500);
    });
}

// ===== CREATE COMBINED FILE =====

async function createCombinedFile() {
    // Validate we have data
    let hasSelectedColumns = false;
    for (const fileInfo of Object.values(AppState.files)) {
        if (Object.keys(fileInfo.selectedCols).length > 0) {
            hasSelectedColumns = true;
            break;
        }
    }

    if (!hasSelectedColumns) {
        showStatus('Please select at least one column from at least one file.', 'error');
        return;
    }

    // Get time settings
    const startDate = new Date(elements.startDate.value + 'T' + elements.startTime.value);
    
    let endDate;
    const rangeMode = document.querySelector('input[name="range-mode"]:checked').value;
    if (rangeMode === 'end-date') {
        endDate = new Date(elements.endDate.value + 'T' + elements.endTime.value);
    } else {
        const days = parseInt(elements.durationDays.value) || 14;
        endDate = new Date(startDate);
        endDate.setDate(endDate.getDate() + days);
    }

    const interval = elements.interval.value;

    // Get alignment options
    const alignmentOptions = {};
    elements.alignmentOptions.querySelectorAll('.alignment-select').forEach(select => {
        alignmentOptions[select.dataset.filename] = select.value;
    });

    // Show progress
    elements.createFileBtn.disabled = true;
    elements.progressContainer.classList.remove('hidden');
    showGlobalLoader('Creating combined file...');
    updateProgress(0, 'Starting file processing...');
    updateGlobalLoader(0, 'Starting file processing...');

    try {
        // Generate timestamps
        const timestamps = DataProcessing.generateTimeIndex(startDate, endDate, interval);
        updateProgress(10, `Generated ${timestamps.length} timestamps...`);
        updateGlobalLoader(10, `Generated ${timestamps.length} timestamps...`);

        // Create combined dataset
        const combined = DataProcessing.createCombinedDataset(
            AppState.files, 
            timestamps, 
            alignmentOptions
        );
        updateProgress(50, 'Data combined, preparing Excel file...');
        updateGlobalLoader(50, 'Data combined, preparing Excel file...');

        // Use uploaded template or create new workbook
        let workbook;
        if (AppState.templateBytes) {
            // Read fresh from raw bytes to preserve ALL formatting
            workbook = XLSX.read(AppState.templateBytes, { 
                type: 'array',
                cellStyles: true,      // Preserve cell styles
                cellFormula: true,     // Preserve formulas
                cellNF: true,          // Preserve number formats
                sheetStubs: true       // Preserve empty cells with formatting
            });
            console.log('Using uploaded template (preserving formatting)');
        } else {
            console.log('Creating new workbook (no template)');
            workbook = XLSX.utils.book_new();
            const ws = XLSX.utils.aoa_to_sheet([[]]);
            XLSX.utils.book_append_sheet(workbook, ws, 'Analysis');
        }

        updateProgress(70, 'Writing data to Excel...');
        updateGlobalLoader(70, 'Writing data to Excel...');

        // Get or create Analysis sheet
        let ws = workbook.Sheets['Analysis'];
        if (!ws) {
            ws = XLSX.utils.aoa_to_sheet([[]]);
            XLSX.utils.book_append_sheet(workbook, ws, 'Analysis');
        }

        // Write headers (Row 10 in Excel = index 9, Column B = index 1)
        const headers = ['Date', ...combined.columns];
        const units = ['', ...combined.units];

        // Row 10 - Headers (B10, C10, D10, ...)
        for (let c = 0; c < headers.length; c++) {
            const cellRef = XLSX.utils.encode_cell({ r: 9, c: c + 1 });
            // Preserve existing cell style if present
            const existingCell = ws[cellRef];
            ws[cellRef] = { 
                t: 's', 
                v: headers[c],
                s: existingCell?.s  // Keep existing style
            };
        }

        // Row 11 - Units (B11, C11, D11, ...)
        for (let c = 0; c < units.length; c++) {
            const cellRef = XLSX.utils.encode_cell({ r: 10, c: c + 1 });
            const existingCell = ws[cellRef];
            ws[cellRef] = { 
                t: 's', 
                v: units[c],
                s: existingCell?.s
            };
        }

        // Row 12 onward - Data
        for (let r = 0; r < combined.data.length; r++) {
            const row = combined.data[r];
            
            // Date column (B12, B13, ...)
            const dateCell = XLSX.utils.encode_cell({ r: 11 + r, c: 1 });
            const existingDateCell = ws[dateCell];
            ws[dateCell] = { 
                t: 'd', 
                v: row.DateTime,
                z: existingDateCell?.z || 'm/d/yy h:mm',  // Use existing format or default
                s: existingDateCell?.s
            };

            // Data columns (C12, D12, ... onwards)
            for (let c = 0; c < combined.columns.length; c++) {
                const cellRef = XLSX.utils.encode_cell({ r: 11 + r, c: c + 2 });
                const value = row[combined.columns[c]];
                const existingCell = ws[cellRef];
                
                if (value !== null && value !== undefined && !isNaN(value)) {
                    ws[cellRef] = { 
                        t: 'n', 
                        v: value,
                        s: existingCell?.s,
                        z: existingCell?.z  // Preserve number format
                    };
                }
            }

            // Update progress periodically
            if (r % 100 === 0) {
                const pct = 70 + (r / combined.data.length) * 25;
                updateProgress(pct, `Writing row ${r + 1} of ${combined.data.length}...`);
                updateGlobalLoader(pct, `Writing row ${r + 1} of ${combined.data.length}...`);
            }
        }

        // Update sheet range
        const lastRow = 11 + combined.data.length;
        const lastCol = 1 + combined.columns.length;
        ws['!ref'] = XLSX.utils.encode_range({
            s: { r: 0, c: 0 },
            e: { r: lastRow, c: lastCol }
        });

        updateProgress(95, 'Generating download...');
        updateGlobalLoader(95, 'Generating download...');

        // Generate file - use bookSST for better string handling
        const wbout = XLSX.write(workbook, { 
            bookType: 'xlsx', 
            type: 'array',
            cellStyles: true  // Try to preserve styles on write
        });
        const blob = new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        
        // Create download link
        const url = URL.createObjectURL(blob);
        
        elements.downloadBtn.onclick = () => {
            const a = document.createElement('a');
            a.href = url;
            a.download = 'Analysis.xlsx';
            a.click();
        };

        elements.downloadContainer.classList.remove('hidden');
        updateProgress(100, '✅ Combined file created!');
        updateGlobalLoader(100, '✅ Combined file created!');

        setTimeout(hideGlobalLoader, 800);

        showStatus('✅ Combined file created! Click the download button to save.', 'success');

    } catch (error) {
        console.error('Error creating combined file:', error);
        showStatus(`Error creating file: ${error.message}`, 'error');
        hideGlobalLoader();
    } finally {
        elements.createFileBtn.disabled = false;
    }
}

// ===== UTILITY FUNCTIONS =====

function showStatus(message, type = 'info') {
    elements.uploadStatus.textContent = message;
    elements.uploadStatus.className = `status-message status-${type}`;
    elements.uploadStatus.classList.remove('hidden');
    
    // Auto-hide after 5 seconds for non-errors
    if (type !== 'error') {
        setTimeout(() => {
            elements.uploadStatus.classList.add('hidden');
        }, 5000);
    }
}

function updateProgress(percent, text) {
    elements.progressFill.style.width = `${percent}%`;
    elements.progressText.textContent = text;
}

function sanitizeId(str) {
    return str.replace(/[^a-zA-Z0-9]/g, '_');
}

function escapeHtml(str) {
    const div = document.createElement('div');
    div.textContent = str;
    return div.innerHTML;
}

// Initialize when DOM is ready
document.addEventListener('DOMContentLoaded', init);
