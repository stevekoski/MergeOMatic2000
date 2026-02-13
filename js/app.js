/**
 * Merge-o-matic 2000 - Main Application
 */

// Application State
const AppState = {
    files: {},          // fileName -> { file, data, columns, headerRow, ... }
    stacks: {},         // stackName -> { files: [], data, columns, dateRange, ... }
    selectedForStack: new Set(),  // Files currently selected for new stack
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
    
    // Stacking elements
    elements.stackSection = document.getElementById('stack-section');
    elements.existingStacks = document.getElementById('existing-stacks');
    elements.stackCreator = document.getElementById('stack-creator');
    elements.stackableFiles = document.getElementById('stackable-files');
    elements.stackWarnings = document.getElementById('stack-warnings');
    elements.stackInfo = document.getElementById('stack-info');
    elements.stackName = document.getElementById('stack-name');
    elements.overlapHandling = document.getElementById('overlap-handling');
    elements.createStackBtn = document.getElementById('create-stack-btn');

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
    
    // Stacking
    elements.createStackBtn.addEventListener('click', createStack);
    elements.stackName.addEventListener('input', updateCreateStackButton);
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
            
            // Detect if data is in long format (needs pivoting)
            const longFormatInfo = FileHandlers.detectLongFormat(result.data, result.columns, dateTimeCol);
            
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
                dupeHandling: 'Average values',
                longFormatInfo: longFormatInfo,  // Store pivot detection info
                isPivoted: false  // Track if user has applied pivot
            };

            if (result.headerRow > 0) {
                console.log(`Detected data starting on line ${result.headerRow + 1} in ${file.name}`);
            }
            
            if (longFormatInfo) {
                console.log(`Detected long format in ${file.name}: ${longFormatInfo.tagCount} unique tags`);
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
    updateStackingSection();

    showStatus(`✅ Uploaded ${Object.keys(AppState.files).length} file(s)!`, 'success');
}

/**
 * Update the file list display
 */
function updateFileList() {
    elements.fileList.innerHTML = '';
    
    // Get set of files that are in stacks
    const stackedFiles = getStackedFiles();

    for (const [fileName, fileInfo] of Object.entries(AppState.files)) {
        const isInStack = stackedFiles.has(fileName);
        const stackName = isInStack ? getStackNameForFile(fileName) : null;
        
        const fileItem = document.createElement('div');
        fileItem.className = 'file-item' + (isInStack ? ' in-stack' : '');
        fileItem.innerHTML = `
            <div>
                <span class="file-item-name">${escapeHtml(fileName)}</span>
                <span class="file-item-size">(${FileHandlers.formatFileSize(fileInfo.file.size)})</span>
                ${isInStack ? `<span class="file-item-stack-badge">In: ${escapeHtml(stackName)}</span>` : ''}
            </div>
            <button class="file-item-remove" data-filename="${escapeHtml(fileName)}" title="Remove file">✕</button>
        `;
        elements.fileList.appendChild(fileItem);
    }

    // Add remove handlers
    elements.fileList.querySelectorAll('.file-item-remove').forEach(btn => {
        btn.addEventListener('click', (e) => {
            const fileName = e.target.dataset.filename;
            
            // Remove from any stacks that contain this file
            removeFileFromAllStacks(fileName);
            
            delete AppState.files[fileName];
            AppState.selectedForStack.delete(fileName);
            
            updateFileList();
            updateFileConfigs();
            updateSectionVisibility();
            updateAlignmentOptions();
            updateGraphColumnOptions();
            updateStackingSection();
        });
    });
}

/**
 * Update file configuration panels
 */
function updateFileConfigs() {
    elements.fileConfigs.innerHTML = '';
    
    const stackedFiles = getStackedFiles();

    // First, add panels for stacks
    for (const [stackName, stackInfo] of Object.entries(AppState.stacks)) {
        const panel = createFileConfigPanel(stackName, stackInfo, true);
        elements.fileConfigs.appendChild(panel);
    }

    // Then, add panels for individual files (that aren't in stacks)
    for (const [fileName, fileInfo] of Object.entries(AppState.files)) {
        if (!stackedFiles.has(fileName)) {
            const panel = createFileConfigPanel(fileName, fileInfo, false);
            elements.fileConfigs.appendChild(panel);
        }
    }
}


/**
 * Create a configuration panel for a file or stack
 */
function createFileConfigPanel(name, info, isStack = false) {
    const panel = document.createElement('div');
    panel.className = 'file-config' + (isStack ? ' is-stack' : '');
    panel.id = `config-${sanitizeId(name)}`;

    // Header
    const header = document.createElement('div');
    header.className = 'file-config-header';
    const icon = isStack ? '📚' : '📈';
    const label = isStack ? 'Stacked data from' : 'Data available from';
    header.innerHTML = `
        <span class="file-config-title">${icon} ${label}: ${escapeHtml(name)}</span>
        <span class="file-config-toggle">▼</span>
    `;
    header.addEventListener('click', () => panel.classList.toggle('open'));

    // Content
    const content = document.createElement('div');
    content.className = 'file-config-content';

    // Data preview table
    const preview = FileHandlers.getPreview(info.data, 2);
    const previewHtml = createDataTable(info.columns, preview);

    // Stack info badge
    let stackInfoHtml = '';
    if (isStack) {
        stackInfoHtml = `
            <div class="stack-success" style="margin-bottom: var(--spacing-md);">
                📊 Contains ${info.files.length} stacked files: ${info.files.join(', ')} 
                <br>Total: ${info.rowCount.toLocaleString()} rows
            </div>
        `;
    }

    // Check for separate Date and Time columns (works for both files and stacks)
    const separateDateTimeCols = FileHandlers.detectSeparateDateTimeColumns(info.columns, info.data);
    
    // Get columns from multiple sources to ensure we have ALL of them
    // 1. From info.columns (the declared columns)
    // 2. From data keys (what's actually in the data)
    const declaredColumns = info.columns || [];
    const dataKeysSet = new Set();
    
    // Get keys from first 10 rows of data to catch all possible columns
    const sampleSize = Math.min(10, info.data.length);
    for (let i = 0; i < sampleSize; i++) {
        if (info.data[i]) {
            Object.keys(info.data[i]).forEach(key => dataKeysSet.add(key));
        }
    }
    const dataColumns = [...dataKeysSet];
    
    // Combine both sources - use declared columns order, then add any missing data keys
    const allColumns = [...declaredColumns];
    for (const col of dataColumns) {
        if (!allColumns.includes(col)) {
            allColumns.push(col);
        }
    }
    
    // Get potential tag and value columns for user selection
    const excludeCols = [info.dateTimeCol, ...(info.dateTimeCols || [])].filter(c => c);
    if (separateDateTimeCols) {
        excludeCols.push(separateDateTimeCols.dateCol, separateDateTimeCols.timeCol);
    }
    
    // Get all available columns (not excluded)
    const availableColumns = allColumns.filter(col => !excludeCols.includes(col));
    
    // Try to detect which are likely tag vs value columns
    const potentialTagCols = FileHandlers.getPotentialTagColumns(allColumns, excludeCols, info.data);
    const potentialValueCols = FileHandlers.getPotentialValueColumns(allColumns, excludeCols, info.data);
    
    // Debug logging
    console.log(`Pivot UI for "${name}":`, {
        declaredColumns,
        dataColumns,
        allColumns,
        availableColumns,
        excludeCols,
        potentialTagCols,
        potentialValueCols,
        separateDateTimeCols
    });

    // Long format (pivot) UI - show for both files and stacks if not already pivoted
    let pivotHtml = '';
    if (!info.isPivoted) {
        // Auto-detected suggestion
        const lf = info.longFormatInfo;
        const suggestedTag = lf?.tagCol || '';
        const suggestedValue = lf?.valueCol || '';
        
        // Build tag column options - ALL available columns, with likely ones first
        let tagOptions = '<option value="">-- Select tag/variable column --</option>';
        const addedTagCols = new Set();
        
        // First add likely tag columns (string-like)
        for (const col of potentialTagCols) {
            const selected = col === suggestedTag ? 'selected' : '';
            tagOptions += `<option value="${escapeHtml(col)}" ${selected}>${escapeHtml(col.trim())}</option>`;
            addedTagCols.add(col);
        }
        // Then add ALL remaining available columns
        for (const col of availableColumns) {
            if (!addedTagCols.has(col)) {
                tagOptions += `<option value="${escapeHtml(col)}">${escapeHtml(col.trim())}</option>`;
            }
        }
        
        // Build value column options - ALL available columns, with likely ones first
        let valueOptions = '<option value="">-- Select value column --</option>';
        const addedValueCols = new Set();
        
        // First add likely value columns (numeric)
        for (const col of potentialValueCols) {
            const selected = col === suggestedValue ? 'selected' : '';
            valueOptions += `<option value="${escapeHtml(col)}" ${selected}>${escapeHtml(col.trim())}</option>`;
            addedValueCols.add(col);
        }
        // Then add ALL remaining available columns
        for (const col of availableColumns) {
            if (!addedValueCols.has(col)) {
                valueOptions += `<option value="${escapeHtml(col)}">${escapeHtml(col.trim())}</option>`;
            }
        }
        
        const autoDetectedMsg = lf ? 
            `<p class="pivot-auto-detected">✨ Auto-detected: Tag column "<strong>${escapeHtml(lf.tagCol.trim())}</strong>" with ${lf.tagCount} unique tags</p>` : 
            '';
        
        const separateDateTimeMsg = separateDateTimeCols ?
            `<p class="pivot-datetime-notice">📅 Separate Date and Time columns detected - they will be combined automatically when pivoting.</p>` :
            '';
        
        pivotHtml = `
            <div class="pivot-notice">
                <div class="pivot-notice-header">
                    <span class="pivot-icon">🔄</span>
                    <span class="pivot-title">Pivot Data (Long → Wide Format)</span>
                </div>
                <p class="pivot-description">
                    If your data has multiple rows per timestamp (e.g., one row per sensor/tag), 
                    you can pivot it to wide format where each tag becomes its own column.
                </p>
                ${autoDetectedMsg}
                ${separateDateTimeMsg}
                <div class="pivot-controls">
                    <div class="form-group">
                        <label>Tag/Variable Name Column</label>
                        <select class="select pivot-tag-select" data-filename="${escapeHtml(name)}" data-isstack="${isStack}">
                            ${tagOptions}
                        </select>
                    </div>
                    <div class="form-group">
                        <label>Value Column</label>
                        <select class="select pivot-value-select" data-filename="${escapeHtml(name)}" data-isstack="${isStack}">
                            ${valueOptions}
                        </select>
                    </div>
                </div>
                <button class="btn btn-primary pivot-btn" data-filename="${escapeHtml(name)}" data-isstack="${isStack}">
                    🔄 Pivot to Wide Format
                </button>
            </div>
        `;
    } else {
        pivotHtml = `
            <div class="pivot-success">
                ✅ Data has been pivoted to wide format (${info.selectableColumns.length} columns from ${info.originalTagCount || 'multiple'} tags)
            </div>
        `;
    }

    // Duplicate warning (only for non-pivoted data, since pivot removes dupes)
    let dupeWarningHtml = '';
    if (info.hasDuplicates && !info.longFormatInfo && !info.isPivoted) {
        dupeWarningHtml = `
            <div class="duplicate-warning">
                <div class="duplicate-warning-title">⚠️ Duplicate timestamps detected</div>
                <div class="form-group">
                    <label>How should duplicates be handled?</label>
                    <select class="select dupe-handling" data-filename="${escapeHtml(name)}">
                        <option value="Average values">Average values</option>
                        <option value="Maximum value">Maximum value</option>
                        <option value="Minimum value">Minimum value</option>
                    </select>
                </div>
            </div>
        `;
    }

    // Column selection with checkboxes
    const columnsHtml = createColumnCheckboxes(name, info, isStack);

    content.innerHTML = `
        ${stackInfoHtml}
        ${pivotHtml}
        <div class="data-preview">${previewHtml}</div>
        ${dupeWarningHtml}
        <h4 style="color: var(--text-primary); margin-bottom: var(--spacing-md);">
            Select data columns to include in the combination
        </h4>
        ${columnsHtml}
        <div class="column-settings" id="column-settings-${sanitizeId(name)}"></div>
    `;

    panel.appendChild(header);
    panel.appendChild(content);

    // Set up event listeners after adding to DOM
    setTimeout(() => {
        setupFileConfigListeners(name, info, isStack);
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
function createColumnCheckboxes(name, info, isStack = false) {
    const selectableCols = info.selectableColumns || [];
    
    if (selectableCols.length === 0) {
        return '<div class="column-checkbox-list"><p class="checkbox-grid-empty">No selectable columns found</p></div>';
    }
    
    let html = `<div class="column-checkbox-list" id="col-list-${sanitizeId(name)}">`;
    for (const col of selectableCols) {
        const checkId = `col-check-${sanitizeId(name)}-${sanitizeId(col)}`;
        html += `
            <div class="column-checkbox-item">
                <input type="checkbox" 
                    id="${checkId}" 
                    data-filename="${escapeHtml(name)}" 
                    data-column="${escapeHtml(col)}"
                    data-isstack="${isStack}"
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
function setupFileConfigListeners(name, info, isStack = false) {
    const panel = document.getElementById(`config-${sanitizeId(name)}`);
    if (!panel) return;

    // Column checkboxes
    panel.querySelectorAll('.column-checkbox').forEach(checkbox => {
        checkbox.addEventListener('change', () => {
            updateSelectedColumns(name, isStack);
        });
    });

    // Duplicate handling
    const dupeSelect = panel.querySelector('.dupe-handling');
    if (dupeSelect) {
        dupeSelect.addEventListener('change', (e) => {
            const target = isStack ? AppState.stacks[name] : AppState.files[name];
            if (target) target.dupeHandling = e.target.value;
        });
    }
    
    // Pivot button handler - works for both files and stacks
    const pivotBtn = panel.querySelector('.pivot-btn');
    if (pivotBtn) {
        pivotBtn.addEventListener('click', () => {
            const tagSelect = panel.querySelector('.pivot-tag-select');
            const valueSelect = panel.querySelector('.pivot-value-select');
            const tagCol = tagSelect?.value;
            const valueCol = valueSelect?.value;
            const isPivotStack = pivotBtn.dataset.isstack === 'true';
            
            applyPivot(name, tagCol, valueCol, isPivotStack);
        });
    }
}

/**
 * Update selected columns based on checkbox state
 */
function updateSelectedColumns(name, isStack = false) {
    const panel = document.getElementById(`config-${sanitizeId(name)}`);
    if (!panel) return;

    // Get the info object (from files or stacks)
    const info = isStack ? AppState.stacks[name] : AppState.files[name];
    if (!info) return;
    
    const checkedBoxes = panel.querySelectorAll('.column-checkbox:checked');
    
    const selectedColumns = Array.from(checkedBoxes).map(cb => cb.dataset.column);
    
    // Update selected columns in state
    const newSelectedCols = {};
    const newUnits = {};
    const newCleanup = {};

    for (const col of selectedColumns) {
        newSelectedCols[col] = info.selectedCols[col] || col;
        newUnits[col] = info.units[col] || '';
        newCleanup[col] = info.cleanup[col] || 'Fill with nearest available value';
    }

    info.selectedCols = newSelectedCols;
    info.units = newUnits;
    info.cleanup = newCleanup;

    // Update column settings UI
    updateColumnSettingsUI(name, selectedColumns, isStack);
    
    // Update graph column options
    updateGraphColumnOptions();
}

/**
 * Update column settings panel UI
 */

/**
 * Update column settings panel UI
 */
function updateColumnSettingsUI(name, selectedColumns, isStack = false) {
    const info = isStack ? AppState.stacks[name] : AppState.files[name];
    const settingsContainer = document.getElementById(`column-settings-${sanitizeId(name)}`);
    
    if (!settingsContainer || !info) return;

    // Build settings HTML
    let html = '';
    for (const col of selectedColumns) {
        html += `
            <div class="column-setting-item" id="col-setting-${sanitizeId(name)}-${sanitizeId(col)}">
                <div class="column-setting-header" onclick="this.parentElement.classList.toggle('open')">
                    <span class="column-setting-name">⚙️ Settings for ${escapeHtml(col)}</span>
                    <span>▼</span>
                </div>
                <div class="column-setting-content">
                    <div class="column-setting-grid">
                        <div class="form-group">
                            <label>Data Column Title</label>
                            <input type="text" class="input col-title" 
                                data-filename="${escapeHtml(name)}" 
                                data-column="${escapeHtml(col)}"
                                data-isstack="${isStack}"
                                value="${escapeHtml(info.selectedCols[col] || col)}">
                        </div>
                        <div class="form-group">
                            <label>Units</label>
                            <input type="text" class="input col-units"
                                data-filename="${escapeHtml(name)}"
                                data-column="${escapeHtml(col)}"
                                data-isstack="${isStack}"
                                value="${escapeHtml(info.units[col] || '')}">
                        </div>
                        <div class="form-group">
                            <label>Missing data handling</label>
                            <select class="select col-cleanup"
                                data-filename="${escapeHtml(name)}"
                                data-column="${escapeHtml(col)}"
                                data-isstack="${isStack}">
                                <option value="Fill with nearest available value" 
                                    ${info.cleanup[col] === 'Fill with nearest available value' ? 'selected' : ''}>
                                    Fill with nearest available value
                                </option>
                                <option value="Fill with a linear interpolation between the nearest values"
                                    ${info.cleanup[col] === 'Fill with a linear interpolation between the nearest values' ? 'selected' : ''}>
                                    Fill with linear interpolation
                                </option>
                                <option value="Delete the entire row of data"
                                    ${info.cleanup[col] === 'Delete the entire row of data' ? 'selected' : ''}>
                                    Delete the entire row
                                </option>
                                <option value="Fill with zero"
                                    ${info.cleanup[col] === 'Fill with zero' ? 'selected' : ''}>
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
            const isStackEl = e.target.dataset.isstack === 'true';
            const target = isStackEl ? AppState.stacks[fn] : AppState.files[fn];
            if (target) {
                target.selectedCols[col] = e.target.value;
                updateGraphColumnOptions();
            }
        });
    });

    settingsContainer.querySelectorAll('.col-units').forEach(input => {
        input.addEventListener('change', (e) => {
            const fn = e.target.dataset.filename;
            const col = e.target.dataset.column;
            const isStackEl = e.target.dataset.isstack === 'true';
            const target = isStackEl ? AppState.stacks[fn] : AppState.files[fn];
            if (target) {
                target.units[col] = e.target.value;
            }
        });
    });

    settingsContainer.querySelectorAll('.col-cleanup').forEach(select => {
        select.addEventListener('change', (e) => {
            const fn = e.target.dataset.filename;
            const col = e.target.dataset.column;
            const isStackEl = e.target.dataset.isstack === 'true';
            const target = isStackEl ? AppState.stacks[fn] : AppState.files[fn];
            if (target) {
                target.cleanup[col] = e.target.value;
            }
        });
    });
}

function updateSectionVisibility() {
    const hasFiles = Object.keys(AppState.files).length > 0;
    const hasStacks = Object.keys(AppState.stacks).length > 0;
    const hasDataSources = hasFiles || hasStacks;

    elements.columnsSection.classList.toggle('hidden', !hasDataSources);
    elements.graphSection.classList.toggle('hidden', !hasDataSources);
    elements.timeSection.classList.toggle('hidden', !hasDataSources);
    elements.downloadSection.classList.toggle('hidden', !hasDataSources);
}

// ===== ALIGNMENT OPTIONS =====

function updateAlignmentOptions() {
    elements.alignmentOptions.innerHTML = '';
    
    const stackedFiles = getStackedFiles();

    // Add stacks first
    for (const stackName of Object.keys(AppState.stacks)) {
        const div = document.createElement('div');
        div.className = 'form-group';
        div.innerHTML = `
            <label>📚 Stack '${escapeHtml(stackName)}':</label>
            <select class="select alignment-select" data-filename="${escapeHtml(stackName)}" data-isstack="true">
                <option value="Fill with the nearest value">Fill with the nearest value</option>
                <option value="Do a linear interpolation from the nearest values">Do a linear interpolation from the nearest values</option>
                <option value="Take an average of the available values within the interval">Take an average of the available values within the interval</option>
            </select>
        `;
        elements.alignmentOptions.appendChild(div);
    }

    // Add individual files (not in stacks)
    for (const fileName of Object.keys(AppState.files)) {
        if (stackedFiles.has(fileName)) continue;
        
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
    let earliest = null;
    let latest = null;

    // Check files
    for (const fileInfo of Object.values(AppState.files)) {
        if (fileInfo.dateRange?.earliest) {
            if (!earliest || fileInfo.dateRange.earliest < earliest) {
                earliest = fileInfo.dateRange.earliest;
            }
        }
        if (fileInfo.dateRange?.latest) {
            if (!latest || fileInfo.dateRange.latest > latest) {
                latest = fileInfo.dateRange.latest;
            }
        }
    }
    
    // Check stacks
    for (const stackInfo of Object.values(AppState.stacks)) {
        if (stackInfo.dateRange?.earliest) {
            if (!earliest || stackInfo.dateRange.earliest < earliest) {
                earliest = stackInfo.dateRange.earliest;
            }
        }
        if (stackInfo.dateRange?.latest) {
            if (!latest || stackInfo.dateRange.latest > latest) {
                latest = stackInfo.dateRange.latest;
            }
        }
    }

    if (earliest) {
        elements.startDate.value = formatDateForInput(earliest);
        elements.startTime.value = earliest.toTimeString().slice(0, 5);
    }
    
    if (latest) {
        elements.endDate.value = formatDateForInput(latest);
        elements.endTime.value = latest.toTimeString().slice(0, 5);
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
    const stackedFiles = getStackedFiles();

    // Add stacks first
    for (const [stackName, stackInfo] of Object.entries(AppState.stacks)) {
        for (const [col, title] of Object.entries(stackInfo.selectedCols || {})) {
            hasColumns = true;
            const checkId = `graph-check-${sanitizeId(stackName)}-${sanitizeId(col)}`;
            const div = document.createElement('div');
            div.className = 'checkbox-grid-item';
            div.innerHTML = `
                <input type="checkbox" 
                    id="${checkId}"
                    data-filename="${escapeHtml(stackName)}"
                    data-column="${escapeHtml(col)}"
                    data-isstack="true"
                    class="graph-column-checkbox">
                <label for="${checkId}">📚 ${escapeHtml(stackName)} - ${escapeHtml(title)}</label>
            `;
            elements.graphColumns.appendChild(div);
            
            div.querySelector('input').addEventListener('change', (e) => {
                div.classList.toggle('checked', e.target.checked);
            });
        }
    }

    // Add individual files (not in stacks)
    for (const [fileName, fileInfo] of Object.entries(AppState.files)) {
        if (stackedFiles.has(fileName)) continue;
        
        for (const [col, title] of Object.entries(fileInfo.selectedCols || {})) {
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
        const name = checkbox.dataset.filename;
        const col = checkbox.dataset.column;
        const isStack = checkbox.dataset.isstack === 'true';
        const info = isStack ? AppState.stacks[name] : AppState.files[name];
        
        if (!info) continue;

        const dateTimeCol = info.dateTimeCol;
        const data = info.data;

        if (!dateTimeCol) {
            console.warn(`No datetime column found for ${name}`);
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
            const displayName = (isStack ? '📚 ' : '') + (info.selectedCols[col] || col);
            traces.push({
                x: x,
                y: y,
                mode: 'lines',  // No markers for better performance
                name: displayName,
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
    // Validate we have data (check both files and stacks)
    let hasSelectedColumns = false;
    const stackedFiles = getStackedFiles();
    
    // Check stacks
    for (const stackInfo of Object.values(AppState.stacks)) {
        if (Object.keys(stackInfo.selectedCols || {}).length > 0) {
            hasSelectedColumns = true;
            break;
        }
    }
    
    // Check individual files (not in stacks)
    if (!hasSelectedColumns) {
        for (const [fileName, fileInfo] of Object.entries(AppState.files)) {
            if (stackedFiles.has(fileName)) continue;
            if (Object.keys(fileInfo.selectedCols || {}).length > 0) {
                hasSelectedColumns = true;
                break;
            }
        }
    }

    if (!hasSelectedColumns) {
        showStatus('Please select at least one column from at least one file or stack.', 'error');
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

        // Build combined data sources (stacks + unstacked files)
        const dataSources = {};
        
        // Add stacks
        for (const [stackName, stackInfo] of Object.entries(AppState.stacks)) {
            dataSources[stackName] = stackInfo;
        }
        
        // Add individual files (not in stacks)
        for (const [fileName, fileInfo] of Object.entries(AppState.files)) {
            if (!stackedFiles.has(fileName)) {
                dataSources[fileName] = fileInfo;
            }
        }

        // Create combined dataset
        const combined = DataProcessing.createCombinedDataset(
            dataSources, 
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

// ===== PIVOT FUNCTIONS =====

/**
 * Apply pivot transformation to convert long format to wide format
 * @param {string} name - File or stack name
 * @param {string} tagCol - User-selected tag column (or null for auto-detect)
 * @param {string} valueCol - User-selected value column (or null for auto-detect)
 * @param {boolean} isStack - Whether this is a stack
 */
function applyPivot(name, tagCol, valueCol, isStack = false) {
    const info = isStack ? AppState.stacks[name] : AppState.files[name];
    
    if (!info) {
        showStatus('Cannot pivot: data not found', 'error');
        return;
    }
    
    // Validate user selections
    if (!tagCol || !valueCol) {
        showStatus('Please select both a tag column and a value column', 'error');
        return;
    }
    
    showGlobalLoader('Pivoting data to wide format...');
    updateGlobalLoader(10, 'Preparing data...');
    
    try {
        let dataToProcess = [...info.data];
        let dateTimeCol = info.dateTimeCol;
        
        // Check if we need to combine Date + Time columns
        const separateDateTimeCols = FileHandlers.detectSeparateDateTimeColumns(info.columns, dataToProcess);
        
        if (separateDateTimeCols) {
            updateGlobalLoader(20, 'Combining Date and Time columns...');
            const combined = FileHandlers.combineDateTimeColumns(
                dataToProcess, 
                separateDateTimeCols.dateCol, 
                separateDateTimeCols.timeCol
            );
            dataToProcess = combined.data;
            dateTimeCol = combined.dateTimeCol;
            console.log(`Combined ${separateDateTimeCols.dateCol} and ${separateDateTimeCols.timeCol} into ${dateTimeCol}`);
        }
        
        updateGlobalLoader(40, 'Pivoting data...');
        
        // Perform the pivot
        const pivoted = FileHandlers.pivotToWideFormat(
            dataToProcess,
            dateTimeCol,
            tagCol,
            valueCol
        );
        
        updateGlobalLoader(70, 'Updating data structure...');
        
        // Update info with pivoted data
        info.data = pivoted.data;
        info.columns = pivoted.columns;
        info.isPivoted = true;
        info.originalTagCount = pivoted.tagCount;
        info.dateTimeCol = dateTimeCol;
        
        // Update dateTimeCols to include the new combined column if applicable
        if (separateDateTimeCols) {
            info.dateTimeCols = [dateTimeCol];
        }
        
        // Recalculate selectable columns
        info.selectableColumns = FileHandlers.getSelectableColumns(
            pivoted.columns, 
            info.dateTimeCols || [dateTimeCol]
        );
        
        // Recalculate date range
        info.dateRange = FileHandlers.getDateRange(pivoted.data, dateTimeCol);
        
        // Check for duplicates (should be none after pivot)
        info.hasDuplicates = FileHandlers.hasDuplicateTimestamps(pivoted.data, dateTimeCol);
        
        // Clear long format info since we've pivoted
        info.longFormatInfo = null;
        
        // Reset selected columns since the column structure changed
        info.selectedCols = {};
        info.units = {};
        info.cleanup = {};
        
        updateGlobalLoader(100, '✅ Data pivoted successfully!');
        
        // Refresh UI
        updateFileList();
        updateFileConfigs();
        updateAlignmentOptions();
        updateDefaultDatesFromData();
        updateGraphColumnOptions();
        updateStackingSection();
        
        setTimeout(hideGlobalLoader, 500);
        
        const typeLabel = isStack ? 'stack' : 'file';
        showStatus(`✅ Pivoted ${typeLabel} "${name}" to wide format: ${pivoted.data.length} rows × ${info.selectableColumns.length} columns`, 'success');
        
    } catch (error) {
        console.error('Error pivoting data:', error);
        showStatus(`Error pivoting data: ${error.message}`, 'error');
        hideGlobalLoader();
    }
}

// ===== STACKING FUNCTIONS =====

/**
 * Get set of files that are in any stack
 */
function getStackedFiles() {
    const stackedFiles = new Set();
    for (const stack of Object.values(AppState.stacks)) {
        for (const fileName of stack.files) {
            stackedFiles.add(fileName);
        }
    }
    return stackedFiles;
}

/**
 * Get the stack name that contains a file
 */
function getStackNameForFile(fileName) {
    for (const [stackName, stack] of Object.entries(AppState.stacks)) {
        if (stack.files.includes(fileName)) {
            return stackName;
        }
    }
    return null;
}

/**
 * Remove a file from all stacks (called when file is deleted)
 */
function removeFileFromAllStacks(fileName) {
    for (const [stackName, stack] of Object.entries(AppState.stacks)) {
        const idx = stack.files.indexOf(fileName);
        if (idx > -1) {
            stack.files.splice(idx, 1);
            // If stack now has less than 2 files, remove it
            if (stack.files.length < 2) {
                delete AppState.stacks[stackName];
            } else {
                // Rebuild stack data
                rebuildStackData(stackName);
            }
        }
    }
}

/**
 * Rebuild stack data after files change
 */
function rebuildStackData(stackName) {
    const stack = AppState.stacks[stackName];
    if (!stack || stack.files.length < 2) return;
    
    const firstFile = AppState.files[stack.files[0]];
    if (!firstFile) return;
    
    // Combine data from all files
    let combinedData = [];
    const dtCol = firstFile.dateTimeCol;
    
    for (const fn of stack.files) {
        const fileData = AppState.files[fn]?.data || [];
        combinedData = combinedData.concat(fileData);
    }
    
    // Sort by datetime
    combinedData.sort((a, b) => {
        const dtA = new Date(a[dtCol]);
        const dtB = new Date(b[dtCol]);
        return dtA - dtB;
    });
    
    // Update stack
    stack.data = combinedData;
    stack.dateRange = FileHandlers.getDateRange(combinedData, dtCol);
    stack.rowCount = combinedData.length;
}

/**
 * Update the stacking section UI
 */
function updateStackingSection() {
    const fileCount = Object.keys(AppState.files).length;
    
    // Only show stacking section if 2+ files
    if (fileCount >= 2) {
        elements.stackSection.classList.remove('hidden');
    } else {
        elements.stackSection.classList.add('hidden');
        return;
    }
    
    // Render existing stacks
    renderExistingStacks();
    
    // Render stackable files for new stack creation
    renderStackableFiles();
    
    // Update button state
    updateCreateStackButton();
}

/**
 * Render existing stacks
 */
function renderExistingStacks() {
    const container = elements.existingStacks;
    container.innerHTML = '';
    
    for (const [stackName, stack] of Object.entries(AppState.stacks)) {
        const div = document.createElement('div');
        div.className = 'existing-stack';
        
        const dateRangeStr = formatDateRange(stack.dateRange);
        
        div.innerHTML = `
            <div class="existing-stack-header">
                <span class="existing-stack-name"><span class="icon">📊</span> ${escapeHtml(stackName)}</span>
                <button class="btn btn-danger btn-sm" data-stack="${escapeHtml(stackName)}">Remove Stack</button>
            </div>
            <div class="existing-stack-files">
                ${stack.files.map(f => `<span class="stack-file-tag">${escapeHtml(f)}</span>`).join('')}
            </div>
            <div class="existing-stack-range">${dateRangeStr}</div>
            <div class="existing-stack-stats">${stack.rowCount.toLocaleString()} total rows • ${stack.selectableColumns.length} data columns</div>
        `;
        
        container.appendChild(div);
        
        // Add remove handler
        div.querySelector('.btn-danger').addEventListener('click', (e) => {
            const name = e.target.dataset.stack;
            delete AppState.stacks[name];
            updateFileList();
            updateFileConfigs();
            updateStackingSection();
            updateAlignmentOptions();
            updateGraphColumnOptions();
            updateDefaultDatesFromData();
        });
    }
}

/**
 * Render stackable files for new stack creation
 */
function renderStackableFiles() {
    const container = elements.stackableFiles;
    container.innerHTML = '';
    
    const stackedFiles = getStackedFiles();
    const availableFiles = Object.keys(AppState.files).filter(fn => !stackedFiles.has(fn));
    
    if (availableFiles.length < 2) {
        container.innerHTML = '<p style="color: var(--text-secondary)">Need at least 2 unstacked files to create a new stack.</p>';
        elements.createStackBtn.disabled = true;
        return;
    }
    
    // Sort by date range
    availableFiles.sort((a, b) => {
        const dateA = AppState.files[a].dateRange?.earliest;
        const dateB = AppState.files[b].dateRange?.earliest;
        if (!dateA && !dateB) return 0;
        if (!dateA) return 1;
        if (!dateB) return -1;
        return dateA - dateB;
    });
    
    // Determine which files are compatible with currently selected files
    const selectedFiles = [...AppState.selectedForStack];
    let referenceColumns = null;
    if (selectedFiles.length > 0) {
        referenceColumns = new Set(AppState.files[selectedFiles[0]].selectableColumns);
    }
    
    for (const fileName of availableFiles) {
        const fileInfo = AppState.files[fileName];
        const isSelected = AppState.selectedForStack.has(fileName);
        
        // Check compatibility with already selected files
        let isCompatible = true;
        if (referenceColumns && !isSelected) {
            const fileCols = new Set(fileInfo.selectableColumns);
            // Compatible if there's significant column overlap
            const overlap = [...referenceColumns].filter(c => fileCols.has(c));
            isCompatible = overlap.length >= Math.min(referenceColumns.size, fileCols.size) * 0.5;
        }
        
        const div = document.createElement('div');
        div.className = 'stackable-file' + (isSelected ? ' selected' : '') + (!isCompatible ? ' incompatible' : '');
        div.dataset.filename = fileName;
        
        const dateRange = formatDateRangeShort(fileInfo.dateRange);
        const selectionOrder = selectedFiles.indexOf(fileName) + 1;
        
        div.innerHTML = `
            ${isSelected ? `<span class="stackable-file-order">${selectionOrder}</span>` : ''}
            <div class="stackable-file-name">${escapeHtml(fileName)}</div>
            <div class="stackable-file-info">${fileInfo.data.length.toLocaleString()} rows • ${fileInfo.selectableColumns.length} columns</div>
            <div class="stackable-file-dates">${dateRange}</div>
        `;
        
        if (isCompatible) {
            div.addEventListener('click', () => toggleFileForStack(fileName));
        }
        
        container.appendChild(div);
    }
    
    // Check for warnings
    updateStackWarnings();
}

/**
 * Toggle file selection for stacking
 */
function toggleFileForStack(fileName) {
    if (AppState.selectedForStack.has(fileName)) {
        AppState.selectedForStack.delete(fileName);
    } else {
        AppState.selectedForStack.add(fileName);
    }
    
    renderStackableFiles();
    updateCreateStackButton();
}

/**
 * Update stack warnings and info
 */
function updateStackWarnings() {
    const warnings = [];
    const infos = [];
    
    const selectedFiles = [...AppState.selectedForStack];
    
    if (selectedFiles.length >= 2) {
        // Check column compatibility
        const columnSets = selectedFiles.map(fn => new Set(AppState.files[fn].selectableColumns));
        let commonColumns = [...columnSets[0]];
        for (let i = 1; i < columnSets.length; i++) {
            commonColumns = commonColumns.filter(c => columnSets[i].has(c));
        }
        
        const allColumns = new Set();
        columnSets.forEach(s => s.forEach(c => allColumns.add(c)));
        const missingInSome = [...allColumns].filter(c => !commonColumns.includes(c));
        
        if (missingInSome.length > 0) {
            warnings.push(`Column mismatch: ${missingInSome.length} column(s) not present in all files will be excluded from the stack.`);
        }
        
        // Check for time gaps or overlaps
        const ranges = selectedFiles
            .map(fn => ({ name: fn, range: AppState.files[fn].dateRange }))
            .filter(r => r.range?.earliest && r.range?.latest)
            .sort((a, b) => a.range.earliest - b.range.earliest);
        
        for (let i = 0; i < ranges.length - 1; i++) {
            const curr = ranges[i];
            const next = ranges[i + 1];
            
            // Check overlap
            if (curr.range.latest > next.range.earliest) {
                warnings.push(`Time overlap detected between "${curr.name}" and "${next.name}". Duplicate handling will be applied.`);
            }
            
            // Check gap (more than 1 hour)
            const gapMs = next.range.earliest - curr.range.latest;
            if (gapMs > 3600000) {
                const gapHours = Math.round(gapMs / 3600000);
                infos.push(`${gapHours} hour gap between "${curr.name}" and "${next.name}".`);
            }
        }
        
        // Show combined stats
        let totalRows = 0;
        let minDate = null, maxDate = null;
        for (const fn of selectedFiles) {
            totalRows += AppState.files[fn].data.length;
            const range = AppState.files[fn].dateRange;
            if (range?.earliest) {
                if (!minDate || range.earliest < minDate) minDate = range.earliest;
            }
            if (range?.latest) {
                if (!maxDate || range.latest > maxDate) maxDate = range.latest;
            }
        }
        
        infos.push(`Stack will contain ${totalRows.toLocaleString()} rows and ${commonColumns.length} columns spanning ${formatDateRange({ earliest: minDate, latest: maxDate })}`);
    }
    
    // Render warnings
    elements.stackWarnings.innerHTML = warnings.map(w => `
        <div class="stack-warning">
            <span class="stack-warning-icon">⚠️</span>
            <span class="stack-warning-text">${escapeHtml(w)}</span>
        </div>
    `).join('');
    
    // Render info
    elements.stackInfo.innerHTML = infos.map(i => `
        <div class="stack-info">ℹ️ ${escapeHtml(i)}</div>
    `).join('');
}

/**
 * Update create stack button state
 */
function updateCreateStackButton() {
    const hasEnoughFiles = AppState.selectedForStack.size >= 2;
    const hasName = elements.stackName.value.trim().length > 0;
    
    elements.createStackBtn.disabled = !(hasEnoughFiles && hasName);
}

/**
 * Create a new stack from selected files
 */
function createStack() {
    const stackName = elements.stackName.value.trim();
    const selectedFiles = [...AppState.selectedForStack];
    const overlapHandling = elements.overlapHandling.value;
    
    if (selectedFiles.length < 2 || !stackName) {
        showStatus('Please select at least 2 files and enter a stack name.', 'error');
        return;
    }
    
    // Check if stack name already exists
    if (AppState.stacks[stackName]) {
        showStatus('A stack with this name already exists. Please choose a different name.', 'error');
        return;
    }
    
    // Sort files by date
    selectedFiles.sort((a, b) => {
        const dateA = AppState.files[a].dateRange?.earliest;
        const dateB = AppState.files[b].dateRange?.earliest;
        if (!dateA && !dateB) return 0;
        if (!dateA) return 1;
        if (!dateB) return -1;
        return dateA - dateB;
    });
    
    // Get common columns
    const columnSets = selectedFiles.map(fn => new Set(AppState.files[fn].selectableColumns));
    let commonColumns = [...columnSets[0]];
    for (let i = 1; i < columnSets.length; i++) {
        commonColumns = commonColumns.filter(c => columnSets[i].has(c));
    }
    
    // Get first file for reference
    const firstFile = AppState.files[selectedFiles[0]];
    const dtCol = firstFile.dateTimeCol;
    
    // Combine data
    let combinedData = [];
    for (const fn of selectedFiles) {
        const fileData = AppState.files[fn].data;
        combinedData = combinedData.concat(fileData);
    }
    
    // Sort by datetime
    combinedData.sort((a, b) => {
        const dtA = new Date(a[dtCol]);
        const dtB = new Date(b[dtCol]);
        return dtA - dtB;
    });
    
    // Handle duplicates based on overlap handling setting
    if (overlapHandling !== 'keep_all') {
        combinedData = DataProcessing.handleDuplicates(combinedData, dtCol, 
            overlapHandling === 'average' ? 'Average values' : 
            overlapHandling === 'first' ? 'Keep first' : 'Keep last');
    }
    
    // Get combined date range
    const dateRange = FileHandlers.getDateRange(combinedData, dtCol);
    
    // Get columns from the actual data keys to ensure they match
    // This handles any case where column names might differ
    const dataKeys = combinedData.length > 0 ? Object.keys(combinedData[0]) : firstFile.columns;
    
    console.log('Creating stack with:', {
        firstFileColumns: firstFile.columns,
        dataKeys: dataKeys,
        commonColumns: commonColumns,
        dtCol: dtCol
    });
    
    // Create stack
    AppState.stacks[stackName] = {
        files: selectedFiles,
        overlapHandling: overlapHandling,
        data: combinedData,
        columns: dataKeys,  // Use actual data keys to ensure match
        selectableColumns: commonColumns,
        dateTimeCol: dtCol,
        dateTimeCols: firstFile.dateTimeCols,
        dateRange: dateRange,
        rowCount: combinedData.length,
        selectedCols: {},
        units: {},
        cleanup: {},
        dupeHandling: 'Average values',
        hasDuplicates: false  // Already handled
    };
    
    // Clear selection
    AppState.selectedForStack.clear();
    elements.stackName.value = '';
    
    // Update UI
    updateFileList();
    updateFileConfigs();
    updateStackingSection();
    updateSectionVisibility();
    updateAlignmentOptions();
    updateDefaultDatesFromData();
    updateGraphColumnOptions();
    
    showStatus(`✅ Created stack "${stackName}" with ${selectedFiles.length} files!`, 'success');
}

/**
 * Format date range for display
 */
function formatDateRange(range) {
    if (!range?.earliest || !range?.latest) return 'Unknown date range';
    
    const opts = { year: 'numeric', month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit' };
    return `${range.earliest.toLocaleDateString('en-US', opts)} → ${range.latest.toLocaleDateString('en-US', opts)}`;
}

/**
 * Format date range short version
 */
function formatDateRangeShort(range) {
    if (!range?.earliest || !range?.latest) return 'Unknown dates';
    
    const opts = { month: 'short', day: 'numeric' };
    const startStr = range.earliest.toLocaleDateString('en-US', opts);
    const endStr = range.latest.toLocaleDateString('en-US', opts);
    
    if (startStr === endStr) {
        return startStr;
    }
    return `${startStr} - ${endStr}`;
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
