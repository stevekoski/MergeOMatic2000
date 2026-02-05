/**
 * File Handlers Module
 * Handles reading and parsing CSV and Excel files
 */

const FileHandlers = {
    /**
     * Read a file and return its contents based on type
     * @param {File} file - The file to read
     * @returns {Promise<{data: Array, columns: Array, headerRow: number}>}
     */
    async readFile(file) {
        const fileName = file.name.toLowerCase();
        const arrayBuffer = await file.arrayBuffer();
        
        if (fileName.endsWith('.csv')) {
            return this.readCSV(arrayBuffer);
        } else if (fileName.endsWith('.xls') || fileName.endsWith('.xlsx')) {
            return this.readExcel(arrayBuffer);
        } else {
            throw new Error(`Unsupported file type: ${file.name}`);
        }
    },

    /**
     * Read CSV file with flexible header detection
     * @param {ArrayBuffer} arrayBuffer - File contents
     * @returns {{data: Array, columns: Array, headerRow: number}}
     */
    readCSV(arrayBuffer) {
        const text = new TextDecoder('utf-8').decode(arrayBuffer);
        const lines = text.split(/\r?\n/);
        const maxCheckLines = 30;
        
        // Find the header row
        let headerRow = 0;
        let detected = false;
        
        for (let i = 0; i < Math.min(maxCheckLines, lines.length); i++) {
            const line = lines[i];
            if (!line.trim()) continue;
            
            // Split by comma or tab
            const parts = line.split(/[,\t]/)
                .map(p => p.trim().replace(/^["']|["']$/g, ''))
                .filter(p => p);
            
            if (parts.length < 2) continue;
            
            // Heuristic: if at least half the entries are non-numeric, assume header
            let nonNumeric = 0;
            for (const p of parts) {
                if (isNaN(parseFloat(p))) {
                    nonNumeric++;
                }
            }
            
            if (nonNumeric / parts.length >= 0.5) {
                headerRow = i;
                detected = true;
                break;
            }
        }
        
        // Parse with Papa Parse, skipping to detected header
        const csvText = lines.slice(headerRow).join('\n');
        const result = Papa.parse(csvText, {
            header: true,
            skipEmptyLines: true,
            dynamicTyping: true
        });
        
        return {
            data: result.data,
            columns: result.meta.fields || [],
            headerRow: headerRow
        };
    },

    /**
     * Read Excel file
     * @param {ArrayBuffer} arrayBuffer - File contents
     * @returns {{data: Array, columns: Array, headerRow: number}}
     */
    readExcel(arrayBuffer) {
        const workbook = XLSX.read(arrayBuffer, { type: 'array', cellDates: true });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        // Convert to JSON with header detection
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { 
            header: 1,
            defval: null,
            raw: false,
            dateNF: 'yyyy-mm-dd hh:mm:ss'
        });
        
        if (jsonData.length === 0) {
            return { data: [], columns: [], headerRow: 0 };
        }
        
        // Find header row using similar heuristics as CSV
        let headerRow = 0;
        const maxCheckLines = Math.min(30, jsonData.length);
        
        for (let i = 0; i < maxCheckLines; i++) {
            const row = jsonData[i];
            if (!row || !Array.isArray(row)) continue;
            
            const parts = row.filter(p => p !== null && p !== undefined && String(p).trim());
            if (parts.length < 2) continue;
            
            let nonNumeric = 0;
            for (const p of parts) {
                if (isNaN(parseFloat(p))) {
                    nonNumeric++;
                }
            }
            
            if (nonNumeric / parts.length >= 0.5) {
                headerRow = i;
                break;
            }
        }
        
        // Get columns from header row
        const columns = jsonData[headerRow]
            .map((col, idx) => col ? String(col).trim() : `Column ${idx + 1}`)
            .filter(col => col);
        
        // Convert remaining rows to objects
        const data = [];
        for (let i = headerRow + 1; i < jsonData.length; i++) {
            const row = jsonData[i];
            if (!row || row.every(cell => cell === null || cell === undefined)) continue;
            
            const obj = {};
            columns.forEach((col, idx) => {
                obj[col] = row[idx] !== undefined ? row[idx] : null;
            });
            data.push(obj);
        }
        
        return {
            data: data,
            columns: columns,
            headerRow: headerRow
        };
    },

    /**
     * Get preview data (first N rows)
     * @param {Array} data - Full dataset
     * @param {number} rows - Number of rows to preview
     * @returns {Array}
     */
    getPreview(data, rows = 2) {
        return data.slice(0, rows);
    },

    /**
     * Format file size for display
     * @param {number} bytes - Size in bytes
     * @returns {string}
     */
    formatFileSize(bytes) {
        if (bytes < 1024) return bytes + ' B';
        if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
        return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
    },

    /**
     * Detect datetime columns in the data
     * @param {Array} columns - Column names
     * @param {Array} data - Data rows
     * @returns {Array} - Column names that appear to be datetime
     */
    detectDateTimeColumns(columns, data) {
        const dateTimeCols = [];
        
        for (const col of columns) {
            const colLower = col.toLowerCase();
            
            // Check column name
            if (colLower.includes('date') || colLower.includes('time')) {
                dateTimeCols.push(col);
                continue;
            }
            
            // Check first few values
            const sampleValues = data.slice(0, 5).map(row => row[col]).filter(v => v != null);
            if (sampleValues.length === 0) continue;
            
            // Try to parse as date
            let dateCount = 0;
            for (const val of sampleValues) {
                const parsed = new Date(val);
                if (!isNaN(parsed.getTime()) && String(val).length > 5) {
                    dateCount++;
                }
            }
            
            if (dateCount >= sampleValues.length * 0.5) {
                dateTimeCols.push(col);
            }
        }
        
        return dateTimeCols;
    },

    /**
     * Check for duplicate timestamps in data
     * @param {Array} data - Data rows
     * @param {string} dateTimeCol - Name of datetime column
     * @returns {boolean}
     */
    hasDuplicateTimestamps(data, dateTimeCol) {
        if (!dateTimeCol || data.length < 2) return false;
        
        const seen = new Set();
        for (const row of data) {
            const val = row[dateTimeCol];
            if (val != null) {
                const key = String(val);
                if (seen.has(key)) return true;
                seen.add(key);
            }
        }
        return false;
    },

    /**
     * Get earliest and latest timestamps from data
     * @param {Array} data - Data rows
     * @param {string} dateTimeCol - Name of datetime column
     * @returns {{earliest: Date|null, latest: Date|null}}
     */
    getDateRange(data, dateTimeCol) {
        if (!dateTimeCol || data.length === 0) {
            return { earliest: null, latest: null };
        }
        
        let earliest = null;
        let latest = null;
        
        for (const row of data) {
            const val = row[dateTimeCol];
            if (val == null) continue;
            
            const date = new Date(val);
            if (isNaN(date.getTime())) continue;
            
            if (earliest === null || date < earliest) {
                earliest = date;
            }
            if (latest === null || date > latest) {
                latest = date;
            }
        }
        
        return { earliest, latest };
    },

    /**
     * Get columns that should be available for selection
     * Excludes datetime columns, index-like columns, and unnamed columns
     * @param {Array} columns - All column names
     * @param {Array} dateTimeCols - Detected datetime columns
     * @returns {Array} - Selectable column names
     */
    getSelectableColumns(columns, dateTimeCols) {
        const excludePatterns = [
            /^index$/i,
            /^idx$/i,
            /^id$/i,
            /^row$/i,
            /^row_num$/i,
            /^row_number$/i,
            /^unnamed:\s*\d+$/i,
            /^unnamed$/i,
            /^column\s*\d+$/i
        ];
        
        return columns.filter(col => {
            const colLower = col.toLowerCase().trim();
            
            // Exclude datetime columns
            if (dateTimeCols.includes(col)) {
                return false;
            }
            
            // Exclude columns matching exclusion patterns
            for (const pattern of excludePatterns) {
                if (pattern.test(colLower)) {
                    return false;
                }
            }
            
            // Exclude columns with date/time in the name
            if (colLower.includes('date') || colLower.includes('time')) {
                return false;
            }
            
            return true;
        });
    }
};

// Export for use in other modules
window.FileHandlers = FileHandlers;
