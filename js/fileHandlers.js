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
        
        // Pre-process: remove trailing commas from each line to avoid extra columns
        const cleanedLines = lines.slice(headerRow).map(line => {
            // Remove trailing commas (and any whitespace after them)
            return line.replace(/,\s*$/, '').replace(/\r$/, '');
        });
        const csvText = cleanedLines.join('\n');
        
        const result = Papa.parse(csvText, {
            header: true,
            skipEmptyLines: true,
            dynamicTyping: false,  // Keep as strings to preserve precision
            transformHeader: (header) => {
                // Clean header names - remove leading special characters and trim
                return header.replace(/^[;:,\s]+/, '').trim();
            }
        });
        
        // Clean up the data - trim whitespace from all values and remove _parsed_extra
        let columns = (result.meta.fields || []).filter(col => col && col !== '_parsed_extra');
        let data = result.data.map(row => {
            const newRow = {};
            for (const col of columns) {
                let value = row[col];
                // Trim string values
                if (typeof value === 'string') {
                    value = value.trim();
                    // Convert numeric strings to numbers
                    if (value !== '' && !isNaN(value)) {
                        const num = parseFloat(value);
                        if (!isNaN(num)) {
                            value = num;
                        }
                    }
                }
                newRow[col] = value;
            }
            return newRow;
        });
        
        // Filter out empty rows
        data = data.filter(row => {
            return Object.values(row).some(v => v !== null && v !== undefined && v !== '');
        });
        
        console.log('CSV parsed:', { columns, rowCount: data.length, sampleRow: data[0] });
        
        return {
            data: data,
            columns: columns,
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
    },

    /**
     * Detect if data is in "long format" (needs pivoting to wide format)
     * Long format has: timestamp, tagname/variable column, value column
     * Same timestamp appears multiple times with different tag values
     * @param {Array} data - Data rows
     * @param {Array} columns - Column names
     * @param {string} dateTimeCol - Detected datetime column
     * @returns {Object|null} - Detection result with tagCol and valueCol, or null if not long format
     */
    detectLongFormat(data, columns, dateTimeCol) {
        if (!dateTimeCol || data.length < 10) return null;
        
        // Common patterns for tag/variable name columns
        const tagPatterns = [
            /^tagname$/i,
            /^tag$/i,
            /^variable$/i,
            /^var$/i,
            /^name$/i,
            /^sensor$/i,
            /^channel$/i,
            /^metric$/i,
            /^parameter$/i,
            /^point$/i,
            /^signal$/i
        ];
        
        // Common patterns for value columns
        const valuePatterns = [
            /^value$/i,
            /^val$/i,
            /^data$/i,
            /^reading$/i,
            /^measurement$/i,
            /^result$/i
        ];
        
        // Find potential tag column
        let tagCol = null;
        for (const col of columns) {
            if (col === dateTimeCol) continue;
            const colTrimmed = col.trim();
            for (const pattern of tagPatterns) {
                if (pattern.test(colTrimmed)) {
                    tagCol = col;
                    break;
                }
            }
            if (tagCol) break;
        }
        
        // Find potential value column
        let valueCol = null;
        for (const col of columns) {
            if (col === dateTimeCol || col === tagCol) continue;
            const colTrimmed = col.trim();
            for (const pattern of valuePatterns) {
                if (pattern.test(colTrimmed)) {
                    valueCol = col;
                    break;
                }
            }
            if (valueCol) break;
        }
        
        // If we didn't find by name, try to detect by data characteristics
        if (!tagCol || !valueCol) {
            // Look for a column with repeating string values (tag names)
            // and a column with numeric values
            for (const col of columns) {
                if (col === dateTimeCol) continue;
                
                const sampleValues = data.slice(0, 100).map(row => row[col]).filter(v => v != null);
                if (sampleValues.length === 0) continue;
                
                const uniqueValues = new Set(sampleValues.map(v => String(v).trim()));
                const numericCount = sampleValues.filter(v => !isNaN(parseFloat(v))).length;
                const numericRatio = numericCount / sampleValues.length;
                
                // Tag column: mostly strings, limited unique values relative to total rows
                if (!tagCol && numericRatio < 0.3 && uniqueValues.size < sampleValues.length * 0.5 && uniqueValues.size > 1) {
                    tagCol = col;
                }
                
                // Value column: mostly numeric
                if (!valueCol && numericRatio > 0.8) {
                    valueCol = col;
                }
            }
        }
        
        if (!tagCol || !valueCol) return null;
        
        // Verify it's actually long format by checking for duplicate timestamps
        const timestampCounts = new Map();
        for (const row of data.slice(0, 500)) {
            const ts = String(row[dateTimeCol]);
            timestampCounts.set(ts, (timestampCounts.get(ts) || 0) + 1);
        }
        
        // If most timestamps appear multiple times, it's likely long format
        const multipleOccurrences = [...timestampCounts.values()].filter(c => c > 1).length;
        const isLongFormat = multipleOccurrences / timestampCounts.size > 0.5;
        
        if (!isLongFormat) return null;
        
        // Get unique tag values
        const uniqueTags = [...new Set(data.map(row => String(row[tagCol]).trim()))].filter(t => t);
        
        return {
            tagCol,
            valueCol,
            uniqueTags,
            tagCount: uniqueTags.length
        };
    },

    /**
     * Pivot long format data to wide format
     * @param {Array} data - Long format data
     * @param {string} dateTimeCol - DateTime column name
     * @param {string} tagCol - Tag/variable name column
     * @param {string} valueCol - Value column
     * @returns {Object} - { data: pivoted data, columns: new column names }
     */
    pivotToWideFormat(data, dateTimeCol, tagCol, valueCol) {
        // Get unique tags (these become column names)
        const uniqueTags = [...new Set(data.map(row => String(row[tagCol]).trim()))].filter(t => t);
        
        // Group by timestamp
        const grouped = new Map();
        
        for (const row of data) {
            const ts = row[dateTimeCol];
            const tag = String(row[tagCol]).trim();
            const value = row[valueCol];
            
            // Create a string key for the timestamp
            const tsKey = String(ts);
            
            if (!grouped.has(tsKey)) {
                grouped.set(tsKey, { [dateTimeCol]: ts });
            }
            
            grouped.get(tsKey)[tag] = value;
        }
        
        // Convert to array
        const pivotedData = [...grouped.values()];
        
        // Sort by timestamp
        pivotedData.sort((a, b) => {
            const dateA = new Date(a[dateTimeCol]);
            const dateB = new Date(b[dateTimeCol]);
            return dateA - dateB;
        });
        
        // New columns: datetime + all unique tags
        const newColumns = [dateTimeCol, ...uniqueTags];
        
        return {
            data: pivotedData,
            columns: newColumns,
            tagCount: uniqueTags.length
        };
    },

    /**
     * Detect separate Date and Time columns that should be combined
     * @param {Array} columns - Column names
     * @param {Array} data - Data rows
     * @returns {Object|null} - { dateCol, timeCol } or null if not found
     */
    detectSeparateDateTimeColumns(columns, data) {
        let dateCol = null;
        let timeCol = null;
        
        for (const col of columns) {
            const colLower = col.toLowerCase().trim();
            
            // Check for date column (but not datetime)
            if (!dateCol && colLower === 'date') {
                dateCol = col;
            }
            
            // Check for time column
            if (!timeCol && colLower === 'time') {
                timeCol = col;
            }
        }
        
        // If we found both, verify they contain appropriate data
        if (dateCol && timeCol && data.length > 0) {
            const sampleDate = data[0][dateCol];
            const sampleTime = data[0][timeCol];
            
            // Basic validation - date should have / or - and time should have :
            if (sampleDate && sampleTime) {
                const dateStr = String(sampleDate);
                const timeStr = String(sampleTime);
                
                if ((dateStr.includes('/') || dateStr.includes('-')) && timeStr.includes(':')) {
                    return { dateCol, timeCol };
                }
            }
        }
        
        return null;
    },

    /**
     * Combine separate Date and Time columns into a single DateTime column
     * @param {Array} data - Data rows
     * @param {string} dateCol - Date column name
     * @param {string} timeCol - Time column name
     * @returns {Object} - { data: modified data, dateTimeCol: new column name }
     */
    combineDateTimeColumns(data, dateCol, timeCol) {
        const dateTimeCol = 'DateTime';
        
        const newData = data.map(row => {
            const newRow = { ...row };
            
            const dateVal = row[dateCol];
            const timeVal = row[timeCol];
            
            if (dateVal && timeVal) {
                // Combine date and time into a single value
                // Handle various date formats
                let dateStr = String(dateVal);
                let timeStr = String(timeVal);
                
                // Try to create a proper datetime
                const combined = `${dateStr} ${timeStr}`;
                const parsed = new Date(combined);
                
                if (!isNaN(parsed.getTime())) {
                    newRow[dateTimeCol] = parsed;
                } else {
                    // If parsing fails, store as string
                    newRow[dateTimeCol] = combined;
                }
            }
            
            return newRow;
        });
        
        return {
            data: newData,
            dateTimeCol: dateTimeCol
        };
    },

    /**
     * Get columns suitable for tag/variable selection (string-like columns)
     * @param {Array} columns - All columns
     * @param {Array} excludeCols - Columns to exclude
     * @param {Array} data - Sample data
     * @returns {Array} - Columns that could be tag columns
     */
    getPotentialTagColumns(columns, excludeCols, data) {
        return columns.filter(col => {
            if (excludeCols.includes(col)) return false;
            
            // Sample the column to see if it's string-like
            const samples = data.slice(0, 50).map(row => row[col]).filter(v => v != null);
            if (samples.length === 0) return true; // Include if no data to check
            
            // Check if mostly non-numeric (strings)
            const numericCount = samples.filter(v => {
                const num = parseFloat(v);
                return !isNaN(num) && String(num) === String(v).trim();
            }).length;
            
            return numericCount / samples.length < 0.5;
        });
    },

    /**
     * Get columns suitable for value selection (numeric columns)
     * @param {Array} columns - All columns
     * @param {Array} excludeCols - Columns to exclude
     * @param {Array} data - Sample data
     * @returns {Array} - Columns that could be value columns (sorted by likelihood)
     */
    getPotentialValueColumns(columns, excludeCols, data) {
        const results = [];
        
        for (const col of columns) {
            if (excludeCols.includes(col)) continue;
            
            // Sample the column to see if it's numeric
            const samples = data.slice(0, 50).map(row => row[col]).filter(v => v != null && v !== '');
            
            if (samples.length === 0) {
                results.push({ col, score: 0.5 }); // Unknown, include with medium score
                continue;
            }
            
            // Check if mostly numeric - handle whitespace-padded values
            let numericCount = 0;
            for (const v of samples) {
                const strVal = String(v).trim();
                const parsed = parseFloat(strVal);
                if (!isNaN(parsed)) {
                    numericCount++;
                }
            }
            
            const numericRatio = numericCount / samples.length;
            
            // Include columns that are at least partially numeric
            if (numericRatio > 0.3) {
                results.push({ col, score: numericRatio });
            }
        }
        
        // Sort by score (most numeric first) and return just column names
        results.sort((a, b) => b.score - a.score);
        return results.map(r => r.col);
    }
};

// Export for use in other modules
window.FileHandlers = FileHandlers;
