/**
 * Data Processing Module
 * Handles data cleaning, alignment, and merging
 */

const DataProcessing = {
    /**
     * Parse interval string to milliseconds
     * @param {string} interval - Interval string like "1min", "5min", "1h"
     * @returns {number} - Milliseconds
     */
    parseInterval(interval) {
        const match = interval.match(/^(\d+)(s|min|h|D)$/);
        if (!match) return 60000; // default 1 minute
        
        const value = parseInt(match[1]);
        const unit = match[2];
        
        switch (unit) {
            case 's': return value * 1000;
            case 'min': return value * 60 * 1000;
            case 'h': return value * 60 * 60 * 1000;
            case 'D': return value * 24 * 60 * 60 * 1000;
            default: return 60000;
        }
    },

    /**
     * Generate array of timestamps between start and end at given interval
     * @param {Date} start - Start datetime
     * @param {Date} end - End datetime
     * @param {string} interval - Interval string
     * @returns {Array<Date>}
     */
    generateTimeIndex(start, end, interval) {
        const intervalMs = this.parseInterval(interval);
        const timestamps = [];
        
        let current = new Date(start);
        while (current <= end) {
            timestamps.push(new Date(current));
            current = new Date(current.getTime() + intervalMs);
        }
        
        return timestamps;
    },

    /**
     * Clean data by handling missing values
     * @param {Array} data - Data rows
     * @param {string} column - Column name
     * @param {string} method - Cleanup method
     * @param {string} dateTimeCol - DateTime column name for sorting
     * @returns {Array} - Cleaned data
     */
    applyCleanup(data, column, method, dateTimeCol) {
        // Sort by datetime first
        const sorted = [...data].sort((a, b) => {
            const dateA = new Date(a[dateTimeCol]);
            const dateB = new Date(b[dateTimeCol]);
            return dateA - dateB;
        });

        switch (method) {
            case 'Fill with nearest available value':
                return this.fillNearest(sorted, column);
            
            case 'Fill with a linear interpolation between the nearest values':
                return this.fillInterpolate(sorted, column, dateTimeCol);
            
            case 'Delete the entire row of data':
                return sorted.filter(row => row[column] != null && row[column] !== '');
            
            case 'Fill with zero':
                return sorted.map(row => ({
                    ...row,
                    [column]: row[column] != null && row[column] !== '' ? row[column] : 0
                }));
            
            default:
                return sorted;
        }
    },

    /**
     * Fill missing values with nearest available (forward then backward fill)
     */
    fillNearest(data, column) {
        const result = [...data];
        
        // Forward fill
        let lastValid = null;
        for (let i = 0; i < result.length; i++) {
            if (result[i][column] != null && result[i][column] !== '') {
                lastValid = result[i][column];
            } else if (lastValid !== null) {
                result[i] = { ...result[i], [column]: lastValid };
            }
        }
        
        // Backward fill for any remaining nulls at start
        lastValid = null;
        for (let i = result.length - 1; i >= 0; i--) {
            if (result[i][column] != null && result[i][column] !== '') {
                lastValid = result[i][column];
            } else if (lastValid !== null) {
                result[i] = { ...result[i], [column]: lastValid };
            }
        }
        
        return result;
    },

    /**
     * Fill missing values with linear interpolation
     */
    fillInterpolate(data, column, dateTimeCol) {
        const result = [...data];
        
        for (let i = 0; i < result.length; i++) {
            if (result[i][column] == null || result[i][column] === '') {
                // Find previous valid value
                let prevIdx = i - 1;
                while (prevIdx >= 0 && (result[prevIdx][column] == null || result[prevIdx][column] === '')) {
                    prevIdx--;
                }
                
                // Find next valid value
                let nextIdx = i + 1;
                while (nextIdx < result.length && (result[nextIdx][column] == null || result[nextIdx][column] === '')) {
                    nextIdx++;
                }
                
                if (prevIdx >= 0 && nextIdx < result.length) {
                    // Linear interpolation
                    const prevVal = parseFloat(result[prevIdx][column]);
                    const nextVal = parseFloat(result[nextIdx][column]);
                    const prevTime = new Date(result[prevIdx][dateTimeCol]).getTime();
                    const nextTime = new Date(result[nextIdx][dateTimeCol]).getTime();
                    const currTime = new Date(result[i][dateTimeCol]).getTime();
                    
                    if (!isNaN(prevVal) && !isNaN(nextVal) && prevTime !== nextTime) {
                        const ratio = (currTime - prevTime) / (nextTime - prevTime);
                        const interpolated = prevVal + ratio * (nextVal - prevVal);
                        result[i] = { ...result[i], [column]: interpolated };
                    }
                } else if (prevIdx >= 0) {
                    // Only have previous - use it
                    result[i] = { ...result[i], [column]: result[prevIdx][column] };
                } else if (nextIdx < result.length) {
                    // Only have next - use it
                    result[i] = { ...result[i], [column]: result[nextIdx][column] };
                }
            }
        }
        
        return result;
    },

    /**
     * Handle duplicate timestamps
     * @param {Array} data - Data rows
     * @param {string} dateTimeCol - DateTime column name
     * @param {string} method - How to handle duplicates (Average, Maximum, Minimum)
     * @returns {Array}
     */
    handleDuplicates(data, dateTimeCol, method) {
        // Group by timestamp
        const groups = new Map();
        
        for (const row of data) {
            const key = String(row[dateTimeCol]);
            if (!groups.has(key)) {
                groups.set(key, []);
            }
            groups.get(key).push(row);
        }
        
        // Aggregate each group
        const result = [];
        for (const [timestamp, rows] of groups) {
            if (rows.length === 1) {
                result.push(rows[0]);
            } else {
                // Aggregate
                const aggregated = { [dateTimeCol]: rows[0][dateTimeCol] };
                const columns = Object.keys(rows[0]).filter(c => c !== dateTimeCol);
                
                for (const col of columns) {
                    const values = rows
                        .map(r => parseFloat(r[col]))
                        .filter(v => !isNaN(v));
                    
                    if (values.length === 0) {
                        aggregated[col] = rows[0][col]; // Keep original for non-numeric
                    } else {
                        switch (method) {
                            case 'Average values':
                                aggregated[col] = values.reduce((a, b) => a + b, 0) / values.length;
                                break;
                            case 'Maximum value':
                                aggregated[col] = Math.max(...values);
                                break;
                            case 'Minimum value':
                                aggregated[col] = Math.min(...values);
                                break;
                            default:
                                aggregated[col] = values[0];
                        }
                    }
                }
                result.push(aggregated);
            }
        }
        
        return result;
    },

    /**
     * Align data to target timestamps
     * @param {Array} data - Source data
     * @param {Array<Date>} targetTimestamps - Target timestamp array
     * @param {string} dateTimeCol - DateTime column name
     * @param {string} valueCol - Value column name
     * @param {string} method - Alignment method
     * @returns {Array} - Values aligned to target timestamps
     */
    alignToTimestamps(data, targetTimestamps, dateTimeCol, valueCol, method) {
        // Build sorted array of {time, value} from source data
        const sourceData = data
            .map(row => ({
                time: new Date(row[dateTimeCol]).getTime(),
                value: row[valueCol]
            }))
            .filter(d => !isNaN(d.time))
            .sort((a, b) => a.time - b.time);
        
        if (sourceData.length === 0) {
            return targetTimestamps.map(() => null);
        }
        
        const result = [];
        
        for (const targetTime of targetTimestamps) {
            const targetMs = targetTime.getTime();
            
            if (method === 'Fill with the nearest value') {
                // Find nearest value
                let nearest = null;
                let nearestDist = Infinity;
                
                for (const d of sourceData) {
                    const dist = Math.abs(d.time - targetMs);
                    if (dist < nearestDist) {
                        nearestDist = dist;
                        nearest = d.value;
                    }
                }
                result.push(nearest);
                
            } else if (method === 'Do a linear interpolation from the nearest values') {
                // Find surrounding values for interpolation
                let before = null;
                let after = null;
                
                for (let i = 0; i < sourceData.length; i++) {
                    if (sourceData[i].time <= targetMs) {
                        before = sourceData[i];
                    }
                    if (sourceData[i].time >= targetMs && after === null) {
                        after = sourceData[i];
                        break;
                    }
                }
                
                if (before === null && after === null) {
                    result.push(null);
                } else if (before === null) {
                    result.push(after.value);
                } else if (after === null) {
                    result.push(before.value);
                } else if (before.time === after.time) {
                    result.push(before.value);
                } else {
                    // Interpolate
                    const ratio = (targetMs - before.time) / (after.time - before.time);
                    const beforeVal = parseFloat(before.value);
                    const afterVal = parseFloat(after.value);
                    
                    if (!isNaN(beforeVal) && !isNaN(afterVal)) {
                        result.push(beforeVal + ratio * (afterVal - beforeVal));
                    } else {
                        result.push(before.value);
                    }
                }
                
            } else if (method === 'Take an average of the available values within the interval') {
                // This requires knowing the interval - for now, use a window approach
                // Find all values within the interval centered on target
                const intervalMs = 60000; // Will be passed properly in full implementation
                const windowStart = targetMs - intervalMs / 2;
                const windowEnd = targetMs + intervalMs / 2;
                
                const windowValues = sourceData
                    .filter(d => d.time >= windowStart && d.time < windowEnd)
                    .map(d => parseFloat(d.value))
                    .filter(v => !isNaN(v));
                
                if (windowValues.length > 0) {
                    result.push(windowValues.reduce((a, b) => a + b, 0) / windowValues.length);
                } else {
                    // Fall back to nearest
                    let nearest = null;
                    let nearestDist = Infinity;
                    for (const d of sourceData) {
                        const dist = Math.abs(d.time - targetMs);
                        if (dist < nearestDist) {
                            nearestDist = dist;
                            nearest = d.value;
                        }
                    }
                    result.push(nearest);
                }
            } else {
                result.push(null);
            }
        }
        
        return result;
    },

    /**
     * Create combined dataset from multiple files
     * @param {Object} filesData - Object with file data and settings
     * @param {Array<Date>} timestamps - Target timestamps
     * @param {Object} alignmentOptions - Alignment method per file
     * @returns {Object} - Combined data with metadata
     */
    createCombinedDataset(filesData, timestamps, alignmentOptions) {
        const combined = {
            timestamps: timestamps,
            columns: [],
            units: [],
            data: timestamps.map(ts => ({ DateTime: ts }))
        };
        
        for (const [fileName, fileInfo] of Object.entries(filesData)) {
            if (!fileInfo.selectedCols || Object.keys(fileInfo.selectedCols).length === 0) {
                continue;
            }
            
            const dateTimeCol = fileInfo.dateTimeCol;
            let processedData = [...fileInfo.data];
            
            // Handle duplicates
            if (fileInfo.hasDuplicates) {
                processedData = this.handleDuplicates(
                    processedData, 
                    dateTimeCol, 
                    fileInfo.dupeHandling || 'Average values'
                );
            }
            
            // Process each selected column
            for (const [origCol, newTitle] of Object.entries(fileInfo.selectedCols)) {
                // Apply cleanup
                const cleanupMethod = fileInfo.cleanup?.[origCol] || 'Fill with nearest available value';
                processedData = this.applyCleanup(processedData, origCol, cleanupMethod, dateTimeCol);
                
                // Align to target timestamps
                const alignmentMethod = alignmentOptions[fileName] || 'Fill with the nearest value';
                const alignedValues = this.alignToTimestamps(
                    processedData, 
                    timestamps, 
                    dateTimeCol, 
                    origCol, 
                    alignmentMethod
                );
                
                // Add to combined data
                combined.columns.push(newTitle);
                combined.units.push(fileInfo.units?.[origCol] || '');
                
                for (let i = 0; i < timestamps.length; i++) {
                    combined.data[i][newTitle] = alignedValues[i];
                }
            }
        }
        
        return combined;
    },

    /**
     * Format date for Excel
     * @param {Date} date
     * @returns {string}
     */
    formatDateForExcel(date) {
        return date.toISOString().replace('T', ' ').substring(0, 19);
    }
};

// Export for use in other modules
window.DataProcessing = DataProcessing;
