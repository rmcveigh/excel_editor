/**
 * @file
 * Excel Worker - Handles Excel parsing and export operations in a Web Worker
 */

// Log initialization to help with debugging
self.console.log('Excel worker initializing...');

// Import SheetJS library in worker context
try {
  importScripts(
    'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js'
  );
  self.console.log('XLSX library loaded successfully');
} catch (error) {
  self.console.error('Error loading XLSX library:', error);
  self.postMessage({
    type: 'worker_ready',
    success: false,
    error: 'Failed to load XLSX library: ' + error.message,
  });
}

/**
 * Excel Worker class to handle all Excel operations
 */
class ExcelWorker {
  constructor() {
    this.isReady = false;
    this.initialize();
  }

  /**
   * Initialize the worker
   */
  initialize() {
    try {
      // Verify XLSX is available
      if (typeof XLSX === 'undefined') {
        throw new Error('XLSX library not available in worker context');
      }

      this.isReady = true;
      this.postMessage({
        type: 'worker_ready',
        success: true,
        message: 'Excel worker initialized successfully',
      });
      self.console.log('Excel worker initialized successfully');
    } catch (error) {
      self.console.error('Worker initialization failed:', error);
      this.postMessage({
        type: 'worker_ready',
        success: false,
        error: error.message,
      });
    }
  }

  /**
   * Send message back to main thread
   */
  postMessage(data) {
    self.postMessage(data);
  }

  /**
   * Handle incoming messages from main thread
   */
  handleMessage(event) {
    const { type, data, taskId } = event.data;

    try {
      switch (type) {
        case 'parse_excel':
          this.parseExcel(data, taskId);
          break;
        case 'parse_csv':
          this.parseCSV(data, taskId);
          break;
        case 'export_excel':
          this.exportExcel(data, taskId);
          break;
        case 'validate_file':
          this.validateFile(data, taskId);
          break;
        default:
          throw new Error(`Unknown operation type: ${type}`);
      }
    } catch (error) {
      this.postMessage({
        type: 'error',
        taskId,
        error: error.message,
        stack: error.stack,
      });
    }
  }

  /**
   * Parse Excel file data
   */
  parseExcel(arrayBuffer, taskId) {
    try {
      // Send progress update
      this.postMessage({
        type: 'progress',
        taskId,
        progress: 10,
        message: 'Reading Excel file...',
      });

      const workbook = XLSX.read(arrayBuffer, {
        type: 'array',
        cellDates: true,
        cellNF: false,
        cellHTML: false,
      });

      if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
        throw new Error('No worksheets found in Excel file.');
      }

      this.postMessage({
        type: 'progress',
        taskId,
        progress: 30,
        message: 'Processing worksheet...',
      });

      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];

      // Convert to JSON with specific options for better performance
      const jsonData = XLSX.utils.sheet_to_json(worksheet, {
        header: 1,
        raw: false,
        defval: '',
        blankrows: false,
      });

      this.postMessage({
        type: 'progress',
        taskId,
        progress: 60,
        message: 'Cleaning data...',
      });

      // Clean and process the data
      const processedData = this.cleanExcelData(jsonData);

      this.postMessage({
        type: 'progress',
        taskId,
        progress: 90,
        message: 'Finalizing...',
      });

      if (processedData.length <= 1) {
        throw new Error('Excel file contains no data rows.');
      }

      // Send the final result
      this.postMessage({
        type: 'parse_complete',
        taskId,
        success: true,
        data: processedData,
        metadata: {
          sheets: workbook.SheetNames,
          activeSheet: sheetName,
          totalRows: processedData.length,
          totalColumns: processedData[0]?.length || 0,
        },
      });
    } catch (error) {
      this.postMessage({
        type: 'parse_complete',
        taskId,
        success: false,
        error: error.message,
      });
    }
  }

  /**
   * Parse CSV file data
   */
  parseCSV(textData, taskId) {
    try {
      this.postMessage({
        type: 'progress',
        taskId,
        progress: 20,
        message: 'Parsing CSV data...',
      });

      // Simple CSV parsing with quote handling
      const lines = textData.split('\n').filter((line) => line.trim());
      const data = [];

      for (let i = 0; i < lines.length; i++) {
        if (i % 100 === 0) {
          // Send progress updates for large files
          this.postMessage({
            type: 'progress',
            taskId,
            progress: 20 + (i / lines.length) * 60,
            message: `Processing row ${i + 1} of ${lines.length}...`,
          });
        }

        const row = this.parseCSVLine(lines[i]);
        data.push(row);
      }

      this.postMessage({
        type: 'progress',
        taskId,
        progress: 90,
        message: 'Finalizing CSV processing...',
      });

      const processedData = this.cleanExcelData(data);

      this.postMessage({
        type: 'parse_complete',
        taskId,
        success: true,
        data: processedData,
        metadata: {
          totalRows: processedData.length,
          totalColumns: processedData[0]?.length || 0,
        },
      });
    } catch (error) {
      this.postMessage({
        type: 'parse_complete',
        taskId,
        success: false,
        error: error.message,
      });
    }
  }

  /**
   * Parse a single CSV line with quote handling
   */
  parseCSVLine(line) {
    const result = [];
    let current = '';
    let inQuotes = false;
    let i = 0;

    while (i < line.length) {
      const char = line[i];
      const nextChar = line[i + 1];

      if (char === '"') {
        if (inQuotes && nextChar === '"') {
          // Escaped quote
          current += '"';
          i += 2;
        } else {
          // Toggle quote state
          inQuotes = !inQuotes;
          i++;
        }
      } else if (char === ',' && !inQuotes) {
        // Field separator
        result.push(current.trim());
        current = '';
        i++;
      } else {
        current += char;
        i++;
      }
    }

    // Add the last field
    result.push(current.trim());
    return result;
  }

  /**
   * Clean and standardize Excel/CSV data
   */
  cleanExcelData(data) {
    if (!Array.isArray(data) || data.length === 0) {
      return [];
    }

    // Remove completely empty rows and standardize cell values
    const cleaned = [];

    for (const row of data) {
      if (!Array.isArray(row)) continue;

      // Check if row has any content
      const hasContent = row.some(
        (cell) =>
          cell !== null && cell !== undefined && String(cell).trim() !== ''
      );

      if (hasContent) {
        // Clean each cell in the row
        const cleanedRow = row.map((cell) => {
          if (cell === null || cell === undefined) {
            return '';
          }

          let cleaned = String(cell).trim();

          // Remove quotes if they wrap the entire value
          if (
            cleaned.length >= 2 &&
            cleaned.startsWith('"') &&
            cleaned.endsWith('"')
          ) {
            cleaned = cleaned.slice(1, -1);
          }

          return cleaned;
        });

        cleaned.push(cleanedRow);
      }
    }

    return cleaned;
  }

  /**
   * Export data to Excel format
   */
  exportExcel(exportData, taskId) {
    try {
      const { data, filename, options = {} } = exportData;

      this.postMessage({
        type: 'progress',
        taskId,
        progress: 10,
        message: 'Preparing export data...',
      });

      if (!Array.isArray(data) || data.length === 0) {
        throw new Error('No data provided for export');
      }

      this.postMessage({
        type: 'progress',
        taskId,
        progress: 30,
        message: 'Creating worksheet...',
      });

      // Create worksheet from array of arrays
      const ws = XLSX.utils.aoa_to_sheet(data);

      // Apply column widths if provided
      if (options.columnWidths) {
        const colWidths = options.columnWidths.map((width) => ({ wch: width }));
        ws['!cols'] = colWidths;
      }

      // Apply cell styles if provided
      if (options.styles) {
        this.applyExcelStyles(ws, options.styles);
      }

      this.postMessage({
        type: 'progress',
        taskId,
        progress: 60,
        message: 'Creating workbook...',
      });

      // Create workbook
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, options.sheetName || 'Export');

      this.postMessage({
        type: 'progress',
        taskId,
        progress: 80,
        message: 'Generating Excel file...',
      });

      // Generate binary data
      const binaryString = XLSX.write(wb, {
        bookType: 'xlsx',
        type: 'binary',
        compression: true,
      });

      // Convert to ArrayBuffer for transfer
      const buffer = new ArrayBuffer(binaryString.length);
      const view = new Uint8Array(buffer);
      for (let i = 0; i < binaryString.length; i++) {
        view[i] = binaryString.charCodeAt(i) & 0xff;
      }

      this.postMessage({
        type: 'progress',
        taskId,
        progress: 95,
        message: 'Finalizing export...',
      });

      // Send the result with transferable object for better performance
      this.postMessage(
        {
          type: 'export_complete',
          taskId,
          success: true,
          filename: filename,
          data: buffer,
          metadata: {
            size: buffer.byteLength,
            rows: data.length,
            columns: data[0]?.length || 0,
          },
        },
        [buffer]
      ); // Transfer ownership of the ArrayBuffer
    } catch (error) {
      this.postMessage({
        type: 'export_complete',
        taskId,
        success: false,
        error: error.message,
      });
    }
  }

  /**
   * Apply Excel styles to worksheet
   */
  applyExcelStyles(worksheet, styles) {
    try {
      // Apply header styles
      if (styles.header) {
        const range = XLSX.utils.decode_range(worksheet['!ref']);
        for (let col = range.s.c; col <= range.e.c; col++) {
          const cellRef = XLSX.utils.encode_cell({ r: 0, c: col });
          if (worksheet[cellRef]) {
            worksheet[cellRef].s = styles.header;
          }
        }
      }

      // Apply data styles
      if (styles.data) {
        const range = XLSX.utils.decode_range(worksheet['!ref']);
        for (let row = 1; row <= range.e.r; row++) {
          for (let col = range.s.c; col <= range.e.c; col++) {
            const cellRef = XLSX.utils.encode_cell({ r: row, c: col });
            if (worksheet[cellRef]) {
              worksheet[cellRef].s = styles.data;
            }
          }
        }
      }
    } catch (error) {
      // Style application is optional, continue without styles
      console.warn('Failed to apply Excel styles:', error);
    }
  }

  /**
   * Validate file before processing
   */
  validateFile(fileData, taskId) {
    try {
      const { name, size, type } = fileData;
      const errors = [];
      const warnings = [];

      // Size validation (10MB limit)
      const maxSize = 10 * 1024 * 1024;
      if (size > maxSize) {
        errors.push(
          `File too large: ${Math.round(size / 1024 / 1024)}MB (max: 10MB)`
        );
      }

      // Type validation
      const validExtensions = ['.xlsx', '.xls', '.csv'];
      const extension = '.' + name.split('.').pop().toLowerCase();
      if (!validExtensions.includes(extension)) {
        errors.push(`Unsupported file type: ${extension}`);
      }

      // Size warnings
      if (size > 5 * 1024 * 1024) {
        warnings.push(
          'Large file detected. Processing may take longer than usual.'
        );
      }

      this.postMessage({
        type: 'validation_complete',
        taskId,
        success: errors.length === 0,
        errors,
        warnings,
        fileInfo: {
          name,
          size,
          type,
          extension,
        },
      });
    } catch (error) {
      this.postMessage({
        type: 'validation_complete',
        taskId,
        success: false,
        error: error.message,
      });
    }
  }
}

// Initialize the worker
const excelWorker = new ExcelWorker();

// Set up message handling
self.onmessage = function (event) {
  excelWorker.handleMessage(event);
};

// Set up error handling
self.onerror = function (error) {
  self.console.error('Worker global error:', error);
  self.postMessage({
    type: 'worker_error',
    error: error.message,
  });
};

// Handle unhandled promise rejections
self.onunhandledrejection = function (event) {
  self.postMessage({
    type: 'worker_error',
    error: 'Unhandled promise rejection: ' + event.reason,
  });
};
