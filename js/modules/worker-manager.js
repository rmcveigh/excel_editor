/**
 * @file
 * Excel Editor Worker Manager Module
 */

export class ExcelEditorWorkerManager {
  constructor(app) {
    this.app = app;
    this.worker = null;
    this.isSupported = this.checkWorkerSupport();
    this.isReady = false;
    this.taskCounter = 0;
    this.pendingTasks = new Map();
    this.workerUrl = null;
  }

  /**
   * Check if Web Workers are supported
   */
  checkWorkerSupport() {
    return typeof Worker !== 'undefined' && typeof window !== 'undefined';
  }

  /**
   * Initialize the worker
   */
  async initialize() {
    if (!this.isSupported) {
      this.app.utilities.logDebug('Web Workers not supported, using fallback');
      return false;
    }

    try {
      // First attempt to load the worker from a separate file
      try {
        const workerPath =
          '/modules/custom/excel_editor/js/workers/excel-worker.js';
        this.worker = new Worker(workerPath);
        this.app.utilities.logDebug('Worker created from external file');
      } catch (pathError) {
        // If loading from external file fails, fall back to inline script
        this.app.utilities.logDebug(
          'Failed to load worker from file, using fallback:',
          pathError
        );

        const workerScript = this.getWorkerScript();
        const blob = new Blob([workerScript], {
          type: 'application/javascript',
        });
        const blobURL = URL.createObjectURL(blob);

        this.worker = new Worker(blobURL);
        this.workerUrl = blobURL; // Store for cleanup later

        this.app.utilities.logDebug('Worker created from blob URL');
      }

      // Set up message handler
      this.worker.onmessage = (event) => this.handleWorkerMessage(event);

      // Set up error handler
      this.worker.onerror = (error) => this.handleWorkerError(error);

      // Wait for worker to be ready
      return new Promise((resolve) => {
        const checkReady = (event) => {
          if (event.data && event.data.type === 'worker_ready') {
            this.isReady = event.data.success;
            this.worker.removeEventListener('message', checkReady);

            // Log the result for debugging
            this.app.utilities.logDebug('Worker ready status:', this.isReady);
            if (!this.isReady && event.data.error) {
              this.app.utilities.logDebug(
                'Worker initialization error:',
                event.data.error
              );
            }

            resolve(this.isReady);
          }
        };

        this.worker.addEventListener('message', checkReady);

        // Timeout after 10 seconds (increased from 5 seconds)
        setTimeout(() => {
          if (!this.isReady) {
            this.app.utilities.logDebug('Worker initialization timed out');
            this.worker.removeEventListener('message', checkReady);
            resolve(false);
          }
        }, 10000);
      });
    } catch (error) {
      this.app.utilities.logDebug('Failed to initialize worker:', error);
      return false;
    }
  }

  /**
   * Get the worker script content
   */
  getWorkerScript() {
    // Return the Excel worker script content
    // In a real implementation, this would either be fetched from a file
    // or embedded. For this example, I'll create a minimal version.
    return `
// Import SheetJS library in worker context
importScripts('https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js');

class ExcelWorker {
  constructor() {
    this.isReady = false;
    this.initialize();
  }

  initialize() {
    try {
      if (typeof XLSX === 'undefined') {
        throw new Error('XLSX library not available in worker context');
      }

      this.isReady = true;
      this.postMessage({
        type: 'worker_ready',
        success: true,
        message: 'Excel worker initialized successfully'
      });
    } catch (error) {
      this.postMessage({
        type: 'worker_ready',
        success: false,
        error: error.message
      });
    }
  }

  postMessage(data) {
    self.postMessage(data);
  }

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
        default:
          throw new Error(\`Unknown operation type: \${type}\`);
      }
    } catch (error) {
      this.postMessage({
        type: 'error',
        taskId,
        error: error.message
      });
    }
  }

  parseExcel(arrayBuffer, taskId) {
    try {
      this.postMessage({
        type: 'progress',
        taskId,
        progress: 10,
        message: 'Reading Excel file...'
      });

      const workbook = XLSX.read(arrayBuffer, {
        type: 'array',
        cellDates: true,
        cellNF: false,
        cellHTML: false
      });

      if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
        throw new Error('No worksheets found in Excel file.');
      }

      this.postMessage({
        type: 'progress',
        taskId,
        progress: 50,
        message: 'Processing worksheet...'
      });

      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];

      const jsonData = XLSX.utils.sheet_to_json(worksheet, {
        header: 1,
        raw: false,
        defval: '',
        blankrows: false
      });

      const processedData = this.cleanData(jsonData);

      if (processedData.length <= 1) {
        throw new Error('Excel file contains no data rows.');
      }

      this.postMessage({
        type: 'parse_complete',
        taskId,
        success: true,
        data: processedData,
        metadata: {
          sheets: workbook.SheetNames,
          activeSheet: sheetName,
          totalRows: processedData.length,
          totalColumns: processedData[0]?.length || 0
        }
      });

    } catch (error) {
      this.postMessage({
        type: 'parse_complete',
        taskId,
        success: false,
        error: error.message
      });
    }
  }

  parseCSV(textData, taskId) {
    try {
      this.postMessage({
        type: 'progress',
        taskId,
        progress: 20,
        message: 'Parsing CSV data...'
      });

      const lines = textData.split('\\n').filter(line => line.trim());
      const data = [];

      for (let i = 0; i < lines.length; i++) {
        const row = this.parseCSVLine(lines[i]);
        data.push(row);
      }

      const processedData = this.cleanData(data);

      this.postMessage({
        type: 'parse_complete',
        taskId,
        success: true,
        data: processedData,
        metadata: {
          totalRows: processedData.length,
          totalColumns: processedData[0]?.length || 0
        }
      });

    } catch (error) {
      this.postMessage({
        type: 'parse_complete',
        taskId,
        success: false,
        error: error.message
      });
    }
  }

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
          current += '"';
          i += 2;
        } else {
          inQuotes = !inQuotes;
          i++;
        }
      } else if (char === ',' && !inQuotes) {
        result.push(current.trim());
        current = '';
        i++;
      } else {
        current += char;
        i++;
      }
    }

    result.push(current.trim());
    return result;
  }

  exportExcel(exportData, taskId) {
    try {
      const { data, filename } = exportData;

      this.postMessage({
        type: 'progress',
        taskId,
        progress: 30,
        message: 'Creating worksheet...'
      });

      const ws = XLSX.utils.aoa_to_sheet(data);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Export');

      this.postMessage({
        type: 'progress',
        taskId,
        progress: 80,
        message: 'Generating Excel file...'
      });

      const binaryString = XLSX.write(wb, {
        bookType: 'xlsx',
        type: 'binary',
        compression: true
      });

      const buffer = new ArrayBuffer(binaryString.length);
      const view = new Uint8Array(buffer);
      for (let i = 0; i < binaryString.length; i++) {
        view[i] = binaryString.charCodeAt(i) & 0xFF;
      }

      this.postMessage({
        type: 'export_complete',
        taskId,
        success: true,
        filename: filename,
        data: buffer,
        metadata: {
          size: buffer.byteLength,
          rows: data.length,
          columns: data[0]?.length || 0
        }
      }, [buffer]);

    } catch (error) {
      this.postMessage({
        type: 'export_complete',
        taskId,
        success: false,
        error: error.message
      });
    }
  }

  cleanData(data) {
    if (!Array.isArray(data) || data.length === 0) {
      return [];
    }

    const cleaned = [];

    for (const row of data) {
      if (!Array.isArray(row)) continue;

      const hasContent = row.some(cell =>
        cell !== null &&
        cell !== undefined &&
        String(cell).trim() !== ''
      );

      if (hasContent) {
        const cleanedRow = row.map(cell => {
          if (cell === null || cell === undefined) {
            return '';
          }

          let cleaned = String(cell).trim();

          if (cleaned.length >= 2 &&
              cleaned.startsWith('"') &&
              cleaned.endsWith('"')) {
            cleaned = cleaned.slice(1, -1);
          }

          return cleaned;
        });

        cleaned.push(cleanedRow);
      }
    }

    return cleaned;
  }
}

const excelWorker = new ExcelWorker();

self.onmessage = function(event) {
  excelWorker.handleMessage(event);
};

self.onerror = function(error) {
  self.postMessage({
    type: 'worker_error',
    error: error.message
  });
};
    `;
  }

  /**
   * Handle messages from the worker
   */
  handleWorkerMessage(event) {
    const { type, taskId, ...data } = event.data;

    const task = this.pendingTasks.get(taskId);
    if (!task) {
      this.app.utilities.logDebug('Received message for unknown task:', taskId);
      return;
    }

    switch (type) {
      case 'progress':
        if (task.onProgress) {
          task.onProgress(data.progress, data.message);
        }
        break;

      case 'parse_complete':
      case 'export_complete':
        this.pendingTasks.delete(taskId);
        if (data.success) {
          task.resolve(data);
        } else {
          task.reject(new Error(data.error));
        }
        break;

      case 'error':
        this.pendingTasks.delete(taskId);
        task.reject(new Error(data.error));
        break;

      case 'worker_error':
        this.app.utilities.logDebug('Worker error:', data.error);
        break;
    }
  }

  /**
   * Handle worker errors
   */
  handleWorkerError(error) {
    this.app.utilities.logDebug('Worker error:', error);
    this.isReady = false;

    // Reject all pending tasks
    for (const [taskId, task] of this.pendingTasks) {
      task.reject(new Error('Worker error: ' + error.message));
    }
    this.pendingTasks.clear();
  }

  /**
   * Generate unique task ID
   */
  generateTaskId() {
    return `task_${Date.now()}_${++this.taskCounter}`;
  }

  /**
   * Send task to worker
   */
  sendTask(type, data, onProgress = null) {
    if (!this.isReady) {
      return Promise.reject(new Error('Worker not ready'));
    }

    const taskId = this.generateTaskId();

    return new Promise((resolve, reject) => {
      // Store task info
      this.pendingTasks.set(taskId, {
        resolve,
        reject,
        onProgress,
        startTime: Date.now(),
      });

      // Send message to worker
      this.worker.postMessage({
        type,
        data,
        taskId,
      });

      // Set timeout for long-running tasks
      setTimeout(() => {
        if (this.pendingTasks.has(taskId)) {
          this.pendingTasks.delete(taskId);
          reject(new Error('Task timeout'));
        }
      }, 60000); // 1 minute timeout
    });
  }

  /**
   * Parse Excel file using worker
   */
  async parseExcel(arrayBuffer, onProgress = null) {
    if (!this.isSupported || !this.isReady) {
      throw new Error('Worker not available for Excel parsing');
    }

    return this.sendTask('parse_excel', arrayBuffer, onProgress);
  }

  /**
   * Parse CSV file using worker
   */
  async parseCSV(textData, onProgress = null) {
    if (!this.isSupported || !this.isReady) {
      throw new Error('Worker not available for CSV parsing');
    }

    return this.sendTask('parse_csv', textData, onProgress);
  }

  /**
   * Export Excel file using worker
   */
  async exportExcel(data, filename, options = {}, onProgress = null) {
    if (!this.isSupported || !this.isReady) {
      throw new Error('Worker not available for Excel export');
    }

    return this.sendTask(
      'export_excel',
      {
        data,
        filename,
        options,
      },
      onProgress
    );
  }

  /**
   * Check if worker operations are available
   */
  isAvailable() {
    return this.isSupported && this.isReady;
  }

  /**
   * Get worker status information
   */
  getStatus() {
    return {
      supported: this.isSupported,
      ready: this.isReady,
      pendingTasks: this.pendingTasks.size,
    };
  }

  /**
   * Terminate the worker
   */
  terminate() {
    if (this.worker) {
      // Reject all pending tasks
      for (const [taskId, task] of this.pendingTasks) {
        task.reject(new Error('Worker terminated'));
      }
      this.pendingTasks.clear();

      this.worker.terminate();
      this.worker = null;
      this.isReady = false;
    }

    if (this.workerUrl) {
      URL.revokeObjectURL(this.workerUrl);
      this.workerUrl = null;
    }
  }
}
