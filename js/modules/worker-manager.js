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
        // A short timeout to give the worker time to load and be ready
        setTimeout(() => {
          this.isReady = true;
          this.app.utilities.logDebug('Worker assumed ready.');
          resolve(true);
        }, 1000);
      });
    } catch (error) {
      this.app.utilities.logDebug('Failed to initialize worker:', error);
      return false;
    }
  }

  /**
   * Get the worker script content for the fallback method.
   */
  getWorkerScript() {
    // This now contains the full, self-contained worker code.
    return `
      importScripts('https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js');

      function parseExcelData(data) {
        const workbook = XLSX.read(data, { type: 'array' });
        if (!workbook.SheetNames || workbook.SheetNames.length === 0) throw new Error('No worksheets found');
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false, defval: '' });
        const trimmed = jsonData.map(row => Array.isArray(row) ? row.map(cell => String(cell || '').trim()) : row);
        const filtered = trimmed.filter(row => Array.isArray(row) && row.some(cell => cell !== null && cell !== undefined && String(cell).trim() !== ''));
        if (filtered.length <= 1) throw new Error('No data rows found');
        return filtered;
      }

      function parseCSVData(text) {
        const lines = text.split('\\n').filter(line => line.trim());
        return lines.map(line => line.split(',').map(cell => cell.trim().replace(/^["']|["']$/g, '')));
      }

      function createXLSX(data) {
        const ws = XLSX.utils.aoa_to_sheet(data);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Export');
        const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        return new Blob([wbout], { type: 'application/octet-stream' });
      }

      self.onmessage = function (event) {
        const { type, data, taskId } = event.data;
        try {
          let result;
          switch (type) {
            case 'parse_excel':
              result = parseExcelData(data);
              self.postMessage({ type: 'parse_complete', taskId, success: true, data: result });
              break;
            case 'parse_csv':
              result = parseCSVData(data);
              self.postMessage({ type: 'parse_complete', taskId, success: true, data: result });
              break;
            case 'export_excel':
              const blob = createXLSX(data.data);
              self.postMessage({ type: 'export_complete', taskId, success: true, data: blob, filename: data.filename });
              break;
            default:
              throw new Error('Unknown operation type');
          }
        } catch (error) {
          self.postMessage({ type: 'error', taskId, error: error.message });
        }
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
    for (const task of this.pendingTasks.values()) {
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
      for (const task of this.pendingTasks.values()) {
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
