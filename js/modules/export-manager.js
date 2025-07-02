/**
 * @file
 * Excel Editor Export Manager Module - Updated with Web Worker Support
 */

export class ExcelEditorExportManager {
  constructor(app) {
    this.app = app;
  }

  /**
   * Exports only the selected rows to an XLSX file.
   * @returns {Promise<void>} - Resolves when export is complete.
   */
  async exportSelected() {
    if (!this.app.data?.selected?.size) {
      this.app.utilities.showMessage('No rows selected for export', 'warning');
      return;
    }

    try {
      const exportData = this.prepareExportData(true);
      const filename = this.generateExportFilename('selected');

      // Determine if we should use worker
      const useWorker = this.shouldUseWorkerForExport(exportData);

      if (useWorker) {
        await this.exportWithWorker(exportData, filename, 'selected');
      } else {
        await this.exportWithMainThread(exportData, filename, 'selected');
      }

      this.app.utilities.showMessage(
        `Exported ${this.app.data.selected.size} selected rows`,
        'success'
      );
    } catch (error) {
      this.app.utilities.handleError('Failed to export selected rows', error);
    }
  }

  /**
   * Exports all visible (filtered) rows to an XLSX file.
   * @returns {Promise<void>} - Resolves when export is complete.
   */
  async exportAll() {
    if (
      !this.app.data?.filtered?.length ||
      this.app.data.filtered.length <= 1
    ) {
      this.app.utilities.showMessage('No data available to export', 'warning');
      return;
    }

    try {
      const exportData = this.prepareExportData(false);
      const filename = this.generateExportFilename('all');

      // Determine if we should use worker
      const useWorker = this.shouldUseWorkerForExport(exportData);

      if (useWorker) {
        await this.exportWithWorker(exportData, filename, 'all');
      } else {
        await this.exportWithMainThread(exportData, filename, 'all');
      }

      const rowCount = this.app.data.filtered.length - 1;
      this.app.utilities.showMessage(`Exported ${rowCount} rows`, 'success');
    } catch (error) {
      this.app.utilities.handleError('Failed to export data', error);
    }
  }

  /**
   * Determine if worker should be used for export
   * @param {Array} exportData - The data to be exported
   * @returns {boolean} - True if worker should be used
   */
  shouldUseWorkerForExport(exportData) {
    // Use worker for larger datasets or when explicitly available
    const dataSize = exportData.length;
    const cellCount = dataSize * (exportData[0]?.length || 0);
    const threshold = 1000; // rows
    const cellThreshold = 10000; // total cells

    return (
      this.app.workerManager &&
      this.app.workerManager.isAvailable() &&
      (dataSize > threshold || cellCount > cellThreshold)
    );
  }

  /**
   * Export using Web Worker
   * @param {Array} exportData - The data to export
   * @param {string} filename - The filename for the export
   * @param {string} exportType - Type of export ('selected' or 'all')
   */
  async exportWithWorker(exportData, filename, exportType) {
    this.app.utilities.logDebug(`Using Web Worker for ${exportType} export`);

    // Show loader with progress support
    this.app.utilities.showProcessLoader('Preparing export...');

    try {
      // Progress callback
      const onProgress = (progress, message) => {
        this.updateExportProgress(progress, message);
      };

      // Export options
      const options = {
        sheetName: 'Export',
        columnWidths: this.calculateColumnWidths(exportData),
        styles: this.getExportStyles(),
      };

      // Use worker for export
      const result = await this.app.workerManager.exportExcel(
        exportData,
        filename,
        options,
        onProgress
      );

      if (!result.success) {
        throw new Error(result.error);
      }

      // Download the file
      await this.downloadFromArrayBuffer(result.data, result.filename);

      this.app.utilities.logDebug('Worker export completed:', result.metadata);
    } finally {
      this.app.utilities.hideProcessLoader();
    }
  }

  /**
   * Export using main thread (fallback)
   * @param {Array} exportData - The data to export
   * @param {string} filename - The filename for the export
   * @param {string} exportType - Type of export ('selected' or 'all')
   */
  async exportWithMainThread(exportData, filename, exportType) {
    this.app.utilities.logDebug(`Using main thread for ${exportType} export`);

    const isLargeDataset = exportData.length > 1000;

    try {
      if (isLargeDataset) {
        this.app.utilities.showProcessLoader('Preparing large export...');
      } else {
        this.app.utilities.showQuickLoader('Preparing export...');
      }

      await this.downloadExport(exportData, filename);
    } finally {
      if (isLargeDataset) {
        this.app.utilities.hideProcessLoader();
      } else {
        this.app.utilities.hideQuickLoader();
      }
    }
  }

  /**
   * Update export progress display
   * @param {number} progress - Progress percentage (0-100)
   * @param {string} message - Progress message
   */
  updateExportProgress(progress, message) {
    const $ = jQuery;

    // Update progress in the process loader
    let progressContainer = $('.excel-editor-overlay-loader .loading-content');

    if (progressContainer.length > 0) {
      // Update existing progress display
      let progressBar = progressContainer.find('.progress');
      let progressText = progressContainer.find('p');

      if (progressBar.length === 0) {
        // Add progress bar if it doesn't exist
        progressContainer.prepend(`
          <progress class="progress is-primary" value="${progress}" max="100" style="margin-bottom: 1rem;"></progress>
        `);
      } else {
        progressBar.attr('value', progress);
      }

      progressText.html(
        `<strong>${
          message || 'Processing...'
        }</strong><br><small>${progress}% complete</small>`
      );
    }
  }

  /**
   * Calculate optimal column widths for export
   * @param {Array} data - The export data
   * @returns {Array} - Array of column widths
   */
  calculateColumnWidths(data) {
    if (!data || data.length === 0) return [];

    const columnWidths = [];
    const maxColumns = data[0]?.length || 0;

    for (let col = 0; col < maxColumns; col++) {
      let maxWidth = 10; // Minimum width

      // Check first few rows to estimate width
      const checkRows = Math.min(data.length, 20);
      for (let row = 0; row < checkRows; row++) {
        const cellValue = String(data[row][col] || '');
        maxWidth = Math.max(maxWidth, cellValue.length);
      }

      // Cap maximum width and apply scaling
      columnWidths.push(Math.min(maxWidth * 1.2, 50));
    }

    return columnWidths;
  }

  /**
   * Get export styles for Excel formatting
   * @returns {Object} - Style definitions
   */
  getExportStyles() {
    return {
      header: {
        font: { bold: true },
        fill: { fgColor: { rgb: 'EEEEEE' } },
        border: {
          top: { style: 'thin' },
          bottom: { style: 'thin' },
          left: { style: 'thin' },
          right: { style: 'thin' },
        },
      },
      data: {
        border: {
          top: { style: 'thin' },
          bottom: { style: 'thin' },
          left: { style: 'thin' },
          right: { style: 'thin' },
        },
      },
    };
  }

  /**
   * Download file from ArrayBuffer (for worker exports)
   * @param {ArrayBuffer} buffer - The file data
   * @param {string} filename - The filename
   */
  async downloadFromArrayBuffer(buffer, filename) {
    try {
      const blob = new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });

      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');

      link.href = url;
      link.setAttribute('download', filename);
      link.style.visibility = 'hidden';

      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);

      // Clean up
      setTimeout(() => URL.revokeObjectURL(url), 1000);
    } catch (error) {
      this.app.utilities.logDebug('ArrayBuffer download error:', error);
      throw new Error('Failed to download exported file');
    }
  }

  /**
   * Prepares the data for export by creating an array of arrays.
   * @param {boolean} selectedOnly - If true, only selected rows are included.
   * @return {Array<Array>} - Array of arrays representing the export data.
   */
  prepareExportData(selectedOnly = false) {
    if (!this.app.data?.filtered?.length) {
      return [[]];
    }

    const exportData = [];
    const headerRow = [];

    // Process headers
    this.app.data.filtered[0].forEach((header, index) => {
      if (!this.app.state.hiddenColumns.has(index)) {
        headerRow.push(header);
      }
    });
    exportData.push(headerRow);

    // Process data rows
    for (let i = 1; i < this.app.data.filtered.length; i++) {
      if (!selectedOnly || this.app.data.selected.has(i)) {
        const dataRow = [];
        this.app.data.filtered[i].forEach((cell, index) => {
          if (!this.app.state.hiddenColumns.has(index)) {
            dataRow.push(cell);
          }
        });
        exportData.push(dataRow);
      }
    }

    return exportData;
  }

  /**
   * Triggers the download of the exported data as an XLSX file (main thread version).
   * @param {Array<Array>} data - The data to export.
   * @param {string} filename - The name of the file to download.
   * @returns {Promise<void>} - Resolves when the file is downloaded.
   */
  async downloadExport(data, filename) {
    try {
      const ws = XLSX.utils.aoa_to_sheet(data);

      // Apply column widths for better formatting
      const columnWidths = this.calculateColumnWidths(data);
      ws['!cols'] = columnWidths.map((width) => ({ wch: width }));

      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Export');

      XLSX.writeFile(wb, filename);
    } catch (error) {
      this.app.utilities.logDebug('XLSX export error:', error);
      throw new Error('Failed to generate Excel file');
    }
  }

  /**
   * Generates a filename for exported data with timestamp.
   * @param {string} type - The type of export (e.g., 'selected', 'all').
   * @return {string} - The generated filename.
   */
  generateExportFilename(type) {
    const timestamp = new Date()
      .toISOString()
      .replace(/[:.]/g, '-')
      .split('T')[0];
    const baseFilename = this.app.config.exportFilenameBase || 'excel_editor';
    return `${baseFilename}_${type}_${timestamp}.xlsx`;
  }

  /**
   * Exports data to CSV format (alternative to XLSX).
   * @param {boolean} selectedOnly - If true, exports only selected rows.
   */
  async exportToCSV(selectedOnly = false) {
    try {
      this.app.utilities.showQuickLoader('Preparing CSV export...');

      const data = this.prepareExportData(selectedOnly);
      if (!data?.length) {
        this.app.utilities.showMessage('No data to export', 'warning');
        return;
      }

      // For large CSV exports, consider using worker
      const useWorker =
        this.shouldUseWorkerForExport(data) &&
        this.app.workerManager &&
        this.app.workerManager.isAvailable();

      if (useWorker) {
        await this.exportCSVWithWorker(data, selectedOnly);
      } else {
        await this.exportCSVMainThread(data, selectedOnly);
      }

      const rowCount = selectedOnly
        ? this.app.data.selected.size
        : this.app.data.filtered.length - 1;
      this.app.utilities.showMessage(
        `Exported ${rowCount} rows to CSV`,
        'success'
      );
    } catch (error) {
      this.app.utilities.handleError('Failed to export CSV', error);
    } finally {
      this.app.utilities.hideQuickLoader();
    }
  }

  /**
   * Export CSV using worker
   */
  async exportCSVWithWorker(data, selectedOnly) {
    // For CSV, we'll process on main thread as it's simpler
    // Worker implementation would require additional CSV formatting logic
    await this.exportCSVMainThread(data, selectedOnly);
  }

  /**
   * Export CSV on main thread
   */
  async exportCSVMainThread(data, selectedOnly) {
    // Convert data to CSV format
    const csvContent = data
      .map((row) =>
        row
          .map((cell) => {
            // Handle commas and quotes in cells
            const cellStr = String(cell || '');
            return cellStr.includes(',') ||
              cellStr.includes('"') ||
              cellStr.includes('\n')
              ? `"${cellStr.replace(/"/g, '""')}"`
              : cellStr;
          })
          .join(',')
      )
      .join('\n');

    // Create download link
    const blob = new Blob([csvContent], {
      type: 'text/csv;charset=utf-8;',
    });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    const filename = this.generateExportFilename(
      selectedOnly ? 'selected' : 'all'
    ).replace('.xlsx', '.csv');

    link.href = url;
    link.setAttribute('download', filename);
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);

    // Clean up
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  /**
   * Exports data with validation check.
   * @param {boolean} selectedOnly - If true, exports only selected rows.
   * @returns {Promise<void>} - Resolves when export is complete.
   */
  async exportWithValidation(selectedOnly = false) {
    // Check if validation is available and has errors
    if (
      this.app.validationManager &&
      this.app.validationManager.hasValidationErrors()
    ) {
      const shouldContinue = this.app.validationManager.validateBeforeExport();
      if (!shouldContinue) {
        return; // User canceled export due to validation errors
      }
    }

    // Proceed with export
    if (selectedOnly) {
      await this.exportSelected();
    } else {
      await this.exportAll();
    }
  }

  /**
   * Exports data with custom format options.
   * @returns {Promise<void>} - Resolves when export is complete.
   */
  async exportWithOptions() {
    const options = await this.showExportOptionsModal();
    if (!options) {
      return; // User canceled
    }

    try {
      if (options.format === 'csv') {
        await this.exportToCSV(options.selectedOnly);
      } else {
        if (options.selectedOnly) {
          await this.exportSelected();
        } else {
          await this.exportAll();
        }
      }
    } catch (error) {
      this.app.utilities.handleError('Export failed', error);
    }
  }

  /**
   * Shows export options modal with enhanced worker information.
   * @returns {Promise<{selectedOnly: boolean, format: string, useWorker: boolean}|null>} - Resolves with selected options or null if canceled.
   */
  showExportOptionsModal() {
    return new Promise((resolve) => {
      const $ = jQuery;
      const selectedCount = this.app.data.selected.size;
      const totalCount = this.app.data.filtered.length - 1;
      const workerAvailable =
        this.app.workerManager && this.app.workerManager.isAvailable();
      const wouldUseWorker = this.shouldUseWorkerForExport(
        this.prepareExportData(false)
      );

      const modalHtml = `
      <div class="modal is-active" id="export-options-modal">
        <div class="modal-background"></div>
        <div class="modal-content">
          <div class="box">
            <h3 class="title is-4">
              <span class="icon"><i class="fas fa-download"></i></span>
              Export Options
            </h3>

            ${
              workerAvailable && wouldUseWorker
                ? `
            <div class="notification is-info is-light mb-4">
              <span class="icon"><i class="fas fa-rocket"></i></span>
              <strong>Performance Mode:</strong> Large dataset detected. Will use background processing for faster export.
            </div>
            `
                : ''
            }

            <div class="field">
              <label class="label">Export Scope</label>
              <div class="control">
                <label class="radio">
                  <input type="radio" name="export-scope" value="selected" ${
                    selectedCount > 0 ? '' : 'disabled'
                  }>
                  Selected rows only (${selectedCount} rows)
                </label>
              </div>
              <div class="control">
                <label class="radio">
                  <input type="radio" name="export-scope" value="all" checked>
                  All visible rows (${totalCount} rows)
                </label>
              </div>
            </div>

            <div class="field">
              <label class="label">Export Format</label>
              <div class="control">
                <label class="radio">
                  <input type="radio" name="export-format" value="xlsx" checked>
                  Excel (.xlsx) - Preserves formatting${
                    workerAvailable ? ' • Background processing' : ''
                  }
                </label>
              </div>
              <div class="control">
                <label class="radio">
                  <input type="radio" name="export-format" value="csv">
                  CSV (.csv) - Text format, compatible with most applications
                </label>
              </div>
            </div>

            <div class="field is-grouped is-grouped-right">
              <div class="control">
                <button class="button" id="cancel-export">Cancel</button>
              </div>
              <div class="control">
                <button class="button is-primary" id="confirm-export">
                  <span class="icon"><i class="fas fa-download"></i></span>
                  <span>Export</span>
                </button>
              </div>
            </div>
          </div>
        </div>
        <button class="modal-close is-large" aria-label="close"></button>
      </div>`;

      const modal = $(modalHtml);
      $('body').append(modal);

      // Close handlers
      modal
        .find('.modal-close, #cancel-export, .modal-background')
        .on('click', () => {
          resolve(null);
          modal.remove();
        });

      // Export handler
      modal.find('#confirm-export').on('click', () => {
        const scope = modal.find('input[name="export-scope"]:checked').val();
        const format = modal.find('input[name="export-format"]:checked').val();

        resolve({
          selectedOnly: scope === 'selected',
          format: format,
          useWorker: workerAvailable && wouldUseWorker,
        });
        modal.remove();
      });
    });
  }

  /**
   * Gets export statistics for the current data.
   */
  getExportStats() {
    const totalRows = this.app.data.original.length - 1;
    const filteredRows = this.app.data.filtered.length - 1;
    const selectedRows = this.app.data.selected.size;
    const visibleColumns =
      this.app.data.filtered[0]?.length - this.app.state.hiddenColumns.size ||
      0;

    return {
      totalRows,
      filteredRows,
      selectedRows,
      visibleColumns,
      hasSelection: selectedRows > 0,
      hasFilters: Object.keys(this.app.state.currentFilters).length > 0,
      hasHiddenColumns: this.app.state.hiddenColumns.size > 0,
      workerAvailable:
        this.app.workerManager && this.app.workerManager.isAvailable(),
      wouldUseWorker: this.shouldUseWorkerForExport(
        this.prepareExportData(false)
      ),
    };
  }

  /**
   * Export with progress tracking and cancellation support
   * @param {boolean} selectedOnly - Whether to export only selected rows
   * @returns {Promise<void>}
   */
  async exportWithProgress(selectedOnly = false) {
    const exportData = this.prepareExportData(selectedOnly);
    const filename = this.generateExportFilename(
      selectedOnly ? 'selected' : 'all'
    );
    const useWorker = this.shouldUseWorkerForExport(exportData);

    if (useWorker) {
      return this.exportWithWorker(
        exportData,
        filename,
        selectedOnly ? 'selected' : 'all'
      );
    } else {
      return this.exportWithMainThread(
        exportData,
        filename,
        selectedOnly ? 'selected' : 'all'
      );
    }
  }
}
