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

    this.app.utilities.showProcessLoader('Preparing export...');

    try {
      const onProgress = (progress, message) => {
        this.updateExportProgress(progress, message);
      };

      const options = {
        sheetName: 'Export',
        columnWidths: this.calculateColumnWidths(exportData),
        styles: this.getExportStyles(),
      };

      const result = await this.app.workerManager.exportExcel(
        exportData,
        filename,
        options,
        onProgress
      );

      if (!result.success) {
        throw new Error(result.error);
      }

      // The worker now returns a Blob directly
      await this.downloadFromBlob(result.data, result.filename);

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
    let progressContainer = $('.excel-editor-overlay-loader .loading-content');

    if (progressContainer.length > 0) {
      let progressBar = progressContainer.find('.progress');
      let progressText = progressContainer.find('p');

      if (progressBar.length === 0) {
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
      const checkRows = Math.min(data.length, 20);
      for (let row = 0; row < checkRows; row++) {
        const cellValue = String(data[row][col] || '');
        maxWidth = Math.max(maxWidth, cellValue.length);
      }
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
   * Download file from a Blob (for worker and main thread exports)
   * @param {Blob} blob - The file data as a Blob.
   * @param {string} filename - The filename.
   */
  async downloadFromBlob(blob, filename) {
    try {
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.setAttribute('download', filename);
      link.style.visibility = 'hidden';
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      setTimeout(() => URL.revokeObjectURL(url), 1000);
    } catch (error) {
      this.app.utilities.logDebug('Blob download error:', error);
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

    this.app.data.filtered[0].forEach((header, index) => {
      if (!this.app.state.hiddenColumns.has(index)) {
        headerRow.push(header);
      }
    });
    exportData.push(headerRow);

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
      // Use the global XLSX object for main-thread operations
      const ws = XLSX.utils.aoa_to_sheet(data);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Export');
      const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      const blob = new Blob([wbout], { type: 'application/octet-stream' });

      await this.downloadFromBlob(blob, filename);
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
}
