/**
 * @file
 * Excel Editor Export Manager Module
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

    this.app.utilities.showQuickLoader('Preparing export...');
    try {
      const exportData = this.prepareExportData(true);
      const filename = this.generateExportFilename('selected');

      await this.downloadExport(exportData, filename);
      this.app.utilities.showMessage(
        `Exported ${this.app.data.selected.size} selected rows`,
        'success'
      );
    } catch (error) {
      this.app.utilities.handleError('Failed to export selected rows', error);
    } finally {
      this.app.utilities.hideQuickLoader();
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

    const isLargeDataset = this.app.data.filtered.length > 1000;

    try {
      if (isLargeDataset) {
        this.app.utilities.showProcessLoader('Preparing large export...');
      } else {
        this.app.utilities.showQuickLoader('Preparing export...');
      }

      const exportData = this.prepareExportData(false);
      const filename = this.generateExportFilename('all');

      await this.downloadExport(exportData, filename);

      const rowCount = this.app.data.filtered.length - 1;
      this.app.utilities.showMessage(`Exported ${rowCount} rows`, 'success');
    } catch (error) {
      this.app.utilities.handleError('Failed to export data', error);
    } finally {
      if (isLargeDataset) {
        this.app.utilities.hideProcessLoader();
      } else {
        this.app.utilities.hideQuickLoader();
      }
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
   * Triggers the download of the exported data as an XLSX file.
   * @param {Array<Array>} data - The data to export.
   * @param {string} filename - The name of the file to download.
   * @returns {Promise<void>} - Resolves when the file is downloaded.
   */
  async downloadExport(data, filename) {
    try {
      const ws = XLSX.utils.aoa_to_sheet(data);
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
  exportToCSV(selectedOnly = false) {
    try {
      this.app.utilities.showQuickLoader('Preparing CSV export...');

      const data = this.prepareExportData(selectedOnly);
      if (!data?.length) {
        this.app.utilities.showMessage('No data to export', 'warning');
        this.app.utilities.hideQuickLoader();
        return;
      }

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
        this.exportToCSV(options.selectedOnly);
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
   * Shows export options modal.
   * @returns {Promise<{selectedOnly: boolean, format: string}|null>} - Resolves with selected options or null if canceled.
   */
  showExportOptionsModal() {
    return new Promise((resolve) => {
      const $ = jQuery;
      const selectedCount = this.app.data.selected.size;
      const totalCount = this.app.data.filtered.length - 1;

      const modalHtml = `
      <div class="modal is-active" id="export-options-modal">
        <div class="modal-background"></div>
        <div class="modal-content">
          <div class="box">
            <h3 class="title is-4">
              <span class="icon"><i class="fas fa-download"></i></span>
              Export Options
            </h3>

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
                  Excel (.xlsx) - Preserves formatting
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
                <button class="button is-primary" id="confirm-export">Export</button>
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
    };
  }
}
