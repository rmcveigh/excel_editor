/**
 * @file
 * Excel Editor Export Manager Module
 *
 * Handles export operations in the Excel Editor.
 */

/* eslint-disable no-console */
(function () {
  'use strict';
  window.ExcelEditorExportManager = function () {
    // =========================================================================
    // EXPORTING
    // =========================================================================

    /**
     * Exports only the selected rows to an XLSX file.
     */
    this.exportSelected = async function () {
      if (!this.data?.selected?.size) {
        this.showMessage('No rows selected for export', 'warning');
        return;
      }

      this.showQuickLoader('Preparing export...');
      try {
        const exportData = this.prepareExportData(true);
        const filename = this._generateExportFilename('selected');

        await this.downloadExport(exportData, filename);
        this.showMessage(
          `Exported ${this.data.selected.size} selected rows`,
          'success'
        );
      } catch (error) {
        this.handleError('Failed to export selected rows', error);
      } finally {
        this.hideQuickLoader();
      }
    };

    /**
     * Exports all visible (filtered) rows to an XLSX file.
     */
    this.exportAll = async function () {
      if (!this.data?.filtered?.length || this.data.filtered.length <= 1) {
        this.showMessage('No data available to export', 'warning');
        return;
      }

      const isLargeDataset = this.data.filtered.length > 1000;

      try {
        if (isLargeDataset) {
          this.showProcessLoader('Preparing large export...');
        } else {
          this.showQuickLoader('Preparing export...');
        }

        const exportData = this.prepareExportData(false);
        const filename = this._generateExportFilename('all');

        await this.downloadExport(exportData, filename);

        const rowCount = this.data.filtered.length - 1;
        this.showMessage(`Exported ${rowCount} rows`, 'success');
      } catch (error) {
        this.handleError('Failed to export data', error);
      } finally {
        if (isLargeDataset) {
          this.hideProcessLoader();
        } else {
          this.hideQuickLoader();
        }
      }
    };

    /**
     * Prepares the data for export by creating an array of arrays.
     * @param {boolean} selectedOnly Whether to include only selected rows.
     * @returns {Array<Array<string>>} The data ready for export.
     */
    this.prepareExportData = function (selectedOnly = false) {
      if (!this.data?.filtered?.length) {
        return [[]];
      }

      const exportData = [];
      const headerRow = [];

      // Process headers
      this.data.filtered[0].forEach((header, index) => {
        if (!this.state.hiddenColumns.has(index)) {
          headerRow.push(header);
        }
      });
      exportData.push(headerRow);

      // Process data rows
      for (let i = 1; i < this.data.filtered.length; i++) {
        if (!selectedOnly || this.data.selected.has(i)) {
          const dataRow = [];
          this.data.filtered[i].forEach((cell, index) => {
            if (!this.state.hiddenColumns.has(index)) {
              dataRow.push(cell);
            }
          });
          exportData.push(dataRow);
        }
      }

      return exportData;
    };

    /**
     * Triggers the download of the exported data as an XLSX file.
     * @param {Array<Array<string>>} data The data to export.
     * @param {string} filename The name of the file to download.
     * @returns {Promise<void>} A promise that resolves when the export is complete.
     */
    this.downloadExport = async function (data, filename) {
      try {
        const ws = XLSX.utils.aoa_to_sheet(data);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Export');
        XLSX.writeFile(wb, filename);
      } catch (error) {
        this.logDebug('XLSX export error:', error);
        throw new Error('Failed to generate Excel file');
      }
    };

    /**
     * [HELPER] Generates a filename for exported data with timestamp.
     * @param {string} type Export type identifier (e.g., 'all', 'selected')
     * @returns {string} Formatted filename with timestamp
     */
    this._generateExportFilename = function (type) {
      const timestamp = new Date()
        .toISOString()
        .replace(/[:.]/g, '-')
        .split('T')[0];
      const baseFilename = this.config.exportFilenameBase || 'excel_editor';
      return `${baseFilename}_${type}_${timestamp}.xlsx`;
    };

    /**
     * Exports data to CSV format (alternative to XLSX).
     * @param {boolean} selectedOnly Whether to include only selected rows.
     */
    this.exportToCSV = function (selectedOnly = false) {
      try {
        this.showQuickLoader('Preparing CSV export...');

        const data = this.prepareExportData(selectedOnly);
        if (!data?.length) {
          this.showMessage('No data to export', 'warning');
          this.hideQuickLoader();
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
        const filename = this._generateExportFilename(
          selectedOnly ? 'selected' : 'all'
        ).replace('.xlsx', '.csv');

        link.href = url;
        link.setAttribute('download', filename);
        link.style.visibility = 'hidden';
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);

        const rowCount = selectedOnly
          ? this.data.selected.size
          : this.data.filtered.length - 1;
        this.showMessage(`Exported ${rowCount} rows to CSV`, 'success');
      } catch (error) {
        this.handleError('Failed to export CSV', error);
      } finally {
        this.hideQuickLoader();
      }
    };
  };
})();
