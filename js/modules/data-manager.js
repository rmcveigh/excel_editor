/**
 * @file
 * Excel Editor Data Manager Module
 *
 * Handles file processing, parsing, data loading, and data management.
 * This module adds data management methods to the ExcelEditor class.
 */

/* eslint-disable no-console */
(function ($) {
  'use strict';

  /**
   * Data Manager module for Excel Editor.
   * This function is called on ExcelEditor instances to add data management methods.
   */
  window.ExcelEditorDataManager = function () {
    // =========================================================================
    // FILE PROCESSING & PARSING
    // =========================================================================

    /**
     * Orchestrates file processing: validation, reading, and parsing.
     * @param {File} file The file to process.
     */
    this.processFile = async function (file) {
      try {
        if (!this.validateFile(file)) return;

        this.showLoading('Processing Excel file...');
        const data = await this.readFile(file);
        let parsedData;

        if (file.name.toLowerCase().endsWith('.csv')) {
          parsedData = this.parseCSV(data);
        } else {
          parsedData = await this.parseExcel(data);
        }

        this.loadData(parsedData);
        this.hideLoading();

        this.showMessage(
          `Successfully loaded ${this.data.original.length - 1} rows from ${
            file.name
          }`,
          'success'
        );
      } catch (error) {
        console.error('Error processing file:', error);
        this.hideLoading();
        this.handleError('Failed to process file', error);
      }
    };

    /**
     * Parses CSV data from an ArrayBuffer.
     * @param {ArrayBuffer} data The raw file data.
     * @returns {Array<Array<string>>} The parsed data.
     */
    this.parseCSV = function (data) {
      const text = new TextDecoder().decode(data);
      const lines = text.split('\n').filter((line) => line.trim());
      return lines.map((line) => {
        return line.split(',').map((cell) => {
          return cell.trim().replace(/^["']|["']$/g, '');
        });
      });
    };

    /**
     * Parses Excel (.xls, .xlsx) data from an ArrayBuffer using SheetJS.
     * @param {ArrayBuffer} data The raw file data.
     * @returns {Promise<Array<Array<string>>>} The parsed data.
     */
    this.parseExcel = async function (data) {
      return new Promise((resolve, reject) => {
        try {
          const workbook = XLSX.read(data, { type: 'array' });
          if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
            return reject(new Error('No worksheets found in Excel file.'));
          }

          const sheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[sheetName];
          const jsonData = XLSX.utils.sheet_to_json(worksheet, {
            header: 1,
            raw: false,
            defval: '',
          });

          const trimmedData = jsonData.map((row) =>
            Array.isArray(row)
              ? row.map((cell) => String(cell || '').trim())
              : row
          );

          const filteredData = trimmedData.filter(
            (row) =>
              Array.isArray(row) &&
              row.some(
                (cell) =>
                  cell !== null &&
                  cell !== undefined &&
                  String(cell).trim() !== ''
              )
          );

          if (filteredData.length <= 1) {
            return reject(new Error('Excel file contains no data rows.'));
          }

          resolve(filteredData);
        } catch (error) {
          reject(
            new Error(
              'Failed to parse Excel file. It might be corrupted or in an unsupported format.'
            )
          );
        }
      });
    };

    // =========================================================================
    // DATA LOADING & MANAGEMENT
    // =========================================================================

    /**
     * Loads the parsed data into the application's state.
     * @param {Array<Array<string>>} data The parsed data from a file.
     */
    this.loadData = function (data) {
      this.logDebug('Loading data into application...', data);

      if (!data || data.length === 0) {
        throw new Error('No data found in file');
      }

      const trimmedData = data.map((row) => {
        if (!Array.isArray(row)) return row;
        return row.map((cell) => String(cell || '').trim());
      });

      this.data.original = this.deepClone(trimmedData);
      this.addEditableColumns();
      this.data.filtered = this.deepClone(this.data.original);
      this.data.selected.clear();
      this.data.dirty = false;

      this.applyDefaultColumnVisibility();
      this.renderInterface();
      this.updateSelectionCount();
    };

    /**
     * Adds the default editable columns ('new_barcode', 'notes', 'actions')
     * to the dataset if they don't already exist.
     * Auto-populates new_barcode with values using either tissue research or generic formatter.
     */
    this.addEditableColumns = function () {
      if (!this.data.original.length) return;

      const headerRow = this.data.original[0];

      // Check if editable columns already exist
      if (this.config.editableColumns.some((col) => headerRow.includes(col))) {
        return;
      }

      // Detect file type and find column indices BEFORE modifying the header
      const isTissueResearchFile = this.detectTissueResearchFile(headerRow);

      // Find column indices in the original header row (DON'T adjust these)
      let columnIndices = {};
      if (isTissueResearchFile) {
        columnIndices = {
          subjectId: this.findTissueResearchColumn(headerRow, 'subject_id'),
          biopsyType: this.findTissueResearchColumn(
            headerRow,
            'biopsy_necropsy'
          ),
          reqTissueType: this.findTissueResearchColumn(
            headerRow,
            'req_tissue_type'
          ),
          vialTissueType: this.findTissueResearchColumn(
            headerRow,
            'vial_tissue_type'
          ),
          healthStatus: this.findTissueResearchColumn(
            headerRow,
            'health_status'
          ),
        };
        this.logDebug(
          'Original tissue research column indices:',
          columnIndices
        );
      } else {
        columnIndices = {
          subjectId: this.findColumnByType(headerRow, 'subject_id'),
          healthStatus: this.findColumnByType(headerRow, 'health_status'),
        };
        this.logDebug('Original generic column indices:', columnIndices);
      }

      // NOW modify the header row
      headerRow.unshift('new_barcode');
      headerRow.push('notes', 'actions');

      // Populate the data rows using the ORIGINAL indices (data rows haven't been modified yet)
      if (isTissueResearchFile) {
        this.populateTissueResearchBarcodesWithIndices(columnIndices);
      } else {
        this.populateGenericBarcodesWithIndices(columnIndices);
      }

      this.data.dirty = true;
    };

    /**
     * Populates barcodes using the tissue research formatter with pre-calculated indices
     * @param {Object} columnIndices - Pre-calculated column indices (from ORIGINAL header)
     */
    this.populateTissueResearchBarcodesWithIndices = function (columnIndices) {
      let populatedCount = 0;
      let errorCount = 0;

      for (let i = 1; i < this.data.original.length; i++) {
        const row = this.data.original[i];

        try {
          // Generate barcode value using tissue research formatter
          // Use ORIGINAL indices since row hasn't been modified yet
          let barcodeValue = '';
          if (columnIndices.subjectId !== -1 && row[columnIndices.subjectId]) {
            barcodeValue = this.formatTissueResearchBarcode(
              row[columnIndices.subjectId],
              columnIndices.biopsyType !== -1
                ? row[columnIndices.biopsyType]
                : '',
              columnIndices.reqTissueType !== -1
                ? row[columnIndices.reqTissueType]
                : '',
              columnIndices.vialTissueType !== -1
                ? row[columnIndices.vialTissueType]
                : '',
              columnIndices.healthStatus !== -1
                ? row[columnIndices.healthStatus]
                : ''
            );
            populatedCount++;
          }

          // Add new columns: new_barcode (with auto-populated value), notes (empty), actions (empty)
          row.unshift(barcodeValue);
          row.push('', '');
        } catch (error) {
          this.logDebug(`Error generating barcode for row ${i}:`, error);
          // Still add the columns even if barcode generation failed
          row.unshift('');
          row.push('', '');
          errorCount++;
        }
      }

      const message = `Auto-populated ${populatedCount} tissue research barcodes`;
      this.logDebug(
        message + (errorCount > 0 ? ` (${errorCount} errors)` : '')
      );
      this.showMessage(message, 'success', 4000);
    };

    /**
     * Populates barcodes using the generic formatter with pre-calculated indices
     * @param {Object} columnIndices - Pre-calculated column indices (from ORIGINAL header)
     */
    this.populateGenericBarcodesWithIndices = function (columnIndices) {
      let populatedCount = 0;

      for (let i = 1; i < this.data.original.length; i++) {
        const row = this.data.original[i];

        // Generate barcode value using generic formatter
        // Use ORIGINAL indices since row hasn't been modified yet
        let barcodeValue = '';
        if (columnIndices.subjectId !== -1 && row[columnIndices.subjectId]) {
          const healthValue =
            columnIndices.healthStatus !== -1
              ? row[columnIndices.healthStatus]
              : null;
          barcodeValue = this.formatBarcode(
            row[columnIndices.subjectId],
            {
              removeDashes: true,
              removeSpaces: true,
              removeUnderscores: true,
              removeDots: true,
              toUpperCase: true,
              includeContext: true,
            },
            healthValue,
            'health'
          );
          populatedCount++;
        }

        // Add new columns: new_barcode (with auto-populated value), notes (empty), actions (empty)
        row.unshift(barcodeValue);
        row.push('', '');
      }

      // Log what happened for debugging
      if (columnIndices.subjectId !== -1 && populatedCount > 0) {
        let message = `Auto-populated ${populatedCount} generic barcodes from source data`;
        if (columnIndices.healthStatus !== -1) {
          message += ' with context suffixes';
        }
        this.logDebug(
          message +
            ` (Source: original column ${
              columnIndices.subjectId
            }, Context: original column ${
              columnIndices.healthStatus !== -1
                ? columnIndices.healthStatus
                : 'none'
            })`
        );
        this.showMessage(message, 'success', 4000);
      } else {
        this.logDebug(
          'No source column found - new_barcode column added empty'
        );
      }
    };

    /**
     * Detects if this is a tissue research file by checking for required columns
     * @param {Array} headerRow - The header row array
     * @returns {boolean} True if tissue research file detected
     */
    this.detectTissueResearchFile = function (headerRow) {
      const requiredColumns = [
        'subject_id',
        'biopsy_necropsy',
        'req_tissue_type',
        'vial_tissue_type',
        'health_status',
      ];

      const foundColumns = requiredColumns.filter((colType) => {
        const index = this.findTissueResearchColumn
          ? this.findTissueResearchColumn(headerRow, colType)
          : this.findColumnByType(headerRow, colType);
        return index !== -1;
      });

      const isDetected = foundColumns.length >= 4; // Need at least 4 of 5 columns

      this.logDebug(
        `Tissue research file detection: ${isDetected ? 'YES' : 'NO'}`,
        {
          foundColumns,
          totalFound: foundColumns.length,
          required: requiredColumns.length,
        }
      );

      return isDetected;
    };

    /**
     * Populates barcodes using the generic formatter with pre-calculated indices
     * @param {Object} columnIndices - Pre-calculated column indices
     */
    this.populateGenericBarcodesWithIndices = function (columnIndices) {
      let populatedCount = 0;

      for (let i = 1; i < this.data.original.length; i++) {
        const row = this.data.original[i];

        // Generate barcode value using generic formatter
        let barcodeValue = '';
        if (columnIndices.subjectId !== -1 && row[columnIndices.subjectId]) {
          const healthValue =
            columnIndices.healthStatus !== -1
              ? row[columnIndices.healthStatus]
              : null;
          barcodeValue = this.formatBarcode(
            row[columnIndices.subjectId],
            {
              removeDashes: true,
              removeSpaces: true,
              removeUnderscores: true,
              removeDots: true,
              toUpperCase: true,
              includeContext: true,
            },
            healthValue,
            'health'
          );
          populatedCount++;
        }

        // Add new columns: new_barcode (with auto-populated value), notes (empty), actions (empty)
        row.unshift(barcodeValue);
        row.push('', '');
      }

      // Log what happened for debugging
      if (columnIndices.subjectId !== -1 && populatedCount > 0) {
        let message = `Auto-populated ${populatedCount} generic barcodes from source data`;
        if (columnIndices.healthStatus !== -1) {
          message += ' with context suffixes';
        }
        this.logDebug(
          message +
            ` (Source: column ${columnIndices.subjectId - 1}, Context: column ${
              columnIndices.healthStatus !== -1
                ? columnIndices.healthStatus - 1
                : 'none'
            })`
        );
        this.showMessage(message, 'success', 4000);
      } else {
        this.logDebug(
          'No source column found - new_barcode column added empty'
        );
      }
    };

    /**
     * Applies default column visibility based on configuration with intelligent fallbacks.
     * This is called when data is first loaded.
     */
    this.applyDefaultColumnVisibility = function () {
      const { settings } = this.config;
      if (
        settings.hideBehavior !== 'hide_others' ||
        !settings.defaultVisibleColumns?.length
      ) {
        return;
      }

      const defaultColumns = settings.defaultVisibleColumns.map((col) =>
        col.trim().toLowerCase()
      );
      const alwaysVisible = this.config.editableColumns.map((col) =>
        col.toLowerCase()
      );
      const maxColumns = settings.maxVisibleColumns || 50;
      const headerRow = this.data.filtered[0];

      // CHECK: Count how many default columns actually exist
      const matchedColumns = headerRow.filter((header) =>
        defaultColumns.includes(String(header).trim().toLowerCase())
      ).length;

      // If NO defaults found, don't hide anything (show all columns)
      if (matchedColumns === 0) {
        this.logDebug(
          'No configured default columns found - showing all columns'
        );
        this.showMessage(
          'No configured default columns found in this file. Showing all columns.',
          'info',
          5000
        );
        return; // Exit early, leave hiddenColumns empty
      }

      // Rest of existing logic when default columns ARE found
      this.state.hiddenColumns.clear();
      let visibleCount = 0;

      headerRow.forEach((header, index) => {
        const trimmedHeader = String(header).trim().toLowerCase();
        const shouldBeVisible =
          defaultColumns.includes(trimmedHeader) ||
          alwaysVisible.includes(trimmedHeader);
        if (!shouldBeVisible) {
          this.state.hiddenColumns.add(index);
        } else if (visibleCount < maxColumns) {
          visibleCount++;
        } else {
          this.state.hiddenColumns.add(index);
        }
      });

      this.logDebug(
        `Applied default column visibility. Matched ${matchedColumns} configured columns.`
      );
    };

    // =========================================================================
    // DATA UTILITIES
    // =========================================================================

    /**
     * Gets the value from the 'actions' column for a specific row.
     * @param {number} rowIndex The row index.
     * @returns {string|null} The action value.
     */
    this.getActionValue = function (rowIndex) {
      const actionsColumnIndex = this.data.filtered[0].indexOf('actions');
      if (actionsColumnIndex === -1) return null;
      return this.data.filtered[rowIndex][actionsColumnIndex];
    };

    /**
     * Updates the text indicating how many rows are selected.
     */
    this.updateSelectionCount = function () {
      const count = this.data.selected.size;
      this.elements.selectionCount.text(
        `${count} row${count !== 1 ? 's' : ''} selected`
      );
      this.elements.exportBtn
        .prop('disabled', count === 0)
        .toggleClass('is-disabled', count === 0);
    };

    /**
     * Updates the state of the "select all" checkbox.
     */
    this.updateSelectAllCheckbox = function () {
      const totalRows = this.data.filtered.length - 1;
      const selectedRows = this.data.selected.size;
      const $selectAllCheckbox = $('#select-all-checkbox');

      if (selectedRows === 0) {
        $selectAllCheckbox.prop({ checked: false, indeterminate: false });
      } else if (selectedRows === totalRows && totalRows > 0) {
        $selectAllCheckbox.prop({ checked: true, indeterminate: false });
      } else {
        $selectAllCheckbox.prop({ checked: false, indeterminate: true });
      }
    };

    /**
     * Applies CSS classes to rows based on the value in the 'actions' column.
     */
    this.applyRowStyling = function () {
      if (!this.elements.table || !this.elements.table.length) return;

      this.elements.table.find('tbody tr').each((index, row) => {
        const $row = $(row);
        const rowIndex = parseInt($row.data('row'));
        const actionValue = this.getActionValue(rowIndex);
        $row.removeClass('action-relabel action-pending action-discard');
        if (actionValue) $row.addClass(`action-${actionValue}`);
      });
    };

    // =========================================================================
    // DATA VALIDATION & PROCESSING HELPERS
    // =========================================================================

    /**
     * Validates data structure and content quality.
     * @param {Array<Array<string>>} data The data to validate.
     * @returns {object} Validation results with warnings and errors.
     */
    this.validateDataQuality = function (data) {
      const results = {
        isValid: true,
        warnings: [],
        errors: [],
        stats: {
          totalRows: data.length - 1,
          totalColumns: data[0]?.length || 0,
          emptyRows: 0,
          emptyColumns: 0,
        },
      };

      if (!data || data.length === 0) {
        results.isValid = false;
        results.errors.push('No data provided');
        return results;
      }

      if (data.length <= 1) {
        results.isValid = false;
        results.errors.push('File contains only headers, no data rows');
        return results;
      }

      // Check for empty rows
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (
          !Array.isArray(row) ||
          row.every((cell) => !cell || String(cell).trim() === '')
        ) {
          results.stats.emptyRows++;
        }
      }

      // Check for empty columns
      if (data[0]) {
        for (let col = 0; col < data[0].length; col++) {
          let hasData = false;
          for (let row = 1; row < data.length; row++) {
            if (data[row][col] && String(data[row][col]).trim() !== '') {
              hasData = true;
              break;
            }
          }
          if (!hasData) {
            results.stats.emptyColumns++;
          }
        }
      }

      // Add warnings for quality issues
      if (results.stats.emptyRows > 0) {
        results.warnings.push(`Found ${results.stats.emptyRows} empty rows`);
      }
      if (results.stats.emptyColumns > 0) {
        results.warnings.push(
          `Found ${results.stats.emptyColumns} empty columns`
        );
      }

      return results;
    };

    /**
     * Cleans and normalizes data after parsing.
     * @param {Array<Array<string>>} data The raw parsed data.
     * @returns {Array<Array<string>>} The cleaned data.
     */
    this.cleanData = function (data) {
      if (!data || !Array.isArray(data)) return [];

      // Remove completely empty rows
      const cleaned = data.filter((row, index) => {
        if (index === 0) return true; // Always keep header
        return (
          Array.isArray(row) &&
          row.some(
            (cell) =>
              cell !== null && cell !== undefined && String(cell).trim() !== ''
          )
        );
      });

      // Trim all cell values and normalize
      return cleaned.map((row) => {
        if (!Array.isArray(row)) return row;
        return row.map((cell) => {
          if (cell === null || cell === undefined) return '';
          return String(cell).trim();
        });
      });
    };

    /**
     * Prepares data summary for display to user.
     * @param {Array<Array<string>>} data The data to summarize.
     * @returns {object} Data summary information.
     */
    this.getDataSummary = function (data) {
      if (!data || data.length === 0) {
        return { rows: 0, columns: 0, headers: [] };
      }

      return {
        rows: data.length - 1, // Exclude header
        columns: data[0]?.length || 0,
        headers: data[0] || [],
        hasEditableColumns: this.config.editableColumns.some((col) =>
          data[0]?.includes(col)
        ),
        detectedSourceColumn:
          this.findColumnByType(data[0] || [], 'subject_id') !== -1,
        detectedHealthColumn:
          this.findColumnByType(data[0] || [], 'health_status') !== -1,
      };
    };

    this.logDebug('ExcelEditorDataManager module loaded');
  };
})(jQuery);
