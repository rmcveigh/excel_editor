/**
 * @file
 * Excel Editor Data Manager Module - Updated with Web Worker Support
 */

export class ExcelEditorDataManager {
  constructor(app) {
    this.app = app;
  }

  /**
   * Orchestrates file processing: validation, reading, and parsing.
   * @param {File} file - The file to process (CSV or Excel).
   */
  async processFile(file) {
    try {
      if (!this.app.utilities.validateFile(file)) return;

      this.app.utilities.showLoading('Processing Excel file...');

      // Read file as ArrayBuffer for worker compatibility
      const data = await this.app.utilities.readFile(file);
      let parsedData;

      // Determine if we should use worker or fallback
      const useWorker =
        this.app.workerManager && this.app.workerManager.isAvailable();

      if (useWorker) {
        this.app.utilities.logDebug('Using Web Worker for file processing');
        parsedData = await this.processFileWithWorker(file, data);
      } else {
        this.app.utilities.logDebug(
          'Using main thread for file processing (fallback)'
        );
        parsedData = await this.processFileMainThread(file, data);
      }

      this.loadData(parsedData);
      this.app.utilities.hideLoading();

      this.app.utilities.showMessage(
        `Successfully loaded ${this.app.data.original.length - 1} rows from ${
          file.name
        }`,
        'success'
      );
    } catch (error) {
      console.error('Error processing file:', error);
      this.app.utilities.hideLoading();
      this.app.utilities.handleError('Failed to process file', error);
    }
  }

  /**
   * Process file using Web Worker
   * @param {File} file - The file being processed
   * @param {ArrayBuffer} data - The file data
   * @returns {Promise<Array>} - Parsed data
   */
  async processFileWithWorker(file, data) {
    const isCSV = file.name.toLowerCase().endsWith('.csv');

    // Set up progress tracking
    const onProgress = (progress, message) => {
      this.updateLoadingProgress(progress, message);
    };

    try {
      let result;

      if (isCSV) {
        // Convert ArrayBuffer to text for CSV processing
        const textData = new TextDecoder().decode(data);
        result = await this.app.workerManager.parseCSV(textData, onProgress);
      } else {
        result = await this.app.workerManager.parseExcel(data, onProgress);
      }

      if (!result.success) {
        throw new Error(result.error);
      }

      this.app.utilities.logDebug(
        'Worker processing completed:',
        result.metadata
      );
      return result.data;
    } catch (error) {
      this.app.utilities.logDebug(
        'Worker processing failed, falling back to main thread:',
        error
      );
      // Fallback to main thread processing
      return this.processFileMainThread(file, data);
    }
  }

  /**
   * Process file on main thread (fallback)
   * @param {File} file - The file being processed
   * @param {ArrayBuffer} data - The file data
   * @returns {Promise<Array>} - Parsed data
   */
  async processFileMainThread(file, data) {
    const isCSV = file.name.toLowerCase().endsWith('.csv');

    if (isCSV) {
      return this.parseCSV(data);
    } else {
      return await this.parseExcel(data);
    }
  }

  /**
   * Update loading progress display
   * @param {number} progress - Progress percentage (0-100)
   * @param {string} message - Progress message
   */
  updateLoadingProgress(progress, message) {
    const $ = jQuery;

    // Update existing loading message or create progress indicator
    let progressContainer = $('.excel-editor-progress');

    if (progressContainer.length === 0) {
      progressContainer = $(`
        <div class="excel-editor-progress">
          <div class="progress-bar-container">
            <progress class="progress is-primary" value="0" max="100"></progress>
          </div>
          <p class="progress-message">Loading...</p>
        </div>
      `);

      $('.excel-editor-loading').append(progressContainer);
    }

    progressContainer.find('.progress').attr('value', progress);
    progressContainer
      .find('.progress-message')
      .text(message || 'Processing...');
  }

  /**
   * Parses CSV data from an ArrayBuffer (fallback method).
   */
  parseCSV(data) {
    const text = new TextDecoder().decode(data);
    const lines = text.split('\n').filter((line) => line.trim());
    return lines.map((line) => {
      return line.split(',').map((cell) => {
        return cell.trim().replace(/^["']|["']$/g, '');
      });
    });
  }

  /**
   * Parses Excel (.xls, .xlsx) data from an ArrayBuffer using SheetJS (fallback method).
   */
  async parseExcel(data) {
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
  }

  /**
   * Enhanced file processing with batch operations for large datasets
   * @param {Array} data - The parsed data to process
   */
  async processLargeDataset(data) {
    const BATCH_SIZE = 1000;
    const isLarge = data.length > BATCH_SIZE;

    if (!isLarge) {
      return data;
    }

    this.app.utilities.showProcessLoader('Processing large dataset...');

    try {
      // Process data in batches to prevent UI blocking
      const processed = [];

      for (let i = 0; i < data.length; i += BATCH_SIZE) {
        const batch = data.slice(i, i + BATCH_SIZE);

        // Update progress
        const progress = Math.round((i / data.length) * 100);
        this.updateLoadingProgress(
          progress,
          `Processing rows ${i + 1} to ${Math.min(i + BATCH_SIZE, data.length)}`
        );

        // Process batch
        const processedBatch = this.processBatch(batch);
        processed.push(...processedBatch);

        // Yield to browser for UI updates
        await new Promise((resolve) => setTimeout(resolve, 10));
      }

      return processed;
    } finally {
      this.app.utilities.hideProcessLoader();
    }
  }

  /**
   * Process a batch of data rows
   * @param {Array} batch - Batch of rows to process
   * @returns {Array} - Processed batch
   */
  processBatch(batch) {
    return batch.map((row) => {
      if (!Array.isArray(row)) return row;

      return row.map((cell) => {
        // Clean and standardize cell values
        if (cell === null || cell === undefined) {
          return '';
        }

        let cleaned = String(cell).trim();

        // Remove surrounding quotes if present
        if (
          cleaned.length >= 2 &&
          cleaned.startsWith('"') &&
          cleaned.endsWith('"')
        ) {
          cleaned = cleaned.slice(1, -1);
        }

        return cleaned;
      });
    });
  }

  /**
   * Loads the parsed data into the application's state.
   */
  async loadData(data) {
    this.app.utilities.logDebug('Loading data into application...', data);

    if (!data || data.length === 0) {
      throw new Error('No data found in file');
    }

    // Process large datasets efficiently
    const processedData = await this.processLargeDataset(data);

    this.app.data.original = this.app.utilities.deepClone(processedData);
    this.addEditableColumns();
    this.app.data.filtered = this.app.utilities.deepClone(
      this.app.data.original
    );
    this.app.data.selected.clear();
    this.app.data.dirty = false;

    // Reset subject ID link cache for new dataset
    if (this.app.uiRenderer && this.app.uiRenderer.resetSubjectIdLinkCache) {
      this.app.uiRenderer.resetSubjectIdLinkCache();
    }

    this.applyDefaultColumnVisibility();
    await this.app.uiRenderer.renderInterface();
    this.updateSelectionCount();

    // Trigger initial validation after data is loaded and interface is rendered
    setTimeout(() => {
      if (this.app.validationManager) {
        this.app.validationManager.validateExistingBarcodeFields();
      }
    }, 200);
  }

  /**
   * Enhanced method to handle editable columns with better performance
   */
  addEditableColumns() {
    if (!this.app.data.original.length) return;

    const headerRow = this.app.data.original[0];

    // Check if editable columns already exist
    if (
      this.app.config.editableColumns.some((col) => headerRow.includes(col))
    ) {
      return;
    }

    // Use worker for barcode generation if available and dataset is large
    const isLargeDataset = this.app.data.original.length > 500;
    const useWorkerForBarcodes =
      isLargeDataset &&
      this.app.workerManager &&
      this.app.workerManager.isAvailable();

    if (useWorkerForBarcodes) {
      this.addEditableColumnsWithWorker();
    } else {
      this.addEditableColumnsMainThread();
    }
  }

  /**
   * Add editable columns using worker for barcode generation
   */
  async addEditableColumnsWithWorker() {
    // For now, use main thread method as worker barcode generation
    // would require more complex implementation
    this.addEditableColumnsMainThread();
  }

  /**
   * Add editable columns on main thread
   */
  addEditableColumnsMainThread() {
    const headerRow = this.app.data.original[0];

    // Detect file type and find column indices BEFORE modifying the header
    const isTissueResearchFile = this.detectTissueResearchFile(headerRow);

    // Find column indices in the original header row
    let columnIndices = {};
    if (isTissueResearchFile) {
      columnIndices = {
        subjectId: this.app.barcodeSystem.findTissueResearchColumn(
          headerRow,
          'subject_id'
        ),
        biopsyType: this.app.barcodeSystem.findTissueResearchColumn(
          headerRow,
          'biopsy_necropsy'
        ),
        reqTissueType: this.app.barcodeSystem.findTissueResearchColumn(
          headerRow,
          'req_tissue_type'
        ),
        vialTissueType: this.app.barcodeSystem.findTissueResearchColumn(
          headerRow,
          'vial_tissue_type'
        ),
        healthStatus: this.app.barcodeSystem.findTissueResearchColumn(
          headerRow,
          'health_status'
        ),
      };
      this.app.utilities.logDebug(
        'Original tissue research column indices:',
        columnIndices
      );
    } else {
      columnIndices = {
        subjectId: this.app.barcodeSystem.findColumnByType(
          headerRow,
          'subject_id'
        ),
        healthStatus: this.app.barcodeSystem.findColumnByType(
          headerRow,
          'health_status'
        ),
      };
      this.app.utilities.logDebug(
        'Original generic column indices:',
        columnIndices
      );
    }

    // NOW modify the header row
    headerRow.unshift('new_barcode');
    headerRow.push('notes', 'actions');

    // Populate the data rows using the ORIGINAL indices
    if (isTissueResearchFile) {
      this.populateTissueResearchBarcodesWithIndices(columnIndices);
    } else {
      this.populateGenericBarcodesWithIndices(columnIndices);
    }

    this.app.data.dirty = true;
  }

  /**
   * Populates barcodes using the tissue research formatter with pre-calculated indices
   */
  populateTissueResearchBarcodesWithIndices(columnIndices) {
    let populatedCount = 0;
    let errorCount = 0;

    for (let i = 1; i < this.app.data.original.length; i++) {
      const row = this.app.data.original[i];

      try {
        let barcodeValue = '';
        if (columnIndices.subjectId !== -1 && row[columnIndices.subjectId]) {
          barcodeValue = this.app.barcodeSystem.formatTissueResearchBarcode(
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

        row.unshift(barcodeValue);
        row.push('', '');
      } catch (error) {
        this.app.utilities.logDebug(
          `Error generating barcode for row ${i}:`,
          error
        );
        row.unshift('');
        row.push('', '');
        errorCount++;
      }
    }

    const message = `Auto-populated ${populatedCount} tissue research barcodes`;
    this.app.utilities.logDebug(
      message + (errorCount > 0 ? ` (${errorCount} errors)` : '')
    );
    this.app.utilities.showMessage(message, 'success', 4000);
  }

  /**
   * Populates barcodes using the generic formatter with pre-calculated indices
   */
  populateGenericBarcodesWithIndices(columnIndices) {
    let populatedCount = 0;

    for (let i = 1; i < this.app.data.original.length; i++) {
      const row = this.app.data.original[i];

      let barcodeValue = '';
      if (columnIndices.subjectId !== -1 && row[columnIndices.subjectId]) {
        const healthValue =
          columnIndices.healthStatus !== -1
            ? row[columnIndices.healthStatus]
            : null;
        barcodeValue = this.app.barcodeSystem.formatBarcode(
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

      row.unshift(barcodeValue);
      row.push('', '');
    }

    if (columnIndices.subjectId !== -1 && populatedCount > 0) {
      let message = `Auto-populated ${populatedCount} generic barcodes from source data`;
      if (columnIndices.healthStatus !== -1) {
        message += ' with context suffixes';
      }
      this.app.utilities.logDebug(message);
      this.app.utilities.showMessage(message, 'success', 4000);
    } else {
      this.app.utilities.logDebug(
        'No source column found - new_barcode column added empty'
      );
    }
  }

  /**
   * Detects if this is a tissue research file by checking for required columns
   */
  detectTissueResearchFile(headerRow) {
    const requiredColumns = [
      'subject_id',
      'biopsy_necropsy',
      'req_tissue_type',
      'vial_tissue_type',
      'health_status',
    ];

    const foundColumns = requiredColumns.filter((colType) => {
      const index = this.app.barcodeSystem.findTissueResearchColumn
        ? this.app.barcodeSystem.findTissueResearchColumn(headerRow, colType)
        : this.app.barcodeSystem.findColumnByType(headerRow, colType);
      return index !== -1;
    });

    const isDetected = foundColumns.length >= 4;

    this.app.utilities.logDebug(
      `Tissue research file detection: ${isDetected ? 'YES' : 'NO'}`,
      {
        foundColumns,
        totalFound: foundColumns.length,
        required: requiredColumns.length,
      }
    );

    return isDetected;
  }

  /**
   * Applies default column visibility based on configuration.
   */
  applyDefaultColumnVisibility() {
    const { settings } = this.app.config;
    if (
      settings.hideBehavior !== 'hide_others' ||
      !settings.defaultVisibleColumns?.length
    ) {
      return;
    }

    const defaultColumns = settings.defaultVisibleColumns.map((col) =>
      col.trim().toLowerCase()
    );
    const alwaysVisible = this.app.config.editableColumns.map((col) =>
      col.toLowerCase()
    );
    const maxColumns = settings.maxVisibleColumns || 50;
    const headerRow = this.app.data.filtered[0];

    const matchedColumns = headerRow.filter((header) =>
      defaultColumns.includes(String(header).trim().toLowerCase())
    ).length;

    if (matchedColumns === 0) {
      this.app.utilities.logDebug(
        'No configured default columns found - showing all columns'
      );
      this.app.utilities.showMessage(
        'No configured default columns found in this file. Showing all columns.',
        'info',
        5000
      );
      return;
    }

    this.app.state.hiddenColumns.clear();
    let visibleCount = 0;

    headerRow.forEach((header, index) => {
      const trimmedHeader = String(header).trim().toLowerCase();
      const shouldBeVisible =
        defaultColumns.includes(trimmedHeader) ||
        alwaysVisible.includes(trimmedHeader);
      if (!shouldBeVisible) {
        this.app.state.hiddenColumns.add(index);
      } else if (visibleCount < maxColumns) {
        visibleCount++;
      } else {
        this.app.state.hiddenColumns.add(index);
      }
    });

    this.app.utilities.logDebug(
      `Applied default column visibility. Matched ${matchedColumns} configured columns.`
    );
  }

  /**
   * Process loaded data (for draft loading)
   */
  processLoadedData() {
    this.app.data.filtered = this.app.utilities.deepClone(
      this.app.data.original
    );
    this.app.uiRenderer.renderInterface();
    this.updateSelectionCount();
  }

  /**
   * Gets the value from the 'actions' column for a specific row.
   * @param {number} rowIndex - The index of the row to check.
   * @return {string|null} The action value for the row, or null if not found.
   */
  getActionValue(rowIndex) {
    const actionsColumnIndex = this.app.data.filtered[0].indexOf('actions');
    if (actionsColumnIndex === -1) return null;
    return this.app.data.filtered[rowIndex][actionsColumnIndex];
  }

  /**
   * Updates the text indicating how many rows are selected.
   */
  updateSelectionCount() {
    const count = this.app.data.selected.size;
    this.app.elements.selectionCount.text(
      `${count} row${count !== 1 ? 's' : ''} selected`
    );
    this.app.elements.exportBtn
      .prop('disabled', count === 0)
      .toggleClass('is-disabled', count === 0);
  }

  /**
   * Updates the state of the "select all" checkbox.
   */
  updateSelectAllCheckbox() {
    const totalRows = this.app.data.filtered.length - 1;
    const selectedRows = this.app.data.selected.size;
    const $selectAllCheckbox = jQuery('#select-all-checkbox');

    if (selectedRows === 0) {
      $selectAllCheckbox.prop({ checked: false, indeterminate: false });
    } else if (selectedRows === totalRows && totalRows > 0) {
      $selectAllCheckbox.prop({ checked: true, indeterminate: false });
    } else {
      $selectAllCheckbox.prop({ checked: false, indeterminate: true });
    }
  }

  /**
   * Applies CSS classes to rows based on the value in the 'actions' column.
   */
  applyRowStyling() {
    if (!this.app.elements.table || !this.app.elements.table.length) return;

    this.app.elements.table.find('tbody tr').each((index, row) => {
      const $ = jQuery;
      const $row = $(row);
      const rowIndex = parseInt($row.data('row'));
      const actionValue = this.getActionValue(rowIndex);
      $row.removeClass('action-relabel action-pending action-discard');
      if (actionValue) $row.addClass(`action-${actionValue}`);
    });
  }
}
