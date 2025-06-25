/**
 * @file
 * Excel Editor JavaScript - Enhanced version with better architecture and filter fixes
 */

(function ($, Drupal, once, drupalSettings) {
  'use strict';

  /**
   * Excel Editor Application Class
   */
  class ExcelEditor {
    constructor() {
      this.data = {
        original: [],
        filtered: [],
        selected: new Set(),
        dirty: false,
      };

      this.state = {
        hiddenColumns: new Set(),
        currentFilters: {},
        isInitialized: false,
        isLoading: false,
        currentProcessLoader: null,
      };

      this.config = {
        endpoints: drupalSettings?.excelEditor?.endpoints || {},
        settings: drupalSettings?.excelEditor?.settings || {},
        editableColumns: ['new_barcode', 'notes', 'actions'],
        maxFileSize: 10 * 1024 * 1024, // 10MB
        supportedFormats: ['.xlsx', '.xls', '.csv'],
      };

      // Get CSRF token for API calls
      this.csrfToken = null;
      this.getCsrfToken();

      // Debug config only if debug mode is enabled
      this.logDebug('Excel Editor config loaded:', this.config);

      this.elements = {};

      this.init();
    }

    /**
     * Initialize the application
     */
    init() {
      try {
        // Check if required libraries are available
        this.checkDependencies();

        this.cacheElements();
        this.hideLoading(); // Ensure loader is hidden on init
        this.bindEvents();
        this.loadDrafts();
        this.state.isInitialized = true;
        this.logDebug('Excel Editor initialized successfully');
      } catch (error) {
        this.handleError('Failed to initialize Excel Editor', error);
      }
    }

    /**
     * Get CSRF token for API calls
     */
    async getCsrfToken() {
      try {
        const response = await fetch('/session/token');
        if (response.ok) {
          this.csrfToken = await response.text();
          this.logDebug('CSRF token obtained:', this.csrfToken ? 'Yes' : 'No');
        } else {
          console.warn('Failed to get CSRF token:', response.status);
        }
      } catch (error) {
        console.warn('Error getting CSRF token:', error);
      }
    }

    /**
     * Check if required dependencies are loaded
     */
    checkDependencies() {
      const missing = [];

      if (typeof XLSX === 'undefined') {
        missing.push('XLSX (SheetJS)');
      }

      if (typeof jQuery === 'undefined') {
        missing.push('jQuery');
      }

      if (typeof Drupal === 'undefined') {
        missing.push('Drupal');
      }

      if (missing.length > 0) {
        throw new Error(`Missing required libraries: ${missing.join(', ')}`);
      }

      this.logDebug('All dependencies loaded successfully');
    }

    /**
     * Cache DOM elements for better performance
     */
    cacheElements() {
      this.elements = {
        container: $('.excel-editor-container'),
        uploadArea: $('#excel-upload-area'),
        fileInput: $('#excel-file-input'),
        loadingArea: $('.excel-editor-loading'),
        mainArea: $('#excel-editor-main'),
        table: $('#excel-table'),
        tableContainer: $('.excel-editor-table-container'),
        filtersContainer: $('#filter-controls'),
        draftsContainer: $('#drafts-list'),
        selectionCount: $('#selection-count'),

        // Buttons
        saveDraftBtn: $('#save-draft-btn'),
        exportBtn: $('#export-btn'),
        exportAllBtn: $('#export-all-btn'),
        toggleColumnsBtn: $('#toggle-columns-btn'),
        selectAllBtn: $('#select-all-visible-btn'),
        deselectAllBtn: $('#deselect-all-btn'),
      };

      // Debug element caching (only if debug mode is on)
      this.logDebug('Cached elements:', this.elements);

      // Ensure upload area is visible initially and main area is hidden until data is loaded
      if (this.elements.uploadArea.length > 0) {
        this.elements.uploadArea.show();
      }

      if (this.elements.mainArea.length > 0) {
        this.elements.mainArea.hide(); // Hide until data is loaded
      }

      // Fallback: if table container doesn't exist, create it
      if (this.elements.table.length === 0) {
        this.logDebug(
          'Table element not found, looking for table container...'
        );
        const tableContainer = $('.excel-editor-table-container');
        if (tableContainer.length > 0) {
          this.logDebug('Found table container, creating table element');
          tableContainer.html(
            '<table id="excel-table" class="excel-editor-table table is-fullwidth is-striped"></table>'
          );
          this.elements.table = $('#excel-table');
        } else {
          console.error('No table container found! Check your template.');
        }
      }
    }

    /**
     * Bind all event handlers
     */
    bindEvents() {
      // File upload events
      this.elements.fileInput.on('change', (e) => this.handleFileUpload(e));
      this.setupDragDropUpload();

      // Button events
      this.elements.saveDraftBtn.on('click', () => this.saveDraft());
      this.elements.exportBtn.on('click', () => this.exportSelected());
      this.elements.exportAllBtn.on('click', () => this.exportAll());
      this.elements.toggleColumnsBtn.on('click', () =>
        this.showColumnVisibilityModal()
      );
      this.elements.selectAllBtn.on('click', () => this.selectAllVisible());
      this.elements.deselectAllBtn.on('click', () => this.deselectAll());

      // Table events - using enhanced binding method
      this.bindTableEvents();

      // Keyboard shortcuts
      $(document).on('keydown', (e) => this.handleKeyboardShortcuts(e));

      // Window events
      $(window).on('beforeunload', () => this.handleBeforeUnload());
    }

    /**
     * Enhanced table event binding with multiple strategies and debugging
     */
    bindTableEvents() {
      console.log('[Excel Editor] bindTableEvents called');

      // Remove any existing events to prevent conflicts
      this.elements.tableContainer.off('.excelEditor');

      // Bind with detailed logging
      this.elements.tableContainer.on(
        'click.excelEditor',
        '.filter-link',
        (e) => {
          console.log(
            '[Excel Editor] Click event triggered on table container!',
            {
              target: e.target,
              currentTarget: e.currentTarget,
              type: e.type,
              timeStamp: e.timeStamp,
            }
          );

          e.preventDefault();
          e.stopPropagation(); // Prevent any other handlers from interfering

          console.log('[Excel Editor] About to call handleFilterClick...');

          try {
            this.handleFilterClick(e);
            console.log(
              '[Excel Editor] handleFilterClick completed successfully'
            );
          } catch (error) {
            console.error('[Excel Editor] Error in handleFilterClick:', error);
            alert('Error opening filter: ' + error.message);
          }
        }
      );

      // Additional event binding strategies as fallback
      $(document)
        .off('click.excelEditorGlobal')
        .on(
          'click.excelEditorGlobal',
          '.excel-editor-table .filter-link',
          (e) => {
            console.log(
              '[Excel Editor] Global document click handler triggered'
            );
            e.preventDefault();
            e.stopPropagation();

            try {
              this.handleFilterClick(e);
            } catch (error) {
              console.error(
                '[Excel Editor] Error in global filter handler:',
                error
              );
            }
          }
        );

      // Bind other table events
      this.elements.tableContainer.on(
        'change.excelEditor',
        '.excel-editor-cell.editable',
        (e) => this.handleCellEdit(e)
      );
      this.elements.tableContainer.on(
        'change.excelEditor',
        '.row-checkbox',
        (e) => this.handleRowSelection(e)
      );
      this.elements.tableContainer.on(
        'change.excelEditor',
        '#select-all-checkbox',
        (e) => this.handleSelectAllCheckbox(e)
      );

      console.log('[Excel Editor] All table events bound successfully');
    }

    /**
     * Setup drag and drop file upload
     */
    setupDragDropUpload() {
      const uploadArea = this.elements.uploadArea[0];
      if (!uploadArea) return;

      ['dragenter', 'dragover', 'dragleave', 'drop'].forEach((eventName) => {
        uploadArea.addEventListener(eventName, this.preventDefaults, false);
      });

      ['dragenter', 'dragover'].forEach((eventName) => {
        uploadArea.addEventListener(
          eventName,
          () => uploadArea.classList.add('dragover'),
          false
        );
      });

      ['dragleave', 'drop'].forEach((eventName) => {
        uploadArea.addEventListener(
          eventName,
          () => uploadArea.classList.remove('dragover'),
          false
        );
      });

      uploadArea.addEventListener('drop', (e) => this.handleFileDrop(e), false);
    }

    /**
     * Prevent default drag behaviors
     */
    preventDefaults(e) {
      e.preventDefault();
      e.stopPropagation();
    }

    /**
     * Handle file drop
     */
    handleFileDrop(e) {
      const files = e.dataTransfer.files;
      if (files.length > 0) {
        this.processFile(files[0]);
      }
    }

    /**
     * Handle file input change
     */
    handleFileUpload(e) {
      const file = e.target.files[0];
      if (file) {
        this.processFile(file);
      }
    }

    /**
     * Process uploaded file with validation
     */
    async processFile(file) {
      try {
        this.logDebug(
          'Processing file:',
          file.name,
          'Type:',
          file.type,
          'Size:',
          file.size
        );

        // Validate file
        if (!this.validateFile(file)) {
          return;
        }

        this.showLoading('Processing Excel file...');

        // Read file
        this.logDebug('Reading file...');
        const data = await this.readFile(file);
        this.logDebug('File read successfully, data length:', data.byteLength);

        // Parse based on file type
        let parsedData;
        if (file.name.toLowerCase().endsWith('.csv')) {
          this.logDebug('Parsing as CSV...');
          parsedData = this.parseCSV(data);
        } else {
          this.logDebug('Parsing as Excel...');
          parsedData = await this.parseExcel(data);
        }

        this.logDebug('File parsed successfully, rows:', parsedData.length);

        // Process the data
        this.logDebug('Loading data into application...');
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
    }

    /**
     * Validate uploaded file
     */
    validateFile(file) {
      // Check file size
      if (file.size > this.config.maxFileSize) {
        this.showMessage(
          `File too large. Maximum size is ${
            this.config.maxFileSize / (1024 * 1024)
          }MB`,
          'error'
        );
        return false;
      }

      // Check file type
      const extension = '.' + file.name.split('.').pop().toLowerCase();
      if (!this.config.supportedFormats.includes(extension)) {
        this.showMessage(
          `Unsupported file format. Supported formats: ${this.config.supportedFormats.join(
            ', '
          )}`,
          'error'
        );
        return false;
      }

      return true;
    }

    /**
     * Read file as array buffer
     */
    readFile(file) {
      return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => resolve(e.target.result);
        reader.onerror = () => reject(new Error('Failed to read file'));
        reader.readAsArrayBuffer(file);
      });
    }

    /**
     * Parse CSV data
     */
    parseCSV(data) {
      const text = new TextDecoder().decode(data);
      const lines = text.split('\n').filter((line) => line.trim());
      return lines.map((line) => {
        // Simple CSV parsing with enhanced trimming
        return line.split(',').map((cell) => {
          // Remove quotes and trim whitespace
          return cell.trim().replace(/^["']|["']$/g, '');
        });
      });
    }

    /**
     * Parse Excel data using SheetJS
     */
    async parseExcel(data) {
      return new Promise((resolve, reject) => {
        try {
          this.logDebug('Parsing Excel file...');
          this.logDebug('Data type:', typeof data);
          this.logDebug('Data length:', data.byteLength || data.length);

          // Check if XLSX is available
          if (typeof XLSX === 'undefined') {
            console.error('XLSX library not available during parsing.');
            reject(
              new Error(
                'Excel parsing library is not available. Please try refreshing the page, or upload a CSV file instead.'
              )
            );
            return;
          }

          this.logDebug('XLSX library available, version:', XLSX.version);

          // Try parsing with different options if first attempt fails
          let workbook;
          try {
            workbook = XLSX.read(data, { type: 'array' });
          } catch (parseError) {
            this.logDebug(
              'First parse attempt failed, trying with buffer type:',
              parseError
            );
            try {
              workbook = XLSX.read(data, { type: 'buffer' });
            } catch (bufferError) {
              this.logDebug(
                'Buffer parse failed, trying with uint8array:',
                bufferError
              );
              workbook = XLSX.read(new Uint8Array(data), { type: 'array' });
            }
          }

          this.logDebug('Workbook parsed:', workbook);
          this.logDebug('Sheet names:', workbook.SheetNames);

          if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
            reject(
              new Error(
                'No worksheets found in Excel file. Please check that the file is not corrupted.'
              )
            );
            return;
          }

          const sheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[sheetName];
          this.logDebug('Using worksheet:', sheetName);

          // Convert to JSON with headers as first row
          const jsonData = XLSX.utils.sheet_to_json(worksheet, {
            header: 1, // Use first row as headers
            raw: false, // Convert everything to strings for consistency
            defval: '', // Default value for empty cells
            blankrows: false, // Skip blank rows
          });

          this.logDebug('Parsed JSON data rows:', jsonData.length);

          if (!jsonData || jsonData.length === 0) {
            reject(
              new Error(
                'No data found in Excel file. The file may be empty or corrupted.'
              )
            );
            return;
          }

          // TRIM ALL CELL VALUES to prevent duplicate filters
          const trimmedData = jsonData.map((row, rowIndex) => {
            if (!Array.isArray(row)) return row;

            return row.map((cell, cellIndex) => {
              // Trim all string values, preserve other types
              if (typeof cell === 'string') {
                return cell.trim();
              }
              // Convert other types to strings and trim
              return String(cell || '').trim();
            });
          });

          // Filter out completely empty rows and ensure we have at least header + 1 data row
          const filteredData = trimmedData.filter((row, index) => {
            // Always keep the first row (headers)
            if (index === 0) return true;

            // For other rows, check if they have any non-empty content
            return (
              Array.isArray(row) &&
              row.some(
                (cell) =>
                  cell !== null &&
                  cell !== undefined &&
                  String(cell).trim() !== ''
              )
            );
          });

          if (filteredData.length < 1) {
            reject(
              new Error(
                'Excel file appears to be empty or contains no valid data rows.'
              )
            );
            return;
          }

          if (filteredData.length === 1) {
            reject(
              new Error('Excel file only contains headers with no data rows.')
            );
            return;
          }

          this.logDebug(
            'Final filtered and trimmed data:',
            filteredData.length,
            'rows'
          );
          resolve(filteredData);
        } catch (error) {
          console.error('Excel parsing error:', error);

          // Provide more specific error messages
          let errorMessage = 'Failed to parse Excel file: ';
          if (error.message.includes('Invalid file')) {
            errorMessage +=
              'The file appears to be corrupted or not a valid Excel format.';
          } else if (error.message.includes('Password')) {
            errorMessage += 'Password-protected Excel files are not supported.';
          } else if (error.message.includes('Encrypted')) {
            errorMessage += 'Encrypted Excel files are not supported.';
          } else {
            errorMessage += error.message;
          }

          reject(
            new Error(
              errorMessage +
                ' Please try saving the file as .xlsx or upload a CSV instead.'
            )
          );
        }
      });
    }

    /**
     * Load data into the application
     */
    loadData(data) {
      this.logDebug('Loading data into application...', data);

      if (!data || data.length === 0) {
        console.error('No data provided to loadData');
        throw new Error('No data found in file');
      }

      // Trim all data before loading
      const trimmedData = data.map((row) => {
        if (!Array.isArray(row)) return row;
        return row.map((cell) => String(cell || '').trim());
      });

      this.logDebug('Setting original data...');
      this.data.original = this.deepClone(trimmedData);

      this.logDebug('Adding editable columns...');
      this.addEditableColumns();

      this.logDebug('Setting filtered data...');
      this.data.filtered = this.deepClone(this.data.original);

      this.data.selected.clear();
      this.data.dirty = false;

      this.logDebug('Applying default column visibility...');
      this.applyDefaultColumnVisibility();

      this.logDebug('Rendering interface...');
      this.renderInterface();

      this.logDebug('Updating selection count...');
      this.updateSelectionCount();

      this.logDebug('Data loading complete!');
    }

    /**
     * Add editable columns to the data
     */
    addEditableColumns() {
      if (!this.data.original.length) return;

      const headerRow = this.data.original[0];

      // Check if editable columns already exist
      if (this.config.editableColumns.some((col) => headerRow.includes(col))) {
        return;
      }

      // Add new_barcode at the beginning
      headerRow.unshift('new_barcode');

      // Add notes and actions at the end
      headerRow.push('notes', 'actions');

      // Add empty values for existing rows
      for (let i = 1; i < this.data.original.length; i++) {
        this.data.original[i].unshift(''); // barcode
        this.data.original[i].push('', ''); // notes, actions
      }

      this.data.dirty = true;
    }

    /**
     * Apply default column visibility based on configuration
     */
    applyDefaultColumnVisibility() {
      const settings = this.config.settings;

      this.logDebug('Applying default column visibility...');
      this.logDebug('Settings:', settings);

      if (!settings.hideBehavior || settings.hideBehavior !== 'hide_others') {
        this.logDebug(
          'Hide behavior is not "hide_others", skipping column hiding. Current behavior:',
          settings.hideBehavior
        );
        return;
      }

      if (!settings.defaultVisibleColumns?.length) {
        this.logDebug('No default visible columns specified');
        return;
      }

      if (!this.data.filtered.length) {
        this.logDebug('No data available for column visibility');
        return;
      }

      const defaultColumns = settings.defaultVisibleColumns.map((col) =>
        col.trim()
      ); // Trim whitespace
      const alwaysVisible = this.config.editableColumns;
      const maxColumns = settings.maxVisibleColumns || 50;

      this.logDebug('Default columns to show:', defaultColumns);
      this.logDebug('Always visible (editable) columns:', alwaysVisible);
      this.logDebug('Max columns allowed:', maxColumns);

      this.state.hiddenColumns.clear();

      const headerRow = this.data.filtered[0];
      let visibleCount = 0;

      this.logDebug('Available columns in data:', headerRow);

      headerRow.forEach((header, index) => {
        const trimmedHeader = String(header).trim();

        // Check if column should be visible (case-insensitive matching)
        const isInDefaultList = defaultColumns.some(
          (defCol) => defCol.toLowerCase() === trimmedHeader.toLowerCase()
        );
        const isAlwaysVisible = alwaysVisible.some(
          (alwaysCol) => alwaysCol.toLowerCase() === trimmedHeader.toLowerCase()
        );
        const shouldBeVisible = isInDefaultList || isAlwaysVisible;

        this.logDebug(`Column "${trimmedHeader}" (index ${index}):`, {
          isInDefaultList,
          isAlwaysVisible,
          shouldBeVisible,
          visibleCount,
          maxColumns,
        });

        if (!shouldBeVisible) {
          this.state.hiddenColumns.add(index);
          this.logDebug(`  -> HIDING column "${trimmedHeader}"`);
        } else if (visibleCount < maxColumns) {
          visibleCount++;
          this.logDebug(
            `  -> SHOWING column "${trimmedHeader}" (visible count: ${visibleCount})`
          );
        } else {
          // Exceeded max columns limit
          this.state.hiddenColumns.add(index);
          this.logDebug(
            `  -> HIDING column "${trimmedHeader}" (max columns exceeded)`
          );
        }
      });

      this.logDebug(
        'Final hidden columns:',
        Array.from(this.state.hiddenColumns)
      );
      this.logDebug('Total columns hidden:', this.state.hiddenColumns.size);
      this.logDebug(
        'Total visible columns:',
        headerRow.length - this.state.hiddenColumns.size
      );

      if (this.state.hiddenColumns.size > 0) {
        const hiddenColumnNames = Array.from(this.state.hiddenColumns).map(
          (index) => headerRow[index]
        );
        this.logDebug('Hidden column names:', hiddenColumnNames);
        this.showMessage(
          `Applied default column visibility: ${this.state.hiddenColumns.size} columns hidden`,
          'info'
        );
      } else {
        this.logDebug('No columns were hidden');
      }
    }

    /**
     * Render the main interface
     */
    renderInterface() {
      this.logDebug('Rendering interface...');

      this.elements.uploadArea.hide();
      this.elements.mainArea.show();

      this.logDebug('Calling renderTable...');
      this.renderTable();

      this.logDebug('Calling setupFilters...');
      this.setupFilters();

      this.logDebug('Interface rendering complete!');
    }

    /**
     * Render the data table with performance optimization and loading
     */
    async renderTable() {
      this.logDebug('Starting renderTable...');
      this.logDebug('Filtered data length:', this.data.filtered.length);

      if (!this.data.filtered.length) {
        this.logDebug('No filtered data, showing empty message');
        this.elements.tableContainer.html(
          '<p class="has-text-centered">No data available</p>'
        );
        return;
      }

      // Only show quick loader if we're not already showing a process loader
      const shouldShowQuickLoader =
        !this.state.currentProcessLoader && !this.state.isLoading;

      if (shouldShowQuickLoader) {
        this.showQuickLoader('Updating table...');
      }

      // Allow UI to update
      await new Promise((resolve) => setTimeout(resolve, 30));

      try {
        const startTime = performance.now();
        const fragment = document.createDocumentFragment();

        // Create table structure
        const table = document.createElement('table');
        table.className = 'excel-editor-table table is-fullwidth is-striped';
        table.id = 'excel-table';

        this.logDebug('Creating table header...');
        // Create header
        const thead = this.createTableHeader();
        table.appendChild(thead);

        this.logDebug('Creating table body...');
        // Create body
        const tbody = this.createTableBody();
        this.logDebug('Body rows count:', tbody.children.length);
        table.appendChild(tbody);

        fragment.appendChild(table);

        this.logDebug('Replacing table content...');
        // Replace table content
        this.elements.tableContainer.html('');
        this.elements.tableContainer.append(fragment);

        // Re-cache table element
        this.elements.table = $('#excel-table');

        // IMPORTANT: Rebind table events after recreation
        this.bindTableEvents();

        const endTime = performance.now();
        this.logDebug(
          `Table rendered in ${(endTime - startTime).toFixed(2)}ms`
        );

        // Apply row styling based on actions
        this.applyRowStyling();

        // Debug: Check if filter links were created properly
        console.log(
          '[Excel Editor] Filter links after table render:',
          $('.filter-link').length
        );
        $('.filter-link').each(function (index) {
          console.log(`Filter link ${index}:`, $(this).data('column'));
        });

        this.logDebug('Table rendering complete!');
      } finally {
        if (shouldShowQuickLoader) {
          this.hideQuickLoader();
        }
      }
    }

    /**
     * Create table header
     */
    createTableHeader() {
      const thead = document.createElement('thead');
      const headerRow = document.createElement('tr');

      // Selection column
      const selectionTh = document.createElement('th');
      selectionTh.className = 'selection-column';
      selectionTh.innerHTML = `
        <label class="checkbox">
          <input type="checkbox" id="select-all-checkbox" />
        </label>
      `;
      headerRow.appendChild(selectionTh);

      // Data columns
      this.data.filtered[0].forEach((header, index) => {
        if (this.state.hiddenColumns.has(index)) return;

        const th = document.createElement('th');
        th.dataset.column = index;

        const isEditable = this.config.editableColumns.includes(header);
        let columnClass = '';

        if (header === 'new_barcode') columnClass = 'new-barcode-column';
        else if (header === 'notes') columnClass = 'notes-column';
        else if (header === 'actions') columnClass = 'actions-column';

        th.className = columnClass;
        th.innerHTML = `
          ${this.escapeHtml(header)}
          <br><small><a href="#" class="filter-link" data-column="${index}">Filter</a></small>
        `;

        headerRow.appendChild(th);
      });

      thead.appendChild(headerRow);
      return thead;
    }

    /**
     * Create table body
     */
    createTableBody() {
      const tbody = document.createElement('tbody');

      // Skip header row (index 0)
      for (let i = 1; i < this.data.filtered.length; i++) {
        const row = this.createTableRow(i);
        tbody.appendChild(row);
      }

      return tbody;
    }

    /**
     * Create individual table row
     */
    createTableRow(rowIndex) {
      const row = document.createElement('tr');
      row.dataset.row = rowIndex;

      const rowData = this.data.filtered[rowIndex];
      const isSelected = this.data.selected.has(rowIndex);

      if (isSelected) {
        row.classList.add('selected-row');
      }

      // Apply action-based styling
      const actionValue = this.getActionValue(rowIndex);
      if (actionValue) {
        row.classList.add(`action-${actionValue}`);
      }

      // Selection checkbox cell
      const selectionTd = document.createElement('td');
      selectionTd.className = 'selection-column';
      selectionTd.innerHTML = `
        <label class="checkbox">
          <input type="checkbox" class="row-checkbox"
                 data-row="${rowIndex}" ${isSelected ? 'checked' : ''} />
        </label>
      `;
      row.appendChild(selectionTd);

      // Data cells
      rowData.forEach((cell, colIndex) => {
        if (this.state.hiddenColumns.has(colIndex)) return;

        const td = this.createTableCell(rowIndex, colIndex, cell);
        row.appendChild(td);
      });

      return row;
    }

    /**
     * Create individual table cell
     */
    createTableCell(rowIndex, colIndex, cellValue) {
      const td = document.createElement('td');
      const columnName = this.data.filtered[0][colIndex];
      const isEditable = this.config.editableColumns.includes(columnName);

      if (isEditable) {
        td.className = 'editable-column';

        if (columnName === 'actions') {
          td.classList.add('actions-column');
          td.innerHTML = this.createActionsDropdown(
            rowIndex,
            colIndex,
            cellValue
          );
        } else if (columnName === 'notes') {
          td.classList.add('notes-column');
          td.innerHTML = this.createNotesTextarea(
            rowIndex,
            colIndex,
            cellValue
          );
        } else {
          td.classList.add('new-barcode-column');
          td.innerHTML = this.createTextInput(
            rowIndex,
            colIndex,
            cellValue,
            'Enter barcode...'
          );
        }
      } else {
        td.className = 'readonly-cell';
        td.innerHTML = `<span class="excel-editor-readonly">${this.escapeHtml(
          cellValue || ''
        )}</span>`;
      }

      return td;
    }

    /**
     * Create actions dropdown
     */
    createActionsDropdown(rowIndex, colIndex, value) {
      const selected = {
        '': !value ? 'selected' : '',
        relabel: value === 'relabel' ? 'selected' : '',
        pending: value === 'pending' ? 'selected' : '',
        discard: value === 'discard' ? 'selected' : '',
      };

      return `
        <div class="select is-small is-fullwidth">
          <select class="excel-editor-cell editable actions-dropdown"
                  data-row="${rowIndex}" data-col="${colIndex}">
            <option value="" ${selected['']}>${Drupal.t(
        '-- Select Action --'
      )}</option>
            <option value="relabel" ${selected['relabel']}>${Drupal.t(
        'Relabel'
      )}</option>
            <option value="pending" ${selected['pending']}>${Drupal.t(
        'Pending'
      )}</option>
            <option value="discard" ${selected['discard']}>${Drupal.t(
        'Discard'
      )}</option>
          </select>
        </div>
      `;
    }

    /**
     * Create notes textarea
     */
    createNotesTextarea(rowIndex, colIndex, value) {
      return `
        <textarea class="excel-editor-cell editable notes-textarea"
                 data-row="${rowIndex}" data-col="${colIndex}"
                 placeholder="${Drupal.t('Add notes...')}"
                 rows="2">${this.escapeHtml(value || '')}</textarea>
      `;
    }

    /**
     * Create text input
     */
    createTextInput(rowIndex, colIndex, value, placeholder) {
      return `
        <input type="text" class="excel-editor-cell editable"
               data-row="${rowIndex}" data-col="${colIndex}"
               value="${this.escapeHtml(value || '')}"
               placeholder="${Drupal.t(placeholder)}" />
      `;
    }

    /**
     * Get action value for a row
     */
    getActionValue(rowIndex) {
      const actionsColumnIndex = this.data.filtered[0].indexOf('actions');
      if (actionsColumnIndex === -1) return null;

      return this.data.filtered[rowIndex][actionsColumnIndex];
    }

    /**
     * Apply row styling based on actions
     */
    applyRowStyling() {
      this.elements.table.find('tbody tr').each((index, row) => {
        const $row = $(row);
        const rowIndex = parseInt($row.data('row'));
        const actionValue = this.getActionValue(rowIndex);

        // Remove existing action classes
        $row.removeClass('action-relabel action-pending action-discard');

        // Add appropriate class
        if (actionValue) {
          $row.addClass(`action-${actionValue}`);
        }
      });
    }

    /**
     * Handle cell editing
     */
    handleCellEdit(e) {
      const $cell = $(e.target);
      const rowIndex = parseInt($cell.data('row'));
      const colIndex = parseInt($cell.data('col'));
      // Trim the input value to prevent whitespace issues
      const newValue = String($cell.val() || '').trim();

      // Update data
      this.data.filtered[rowIndex][colIndex] = newValue;
      this.data.dirty = true;

      // Apply row styling if this was an action change
      const columnName = this.data.filtered[0][colIndex];
      if (columnName === 'actions') {
        this.applyRowStyling();
      }

      // Debounced save indication using Drupal's debounce
      const debouncedSaveIndication = Drupal.debounce(() => {
        this.showMessage(
          'Changes detected. Remember to save your draft.',
          'info',
          3000
        );
      }, 1000);

      debouncedSaveIndication();
    }

    /**
     * Handle row selection
     */
    handleRowSelection(e) {
      const $checkbox = $(e.target);
      const rowIndex = parseInt($checkbox.data('row'));
      const isChecked = $checkbox.is(':checked');

      if (isChecked) {
        this.data.selected.add(rowIndex);
        $checkbox.closest('tr').addClass('selected-row');
      } else {
        this.data.selected.delete(rowIndex);
        $checkbox.closest('tr').removeClass('selected-row');
      }

      this.updateSelectionCount();
      this.updateSelectAllCheckbox();
    }

    /**
     * Update selection count display
     */
    updateSelectionCount() {
      const count = this.data.selected.size;
      this.elements.selectionCount.text(
        `${count} row${count !== 1 ? 's' : ''} selected`
      );

      // Enable/disable export button
      this.elements.exportBtn
        .prop('disabled', count === 0)
        .toggleClass('is-disabled', count === 0);
    }

    /**
     * Update select all checkbox state
     */
    updateSelectAllCheckbox() {
      const totalRows = this.data.filtered.length - 1; // Exclude header
      const selectedRows = this.data.selected.size;

      const $selectAllCheckbox = $('#select-all-checkbox');

      if (selectedRows === 0) {
        $selectAllCheckbox.prop('checked', false).prop('indeterminate', false);
      } else if (selectedRows === totalRows) {
        $selectAllCheckbox.prop('checked', true).prop('indeterminate', false);
      } else {
        $selectAllCheckbox.prop('checked', false).prop('indeterminate', true);
      }
    }

    /**
     * Clean existing data by trimming all cell values.
     */
    cleanExistingData() {
      if (!this.data.original || !this.data.original.length) {
        console.log('No data to clean');
        return;
      }

      console.log('Cleaning existing data by trimming all cell values...');

      // Clean original data
      this.data.original = this.data.original.map((row) => {
        if (!Array.isArray(row)) return row;
        return row.map((cell) => String(cell || '').trim());
      });

      // Clean filtered data
      this.data.filtered = this.data.filtered.map((row) => {
        if (!Array.isArray(row)) return row;
        return row.map((cell) => String(cell || '').trim());
      });

      // Re-render everything
      this.renderTable();
      this.setupFilters();

      console.log('Data cleaning complete!');
      this.showMessage('Data cleaned - all cell values trimmed', 'success');
    }

    /**
     * Select all visible rows
     */
    selectAllVisible() {
      this.elements.tableContainer
        .find('.row-checkbox')
        .each((index, checkbox) => {
          const $checkbox = $(checkbox);
          const rowIndex = parseInt($checkbox.data('row'));

          if (!$checkbox.is(':checked')) {
            $checkbox.prop('checked', true);
            this.data.selected.add(rowIndex);
            $checkbox.closest('tr').addClass('selected-row');
          }
        });

      this.updateSelectionCount();
      this.updateSelectAllCheckbox();
    }

    /**
     * Deselect all rows
     */
    deselectAll() {
      this.data.selected.clear();
      this.elements.tableContainer.find('.row-checkbox').prop('checked', false);
      this.elements.tableContainer.find('tr').removeClass('selected-row');

      this.updateSelectionCount();
      this.updateSelectAllCheckbox();
    }

    /**
     * Setup filters interface
     */
    setupFilters() {
      if (!this.data.filtered.length) return;

      let statusMessages = '';

      // Hidden columns notification
      if (this.state.hiddenColumns.size > 0) {
        statusMessages += this.createHiddenColumnsNotification();
      }

      // Default columns notification
      if (this.shouldShowDefaultColumnsNotification()) {
        statusMessages += this.createDefaultColumnsNotification();
      }

      this.elements.filtersContainer.html(`
        <div class="field" id="active-filters-container" style="display: none;">
          <label class="label">${Drupal.t('Active Filters:')}</label>
          <div class="control" id="active-filters">
            <!-- Active filters will be added here -->
          </div>
          <div class="control">
            <button class="button is-small is-light" id="clear-all-filters-btn">
              <span class="icon is-small"><i class="fas fa-times"></i></span>
              <span>${Drupal.t('Clear All Filters')}</span>
            </button>
          </div>
        </div>
        ${statusMessages}
      `);

      this.bindFilterEvents();
    }

    /**
     * Create hidden columns notification
     */
    createHiddenColumnsNotification() {
      return `
        <div class="field">
          <div class="notification is-info is-light">
            <span class="icon"><i class="fas fa-eye-slash"></i></span>
            ${this.state.hiddenColumns.size} column${
        this.state.hiddenColumns.size !== 1 ? 's' : ''
      } hidden.
            <button class="button is-small is-light ml-2" id="show-column-settings">
              <span class="icon is-small"><i class="fas fa-eye"></i></span>
              <span>${Drupal.t('Manage Columns')}</span>
            </button>
          </div>
        </div>
      `;
    }

    /**
     * Create default columns notification
     */
    createDefaultColumnsNotification() {
      const settings = this.config.settings;
      return `
        <div class="field">
          <div class="notification is-primary is-light">
            <span class="icon"><i class="fas fa-cog"></i></span>
            ${Drupal.t('Default column visibility applied')} (${
        settings.defaultVisibleColumns.length
      } ${Drupal.t('columns configured')}).
            <button class="button is-small is-light ml-2" id="reset-to-defaults">
              <span class="icon is-small"><i class="fas fa-undo"></i></span>
              <span>${Drupal.t('Reset to Defaults')}</span>
            </button>
            <button class="button is-small is-light ml-2" id="show-all-override">
              <span class="icon is-small"><i class="fas fa-eye"></i></span>
              <span>${Drupal.t('Show All')}</span>
            </button>
          </div>
        </div>
      `;
    }

    /**
     * Check if default columns notification should be shown
     */
    shouldShowDefaultColumnsNotification() {
      const settings = this.config.settings;
      return (
        settings.hideBehavior === 'hide_others' &&
        settings.defaultVisibleColumns?.length > 0
      );
    }

    /**
     * Bind filter-related events
     */
    bindFilterEvents() {
      $('#clear-all-filters-btn').on('click', () => this.clearAllFilters());
      $('#show-column-settings').on('click', () =>
        this.showColumnVisibilityModal()
      );
      $('#reset-to-defaults').on('click', () => this.resetToDefaultColumns());
      $('#show-all-override').on('click', () => this.showAllColumnsOverride());
    }

    /**
     * Handle select all checkbox
     */
    handleSelectAllCheckbox(e) {
      const isChecked = $(e.target).is(':checked');

      if (isChecked) {
        this.selectAllVisible();
      } else {
        this.deselectAll();
      }
    }

    /**
     * Enhanced handleFilterClick with extensive debugging and error handling
     */
    handleFilterClick(e) {
      console.log('[Excel Editor] handleFilterClick method entered', {
        event: e,
        target: e.target,
        currentTarget: e.currentTarget,
      });

      e.preventDefault();

      const $target = $(e.target);
      let columnIndex = $target.data('column');

      console.log('[Excel Editor] Initial column index:', columnIndex);

      // Fallback: if data-column not found on target, try parent elements
      if (columnIndex === undefined || columnIndex === null) {
        const $link = $target.closest('.filter-link');
        columnIndex = $link.data('column');
        console.log(
          '[Excel Editor] Column index from closest .filter-link:',
          columnIndex
        );
      }

      // Another fallback: parse from nearby th element
      if (columnIndex === undefined || columnIndex === null) {
        const $th = $target.closest('th');
        columnIndex = $th.data('column');
        console.log('[Excel Editor] Column index from th:', columnIndex);
      }

      console.log('[Excel Editor] Final column index:', columnIndex);

      if (columnIndex === undefined || columnIndex === null) {
        console.error(
          '[Excel Editor] Could not determine column index for filter',
          {
            target: e.target,
            targetData: $target.data(),
            closestLink: $target.closest('.filter-link').data(),
            closestTh: $target.closest('th').data(),
          }
        );
        alert(
          'Error: Could not determine column for filtering. See console for details.'
        );
        return;
      }

      console.log(
        '[Excel Editor] About to call showColumnFilter with index:',
        columnIndex
      );

      try {
        this.showColumnFilter(columnIndex);
        console.log('[Excel Editor] showColumnFilter completed successfully');
      } catch (error) {
        console.error('[Excel Editor] Error in showColumnFilter:', error);
        alert('Error showing filter modal: ' + error.message);
      }
    }

    /**
     * Enhanced showColumnFilter with better error handling and debugging
     */
    showColumnFilter(columnIndex) {
      console.log(
        '[Excel Editor] showColumnFilter called with columnIndex:',
        columnIndex
      );

      // Validate inputs
      if (!this.data.filtered || !this.data.filtered.length) {
        console.error('[Excel Editor] No data available for filtering');
        alert('No data available for filtering');
        return;
      }

      if (columnIndex < 0 || columnIndex >= this.data.filtered[0].length) {
        console.error(
          '[Excel Editor] Invalid column index:',
          columnIndex,
          'Available columns:',
          this.data.filtered[0].length
        );
        alert('Invalid column selected');
        return;
      }

      const header = this.data.filtered[0][columnIndex];
      console.log('[Excel Editor] Creating filter for column:', header);

      // Show loader while getting unique values
      this.showQuickLoader('Loading filter options...');

      // Use setTimeout to allow loader to show
      setTimeout(() => {
        try {
          const uniqueValues = this.getUniqueColumnValues(columnIndex);
          console.log(
            '[Excel Editor] Unique values for column:',
            uniqueValues.length,
            'values'
          );

          // Remove any existing filter modals
          $('.modal#filter-modal').remove();
          console.log('[Excel Editor] Removed existing modals');

          // Create checkbox options HTML
          const checkboxOptionsHtml = uniqueValues
            .map((val, index) => {
              const value = val || '';
              const displayValue = value === '' ? '(empty)' : value;
              const isChecked = this.isValueSelectedInFilter(
                columnIndex,
                value
              );

              return `
            <div class="column is-half">
              <label class="checkbox filter-checkbox-item">
                <input type="checkbox"
                       value="${this.escapeHtml(value)}"
                       ${isChecked ? 'checked' : ''}
                       class="filter-value-checkbox">
                <span class="filter-checkbox-label">${this.escapeHtml(
                  displayValue
                )}</span>
              </label>
            </div>
          `;
            })
            .join('');

          const modalHtml = `
        <div class="modal is-active" id="filter-modal" style="display: flex !important; z-index: 9999;">
          <div class="modal-background"></div>
          <div class="modal-content">
            <div class="box">
              <h3 class="title is-4">
                <span class="icon"><i class="fas fa-filter"></i></span>
                Filter: ${this.escapeHtml(header)}
              </h3>
              <p class="subtitle is-6">Column ${columnIndex + 1} with ${
            uniqueValues.length
          } unique values</p>

              <div class="tabs" id="filter-tabs">
                <ul>
                  <li class="is-active"><a data-tab="quick">Quick Filter</a></li>
                  <li><a data-tab="advanced">Advanced Filter</a></li>
                </ul>
              </div>

              <!-- Quick Filter Tab with Checkboxes -->
              <div class="tab-content" id="quick-filter-tab">
                <div class="field">
                  <div class="field is-grouped is-grouped-multiline mb-3">
                    <div class="control">
                      <button class="button is-small is-info" id="select-all-values">
                        <span class="icon is-small"><i class="fas fa-check-square"></i></span>
                        <span>Select All</span>
                      </button>
                    </div>
                    <div class="control">
                      <button class="button is-small is-light" id="deselect-all-values">
                        <span class="icon is-small"><i class="fas fa-square"></i></span>
                        <span>Deselect All</span>
                      </button>
                    </div>
                    <div class="control">
                      <button class="button is-small is-warning" id="invert-selection">
                        <span class="icon is-small"><i class="fas fa-exchange-alt"></i></span>
                        <span>Invert Selection</span>
                      </button>
                    </div>
                  </div>

                  <div class="field">
                    <input class="input is-small" type="text" id="filter-search"
                           placeholder="Search values..." />
                  </div>

                  <div class="filter-values-container" style="max-height: 300px; overflow-y: auto; border: 1px solid #dbdbdb; border-radius: 4px; padding: 1rem;">
                    <div class="columns is-multiline" id="filter-checkboxes">
                      ${checkboxOptionsHtml}
                    </div>
                  </div>

                  <p class="help mt-2">
                    <span id="selected-count">0</span> of ${
                      uniqueValues.length
                    } values selected
                  </p>
                </div>
              </div>

              <!-- Advanced Filter Tab -->
              <div class="tab-content" id="advanced-filter-tab" style="display: none;">
                <div class="field">
                  <label class="label">Filter Type:</label>
                  <div class="control">
                    <div class="select is-fullwidth">
                      <select id="filter-type">
                        <option value="equals">Equals</option>
                        <option value="contains">Contains</option>
                        <option value="starts">Starts with</option>
                        <option value="ends">Ends with</option>
                        <option value="not_equals">Not equals</option>
                        <option value="not_contains">Does not contain</option>
                        <option value="empty">Is empty</option>
                        <option value="not_empty">Is not empty</option>
                      </select>
                    </div>
                  </div>
                </div>

                <div class="field" id="filter-value-field">
                  <label class="label">Filter Value:</label>
                  <div class="control">
                    <input class="input" type="text" id="filter-value" placeholder="Enter filter value...">
                  </div>
                </div>
              </div>

              <!-- Modal Actions -->
              <div class="field is-grouped is-grouped-right">
                <div class="control">
                  <button class="button" id="clear-column-filter">
                    <span class="icon"><i class="fas fa-times"></i></span>
                    <span>Clear Filter</span>
                  </button>
                </div>
                <div class="control">
                  <button class="button" id="cancel-filter">Cancel</button>
                </div>
                <div class="control">
                  <button class="button is-primary" id="apply-filter">
                    <span class="icon"><i class="fas fa-check"></i></span>
                    <span>Apply Filter</span>
                  </button>
                </div>
              </div>
            </div>
          </div>
          <button class="modal-close is-large" aria-label="close"></button>
        </div>
      `;

          console.log('[Excel Editor] Creating modal jQuery object');
          const modal = $(modalHtml);

          console.log('[Excel Editor] Appending modal to body');
          $('body').append(modal);

          // Verify modal was added and is visible
          const $modalCheck = $('#filter-modal');
          console.log('[Excel Editor] Modal verification:', {
            found: $modalCheck.length,
            isVisible: $modalCheck.is(':visible'),
            hasClass: $modalCheck.hasClass('is-active'),
            display: $modalCheck.css('display'),
            zIndex: $modalCheck.css('z-index'),
          });

          // Force visibility if needed
          if (!$modalCheck.is(':visible')) {
            console.log('[Excel Editor] Modal not visible, forcing display');
            $modalCheck
              .css({
                display: 'flex !important',
                position: 'fixed',
                top: '0',
                left: '0',
                width: '100%',
                height: '100%',
                'z-index': '9999',
              })
              .addClass('is-active')
              .show();
          }

          console.log('[Excel Editor] Binding modal events');
          this.bindFilterModalEvents(modal, columnIndex, header);

          // Update selected count
          this.updateFilterSelectedCount(modal);

          console.log('[Excel Editor] showColumnFilter completed');
        } finally {
          this.hideQuickLoader();
        }
      }, 100);
    }

    /**
     * Update the selected count in filter modal
     */
    updateFilterSelectedCount(modal) {
      const checkedBoxes = modal.find('.filter-value-checkbox:checked');
      const totalBoxes = modal.find('.filter-value-checkbox');

      modal.find('#selected-count').text(checkedBoxes.length);

      // Update select/deselect all button states
      const selectAllBtn = modal.find('#select-all-values');
      const deselectAllBtn = modal.find('#deselect-all-values');

      if (checkedBoxes.length === 0) {
        selectAllBtn.removeClass('is-light').addClass('is-info');
        deselectAllBtn.removeClass('is-info').addClass('is-light');
      } else if (checkedBoxes.length === totalBoxes.length) {
        selectAllBtn.removeClass('is-info').addClass('is-light');
        deselectAllBtn.removeClass('is-light').addClass('is-info');
      } else {
        selectAllBtn.removeClass('is-light').addClass('is-info');
        deselectAllBtn.removeClass('is-light').addClass('is-info');
      }
    }

    /**
     * Show column visibility modal
     */
    showColumnVisibilityModal() {
      if (!this.data.filtered.length) {
        this.showMessage('No data loaded', 'warning');
        return;
      }

      const headers = this.data.filtered[0];

      // Generate column checkboxes
      const checkboxesHtml = headers
        .map((header, index) => {
          const isVisible = !this.state.hiddenColumns.has(index);
          const isEditable = this.config.editableColumns.includes(header);
          return `
          <div class="column is-half">
            <label class="checkbox">
              <input type="checkbox"
                     class="column-visibility-checkbox"
                     data-column-index="${index}"
                     ${isVisible ? 'checked' : ''}>
              <span class="column-name">${this.escapeHtml(header)}</span>
              ${
                isEditable
                  ? `<span class="tag is-small is-info ml-2">${Drupal.t(
                      'Editable'
                    )}</span>`
                  : ''
              }
            </label>
          </div>
        `;
        })
        .join('');

      const modalHtml = `
        <div class="modal is-active" id="column-visibility-modal">
          <div class="modal-background"></div>
          <div class="modal-content">
            <div class="box">
              <h3 class="title is-4">Manage Column Visibility</h3>

              <div class="field is-grouped is-grouped-multiline">
                <div class="control">
                  <button class="button is-small is-info" id="show-all-columns">
                    <span class="icon"><i class="fas fa-eye"></i></span>
                    <span>Show All</span>
                  </button>
                </div>
                <div class="control">
                  <button class="button is-small is-warning" id="hide-non-editable">
                    <span class="icon"><i class="fas fa-eye-slash"></i></span>
                    <span>Hide Non-Editable</span>
                  </button>
                </div>
                <div class="control">
                  <button class="button is-small is-primary" id="show-only-editable">
                    <span class="icon"><i class="fas fa-edit"></i></span>
                    <span>Show Only Editable</span>
                  </button>
                </div>
              </div>

              <div class="column-checkboxes columns is-multiline" id="column-checkboxes">
                ${checkboxesHtml}
              </div>

              <div class="field is-grouped is-grouped-right">
                <div class="control">
                  <button class="button" id="cancel-column-visibility">Cancel</button>
                </div>
                <div class="control">
                  <button class="button is-primary" id="apply-column-visibility">
                    <span class="icon"><i class="fas fa-check"></i></span>
                    <span>Apply Changes</span>
                  </button>
                </div>
              </div>
            </div>
          </div>
          <button class="modal-close is-large" aria-label="close"></button>
        </div>
      `;

      const modal = $(modalHtml);
      $('body').append(modal);

      this.bindColumnModalEvents(modal, headers);
    }

    /**
     * Bind column modal events
     */
    bindColumnModalEvents(modal, headers) {
      // Close modal events
      modal
        .find('.modal-close, #cancel-column-visibility, .modal-background')
        .on('click', () => {
          modal.remove();
        });

      // Quick action buttons
      modal.find('#show-all-columns').on('click', () => {
        modal.find('.column-visibility-checkbox').prop('checked', true);
      });

      modal.find('#hide-non-editable').on('click', () => {
        modal.find('.column-visibility-checkbox').each((index, checkbox) => {
          const colIndex = parseInt($(checkbox).data('column-index'));
          const header = headers[colIndex];
          const isEditable = this.config.editableColumns.includes(header);
          $(checkbox).prop('checked', isEditable);
        });
      });

      modal.find('#show-only-editable').on('click', () => {
        modal.find('.column-visibility-checkbox').each((index, checkbox) => {
          const colIndex = parseInt($(checkbox).data('column-index'));
          const header = headers[colIndex];
          const isEditable = this.config.editableColumns.includes(header);
          $(checkbox).prop('checked', isEditable);
        });
      });

      // Apply changes
      modal.find('#apply-column-visibility').on('click', () => {
        this.applyColumnVisibilityChanges(modal);
        modal.remove();
      });
    }

    /**
     * Apply column visibility changes from modal with loader
     */
    async applyColumnVisibilityChanges(modal) {
      this.showProcessLoader('Updating column visibility...');

      // Allow UI to update
      await new Promise((resolve) => setTimeout(resolve, 50));

      try {
        // Update hidden columns set
        this.state.hiddenColumns.clear();

        modal.find('.column-visibility-checkbox').each((index, checkbox) => {
          const colIndex = parseInt($(checkbox).data('column-index'));
          const isChecked = $(checkbox).is(':checked');

          if (!isChecked) {
            this.state.hiddenColumns.add(colIndex);
          }
        });

        // Re-render table and filters
        await this.renderTable();
        this.setupFilters();

        // Show feedback message
        const hiddenCount = this.state.hiddenColumns.size;
        if (hiddenCount > 0) {
          this.showMessage(
            `${hiddenCount} column${
              hiddenCount !== 1 ? 's' : ''
            } hidden from view`,
            'info'
          );
        } else {
          this.showMessage('All columns are now visible', 'success');
        }
      } finally {
        this.hideProcessLoader();
      }
    }

    /**
     * Reset to default columns
     */
    resetToDefaultColumns() {
      if (!this.data.filtered.length) {
        this.showMessage('No data loaded', 'warning');
        return;
      }

      const settings = this.config.settings;
      if (
        settings.hideBehavior !== 'hide_others' ||
        !settings.defaultVisibleColumns?.length
      ) {
        this.showMessage('Default column hiding is not configured', 'warning');
        return;
      }

      this.applyDefaultColumnVisibility();
      this.renderTable();
      this.setupFilters();
      this.showMessage(
        `Reset to default column visibility: ${this.state.hiddenColumns.size} columns hidden`,
        'success'
      );
    }

    /**
     * Check if a value is currently selected in the filter
     */
    isValueSelectedInFilter(columnIndex, value) {
      if (!this.state.currentFilters[columnIndex]) {
        return true; // If no filter, show all as selected
      }

      const filter = this.state.currentFilters[columnIndex];
      if (filter.type === 'quick' && filter.selected) {
        return filter.selected.includes(value);
      }

      return false;
    }

    /**
     * Show all columns override
     */
    showAllColumnsOverride() {
      this.state.hiddenColumns.clear();
      this.renderTable();
      this.setupFilters();
      this.showMessage('All columns are now visible', 'success');
    }

    /**
     * Get unique values for a column
     */
    getUniqueColumnValues(columnIndex) {
      const values = new Set();

      // Skip header row (index 0)
      for (let i = 1; i < this.data.filtered.length; i++) {
        const rawValue = this.data.filtered[i][columnIndex];
        // Trim the value to prevent duplicates from whitespace
        const trimmedValue = String(rawValue || '').trim();
        values.add(trimmedValue);
      }

      return Array.from(values).sort();
    }

    /**
     * Bind filter modal events
     */
    bindFilterModalEvents(modal, columnIndex, header) {
      // Tab switching
      modal.find('[data-tab]').on('click', (e) => {
        e.preventDefault();
        const tabName = $(e.target).data('tab');

        modal.find('[data-tab]').parent().removeClass('is-active');
        $(e.target).parent().addClass('is-active');

        modal.find('.tab-content').hide();
        modal.find(`#${tabName}-filter-tab`).show();
      });

      // Checkbox change events
      modal.find('.filter-value-checkbox').on('change', () => {
        this.updateFilterSelectedCount(modal);
      });

      // Select all values
      modal.find('#select-all-values').on('click', () => {
        modal.find('.filter-value-checkbox').prop('checked', true);
        this.updateFilterSelectedCount(modal);
      });

      // Deselect all values
      modal.find('#deselect-all-values').on('click', () => {
        modal.find('.filter-value-checkbox').prop('checked', false);
        this.updateFilterSelectedCount(modal);
      });

      // Invert selection
      modal.find('#invert-selection').on('click', () => {
        modal.find('.filter-value-checkbox').each(function () {
          $(this).prop('checked', !$(this).prop('checked'));
        });
        this.updateFilterSelectedCount(modal);
      });

      // Search functionality
      modal.find('#filter-search').on('input', (e) => {
        const searchTerm = $(e.target).val().toLowerCase();

        modal.find('.filter-checkbox-item').each(function () {
          const label = $(this)
            .find('.filter-checkbox-label')
            .text()
            .toLowerCase();
          const shouldShow = label.includes(searchTerm);
          $(this).closest('.column').toggle(shouldShow);
        });
      });

      // Filter type changes
      modal.find('#filter-type').on('change', (e) => {
        const filterType = $(e.target).val();
        const valueField = modal.find('#filter-value-field');

        if (filterType === 'empty' || filterType === 'not_empty') {
          valueField.hide();
        } else {
          valueField.show();
        }
      });

      // Modal close events
      modal
        .find('.modal-close, #cancel-filter, .modal-background')
        .on('click', () => {
          modal.remove();
        });

      // Clear filter
      modal.find('#clear-column-filter').on('click', async () => {
        this.showProcessLoader('Clearing filter...');

        try {
          delete this.state.currentFilters[columnIndex];
          await this.applyFilters();
          this.updateActiveFiltersDisplay();
          modal.remove();
          this.showMessage(`Filter cleared for ${header}`, 'success');
        } finally {
          this.hideProcessLoader();
        }
      });

      // Apply filter
      modal.find('#apply-filter').on('click', async () => {
        await this.applyFilterFromModal(modal, columnIndex);
        modal.remove();
      });
    }

    /**
     * Apply filter from modal
     */
    async applyFilterFromModal(modal, columnIndex) {
      this.showProcessLoader('Applying filter...');

      try {
        const activeTab = modal.find('.tabs .is-active [data-tab]').data('tab');

        if (activeTab === 'quick') {
          // Quick filter using checkboxes
          const selectedValues = [];
          modal.find('.filter-value-checkbox:checked').each(function () {
            selectedValues.push($(this).val());
          });

          if (selectedValues.length > 0) {
            this.state.currentFilters[columnIndex] = {
              type: 'quick',
              selected: selectedValues,
            };
          } else {
            delete this.state.currentFilters[columnIndex];
          }
        } else {
          // Advanced filter
          const filterType = modal.find('#filter-type').val();
          const filterValue = modal.find('#filter-value').val();

          if (
            filterType === 'empty' ||
            filterType === 'not_empty' ||
            filterValue
          ) {
            this.state.currentFilters[columnIndex] = {
              type: filterType,
              value: filterValue,
            };
          } else {
            delete this.state.currentFilters[columnIndex];
          }
        }

        await this.applyFilters();
        this.updateActiveFiltersDisplay();

        const header = this.data.filtered[0][columnIndex];
        this.showMessage(`Filter applied to ${header}`, 'success');
      } finally {
        this.hideProcessLoader();
      }
    }

    /**
     * Apply all filters to data with loading indicator
     */
    async applyFilters() {
      if (Object.keys(this.state.currentFilters).length === 0) {
        this.data.filtered = this.deepClone(this.data.original);
        this.renderTable();
        return;
      }

      this.showProcessLoader('Applying filters...');

      // Use setTimeout to allow UI to update before heavy processing
      await new Promise((resolve) => setTimeout(resolve, 50));

      try {
        const startTime = performance.now();

        // Start with original data
        this.data.filtered = [this.data.original[0]]; // Keep header

        // Filter each row
        for (let i = 1; i < this.data.original.length; i++) {
          const row = this.data.original[i];
          let includeRow = true;

          // Check each filter
          for (const [columnIndex, filter] of Object.entries(
            this.state.currentFilters
          )) {
            const colIndex = parseInt(columnIndex);
            const cellValue = row[colIndex] || '';

            if (!this.rowMatchesFilter(cellValue, filter)) {
              includeRow = false;
              break;
            }
          }

          if (includeRow) {
            this.data.filtered.push(row);
          }
        }

        const endTime = performance.now();
        this.logDebug(
          `Filters applied in ${(endTime - startTime).toFixed(2)}ms`
        );

        // Clear selection since row indices have changed
        this.data.selected.clear();

        await this.renderTable();
        this.updateSelectionCount();

        const filteredCount = this.data.filtered.length - 1; // Exclude header
        const totalCount = this.data.original.length - 1;

        if (filteredCount < totalCount) {
          this.showMessage(
            `Showing ${filteredCount} of ${totalCount} rows`,
            'info'
          );
        }
      } finally {
        this.hideProcessLoader();
      }
    }

    /**
     * Check if a row matches a filter
     */
    rowMatchesFilter(cellValue, filter) {
      // Trim both the cell value and filter value for comparison
      const value = String(cellValue || '')
        .trim()
        .toLowerCase();

      switch (filter.type) {
        case 'quick':
          // Trim the selected values for comparison
          // eslint-disable-next-line no-case-declarations
          const trimmedSelected = filter.selected.map((val) =>
            String(val || '').trim()
          );
          return trimmedSelected.includes(String(cellValue || '').trim());

        case 'equals':
          return (
            value ===
            String(filter.value || '')
              .trim()
              .toLowerCase()
          );

        case 'contains':
          return value.includes(
            String(filter.value || '')
              .trim()
              .toLowerCase()
          );

        case 'starts':
          return value.startsWith(
            String(filter.value || '')
              .trim()
              .toLowerCase()
          );

        case 'ends':
          return value.endsWith(
            String(filter.value || '')
              .trim()
              .toLowerCase()
          );

        case 'not_equals':
          return (
            value !==
            String(filter.value || '')
              .trim()
              .toLowerCase()
          );

        case 'not_contains':
          return !value.includes(
            String(filter.value || '')
              .trim()
              .toLowerCase()
          );

        case 'empty':
          return !cellValue || String(cellValue).trim() === '';

        case 'not_empty':
          return cellValue && String(cellValue).trim() !== '';

        default:
          return true;
      }
    }

    /**
     * Enhanced debug logging specifically for draft operations
     */
    logDraftDebug(operation, data = null) {
      const isDraftDebug =
        this.config.settings.debug ||
        window.location.search.includes('debug=1') ||
        localStorage.getItem('excel_editor_debug') === 'true' ||
        localStorage.getItem('excel_editor_draft_debug') === 'true';

      if (isDraftDebug) {
        console.log(`[Excel Editor - Drafts] ${operation}`, data);
      }
    }

    /**
     * Show process-specific loading overlay
     */
    showProcessLoader(message = 'Processing...', target = null) {
      // Remove any existing process loaders
      this.hideProcessLoader();

      const loaderId = 'process-loader-' + Date.now();
      const loader = $(`
        <div class="excel-editor-overlay-loader" id="${loaderId}">
          <div class="loading-content">
            <div class="excel-editor-spinner"></div>
            <p><strong>${this.escapeHtml(message)}</strong></p>
          </div>
        </div>
      `);

      // Append to body for full-screen coverage
      $('body').append(loader);
      this.state.currentProcessLoader = loaderId;

      this.logDebug('Process loader shown:', message);
    }

    /**
     * Hide process-specific loading overlay
     */
    hideProcessLoader() {
      if (this.state.currentProcessLoader) {
        $(`#${this.state.currentProcessLoader}`).remove();
        this.state.currentProcessLoader = null;
      }
      // Also remove any stray loaders
      $('.excel-editor-overlay-loader').remove();

      this.logDebug('Process loader hidden');
    }

    /**
     * Clear all loaders as a safety net
     */
    clearAllLoaders() {
      this.logDebug('Clearing all loaders...');

      // Hide main loading
      this.hideLoading();

      // Hide process loader
      this.hideProcessLoader();

      // Hide quick loader
      this.hideQuickLoader();

      // Remove any remaining loader elements
      $('.excel-editor-loading').removeClass('active').hide();
      $('.excel-editor-overlay-loader').remove();
      $('.excel-editor-quick-loader').remove();

      // Reset state
      this.state.currentProcessLoader = null;
      this.state.isLoading = false;

      this.logDebug('All loaders cleared');
    }

    /**
     * Show quick notification loader (top-right corner)
     */
    showQuickLoader(message = 'Working...') {
      this.hideQuickLoader();

      const loader = $(`
        <div class="excel-editor-quick-loader" id="quick-loader">
          <div class="spinner"></div>
          <span>${this.escapeHtml(message)}</span>
        </div>
      `);

      $('body').append(loader);

      setTimeout(() => {
        loader.addClass('slide-in');
      }, 10);
    }

    /**
     * Hide quick notification loader
     */
    hideQuickLoader() {
      const loader = $('#quick-loader');
      if (loader.length) {
        loader.addClass('slide-out');
        setTimeout(() => loader.remove(), 300);
      }
    }

    /**
     * Update active filters display
     */
    updateActiveFiltersDisplay() {
      const filtersContainer = $('#active-filters');
      const containerWrapper = $('#active-filters-container');

      if (Object.keys(this.state.currentFilters).length === 0) {
        containerWrapper.hide();
        return;
      }

      const filterTags = Object.entries(this.state.currentFilters)
        .map(([columnIndex, filter]) => {
          const header = this.data.filtered[0][parseInt(columnIndex)];
          const filterDescription = this.getFilterDescription(filter);

          return `
          <span class="tag is-info">
            <strong>${this.escapeHtml(header)}</strong>: ${filterDescription}
            <button class="delete is-small ml-1" data-column="${columnIndex}"></button>
          </span>
        `;
        })
        .join(' ');

      filtersContainer.html(filterTags);
      containerWrapper.show();

      // Bind remove filter events
      filtersContainer.find('.delete').on('click', (e) => {
        const columnIndex = $(e.target).data('column');
        delete this.state.currentFilters[columnIndex];
        this.applyFilters();
        this.updateActiveFiltersDisplay();
      });
    }

    /**
     * Get filter description for display
     */
    getFilterDescription(filter) {
      switch (filter.type) {
        case 'quick':
          return `${filter.selected.length} selected`;
        case 'equals':
          return `= "${this.escapeHtml(filter.value)}"`;
        case 'contains':
          return `contains "${this.escapeHtml(filter.value)}"`;
        case 'starts':
          return `starts with "${this.escapeHtml(filter.value)}"`;
        case 'ends':
          return `ends with "${this.escapeHtml(filter.value)}"`;
        case 'not_equals':
          return `≠ "${this.escapeHtml(filter.value)}"`;
        case 'not_contains':
          return `doesn't contain "${this.escapeHtml(filter.value)}"`;
        case 'empty':
          return 'is empty';
        case 'not_empty':
          return 'is not empty';
        default:
          return 'unknown filter';
      }
    }

    /**
     * Clear all filters
     */
    clearAllFilters() {
      this.state.currentFilters = {};
      this.applyFilters();
      this.updateActiveFiltersDisplay();
      this.showMessage('All filters cleared', 'success');
    }

    /**
     * Save draft
     */
    async saveDraft() {
      if (!this.data.original.length) {
        this.showMessage('No data to save', 'warning');
        return;
      }

      try {
        this.showLoading('Saving draft...');

        // Prepare draft data
        const draftData = {
          name: `Draft ${new Date().toLocaleString()}`, // Add a default name
          data: this.data.original,
          filters: this.state.currentFilters,
          hiddenColumns: Array.from(this.state.hiddenColumns),
          selected: Array.from(this.data.selected),
          timestamp: new Date().toISOString(),
        };

        this.logDebug('Saving draft data:', {
          dataRows: draftData.data.length,
          filtersCount: Object.keys(draftData.filters).length,
          hiddenColumnsCount: draftData.hiddenColumns.length,
          selectedRowsCount: draftData.selected.length,
        });

        const response = await this.apiCall(
          'POST',
          this.config.endpoints.saveDraft,
          draftData
        );

        if (response.success) {
          this.data.dirty = false;
          this.showMessage('Draft saved successfully', 'success');
          this.loadDrafts(); // Refresh drafts list
        } else {
          throw new Error(response.message || 'Failed to save draft');
        }
      } catch (error) {
        console.error('Save draft error details:', error);

        // Provide more helpful error messages
        let userMessage = 'Failed to save draft: ';

        if (error.message.includes('Access denied')) {
          userMessage +=
            'You do not have permission to save drafts. Please check with your administrator.';
        } else if (error.message.includes('Endpoint not found')) {
          userMessage +=
            'The save functionality is not properly configured. Please contact support.';
        } else if (
          error.message.includes('CSRF') ||
          error.message.includes('token')
        ) {
          userMessage +=
            'Security token expired. Please refresh the page and try again.';
        } else if (error.message.includes('non-JSON response')) {
          userMessage += 'Server configuration issue. Please contact support.';
        } else {
          userMessage += error.message;
        }

        this.handleError(userMessage, error);
      } finally {
        this.hideLoading();
      }
    }

    /**
     * Test API endpoints (can be called from console)
     */
    async testApiEndpoints() {
      console.log('=== TESTING API ENDPOINTS ===');

      // Test CSRF token
      console.log('CSRF Token:', this.csrfToken ? 'Available' : 'Missing');
      if (!this.csrfToken) {
        console.log('Attempting to get CSRF token...');
        await this.getCsrfToken();
        console.log(
          'CSRF Token after request:',
          this.csrfToken ? 'Available' : 'Still missing'
        );
      }

      // Test endpoints configuration
      console.log('Configured endpoints:', this.config.endpoints);

      // Test list drafts (GET request - should work without CSRF)
      try {
        console.log('Testing list drafts...');
        const draftsResponse = await this.apiCall(
          'GET',
          this.config.endpoints.listDrafts
        );
        console.log('List drafts response:', draftsResponse);
      } catch (error) {
        console.error('List drafts failed:', error.message);
      }

      // Test save draft with minimal data
      try {
        console.log('Testing save draft...');
        const testData = {
          name: 'Test Draft',
          data: [
            ['header1', 'header2'],
            ['test1', 'test2'],
          ],
          filters: {},
          hiddenColumns: [],
          selected: [],
          timestamp: new Date().toISOString(),
        };

        const saveResponse = await this.apiCall(
          'POST',
          this.config.endpoints.saveDraft,
          testData
        );
        console.log('Save draft response:', saveResponse);
      } catch (error) {
        console.error('Save draft failed:', error.message);
      }

      console.log('=== END API TEST ===');
    }

    /**
     * Export selected rows
     */
    async exportSelected() {
      if (this.data.selected.size === 0) {
        this.showMessage('No rows selected for export', 'warning');
        return;
      }

      try {
        const exportData = this.prepareExportData(true);
        await this.downloadExport(exportData, 'selected_data.xlsx');
        this.showMessage(
          `Exported ${this.data.selected.size} selected rows`,
          'success'
        );
      } catch (error) {
        this.handleError('Failed to export selected rows', error);
      }
    }

    /**
     * Export all visible rows
     */
    async exportAll() {
      try {
        const exportData = this.prepareExportData(false);
        await this.downloadExport(exportData, 'all_data.xlsx');

        const rowCount = this.data.filtered.length - 1; // Exclude header
        this.showMessage(`Exported ${rowCount} rows`, 'success');
      } catch (error) {
        this.handleError('Failed to export all data', error);
      }
    }

    /**
     * Prepare data for export
     */
    prepareExportData(selectedOnly = false) {
      const exportData = [];

      // Add header row with visible columns only
      const headerRow = [];
      this.data.filtered[0].forEach((header, index) => {
        if (!this.state.hiddenColumns.has(index)) {
          headerRow.push(header);
        }
      });
      exportData.push(headerRow);

      // Add data rows
      for (let i = 1; i < this.data.filtered.length; i++) {
        const shouldInclude = selectedOnly ? this.data.selected.has(i) : true;

        if (shouldInclude) {
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
    }

    /**
     * Download export data as Excel file
     */
    async downloadExport(data, filename) {
      // Create worksheet
      const ws = XLSX.utils.aoa_to_sheet(data);

      // Create workbook
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Export');

      // Download file
      XLSX.writeFile(wb, filename);
    }

    /**
     * Load drafts list
     */
    async loadDrafts() {
      try {
        const response = await this.apiCall(
          'GET',
          this.config.endpoints.listDrafts
        );

        if (response.success && response.drafts) {
          this.renderDrafts(response.drafts);
        } else if (response.drafts) {
          // Handle case where success flag might be missing but we have drafts
          this.renderDrafts(response.drafts);
        } else {
          this.logDebug('No drafts returned from server');
          this.renderDrafts([]);
        }
      } catch (error) {
        this.logDebug('Failed to load drafts:', error);
        // Don't show error message for drafts loading failure - just log it
        this.renderDrafts([]);
      }
    }

    /**
     * Render drafts list
     */
    renderDrafts(drafts) {
      if (!drafts || drafts.length === 0) {
        this.elements.draftsContainer.html(
          `<p class="has-text-grey">${Drupal.t('No drafts found')}</p>`
        );
        return;
      }

      const draftsHtml = drafts
        .map(
          (draft) => `
        <div class="excel-editor-draft-item">
          <div>
            <strong>${this.escapeHtml(draft.name || 'Untitled Draft')}</strong>
            <br>
            <small class="has-text-grey">
              ${new Date(draft.created).toLocaleDateString()} -
              ${draft.rows || 0} rows
            </small>
          </div>
          <div class="field is-grouped">
            <div class="control">
              <button class="button is-small is-info load-draft-btn" data-draft-id="${
                draft.id
              }">
                <span class="icon is-small"><i class="fas fa-upload"></i></span>
                <span>${Drupal.t('Load')}</span>
              </button>
            </div>
            <div class="control">
              <button class="button is-small is-danger delete-draft-btn" data-draft-id="${
                draft.id
              }">
                <span class="icon is-small"><i class="fas fa-trash"></i></span>
              </button>
            </div>
          </div>
        </div>
      `
        )
        .join('');

      this.elements.draftsContainer.html(draftsHtml);

      // Bind draft actions
      this.elements.draftsContainer.find('.load-draft-btn').on('click', (e) => {
        const draftId = $(e.target).closest('button').data('draft-id');
        this.loadDraft(draftId);
      });

      this.elements.draftsContainer
        .find('.delete-draft-btn')
        .on('click', (e) => {
          const draftId = $(e.target).closest('button').data('draft-id');
          this.deleteDraft(draftId);
        });
    }

    /**
     * Load a specific draft
     */
    async loadDraft(draftId) {
      try {
        this.showLoading('Loading draft...');

        const response = await this.apiCall(
          'GET',
          `${this.config.endpoints.loadDraft}${draftId}`
        );

        if (response.success && response.data) {
          this.loadDraftData(response.data);
          this.showMessage('Draft loaded successfully', 'success');
        } else {
          throw new Error(response.message || 'Failed to load draft');
        }
      } catch (error) {
        this.handleError('Failed to load draft', error);
      } finally {
        this.hideLoading();
      }
    }

    /**
     * Load draft data into the application
     */
    loadDraftData(draftData) {
      this.data.original = draftData.data || [];
      this.data.filtered = this.deepClone(this.data.original);
      this.state.currentFilters = draftData.filters || {};
      this.state.hiddenColumns = new Set(draftData.hiddenColumns || []);
      this.data.selected = new Set(draftData.selected || []);
      this.data.dirty = false;

      this.renderInterface();
      this.applyFilters();
      this.updateActiveFiltersDisplay();
      this.updateSelectionCount();
    }

    /**
     * Delete a draft
     */
    async deleteDraft(draftId) {
      if (!confirm(Drupal.t('Are you sure you want to delete this draft?'))) {
        return;
      }

      try {
        const response = await this.apiCall(
          'DELETE',
          `${this.config.endpoints.deleteDraft}${draftId}`
        );

        if (response.success) {
          this.showMessage('Draft deleted successfully', 'success');
          this.loadDrafts(); // Refresh list
        } else {
          throw new Error(response.message || 'Failed to delete draft');
        }
      } catch (error) {
        this.handleError('Failed to delete draft', error);
      }
    }

    /**
     * Handle keyboard shortcuts
     */
    handleKeyboardShortcuts(e) {
      // Ctrl+S or Cmd+S - Save draft
      if ((e.ctrlKey || e.metaKey) && e.key === 's') {
        e.preventDefault();
        this.saveDraft();
      }

      // Ctrl+A or Cmd+A - Select all (when table is focused)
      if (
        (e.ctrlKey || e.metaKey) &&
        e.key === 'a' &&
        $(e.target).closest('.excel-editor-table').length
      ) {
        e.preventDefault();
        this.selectAllVisible();
      }

      // Escape - Clear selection
      if (e.key === 'Escape') {
        this.deselectAll();
      }
    }

    /**
     * Handle before page unload
     */
    handleBeforeUnload() {
      if (this.data.dirty) {
        return Drupal.t(
          'You have unsaved changes. Are you sure you want to leave?'
        );
      }
    }

    /**
     * Make API calls with error handling
     */
    async apiCall(method, url, data = null) {
      // Ensure we have CSRF token for write operations
      if (
        (method === 'POST' || method === 'PUT' || method === 'DELETE') &&
        !this.csrfToken
      ) {
        await this.getCsrfToken();
      }

      const options = {
        method: method,
        headers: {
          'Content-Type': 'application/json',
          'X-Requested-With': 'XMLHttpRequest',
        },
        credentials: 'same-origin', // Include cookies for session
      };

      // Add CSRF token for write operations
      if (
        (method === 'POST' || method === 'PUT' || method === 'DELETE') &&
        this.csrfToken
      ) {
        options.headers['X-CSRF-Token'] = this.csrfToken;
      }

      if (data) {
        options.body = JSON.stringify(data);
      }

      this.logDebug('Making API call:', {
        method,
        url,
        hasData: !!data,
        hasCsrf: !!this.csrfToken,
      });

      try {
        const response = await fetch(url, options);

        this.logDebug('API response status:', response.status);

        if (!response.ok) {
          // Try to get error details from response
          let errorMessage = `HTTP ${response.status}: ${response.statusText}`;

          try {
            const contentType = response.headers.get('content-type');
            if (contentType && contentType.includes('application/json')) {
              const errorData = await response.json();
              if (errorData.message) {
                errorMessage = errorData.message;
              }
            } else {
              // If it's HTML, we're probably getting an error page
              const htmlText = await response.text();
              if (
                htmlText.includes('<!DOCTYPE') ||
                htmlText.includes('<html')
              ) {
                if (response.status === 403) {
                  errorMessage = 'Access denied. Please check permissions.';
                } else if (response.status === 404) {
                  errorMessage =
                    'Endpoint not found. Please check module routing.';
                } else {
                  errorMessage = `Server returned an error page (${response.status})`;
                }
              }
            }
          } catch (parseError) {
            this.logDebug('Could not parse error response:', parseError);
          }

          throw new Error(errorMessage);
        }

        const contentType = response.headers.get('content-type');
        if (contentType && contentType.includes('application/json')) {
          const result = await response.json();
          this.logDebug('API response data:', result);
          return result;
        } else {
          // If we're not getting JSON, something is wrong
          const text = await response.text();
          this.logDebug('Non-JSON response received:', text.substring(0, 200));
          throw new Error(
            'Server returned non-JSON response. Check if the endpoint is configured correctly.'
          );
        }
      } catch (error) {
        this.logDebug('API call failed:', error);
        throw error;
      }
    }

    /**
     * Show loading indicator
     */
    showLoading(message = 'Loading...') {
      this.state.isLoading = true;
      this.elements.loadingArea.find('p').text(message);
      this.elements.loadingArea.addClass('active').show();
    }

    /**
     * Hide loading indicator
     */
    hideLoading() {
      this.state.isLoading = false;

      if (this.elements.loadingArea && this.elements.loadingArea.length > 0) {
        this.elements.loadingArea.removeClass('active').hide();
      }

      // Also hide any loading elements that might be visible by default
      $('.excel-editor-loading').removeClass('active').hide();
    }

    /**
     * Show user messages
     */
    showMessage(message, type = 'info', duration = 5000) {
      const alertClass =
        {
          success: 'is-success',
          error: 'is-danger',
          warning: 'is-warning',
          info: 'is-info',
        }[type] || 'is-info';

      const messageElement = $(`
        <div class="notification ${alertClass} excel-editor-message">
          <button class="delete"></button>
          ${this.escapeHtml(message)}
        </div>
      `);

      // Add to page
      this.elements.container.prepend(messageElement);

      // Bind close button
      messageElement.find('.delete').on('click', () => {
        messageElement.fadeOut(() => messageElement.remove());
      });

      // Auto-remove after duration
      if (duration > 0) {
        setTimeout(() => {
          messageElement.fadeOut(() => messageElement.remove());
        }, duration);
      }
    }

    /**
     * Handle errors with user-friendly messages
     */
    handleError(message, error = null) {
      this.logDebug(message, error);

      let userMessage = message;
      if (error && error.message) {
        userMessage += `: ${error.message}`;
      }

      this.showMessage(userMessage, 'error');
    }

    /**
     * Test function to be called from console
     */
    testDirectFilter() {
      console.log('[Excel Editor] Testing direct filter call');
      if (this.data.filtered && this.data.filtered.length > 0) {
        this.showColumnFilter(0); // Test with first column
      } else {
        console.log('[Excel Editor] No data loaded for test');
      }
    }

    /**
     * Debug function to check column configuration (accessible from browser console)
     */
    debugColumnConfig() {
      console.log('=== COLUMN CONFIGURATION DEBUG ===');
      console.log('Settings from Drupal:', this.config.settings);

      if (this.data.filtered.length > 0) {
        console.log('Available columns in loaded data:', this.data.filtered[0]);

        if (this.config.settings.defaultVisibleColumns) {
          console.log(
            'Configured default visible columns:',
            this.config.settings.defaultVisibleColumns
          );

          console.log('Column matching check:');
          this.config.settings.defaultVisibleColumns.forEach((configCol) => {
            const matches = this.data.filtered[0].filter(
              (dataCol) =>
                String(dataCol).trim().toLowerCase() ===
                configCol.trim().toLowerCase()
            );
            console.log(`  "${configCol}" -> Found matches:`, matches);
          });
        }

        console.log('Hide behavior:', this.config.settings.hideBehavior);
        console.log(
          'Currently hidden column indices:',
          Array.from(this.state.hiddenColumns)
        );
        console.log(
          'Currently hidden column names:',
          Array.from(this.state.hiddenColumns).map(
            (index) => this.data.filtered[0][index]
          )
        );
      } else {
        console.log('No data loaded yet');
      }
      console.log('=== END DEBUG ===');
    }

    /**
     * Debug logging
     */
    logDebug(message, data = null) {
      // Enable debug mode if URL contains debug parameter or if explicitly set
      const isDebugMode =
        this.config.settings.debug ||
        window.location.search.includes('debug=1') ||
        localStorage.getItem('excel_editor_debug') === 'true';

      if (isDebugMode) {
        console.log(`[Excel Editor] ${message}`, data);
      }
    }

    /**
     * Deep clone objects/arrays
     */
    deepClone(obj) {
      return JSON.parse(JSON.stringify(obj));
    }

    /**
     * Escape HTML to prevent XSS
     */
    escapeHtml(text) {
      const div = document.createElement('div');
      div.textContent = text;
      return div.innerHTML;
    }
  }

  /**
   * Dynamically load XLSX library if not available
   */
  function loadXLSXLibrary() {
    return new Promise((resolve, reject) => {
      if (typeof XLSX !== 'undefined') {
        resolve();
        return;
      }

      // Try multiple CDN sources
      const cdnSources = [
        'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js',
        'https://cdn.sheetjs.com/xlsx-0.18.5/package/dist/xlsx.full.min.js',
        'https://unpkg.com/xlsx@0.18.5/dist/xlsx.full.min.js',
      ];

      let currentIndex = 0;

      function tryLoadFromCDN() {
        if (currentIndex >= cdnSources.length) {
          reject(new Error('Failed to load XLSX library from all CDN sources'));
          return;
        }

        const script = document.createElement('script');
        script.src = cdnSources[currentIndex];

        script.onload = () => {
          if (typeof XLSX !== 'undefined') {
            console.log(
              'XLSX library loaded dynamically from:',
              cdnSources[currentIndex]
            );
            resolve();
          } else {
            console.warn(
              'XLSX library loaded but not available, trying next CDN...'
            );
            currentIndex++;
            tryLoadFromCDN();
          }
        };

        script.onerror = () => {
          console.warn('Failed to load from:', cdnSources[currentIndex]);
          currentIndex++;
          tryLoadFromCDN();
        };

        document.head.appendChild(script);
      }

      tryLoadFromCDN();
    });
  }

  /**
   * Initialize Excel Editor with library loading
   */
  async function initializeExcelEditor(element) {
    try {
      // Show loading message
      const loadingMsg = $(`
        <div class="notification is-info excel-editor-init-loading">
          <span class="icon"><i class="fas fa-spinner fa-spin"></i></span>
          Loading Excel processing library...
        </div>
      `);
      $(element).prepend(loadingMsg);

      // Try to load XLSX library
      await loadXLSXLibrary();

      // Remove loading message
      loadingMsg.remove();

      // Initialize Excel Editor application
      const app = new ExcelEditor();

      // Store reference on the element for potential external access
      element.excelEditor = app;

      // Make debug function globally accessible
      window.excelEditorDebug = () => app.debugColumnConfig();

      // Make test function globally accessible
      window.testExcelEditorFilter = () => app.testDirectFilter();

      console.log('Excel Editor initialized successfully with XLSX library');
      console.log(
        'Run "excelEditorDebug()" in console to debug column configuration'
      );
      console.log(
        'Run "testExcelEditorFilter()" in console to test filter functionality'
      );
    } catch (error) {
      console.error('Failed to initialize Excel Editor:', error);

      // Remove loading message and show error
      $('.excel-editor-init-loading').remove();

      $(element).prepend(`
        <div class="notification is-warning">
          <button class="delete"></button>
          <strong>Excel Library Loading Issue:</strong> ${error.message}
          <br><small>You can still upload and work with CSV files. To use Excel files, please refresh the page or contact support.</small>
          <br><br>
          <button class="button is-small is-info" onclick="window.location.reload()">
            <span class="icon is-small"><i class="fas fa-refresh"></i></span>
            <span>Refresh Page</span>
          </button>
        </div>
      `);

      // Bind close button for the notification
      $(element)
        .find('.notification .delete')
        .on('click', function () {
          $(this).parent().remove();
        });

      // Initialize Excel Editor anyway (will work with CSV files)
      try {
        const app = new ExcelEditor();
        element.excelEditor = app;
      } catch (initError) {
        console.error('Failed to initialize Excel Editor at all:', initError);
      }
    }
  }

  /**
   * Drupal behavior to initialize Excel Editor
   */
  Drupal.behaviors.excelEditor = {
    attach: function (context, settings) {
      once('excel-editor', '.excel-editor-container', context).forEach(
        function (element) {
          initializeExcelEditor(element);
        }
      );
    },
  };
})(jQuery, Drupal, once, drupalSettings);
