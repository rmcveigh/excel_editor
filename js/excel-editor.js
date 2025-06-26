/**
 * @file
 * Excel Editor JavaScript - Enhanced version with better architecture and filter fixes
 *
 * This file contains the main ExcelEditor class which controls all the
 * frontend functionality for the Excel Editor module.
 */

/* eslint-disable no-console */
(function ($, Drupal, once, drupalSettings) {
  'use strict';

  /**
   * The main application class for the Excel Editor.
   * This class encapsulates all the properties and methods needed to run the editor.
   */
  class ExcelEditor {
    // =========================================================================
    // INITIALIZATION & CONFIGURATION
    // =========================================================================

    /**
     * The constructor initializes the application state, configuration,
     * and kicks off the main initialization process.
     */
    constructor() {
      /**
       * Holds the application's data, including original, filtered,
       * and selected data sets.
       * @type {{original: Array, filtered: Array, selected: Set, dirty: boolean}}
       */
      this.data = {
        original: [],
        filtered: [],
        selected: new Set(),
        dirty: false,
      };

      /**
       * Holds the application's current state, such as hidden columns,
       * filters, and loading statuses.
       * @type {{hiddenColumns: Set, currentFilters: {}, isInitialized: boolean, isLoading: boolean, currentProcessLoader: null, currentDraftId: null, currentDraftName: string}}
       */
      this.state = {
        hiddenColumns: new Set(),
        currentFilters: {},
        isInitialized: false,
        isLoading: false,
        currentProcessLoader: null,
        currentDraftId: null,
        currentDraftName: '',
      };

      /**
       * Stores configuration passed from Drupal's settings and some
       * hardcoded values.
       */
      this.config = {
        endpoints: drupalSettings?.excelEditor?.endpoints || {},
        settings: drupalSettings?.excelEditor?.settings || {},
        editableColumns: ['new_barcode', 'notes', 'actions'],
        maxFileSize: 10 * 1024 * 1024, // 10MB
        supportedFormats: ['.xlsx', '.xls', '.csv'],
        autosaveInterval: 5 * 60 * 1000, // 5 minutes
      };

      /**
       * Holds the timer ID for the autosave functionality.
       * @type {number|null}
       */
      this.autosaveTimer = null;

      /**
       * Stores the CSRF token required for authenticated POST/DELETE requests.
       * @type {string|null}
       */
      this.csrfToken = null;
      this.getCsrfToken();

      this.logDebug('Excel Editor config loaded:', this.config);

      /**
       * A cache for frequently accessed DOM elements.
       * @type {Object}
       */
      this.elements = {};

      this.init();
    }

    /**
     * The main initialization method.
     * Checks dependencies, caches elements, binds events, and loads drafts.
     */
    init() {
      try {
        this.checkDependencies();
        this.cacheElements();
        this.hideLoading(); // Ensure loader is hidden on init
        this.bindEvents();
        this.loadDrafts();
        if (this.config.settings.autosave_enabled) {
          this.startAutosave();
        }
        this.state.isInitialized = true;
        this.logDebug('Excel Editor initialized successfully');
      } catch (error) {
        this.handleError('Failed to initialize Excel Editor', error);
      }
    }

    /**
     * Fetches the CSRF token from Drupal's session endpoint.
     * This is a helper function to centralize token fetching.
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
     * Checks if required third-party libraries (like SheetJS) are available.
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
     * Caches jQuery selectors for DOM elements to improve performance.
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
        saveDraftBtn: $('#save-draft-btn'),
        exportBtn: $('#export-btn'),
        exportAllBtn: $('#export-all-btn'),
        toggleColumnsBtn: $('#toggle-columns-btn'),
        selectAllBtn: $('#select-all-visible-btn'),
        deselectAllBtn: $('#deselect-all-btn'),
      };
      this.logDebug('Cached elements:', this.elements);
    }

    // =========================================================================
    // EVENT BINDING
    // =========================================================================

    /**
     * Binds all application-level event handlers.
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
     * Binds events that happen within the main data table, using delegation
     * to handle dynamically added rows.
     */
    bindTableEvents() {
      this.elements.tableContainer.off('.excelEditor');
      this.elements.tableContainer.on(
        'click.excelEditor',
        '.filter-link',
        (e) => {
          e.preventDefault();
          e.stopPropagation();
          this.handleFilterClick(e);
        }
      );
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
    }

    /**
     * Sets up the drag-and-drop file upload area.
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
     * Prevent default drag behaviors.
     */
    preventDefaults(e) {
      e.preventDefault();
      e.stopPropagation();
    }

    // =========================================================================
    // EVENT HANDLERS
    // =========================================================================

    /**
     * Handles the file drop event.
     * @param {DragEvent} e The drop event.
     */
    handleFileDrop(e) {
      const files = e.dataTransfer.files;
      if (files.length > 0) {
        this.processFile(files[0]);
      }
    }

    /**
     * Handles the file input change event.
     * @param {Event} e The change event.
     */
    handleFileUpload(e) {
      const file = e.target.files[0];
      if (file) {
        this.processFile(file);
      }
    }

    /**
     * Handles clicks on the "Filter" link in table headers.
     * @param {Event} e The click event.
     */
    handleFilterClick(e) {
      const $target = $(e.target);
      let columnIndex = $target.data('column');
      if (columnIndex === undefined || columnIndex === null) {
        columnIndex = $target.closest('[data-column]').data('column');
      }
      if (columnIndex !== undefined && columnIndex !== null) {
        this.showColumnFilter(columnIndex);
      }
    }

    /**
     * Handles editing of a cell's value.
     * @param {Event} e The change event from an input/select/textarea.
     */
    handleCellEdit(e) {
      const $cell = $(e.target);
      const rowIndex = parseInt($cell.data('row'));
      const colIndex = parseInt($cell.data('col'));
      const newValue = String($cell.val() || '').trim();

      this.data.filtered[rowIndex][colIndex] = newValue;
      this.data.dirty = true;

      const columnName = this.data.filtered[0][colIndex];
      if (columnName === 'actions') {
        this.applyRowStyling();
      }

      Drupal.debounce(() => {
        this.showMessage(
          'Changes detected. Remember to save your draft.',
          'info',
          3000
        );
      }, 1000)();
    }

    /**
     * Handles row selection via checkbox.
     * @param {Event} e The change event from a row checkbox.
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
     * Handles the "select all" checkbox in the table header.
     * @param {Event} e The change event from the main checkbox.
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
     * Handles keyboard shortcuts.
     * @param {KeyboardEvent} e The keydown event.
     */
    handleKeyboardShortcuts(e) {
      if ((e.ctrlKey || e.metaKey) && e.key === 's') {
        e.preventDefault();
        this.saveDraft();
      }
    }

    /**
     * Shows a confirmation message if there are unsaved changes before leaving the page.
     */
    handleBeforeUnload() {
      if (this.data.dirty) {
        return Drupal.t(
          'You have unsaved changes. Are you sure you want to leave?'
        );
      }
    }

    // =========================================================================
    // FILE HANDLING & PARSING
    // =========================================================================

    /**
     * Orchestrates file processing: validation, reading, and parsing.
     * @param {File} file The file to process.
     */
    async processFile(file) {
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
    }

    /**
     * Validates a file based on size and format.
     * @param {File} file The file to validate.
     * @returns {boolean} True if the file is valid.
     */
    validateFile(file) {
      if (file.size > this.config.maxFileSize) {
        this.showMessage(
          `File too large. Maximum size is ${
            this.config.maxFileSize / (1024 * 1024)
          }MB`,
          'error'
        );
        return false;
      }
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
     * Reads a file into an ArrayBuffer.
     * @param {File} file The file to read.
     * @returns {Promise<ArrayBuffer>}
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
     * Parses CSV data from an ArrayBuffer.
     * @param {ArrayBuffer} data The raw file data.
     * @returns {Array<Array<string>>} The parsed data.
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
     * Parses Excel (.xls, .xlsx) data from an ArrayBuffer using SheetJS.
     * @param {ArrayBuffer} data The raw file data.
     * @returns {Promise<Array<Array<string>>>} The parsed data.
     */
    async parseExcel(data) {
      // Implementation is complex, but well-contained here.
      // It reads the workbook, gets the first sheet, converts to JSON,
      // trims all data, and filters out empty rows.
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
            // Header only or empty
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

    // =========================================================================
    // DATA MANAGEMENT
    // =========================================================================

    /**
     * Loads the parsed data into the application's state.
     * @param {Array<Array<string>>} data The parsed data from a file.
     */
    loadData(data) {
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
    }

    /**
     * Adds the default editable columns ('new_barcode', 'notes', 'actions')
     * to the dataset if they don't already exist.
     */
    addEditableColumns() {
      if (!this.data.original.length) return;
      const headerRow = this.data.original[0];
      if (this.config.editableColumns.some((col) => headerRow.includes(col))) {
        return;
      }
      headerRow.unshift('new_barcode');
      headerRow.push('notes', 'actions');
      for (let i = 1; i < this.data.original.length; i++) {
        this.data.original[i].unshift('');
        this.data.original[i].push('', '');
      }
      this.data.dirty = true;
    }

    /**
     * Applies default column visibility based on configuration.
     * This is called when data is first loaded.
     */
    applyDefaultColumnVisibility() {
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

      this.state.hiddenColumns.clear();
      const headerRow = this.data.filtered[0];
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
    }

    /**
     * Creates a deep clone of an object or array.
     * @param {*} obj The object or array to clone.
     * @returns {*} A deep copy.
     */
    deepClone(obj) {
      return JSON.parse(JSON.stringify(obj));
    }

    // =========================================================================
    // UI RENDERING & UPDATES
    // =========================================================================

    /**
     * Renders the main interface after data is loaded.
     */
    renderInterface() {
      this.elements.uploadArea.hide();
      this.elements.mainArea.show();
      this.renderTable();
      this.setupFilters();
    }

    /**
     * Renders the main data table.
     */
    async renderTable() {
      if (!this.data.filtered.length) {
        this.elements.tableContainer.html(
          '<p class="has-text-centered">No data available</p>'
        );
        return;
      }
      const fragment = document.createDocumentFragment();
      const table = document.createElement('table');
      table.className = 'excel-editor-table table is-fullwidth is-striped';
      table.id = 'excel-table';

      table.appendChild(this.createTableHeader());
      table.appendChild(this.createTableBody());
      fragment.appendChild(table);

      this.elements.tableContainer.html(fragment);
      this.elements.table = $('#excel-table');
      this.bindTableEvents();
      this.applyRowStyling();
    }

    /**
     * Creates the table header (<thead>).
     * @returns {HTMLTableSectionElement} The created thead element.
     */
    createTableHeader() {
      const thead = document.createElement('thead');
      const headerRow = document.createElement('tr');
      headerRow.innerHTML =
        '<th class="selection-column"><label class="checkbox"><input type="checkbox" id="select-all-checkbox" /></label></th>';

      this.data.filtered[0].forEach((header, index) => {
        if (!this.state.hiddenColumns.has(index)) {
          const th = document.createElement('th');
          th.dataset.column = index;
          th.innerHTML = `${this.escapeHtml(
            header
          )}<br><small><a href="#" class="filter-link" data-column="${index}">Filter</a></small>`;
          headerRow.appendChild(th);
        }
      });

      thead.appendChild(headerRow);
      return thead;
    }

    /**
     * Creates the table body (<tbody>).
     * @returns {HTMLTableSectionElement} The created tbody element.
     */
    createTableBody() {
      const tbody = document.createElement('tbody');
      for (let i = 1; i < this.data.filtered.length; i++) {
        tbody.appendChild(this.createTableRow(i));
      }
      return tbody;
    }

    /**
     * Creates an individual table row (<tr>).
     * @param {number} rowIndex The index of the row.
     * @returns {HTMLTableRowElement} The created tr element.
     */
    createTableRow(rowIndex) {
      const row = document.createElement('tr');
      row.dataset.row = rowIndex;
      const rowData = this.data.filtered[rowIndex];
      const isSelected = this.data.selected.has(rowIndex);
      if (isSelected) row.classList.add('selected-row');

      row.innerHTML = `<td class="selection-column"><label class="checkbox"><input type="checkbox" class="row-checkbox" data-row="${rowIndex}" ${
        isSelected ? 'checked' : ''
      } /></label></td>`;

      rowData.forEach((cell, colIndex) => {
        if (!this.state.hiddenColumns.has(colIndex)) {
          row.appendChild(this.createTableCell(rowIndex, colIndex, cell));
        }
      });
      return row;
    }

    /**
     * Creates an individual table cell (<td>).
     * @param {number} rowIndex The row index.
     * @param {number} colIndex The column index.
     * @param {string} cellValue The value of the cell.
     * @returns {HTMLTableCellElement} The created td element.
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

    createActionsDropdown(rowIndex, colIndex, value) {
      const selected = {
        '': !value ? 'selected' : '',
        relabel: value === 'relabel' ? 'selected' : '',
        pending: value === 'pending' ? 'selected' : '',
        discard: value === 'discard' ? 'selected' : '',
      };
      return `<div class="select is-small is-fullwidth"><select class="excel-editor-cell editable actions-dropdown" data-row="${rowIndex}" data-col="${colIndex}"><option value="" ${
        selected['']
      }>${Drupal.t('-- Select Action --')}</option><option value="relabel" ${
        selected['relabel']
      }>${Drupal.t('Relabel')}</option><option value="pending" ${
        selected['pending']
      }>${Drupal.t('Pending')}</option><option value="discard" ${
        selected['discard']
      }>${Drupal.t('Discard')}</option></select></div>`;
    }

    createNotesTextarea(rowIndex, colIndex, value) {
      return `<textarea class="excel-editor-cell editable notes-textarea" data-row="${rowIndex}" data-col="${colIndex}" placeholder="${Drupal.t(
        'Add notes...'
      )}" rows="2">${this.escapeHtml(value || '')}</textarea>`;
    }

    createTextInput(rowIndex, colIndex, value, placeholder) {
      return `<input type="text" class="excel-editor-cell editable" data-row="${rowIndex}" data-col="${colIndex}" value="${this.escapeHtml(
        value || ''
      )}" placeholder="${Drupal.t(placeholder)}" />`;
    }

    /**
     * Gets the value from the 'actions' column for a specific row.
     * @param {number} rowIndex The row index.
     * @returns {string|null} The action value.
     */
    getActionValue(rowIndex) {
      const actionsColumnIndex = this.data.filtered[0].indexOf('actions');
      if (actionsColumnIndex === -1) return null;
      return this.data.filtered[rowIndex][actionsColumnIndex];
    }

    /**
     * Applies CSS classes to rows based on the value in the 'actions' column.
     */
    applyRowStyling() {
      this.elements.table.find('tbody tr').each((index, row) => {
        const $row = $(row);
        const rowIndex = parseInt($row.data('row'));
        const actionValue = this.getActionValue(rowIndex);
        $row.removeClass('action-relabel action-pending action-discard');
        if (actionValue) $row.addClass(`action-${actionValue}`);
      });
    }

    /**
     * Updates the text indicating how many rows are selected.
     */
    updateSelectionCount() {
      const count = this.data.selected.size;
      this.elements.selectionCount.text(
        `${count} row${count !== 1 ? 's' : ''} selected`
      );
      this.elements.exportBtn
        .prop('disabled', count === 0)
        .toggleClass('is-disabled', count === 0);
    }

    /**
     * Updates the state of the "select all" checkbox.
     */
    updateSelectAllCheckbox() {
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
    }

    // =========================================================================
    // SELECTION
    // =========================================================================

    selectAllVisible() {
      this.elements.tableContainer
        .find('.row-checkbox:visible')
        .each((index, checkbox) => {
          const $checkbox = $(checkbox);
          if (!$checkbox.is(':checked')) {
            $checkbox.prop('checked', true).trigger('change');
          }
        });
    }

    deselectAll() {
      this.data.selected.clear();
      this.elements.tableContainer.find('.row-checkbox').prop('checked', false);
      this.elements.tableContainer.find('tr').removeClass('selected-row');
      this.updateSelectionCount();
      this.updateSelectAllCheckbox();
    }

    // =========================================================================
    // FILTERING
    // =========================================================================

    /**
     * Sets up the filter controls area above the table.
     */
    setupFilters() {
      if (!this.data.filtered.length) return;
      let statusMessages = '';
      if (this.state.hiddenColumns.size > 0) {
        statusMessages += `<div class="field"><div class="notification is-info is-light"><span class="icon"><i class="fas fa-eye-slash"></i></span> ${
          this.state.hiddenColumns.size
        } column${
          this.state.hiddenColumns.size !== 1 ? 's' : ''
        } hidden. <button class="button is-small is-light ml-2" id="show-column-settings"><span>Manage Columns</span></button></div></div>`;
      }
      if (
        this.config.settings.hideBehavior === 'hide_others' &&
        this.config.settings.defaultVisibleColumns?.length > 0
      ) {
        statusMessages += `<div class="field"><div class="notification is-primary is-light"><span class="icon"><i class="fas fa-cog"></i></span> Default column visibility applied. <button class="button is-small is-light ml-2" id="reset-to-defaults"><span>Reset to Defaults</span></button> <button class="button is-small is-light ml-2" id="show-all-override"><span>Show All</span></button></div></div>`;
      }
      this.elements.filtersContainer.html(
        `${statusMessages} <div class="field mb-2" id="active-filters-container" style="display: none;"><label class="label">Active Filters:</label><div class="control" id="active-filters"></div><div class="control mt-2"><button class="button is-small is-light" id="clear-all-filters-btn"><span>Clear All Filters</span></button></div></div>`
      );
      this.bindFilterEvents();
    }

    /**
     * Binds events for the filter control area.
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
     * Shows the filter modal for a specific column.
     * @param {number} columnIndex The index of the column to filter.
     */
    showColumnFilter(columnIndex) {
      this.showQuickLoader('Loading filter options...');
      setTimeout(() => {
        try {
          const header = this.data.filtered[0][columnIndex];
          const uniqueValues = this.getUniqueColumnValues(columnIndex);
          const modalHtml = this._buildFilterModalHtml(
            header,
            uniqueValues,
            columnIndex
          );
          $('body').append(modalHtml);
          this.bindFilterModalEvents($('#filter-modal'), columnIndex, header);
          this.updateFilterSelectedCount($('#filter-modal'));
        } finally {
          this.hideQuickLoader();
        }
      }, 50);
    }

    /**
     * [HELPER] Builds the HTML string for the filter modal.
     * @param {string} header The column header text.
     * @param {Array<string>} uniqueValues The unique values in the column.
     * @param {number} columnIndex The index of the column.
     * @returns {string} The complete HTML string for the modal.
     */
    _buildFilterModalHtml(header, uniqueValues, columnIndex) {
      const checkboxOptionsHtml = uniqueValues
        .map((val) => {
          const value = val || '';
          const displayValue = value === '' ? '(empty)' : value;
          const isChecked = this.isValueSelectedInFilter(columnIndex, value);
          return `<div class="column is-half"><label class="checkbox filter-checkbox-item"><input type="checkbox" value="${this.escapeHtml(
            value
          )}" ${
            isChecked ? 'checked' : ''
          } class="filter-value-checkbox mr-1"><span class="filter-checkbox-label">${this.escapeHtml(
            displayValue
          )}</span></label></div>`;
        })
        .join('');

      return `<div class="modal is-active" id="filter-modal" style="display: flex !important; z-index: 99999;">
                <div class="modal-background"></div>
                <div class="modal-content">
                  <div class="box">
                    <h3 class="title is-4"><span class="icon"><i class="fas fa-filter"></i></span> Filter: ${this.escapeHtml(
                      header
                    )}</h3>
                    <p class="subtitle is-6">Column ${columnIndex + 1} with ${
        uniqueValues.length
      } unique values</p>
                    <div class="field is-grouped is-grouped-multiline mb-3">
                      <div class="control">
                        <button class="button is-small is-info" id="select-all-values">Select All</button>
                      </div>
                      <div class="control">
                        <button class="button is-small is-light" id="deselect-all-values">Deselect All</button>
                      </div>
                       <div class="control">
                        <button class="button is-small" id="invert-selection">Invert</button>
                      </div>
                    </div>
                    <div class="field">
                        <input class="input is-small" type="text" id="filter-search" placeholder="Search values..." />
                    </div>
                    <div class="filter-values-container">
                      <div class="columns is-multiline" id="filter-checkboxes">${checkboxOptionsHtml}</div>
                    </div>
                    <p class="help mt-2"><span id="selected-count">0</span> of ${
                      uniqueValues.length
                    } values selected</p>
                    <div class="field is-grouped is-grouped-right">
                      <div class="control"><button class="button" id="clear-column-filter">Clear Filter</button></div>
                      <div class="control"><button class="button" id="cancel-filter">Cancel</button></div>
                      <div class="control"><button class="button is-primary" id="apply-filter">Apply Filter</button></div>
                    </div>
                  </div>
                </div>
                <button class="modal-close is-large" aria-label="close"></button>
              </div>`;
    }

    /**
     * Binds events for the filter modal (buttons, search, etc.).
     * @param {jQuery} modal The jQuery object for the modal.
     * @param {number} columnIndex The column index being filtered.
     * @param {string} header The name of the column.
     */
    bindFilterModalEvents(modal, columnIndex, header) {
      modal
        .find('.modal-close, #cancel-filter, .modal-background')
        .on('click', () => modal.remove());

      modal.find('#select-all-values').on('click', () => {
        modal.find('.filter-value-checkbox').prop('checked', true);
        this.updateFilterSelectedCount(modal);
      });

      modal.find('#deselect-all-values').on('click', () => {
        modal.find('.filter-value-checkbox').prop('checked', false);
        this.updateFilterSelectedCount(modal);
      });

      modal.find('#invert-selection').on('click', () => {
        modal.find('.filter-value-checkbox').each(function () {
          $(this).prop('checked', !$(this).prop('checked'));
        });
        this.updateFilterSelectedCount(modal);
      });

      modal.find('#filter-search').on('input', (e) => {
        const searchTerm = $(e.target).val().toLowerCase();
        modal.find('.filter-checkbox-item').each(function () {
          const label = $(this)
            .find('.filter-checkbox-label')
            .text()
            .toLowerCase();
          $(this).closest('.column').toggle(label.includes(searchTerm));
        });
      });

      modal.find('.filter-value-checkbox').on('change', () => {
        this.updateFilterSelectedCount(modal);
      });

      modal.find('#clear-column-filter').on('click', async () => {
        delete this.state.currentFilters[columnIndex];
        await this.applyFilters();
        this.updateActiveFiltersDisplay();
        modal.remove();
      });

      modal.find('#apply-filter').on('click', async () => {
        await this.applyFilterFromModal(modal, columnIndex);
        modal.remove();
      });
    }

    /**
     * Applies the filter settings from the filter modal.
     * @param {jQuery} modal The modal element.
     * @param {number} columnIndex The column index being filtered.
     */
    async applyFilterFromModal(modal, columnIndex) {
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

      await this.applyFilters();
      this.updateActiveFiltersDisplay();
    }

    /**
     * Applies all active filters to the original data to create the filtered dataset.
     */
    async applyFilters() {
      if (Object.keys(this.state.currentFilters).length === 0) {
        this.data.filtered = this.deepClone(this.data.original);
      } else {
        this.data.filtered = [this.data.original[0]]; // Keep header
        for (let i = 1; i < this.data.original.length; i++) {
          const row = this.data.original[i];
          if (
            Object.entries(this.state.currentFilters).every(
              ([colIndex, filter]) =>
                this.rowMatchesFilter(row[parseInt(colIndex)], filter)
            )
          ) {
            this.data.filtered.push(row);
          }
        }
      }
      this.data.selected.clear();
      await this.renderTable();
      this.updateSelectionCount();
    }

    /**
     * Checks if a single cell's value matches a given filter.
     * @param {string} cellValue The value of the cell.
     * @param {object} filter The filter object.
     * @returns {boolean} True if the cell matches the filter.
     */
    rowMatchesFilter(cellValue, filter) {
      const value = String(cellValue || '')
        .trim()
        .toLowerCase();
      switch (filter.type) {
        case 'quick':
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
     * Updates the display of active filters above the table.
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
          const header = this.data.original[0][parseInt(columnIndex)];
          const filterDescription = this.getFilterDescription(filter);
          return `<span class="tag is-info"><strong>${this.escapeHtml(
            header
          )}</strong>: ${filterDescription}<button class="delete is-small ml-1" data-column="${columnIndex}"></button></span>`;
        })
        .join(' ');

      filtersContainer.html(filterTags);
      containerWrapper.show();

      filtersContainer.find('.delete').on('click', (e) => {
        const columnIndex = $(e.target).data('column');
        delete this.state.currentFilters[columnIndex];
        this.applyFilters();
        this.updateActiveFiltersDisplay();
      });
    }

    /**
     * Gets a human-readable description of a filter.
     * @param {object} filter The filter object.
     * @returns {string} The filter description.
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
     * Clears all active filters and re-renders the table.
     */
    clearAllFilters() {
      this.state.currentFilters = {};
      this.applyFilters();
      this.updateActiveFiltersDisplay();
      this.showMessage('All filters cleared', 'success');
    }

    /**
     * Gets unique values for a column to populate the filter modal.
     * @param {number} columnIndex The index of the column.
     * @returns {Array<string>} A sorted array of unique values.
     */
    getUniqueColumnValues(columnIndex) {
      const values = new Set();
      for (let i = 1; i < this.data.original.length; i++) {
        const rawValue = this.data.original[i][columnIndex];
        values.add(String(rawValue || '').trim());
      }
      return Array.from(values).sort();
    }

    /**
     * Checks if a value is currently selected in a column's filter.
     * @param {number} columnIndex The index of the column.
     * @param {string} value The value to check.
     * @returns {boolean} True if the value is selected (or if there's no filter).
     */
    isValueSelectedInFilter(columnIndex, value) {
      if (!this.state.currentFilters[columnIndex]) {
        return true; // If no filter, all are considered selected
      }
      const filter = this.state.currentFilters[columnIndex];
      if (filter.type === 'quick' && filter.selected) {
        return filter.selected.includes(value);
      }
      return false;
    }

    /**
     * Updates the selected count display in the filter modal.
     * @param {jQuery} modal The jQuery object for the modal.
     */
    updateFilterSelectedCount(modal) {
      const checkedBoxes = modal.find('.filter-value-checkbox:checked');
      modal.find('#selected-count').text(checkedBoxes.length);
    }

    // =========================================================================
    // COLUMN VISIBILITY
    // =========================================================================

    /**
     * Shows the modal for managing column visibility.
     */
    showColumnVisibilityModal() {
      if (!this.data.filtered.length) {
        this.showMessage('No data loaded', 'warning');
        return;
      }

      const modalHtml = this._buildColumnVisibilityModalHtml();
      const modal = $(modalHtml);
      $('body').append(modal);

      this.bindColumnModalEvents(modal);
    }

    /**
     * [HELPER] Builds the HTML string for the column visibility modal.
     * @returns {string} The complete HTML string for the modal.
     */
    _buildColumnVisibilityModalHtml() {
      const headers = this.data.filtered[0];
      const checkboxesHtml = headers
        .map((header, index) => {
          const isVisible = !this.state.hiddenColumns.has(index);
          const isEditable = this.config.editableColumns.includes(header);
          return `<div class="column is-half"><label class="checkbox"><input type="checkbox" class="column-visibility-checkbox" data-column-index="${index}" ${
            isVisible ? 'checked' : ''
          }><span class="column-name">${this.escapeHtml(header)}</span>${
            isEditable
              ? `<span class="tag is-small is-info ml-2">${Drupal.t(
                  'Editable'
                )}</span>`
              : ''
          }</label></div>`;
        })
        .join('');

      return `<div class="modal is-active" id="column-visibility-modal"><div class="modal-background"></div><div class="modal-content"><div class="box"><h3 class="title is-4">Manage Column Visibility</h3><div class="field is-grouped"><button class="button is-small" id="show-all-columns">Show All</button><button class="button is-small" id="show-only-editable">Show Only Editable</button></div><div class="column-checkboxes columns is-multiline">${checkboxesHtml}</div><div class="field is-grouped is-grouped-right"><button class="button" id="cancel-column-visibility">Cancel</button><button class="button is-primary" id="apply-column-visibility">Apply</button></div></div></div><button class="modal-close is-large" aria-label="close"></button></div>`;
    }

    /**
     * Binds events for the column visibility modal.
     * @param {jQuery} modal The jQuery object for the modal.
     */
    bindColumnModalEvents(modal) {
      modal
        .find('.modal-close, #cancel-column-visibility, .modal-background')
        .on('click', () => modal.remove());
      modal
        .find('#show-all-columns')
        .on('click', () =>
          modal.find('.column-visibility-checkbox').prop('checked', true)
        );
      modal.find('#show-only-editable').on('click', () => {
        modal.find('.column-visibility-checkbox').each((index, checkbox) => {
          const colIndex = parseInt($(checkbox).data('column-index'));
          const header = this.data.filtered[0][colIndex];
          $(checkbox).prop(
            'checked',
            this.config.editableColumns.includes(header)
          );
        });
      });
      modal.find('#apply-column-visibility').on('click', () => {
        this.applyColumnVisibilityChanges(modal);
        modal.remove();
      });
    }

    /**
     * Applies the column visibility changes selected in the modal.
     * @param {jQuery} modal The jQuery object for the modal.
     */
    async applyColumnVisibilityChanges(modal) {
      this.showProcessLoader('Updating column visibility...');
      await new Promise((resolve) => setTimeout(resolve, 50));
      try {
        this.state.hiddenColumns.clear();
        modal.find('.column-visibility-checkbox').each((index, checkbox) => {
          if (!$(checkbox).is(':checked')) {
            this.state.hiddenColumns.add(
              parseInt($(checkbox).data('column-index'))
            );
          }
        });
        await this.renderTable();
        this.setupFilters();
      } finally {
        this.hideProcessLoader();
      }
    }

    /**
     * Resets the visible columns to the defaults specified in the module settings.
     */
    resetToDefaultColumns() {
      if (!this.data.filtered.length) return;
      this.applyDefaultColumnVisibility();
      this.renderTable();
      this.setupFilters();
      this.showMessage('Columns reset to default visibility.', 'success');
    }

    /**
     * Overrides any settings and makes all columns visible.
     */
    showAllColumnsOverride() {
      this.state.hiddenColumns.clear();
      this.renderTable();
      this.setupFilters();
      this.showMessage('All columns are now visible.', 'success');
    }

    // =========================================================================
    // API & DRAFT MANAGEMENT
    // =========================================================================

    /**
     * Starts the autosave timer.
     */
    startAutosave() {
      this.stopAutosave(); // Clear any existing timer
      this.autosaveTimer = setInterval(() => {
        this.autosaveDraft();
      }, this.config.autosaveInterval);
      this.logDebug(
        `Autosave started with interval: ${this.config.autosaveInterval}ms`
      );
    }

    /**
     * Stops the autosave timer.
     */
    stopAutosave() {
      if (this.autosaveTimer) {
        clearInterval(this.autosaveTimer);
        this.autosaveTimer = null;
        this.logDebug('Autosave stopped.');
      }
    }

    /**
     * Performs the autosave operation.
     */
    async autosaveDraft() {
      if (!this.data.dirty || !this.state.currentDraftId) {
        return;
      }

      this.logDebug(`Autosaving draft ID: ${this.state.currentDraftId}`);
      this.showQuickLoader('Autosaving...');

      try {
        const draftData = {
          draft_id: this.state.currentDraftId,
          name: this.state.currentDraftName,
          data: this.data.original,
          filters: this.state.currentFilters,
          hiddenColumns: Array.from(this.state.hiddenColumns),
          selected: Array.from(this.data.selected),
          timestamp: new Date().toISOString(),
        };

        const response = await this.apiCall(
          'POST',
          this.config.endpoints.saveDraft,
          draftData
        );

        if (response.success) {
          this.data.dirty = false;
          this.showQuickLoader('Draft autosaved.', 'success');
          setTimeout(() => this.hideQuickLoader(), 2000);
          this.loadDrafts();
        } else {
          throw new Error(response.message || 'Autosave failed');
        }
      } catch (error) {
        this.handleError('Autosave failed', error);
      }
    }

    /**
     * [HELPER] Creates and displays a modal to prompt the user for a draft name.
     * @returns {Promise<string|null>} A promise that resolves with the draft name, or null if canceled.
     */
    _promptForDraftName() {
      return new Promise((resolve) => {
        $('.modal#save-draft-modal').remove();

        const defaultName = `Draft ${new Date().toLocaleString()}`;
        const modalHtml = `
          <div class="modal is-active" id="save-draft-modal">
            <div class="modal-background"></div>
            <div class="modal-content">
              <div class="box">
                <h3 class="title is-4">Save Draft</h3>
                <div class="field">
                  <label class="label">Draft Name</label>
                  <div class="control">
                    <input class="input" id="draft-name-input" type="text" value="${this.escapeHtml(
                      defaultName
                    )}">
                  </div>
                </div>
                <div class="field is-grouped is-grouped-right">
                  <div class="control">
                    <button class="button" id="cancel-save-draft">Cancel</button>
                  </div>
                  <div class="control">
                    <button class="button is-primary" id="confirm-save-draft">Save</button>
                  </div>
                </div>
              </div>
            </div>
          </div>`;

        const modal = $(modalHtml);
        $('body').append(modal);
        const nameInput = modal.find('#draft-name-input').focus().select();

        const closeModal = () => modal.remove();

        modal.find('#confirm-save-draft').on('click', () => {
          const draftName = nameInput.val().trim();
          if (draftName) {
            resolve(draftName);
            closeModal();
          } else {
            nameInput.addClass('is-danger');
          }
        });

        modal.find('#cancel-save-draft, .modal-background').on('click', () => {
          resolve(null);
          closeModal();
        });
      });
    }

    /**
     * Shows a modal to ask the user for a draft name before saving manually.
     */
    async saveDraft() {
      if (!this.data.original.length) {
        this.showMessage('No data to save', 'warning');
        return;
      }
      const draftName = await this._promptForDraftName();
      if (draftName) {
        this.showLoading('Saving draft...');
        try {
          const draftData = {
            name: draftName,
            data: this.data.original,
            filters: this.state.currentFilters,
            hiddenColumns: Array.from(this.state.hiddenColumns),
            selected: Array.from(this.data.selected),
            timestamp: new Date().toISOString(),
          };
          const response = await this.apiCall(
            'POST',
            this.config.endpoints.saveDraft,
            draftData
          );
          if (response.success && response.draft_id) {
            this.data.dirty = false;
            this.state.currentDraftId = response.draft_id;
            this.state.currentDraftName = draftName;
            this.showMessage(
              `Draft "${this.escapeHtml(draftName)}" saved successfully`,
              'success'
            );
            this.loadDrafts();
          } else {
            throw new Error(response.message || 'Failed to save draft');
          }
        } catch (error) {
          this.handleError('Failed to save draft', error);
        } finally {
          this.hideLoading();
        }
      }
    }

    /**
     * A centralized helper function for making API calls to the Drupal backend.
     * @param {string} method The HTTP method (GET, POST, DELETE).
     * @param {string} url The API endpoint URL.
     * @param {Object|null} data The data to send in the request body.
     * @returns {Promise<Object>} The JSON response from the server.
     */
    async apiCall(method, url, data = null) {
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
        credentials: 'same-origin',
      };
      if (
        (method === 'POST' || method === 'PUT' || method === 'DELETE') &&
        this.csrfToken
      ) {
        options.headers['X-CSRF-Token'] = this.csrfToken;
      }
      if (data) {
        options.body = JSON.stringify(data);
      }
      const response = await fetch(url, options);
      if (!response.ok) {
        let errorMessage = `HTTP ${response.status}: ${response.statusText}`;
        try {
          const errorData = await response.json();
          if (errorData.message) errorMessage = errorData.message;
        } catch (e) {
          // Ignore if response is not JSON
        }
        throw new Error(errorMessage);
      }
      return response.json();
    }

    /**
     * Loads a list of the user's drafts from the server.
     */
    async loadDrafts() {
      try {
        const response = await this.apiCall(
          'GET',
          this.config.endpoints.listDrafts
        );
        if (response.success && response.drafts) {
          this.renderDrafts(response.drafts);
        }
      } catch (error) {
        this.logDebug('Failed to load drafts:', error);
      }
    }

    /**
     * Renders the list of drafts into the UI.
     * @param {Array} drafts The array of draft objects from the server.
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
          (draft) =>
            `<div class="excel-editor-draft-item"><div><strong>${this.escapeHtml(
              draft.name || 'Untitled Draft'
            )}</strong><br><small class="has-text-grey">${new Date(
              draft.changed * 1000
            ).toLocaleString()}</small></div><div class="field is-grouped"><div class="control"><button class="button is-small is-info load-draft-btn" data-draft-id="${
              draft.id
            }"><span>${Drupal.t(
              'Load'
            )}</span></button></div><div class="control"><button class="button is-small is-danger delete-draft-btn" data-draft-id="${
              draft.id
            }"><span class="icon is-small"><i class="fas fa-trash"></i></span></button></div></div></div>`
        )
        .join('');
      this.elements.draftsContainer.html(draftsHtml);
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
     * Loads the data from a specific draft into the editor.
     * @param {number} draftId The ID of the draft to load.
     */
    async loadDraft(draftId) {
      this.showLoading('Loading draft...');
      try {
        const response = await this.apiCall(
          'GET',
          `${this.config.endpoints.loadDraft}${draftId}`
        );
        if (response.success && response.data) {
          this.state.currentDraftId = response.id;
          this.state.currentDraftName = response.name;
          this.loadDraftData(response.data);
          this.showMessage(
            `Draft "${this.escapeHtml(response.name)}" loaded successfully`,
            'success'
          );
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
     * Loads draft data into the application state.
     * @param {object} draftData The draft data object from the server.
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
     * Deletes a draft from the server.
     * @param {number} draftId The ID of the draft to delete.
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
          this.loadDrafts();
        } else {
          throw new Error(response.message || 'Failed to delete draft');
        }
      } catch (error) {
        this.handleError('Failed to delete draft', error);
      }
    }

    // =========================================================================
    // EXPORTING
    // =========================================================================

    /**
     * Exports only the selected rows to an XLSX file.
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
     * Exports all visible (filtered) rows to an XLSX file.
     */
    async exportAll() {
      try {
        const exportData = this.prepareExportData(false);
        await this.downloadExport(exportData, 'all_data.xlsx');
        const rowCount = this.data.filtered.length - 1;
        this.showMessage(`Exported ${rowCount} rows`, 'success');
      } catch (error) {
        this.handleError('Failed to export all data', error);
      }
    }

    /**
     * Prepares the data for export by creating an array of arrays.
     * @param {boolean} selectedOnly Whether to include only selected rows.
     * @returns {Array<Array<string>>} The data ready for export.
     */
    prepareExportData(selectedOnly = false) {
      const exportData = [];
      const headerRow = [];
      this.data.filtered[0].forEach((header, index) => {
        if (!this.state.hiddenColumns.has(index)) {
          headerRow.push(header);
        }
      });
      exportData.push(headerRow);
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
    }

    /**
     * Triggers the download of the exported data as an XLSX file.
     * @param {Array<Array<string>>} data The data to export.
     * @param {string} filename The name of the file to download.
     */
    async downloadExport(data, filename) {
      const ws = XLSX.utils.aoa_to_sheet(data);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Export');
      XLSX.writeFile(wb, filename);
    }

    // =========================================================================
    // UTILITIES (Loaders, Messages, Error Handling, etc.)
    // =========================================================================

    /**
     * Shows the main loading indicator.
     * @param {string} message The message to display.
     */
    showLoading(message = 'Loading...') {
      this.state.isLoading = true;
      this.elements.loadingArea.find('p').text(message);
      this.elements.loadingArea.addClass('active').show();
    }

    /**
     * Hides the main loading indicator.
     */
    hideLoading() {
      this.state.isLoading = false;
      if (this.elements.loadingArea && this.elements.loadingArea.length > 0) {
        this.elements.loadingArea.removeClass('active').hide();
      }
      $('.excel-editor-loading').removeClass('active').hide();
    }

    /**
     * Shows a temporary process loader for intensive operations.
     * @param {string} message The message to display.
     */
    showProcessLoader(message = 'Processing...') {
      this.hideProcessLoader();
      const loaderId = 'process-loader-' + Date.now();
      const loader = $(
        `<div class="excel-editor-overlay-loader" id="${loaderId}"><div class="loading-content"><div class="excel-editor-spinner"></div><p><strong>${this.escapeHtml(
          message
        )}</strong></p></div></div>`
      );
      $('body').append(loader);
      this.state.currentProcessLoader = loaderId;
    }

    /**
     * Hides the process loader.
     */
    hideProcessLoader() {
      if (this.state.currentProcessLoader) {
        $(`#${this.state.currentProcessLoader}`).remove();
        this.state.currentProcessLoader = null;
      }
      $('.excel-editor-overlay-loader').remove();
    }

    /**
     * Shows a small, temporary loader in the corner of the screen.
     * @param {string} message The message to display.
     */
    showQuickLoader(message = 'Working...') {
      this.hideQuickLoader();
      const loader = $(
        `<div class="excel-editor-quick-loader" id="quick-loader"><div class="spinner"></div><span>${this.escapeHtml(
          message
        )}</span></div>`
      );
      $('body').append(loader);
      setTimeout(() => loader.addClass('slide-in'), 10);
    }

    /**
     * Hides the quick loader.
     */
    hideQuickLoader() {
      const loader = $('#quick-loader');
      if (loader.length) {
        loader.addClass('slide-out');
        setTimeout(() => loader.remove(), 300);
      }
    }

    /**
     * Displays a notification message to the user.
     * @param {string} message The message content.
     * @param {string} type The type of message (success, error, warning, info).
     * @param {number} duration How long to display the message in ms.
     */
    showMessage(message, type = 'info', duration = 5000) {
      const alertClass =
        {
          success: 'is-success',
          error: 'is-danger',
          warning: 'is-warning',
          info: 'is-info',
        }[type] || 'is-info';
      const messageElement = $(
        `<div class="notification ${alertClass} excel-editor-message"><button class="delete"></button>${this.escapeHtml(
          message
        )}</div>`
      );
      this.elements.container.prepend(messageElement);
      messageElement
        .find('.delete')
        .on('click', () =>
          messageElement.fadeOut(() => messageElement.remove())
        );
      if (duration > 0) {
        setTimeout(
          () => messageElement.fadeOut(() => messageElement.remove()),
          duration
        );
      }
    }

    /**
     * Centralized error handler. Logs to console and shows a user message.
     * @param {string} message The user-facing message.
     * @param {Error|null} error The caught error object.
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
     * Centralized debug logger. Only logs if debug mode is enabled.
     * @param {string} message The debug message.
     * @param {*} [data=null] Additional data to log.
     */
    logDebug(message, data = null) {
      const isDebugMode =
        this.config.settings.debug ||
        window.location.search.includes('debug=1') ||
        localStorage.getItem('excel_editor_debug') === 'true';
      if (isDebugMode) {
        console.log(`[Excel Editor] ${message}`, data);
      }
    }

    /**
     * Escapes HTML to prevent XSS vulnerabilities.
     * @param {string} text The text to escape.
     * @returns {string} The escaped HTML string.
     */
    escapeHtml(text) {
      const div = document.createElement('div');
      div.textContent = text;
      return div.innerHTML;
    }
  } // End of ExcelEditor class

  /**
   * Drupal behavior to initialize the Excel Editor on the page.
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

  /**
   * Helper to initialize the editor, including loading SheetJS if needed.
   * @param {HTMLElement} element The container element for the editor.
   */
  async function initializeExcelEditor(element) {
    try {
      const loadingMsg = $(
        `<div class="notification is-info excel-editor-init-loading"><span class="icon"><i class="fas fa-spinner fa-spin"></i></span> Loading Excel processing library...</div>`
      );
      $(element).prepend(loadingMsg);
      await loadXLSXLibrary();
      loadingMsg.remove();
      const app = new ExcelEditor();
      element.excelEditor = app;
    } catch (error) {
      console.error('Failed to initialize Excel Editor:', error);
      $('.excel-editor-init-loading').remove();
      $(element).prepend(
        `<div class="notification is-warning"><button class="delete"></button><strong>Excel Library Loading Issue:</strong> ${error.message}</div>`
      );
    }
  }

  /**
   * Dynamically loads the XLSX library from a CDN if it's not already present.
   * @returns {Promise<void>}
   */
  function loadXLSXLibrary() {
    return new Promise((resolve, reject) => {
      if (typeof XLSX !== 'undefined') {
        resolve();
        return;
      }
      const script = document.createElement('script');
      script.src =
        'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
      script.onload = resolve;
      script.onerror = () =>
        reject(new Error('Failed to load XLSX library from CDN.'));
      document.head.appendChild(script);
    });
  }
})(jQuery, Drupal, once, drupalSettings);
