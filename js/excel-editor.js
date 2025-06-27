/**
 * @file
 * Excel Editor JavaScript - Main class and initialization
 *
 * This is the core class that orchestrates all Excel Editor functionality.
 * Individual features are implemented in separate module files.
 */

/* eslint-disable no-console */
(function ($, Drupal, once, drupalSettings) {
  'use strict';

  /**
   * The main application class for the Excel Editor.
   * This class coordinates all the modules and manages application state.
   */
  class ExcelEditor {
    /**
     * Constructor initializes the application state and modules.
     */
    constructor() {
      // Application data state
      this.data = {
        original: [],
        filtered: [],
        selected: new Set(),
        dirty: false,
      };

      // Application UI and filter state
      this.state = {
        hiddenColumns: new Set(),
        currentFilters: {},
        isInitialized: false,
        isLoading: false,
        currentProcessLoader: null,
        currentDraftId: null,
        currentDraftName: '',
      };

      // Configuration from Drupal settings
      this.config = {
        endpoints: drupalSettings?.excelEditor?.endpoints || {},
        settings: drupalSettings?.excelEditor?.settings || {},
        editableColumns: ['new_barcode', 'notes', 'actions'],
        maxFileSize: 10 * 1024 * 1024, // 10MB
        supportedFormats: ['.xlsx', '.xls', '.csv'],
        autosaveInterval: 5 * 60 * 1000, // 5 minutes
      };

      // Autosave timer
      this.autosaveTimer = null;

      // CSRF token for API calls
      this.csrfToken = null;

      // DOM element cache
      this.elements = {};

      // Initialize modules - FIXED ORDER: BarcodeSystem before DataManager
      if (typeof ExcelEditorUtilities !== 'undefined') {
        ExcelEditorUtilities.call(this);
      }
      if (typeof ExcelEditorBarcodeSystem !== 'undefined') {
        ExcelEditorBarcodeSystem.call(this);
      }
      if (typeof ExcelEditorDataManager !== 'undefined') {
        ExcelEditorDataManager.call(this);
      }
      if (typeof ExcelEditorUIRenderer !== 'undefined') {
        ExcelEditorUIRenderer.call(this);
      }
      if (typeof ExcelEditorFilterManager !== 'undefined') {
        ExcelEditorFilterManager.call(this);
      }
      if (typeof ExcelEditorColumnManager !== 'undefined') {
        ExcelEditorColumnManager.call(this);
      }
      if (typeof ExcelEditorValidationManager !== 'undefined') {
        ExcelEditorValidationManager.call(this);
      }
      if (typeof ExcelEditorDraftManager !== 'undefined') {
        ExcelEditorDraftManager.call(this);
      }
      if (typeof ExcelEditorExportManager !== 'undefined') {
        ExcelEditorExportManager.call(this);
      }

      // Start initialization
      this.init();
    }

    /**
     * Main initialization method.
     */
    init() {
      try {
        this.checkDependencies();
        this.cacheElements();
        this.getCsrfToken();
        this.hideLoading();
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
     * Check if required dependencies are available.
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

      // Check if all modules loaded
      const requiredModules = [
        'ExcelEditorUtilities',
        'ExcelEditorDataManager',
        'ExcelEditorBarcodeSystem',
        'ExcelEditorUIRenderer',
        'ExcelEditorFilterManager',
        'ExcelEditorColumnManager',
        'ExcelEditorValidationManager',
        'ExcelEditorDraftManager',
        'ExcelEditorExportManager',
      ];

      requiredModules.forEach((moduleName) => {
        if (typeof window[moduleName] === 'undefined') {
          missing.push(`${moduleName} module`);
        }
      });

      if (missing.length > 0) {
        throw new Error(`Missing required dependencies: ${missing.join(', ')}`);
      }

      this.logDebug('All dependencies and modules loaded successfully');
    }

    /**
     * Cache DOM elements for performance.
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
        resetBarcodesBtn: $('#reset-barcodes-btn'),
        validateBarcodesBtn: $('#validate-barcodes-btn'),
      };
      this.logDebug('Cached DOM elements');
    }

    /**
     * Bind all application events.
     */
    bindEvents() {
      // File upload events
      this.elements.fileInput.on('change', (e) => this.handleFileUpload(e));
      this.setupDragDropUpload();

      // Toolbar button events
      this.elements.saveDraftBtn.on('click', () => this.saveDraft());
      this.elements.exportBtn.on('click', () => this.exportSelected());
      this.elements.exportAllBtn.on('click', () => this.exportAll());
      this.elements.toggleColumnsBtn.on('click', () =>
        this.showColumnVisibilityModal()
      );
      this.elements.selectAllBtn.on('click', () => this.selectAllVisible());
      this.elements.deselectAllBtn.on('click', () => this.deselectAll());
      this.elements.resetBarcodesBtn.on('click', () => this.resetBarcodes());
      this.elements.validateBarcodesBtn.on('click', () => this.showValidationSummary());

      // Table events
      this.bindTableEvents();

      // Keyboard shortcuts
      $(document).on('keydown', (e) => this.handleKeyboardShortcuts(e));

      // Window events
      $(window).on('beforeunload', () => this.handleBeforeUnload());
    }

    /**
     * Bind table-specific events using delegation.
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
        (e) => {
          this.handleCellEdit(e);
        }
      );

      this.elements.tableContainer.on(
        'change.excelEditor',
        '.row-checkbox',
        (e) => {
          this.handleRowSelection(e);
        }
      );

      this.elements.tableContainer.on(
        'change.excelEditor',
        '#select-all-checkbox',
        (e) => {
          this.handleSelectAllCheckbox(e);
        }
      );
    }

    /**
     * Set up drag and drop file upload.
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
    // EVENT HANDLERS (Core events, modules handle their specific events)
    // =========================================================================

    /**
     * Handle file drop event.
     */
    handleFileDrop(e) {
      const files = e.dataTransfer.files;
      if (files.length > 0) {
        this.processFile(files[0]);
      }
    }

    /**
     * Handle file input change.
     */
    handleFileUpload(e) {
      const file = e.target.files[0];
      if (file) {
        this.processFile(file);
      }
    }

    /**
     * Handle filter link clicks.
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
     * Handle cell editing.
     */
    handleCellEdit(e) {
      const $cell = $(e.target);
      const rowIndex = parseInt($cell.data('row'));
      const colIndex = parseInt($cell.data('col'));
      const newValue = String($cell.val() || '').trim();

      this.data.filtered[rowIndex][colIndex] = newValue;
      this.data.dirty = true;

      const columnName = this.data.filtered[0][colIndex];

      // Validate new_barcode fields
      if (columnName === 'new_barcode') {
        this.validateBarcodeField($cell, newValue, rowIndex);
      }

      if (columnName === 'actions') {
        this.applyRowStyling();
      }
    }

    /**
     * Handle row selection via checkbox.
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
     * Handle select all checkbox.
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
     * Handle keyboard shortcuts.
     */
    handleKeyboardShortcuts(e) {
      if ((e.ctrlKey || e.metaKey) && e.key === 's') {
        e.preventDefault();
        this.saveDraft();
      }
    }

    /**
     * Handle before page unload.
     */
    handleBeforeUnload() {
      if (this.data.dirty) {
        return Drupal.t(
          'You have unsaved changes. Are you sure you want to leave?'
        );
      }
    }

    // =========================================================================
    // SELECTION METHODS (Simple enough to keep in main class)
    // =========================================================================

    /**
     * Select all visible rows.
     */
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

    /**
     * Deselect all rows.
     */
    deselectAll() {
      this.data.selected.clear();
      this.elements.tableContainer.find('.row-checkbox').prop('checked', false);
      this.elements.tableContainer.find('tr').removeClass('selected-row');
      this.updateSelectionCount();
      this.updateSelectAllCheckbox();
    }

    // =========================================================================
    // MODULE INTEGRATION HELPERS
    // =========================================================================

    /**
     * Helper to get the current module context for debugging.
     */
    getModuleContext() {
      return {
        data: this.data,
        state: this.state,
        config: this.config,
        elements: this.elements,
      };
    }

    /**
     * Helper to safely call module methods.
     */
    callModuleMethod(methodName, ...args) {
      if (typeof this[methodName] === 'function') {
        return this[methodName].apply(this, args);
      } else {
        this.logDebug(
          `Method ${methodName} not found. Check if required module is loaded.`
        );
        return null;
      }
    }
  } // End of ExcelEditor class

  // =========================================================================
  // DRUPAL INTEGRATION
  // =========================================================================

  /**
   * Drupal behavior to initialize Excel Editor.
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
   * Initialize the Excel Editor with proper dependency loading.
   */
  async function initializeExcelEditor(element) {
    try {
      // Show loading message
      const loadingMsg = $(
        `<div class="notification is-info excel-editor-init-loading">
          <span class="icon"><i class="fas fa-spinner fa-spin"></i></span>
          Loading Excel processing library...
        </div>`
      );
      $(element).prepend(loadingMsg);

      // Load SheetJS if needed
      await loadXLSXLibrary();

      // Remove loading message
      loadingMsg.remove();

      // Create and store the application instance
      const app = new ExcelEditor();
      element.excelEditor = app;

      // Make it globally accessible for debugging (in development only)
      if (window.location.search.includes('debug=1')) {
        window.excelEditorApp = app;
      }
    } catch (error) {
      console.error('Failed to initialize Excel Editor:', error);
      $('.excel-editor-init-loading').remove();
      $(element).prepend(
        `<div class="notification is-warning">
          <button class="delete"></button>
          <strong>Excel Library Loading Issue:</strong> ${error.message}
        </div>`
      );
    }
  }

  /**
   * Load XLSX library from CDN if not already present.
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

  // Export class for potential external use
  window.ExcelEditor = ExcelEditor;
})(jQuery, Drupal, once, drupalSettings);
