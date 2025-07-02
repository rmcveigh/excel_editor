/**
 * @file
 * Excel Editor JavaScript - Main class and initialization with Web Worker support
 */

import { ExcelEditorUtilities } from './modules/utilities.js';
import { ExcelEditorWorkerManager } from './modules/worker-manager.js';
import { ExcelEditorBarcodeSystem } from './modules/barcode-system.js';
import { ExcelEditorDataManager } from './modules/data-manager.js';
import { ExcelEditorUIRenderer } from './modules/ui-renderer.js';
import { ExcelEditorFilterManager } from './modules/filter-manager.js';
import { ExcelEditorColumnManager } from './modules/column-manager.js';
import { ExcelEditorValidationManager } from './modules/validation-manager.js';
import { ExcelEditorDraftManager } from './modules/draft-manager.js';
import { ExcelEditorExportManager } from './modules/export-manager.js';

/**
 * The main application class for the Excel Editor.
 */
class ExcelEditor {
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
      dogEntityUrlCache: {},
      subjectLinksFullyFetched: false,
    };

    // Configuration from Drupal settings
    this.config = {
      endpoints: window.drupalSettings?.excelEditor?.endpoints || {},
      settings: window.drupalSettings?.excelEditor?.settings || {},
      editableColumns: ['new_barcode', 'notes', 'actions'],
      maxFileSize: 10 * 1024 * 1024, // 10MB
      supportedFormats: ['.xlsx', '.xls', '.csv'],
      autosaveInterval: 5 * 60 * 1000, // 5 minutes
      base_path: window.drupalSettings?.path?.baseUrl || '/',
    };

    // Autosave timer
    this.autosaveTimer = null;

    // CSRF token for API calls
    this.csrfToken = null;

    // DOM element cache
    this.elements = {};

    // Initialize modules
    this.utilities = new ExcelEditorUtilities(this);
    this.workerManager = new ExcelEditorWorkerManager(this);
    this.barcodeSystem = new ExcelEditorBarcodeSystem(this);
    this.dataManager = new ExcelEditorDataManager(this);
    this.uiRenderer = new ExcelEditorUIRenderer(this);
    this.filterManager = new ExcelEditorFilterManager(this);
    this.columnManager = new ExcelEditorColumnManager(this);
    this.validationManager = new ExcelEditorValidationManager(this);
    this.draftManager = new ExcelEditorDraftManager(this);
    this.exportManager = new ExcelEditorExportManager(this);

    // Start initialization
    this.init();
  }

  /**
   * Main initialization method.
   */
  async init() {
    try {
      this.checkDependencies();
      this.cacheElements();
      this.utilities.getCsrfToken();
      this.utilities.hideLoading();

      // Initialize worker manager (non-blocking)
      await this.initializeWorkerManager();

      this.bindEvents();
      this.draftManager.loadDrafts();

      if (this.config.settings.autosave_enabled) {
        this.utilities.startAutosave();
      }

      this.state.isInitialized = true;
      this.utilities.logDebug('Excel Editor initialized successfully');
    } catch (error) {
      this.utilities.handleError('Failed to initialize Excel Editor', error);
    }
  }

  /**
   * Initialize the worker manager
   */
  async initializeWorkerManager() {
    try {
      this.utilities.logDebug('Initializing Web Worker...');

      const workerReady = await this.workerManager.initialize();

      if (workerReady) {
        this.utilities.logDebug('Web Worker initialized successfully');
      } else {
        this.utilities.logDebug(
          'Web Worker initialization failed, using fallback'
        );
      }
    } catch (error) {
      this.utilities.logDebug('Web Worker initialization error:', error);
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

    if (missing.length > 0) {
      throw new Error(`Missing required dependencies: ${missing.join(', ')}`);
    }

    this.utilities.logDebug('All dependencies loaded successfully');
  }

  /**
   * Cache DOM elements for performance.
   */
  cacheElements() {
    const $ = jQuery;
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
    this.utilities.logDebug('Cached DOM elements');
  }

  /**
   * Bind all application events.
   */
  bindEvents() {
    const $ = jQuery;

    // File upload events
    this.elements.fileInput.on('change', (e) => this.handleFileUpload(e));
    this.setupDragDropUpload();

    // Toolbar button events
    this.elements.saveDraftBtn.on('click', () => this.draftManager.saveDraft());
    this.elements.exportBtn.on('click', () =>
      this.exportManager.exportSelected()
    );
    this.elements.exportAllBtn.on('click', () =>
      this.exportManager.exportAll()
    );
    this.elements.toggleColumnsBtn.on('click', () =>
      this.columnManager.showColumnVisibilityModal()
    );
    this.elements.selectAllBtn.on('click', () => this.selectAllVisible());
    this.elements.deselectAllBtn.on('click', () => this.deselectAll());
    this.elements.resetBarcodesBtn.on('click', () =>
      this.barcodeSystem.resetBarcodes()
    );
    this.elements.validateBarcodesBtn.on('click', () =>
      this.validationManager.showValidationSummary()
    );

    // Table events
    this.bindTableEvents();

    // Keyboard shortcuts
    $(document).on('keydown', (e) => this.handleKeyboardShortcuts(e));

    // Window events
    $(window).on('beforeunload', () => this.handleBeforeUnload());

    // Performance monitoring events (debug mode only)
    if (this.config.settings.debug) {
      this.bindPerformanceEvents();
    }
  }

  /**
   * Bind performance monitoring events
   */
  bindPerformanceEvents() {
    // Monitor worker performance
    if (this.workerManager) {
      // Add debug info to export buttons
      this.elements.exportBtn.attr(
        'title',
        this.workerManager.isAvailable()
          ? 'Export (with background processing)'
          : 'Export (standard processing)'
      );

      this.elements.exportAllBtn.attr(
        'title',
        this.workerManager.isAvailable()
          ? 'Export All (with background processing)'
          : 'Export All (standard processing)'
      );
    }
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
  // EVENT HANDLERS
  // =========================================================================

  /**
   * Handle file drop event.
   * @param {DragEvent} e - The drag event containing the dropped files.
   */
  handleFileDrop(e) {
    const files = e.dataTransfer.files;
    if (files.length > 0) {
      this.dataManager.processFile(files[0]);
    }
  }

  /**
   * Handle file input change.
   * @param {Event} e - The change event from the file input.
   */
  handleFileUpload(e) {
    const file = e.target.files[0];
    if (file) {
      this.dataManager.processFile(file);
    }
  }

  /**
   * Handle filter link clicks.
   * @param {Event} e - The click event on a filter link.
   */
  handleFilterClick(e) {
    const $ = jQuery;
    const $target = $(e.target);
    let columnIndex = $target.data('column');
    if (columnIndex === undefined || columnIndex === null) {
      columnIndex = $target.closest('[data-column]').data('column');
    }
    if (columnIndex !== undefined && columnIndex !== null) {
      this.filterManager.showColumnFilter(columnIndex);
    }
  }

  /**
   * Handle cell editing.
   * @param {Event} e - The change event from a cell input.
   */
  handleCellEdit(e) {
    const $ = jQuery;
    const $cell = $(e.target);
    const rowIndex = parseInt($cell.data('row'));
    const colIndex = parseInt($cell.data('col'));
    const newValue = String($cell.val() || '').trim();

    this.data.filtered[rowIndex][colIndex] = newValue;
    this.data.dirty = true;

    const columnName = this.data.filtered[0][colIndex];

    // Validate new_barcode fields
    if (columnName === 'new_barcode') {
      this.validationManager.validateBarcodeField($cell, newValue, rowIndex);
    }

    if (columnName === 'actions') {
      this.dataManager.applyRowStyling();
    }
  }

  /**
   * Handle row selection via checkbox.
   * @param {Event} e - The change event from a row checkbox.
   */
  handleRowSelection(e) {
    const $ = jQuery;
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

    this.dataManager.updateSelectionCount();
    this.dataManager.updateSelectAllCheckbox();
  }

  /**
   * Handle select all checkbox.
   * @param {Event} e - The change event from the select all checkbox.
   */
  handleSelectAllCheckbox(e) {
    const $ = jQuery;
    const isChecked = $(e.target).is(':checked');
    if (isChecked) {
      this.selectAllVisible();
    } else {
      this.deselectAll();
    }
  }

  /**
   * Handle keyboard shortcuts.
   * @param {KeyboardEvent} e - The keyboard event.
   */
  handleKeyboardShortcuts(e) {
    if ((e.ctrlKey || e.metaKey) && e.key === 's') {
      e.preventDefault();
      this.draftManager.saveDraft();
    }

    // Debug shortcut to show worker status
    if (
      this.config.settings.debug &&
      (e.ctrlKey || e.metaKey) &&
      e.shiftKey &&
      e.key === 'W'
    ) {
      e.preventDefault();
      this.showWorkerDebugInfo();
    }
  }

  /**
   * Show worker debug information
   */
  showWorkerDebugInfo() {
    if (!this.config.settings.debug) return;

    const status = this.workerManager.getStatus();
    const message = `Worker Status:
    • Supported: ${status.supported}
    • Ready: ${status.ready}
    • Pending Tasks: ${status.pendingTasks}
    • Available: ${this.workerManager.isAvailable()}`;

    alert(message);
  }

  /**
   * Handle before page unload.
   * @returns {string|undefined} - Confirmation message if there are unsaved changes.
   */
  handleBeforeUnload() {
    if (this.data.dirty) {
      return Drupal.t(
        'You have unsaved changes. Are you sure you want to leave?'
      );
    }
  }

  // =========================================================================
  // SELECTION METHODS
  // =========================================================================

  /**
   * Select all visible rows.
   */
  selectAllVisible() {
    this.elements.tableContainer
      .find('.row-checkbox:visible')
      .each((index, checkbox) => {
        const $ = jQuery;
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
    this.dataManager.updateSelectionCount();
    this.dataManager.updateSelectAllCheckbox();
  }

  // =========================================================================
  // PERFORMANCE MONITORING
  // =========================================================================

  /**
   * Monitor operation performance
   * @param {string} operation - Name of the operation
   * @param {Function} fn - Function to execute
   * @returns {Promise<any>} - Result of the operation
   */
  async monitorPerformance(operation, fn) {
    if (!this.config.settings.debug) {
      return fn();
    }

    const startTime = performance.now();
    this.utilities.logDebug(`Starting operation: ${operation}`);

    try {
      const result = await fn();
      const endTime = performance.now();
      const duration = Math.round(endTime - startTime);

      this.utilities.logDebug(
        `Operation ${operation} completed in ${duration}ms`
      );

      // Show performance notification for long operations
      if (duration > 2000) {
        this.utilities.showMessage(
          `${operation} completed in ${(duration / 1000).toFixed(1)}s`,
          'info',
          3000
        );
      }

      return result;
    } catch (error) {
      const endTime = performance.now();
      const duration = Math.round(endTime - startTime);

      this.utilities.logDebug(
        `Operation ${operation} failed after ${duration}ms:`,
        error
      );
      throw error;
    }
  }

  // =========================================================================
  // CLEANUP
  // =========================================================================

  /**
   * Clean up resources when the application is destroyed
   */
  destroy() {
    // Terminate worker
    if (this.workerManager) {
      this.workerManager.terminate();
    }

    // Stop autosave
    this.utilities.stopAutosave();

    // Clear intervals and timeouts
    if (this.autosaveTimer) {
      clearInterval(this.autosaveTimer);
    }

    // Remove event listeners
    const $ = jQuery;
    $(document).off('keydown');
    $(window).off('beforeunload');

    this.utilities.logDebug('Excel Editor resources cleaned up');
  }
}

// =========================================================================
// DRUPAL INTEGRATION
// =========================================================================

/**
 * Initialize the Excel Editor with proper dependency loading.
 * @param {HTMLElement} element - The container element for the Excel Editor.
 * @returns {Promise<void>} - A promise that resolves when the editor is initialized.
 */
async function initializeExcelEditor(element) {
  try {
    const $ = jQuery;

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

    // Cleanup on page unload
    $(window).on('beforeunload', () => {
      if (app && typeof app.destroy === 'function') {
        app.destroy();
      }
    });

    // Make it globally accessible for debugging (in development only)
    if (window.location.search.includes('debug=1')) {
      window.excelEditorApp = app;
    }
  } catch (error) {
    console.error('Failed to initialize Excel Editor:', error);
    jQuery('.excel-editor-init-loading').remove();
    jQuery(element).prepend(
      `<div class="notification is-warning">
        <button class="delete"></button>
        <strong>Excel Library Loading Issue:</strong> ${error.message}
      </div>`
    );
  }
}

/**
 * Load XLSX library from CDN if not already present.
 * @returns {Promise<void>} - A promise that resolves when the library is loaded.
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

/**
 * Drupal behavior to initialize Excel Editor.
 */
Drupal.behaviors.excelEditor = {
  attach: function (context, settings) {
    once('excel-editor', '.excel-editor-container', context).forEach(function (
      element
    ) {
      initializeExcelEditor(element);
    });
  },

  detach: function (context, settings, trigger) {
    // Clean up when elements are removed
    if (trigger === 'unload') {
      once
        .remove('excel-editor', '.excel-editor-container', context)
        .forEach(function (element) {
          if (
            element.excelEditor &&
            typeof element.excelEditor.destroy === 'function'
          ) {
            element.excelEditor.destroy();
          }
        });
    }
  },
};

// Export class for potential external use
window.ExcelEditor = ExcelEditor;
