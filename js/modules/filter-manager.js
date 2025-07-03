/**
 * @file
 * Bulletproof Excel Editor Filter Manager Module - Prevents filter disappearing
 */

export class ExcelEditorFilterManager {
  constructor(app) {
    this.app = app;
    this.filtersInitialized = false;
    this.activeFiltersVisible = false;
  }

  /**
   * Sets up the filter controls area - now bulletproof against multiple calls
   */
  setupFilters() {
    if (!this.app.data.filtered.length) return;

    // Always ensure the structure exists
    this.ensureFilterStructureExists();

    // Update dynamic content without affecting the rest of the structure
    this.updateValidationControls();

    // Update column visibility status when columns change
    this.updateColumnVisibilityStatus();

    this.bindFilterEvents();

    // Preserve active filters visibility if they should be shown
    if (Object.keys(this.app.state.currentFilters).length > 0) {
      this.showActiveFilters();
    }
  }

  /**
   * Ensures the filter structure exists but doesn't recreate if already there
   */
  ensureFilterStructureExists() {
    const $ = jQuery;

    // Check if the structure already exists
    if ($('#active-filters-container').length > 0) {
      // Structure exists, don't recreate it
      return;
    }

    // Structure doesn't exist, create it
    this.createFilterStructure();
    this.filtersInitialized = true;
  }

  /**
   * Updates just the column visibility status without recreating the entire structure
   */
  updateColumnVisibilityStatus() {
    const $ = jQuery;

    // Find the existing column visibility notification
    const existingNotification = $('#column-visibility-notification');

    // Always generate the new content since we always want to show column count
    const newContent = this.buildColumnVisibilityNotification();

    if (existingNotification.length) {
      // Update the existing notification
      existingNotification.replaceWith(newContent);
    } else {
      // No existing notification, add it before the validation controls
      const validationContainer = $('#validation-controls-container');
      if (validationContainer.length) {
        validationContainer.before(newContent);
      } else {
        // Fallback: add to the filter controls container
        this.app.elements.filtersContainer.prepend(newContent);
      }
    }

    // Re-bind column management events
    this.bindColumnManagementEvents();
  }

  /**
   * Builds the column visibility notification HTML - ALWAYS shows column count
   */
  buildColumnVisibilityNotification() {
    // Get column counts
    const totalColumns = this.app.data.filtered[0]?.length || 0;
    const hiddenColumns = this.app.state.hiddenColumns.size;
    const visibleColumns = totalColumns - hiddenColumns;

    // Start building the notification - always show column info
    let messageContent = '';
    let notificationClass = 'is-info';
    let icon = 'eye';

    // Always show the column count
    messageContent = `${visibleColumns} of ${totalColumns} columns visible`;

    // Adjust styling based on state
    if (hiddenColumns > 0) {
      messageContent += ` (${hiddenColumns} hidden)`;
      notificationClass = 'is-warning';
      icon = 'eye-slash';
    }

    // Create the notification with updated content
    return `<div class="field" id="column-visibility-notification">
      <div class="notification ${notificationClass} is-light">
        <label class="label has-text-link">
          <span class="icon"><i class="fas fa-${icon}"></i></span>
          ${messageContent}
        </label>
        <div class="buttons mt-2 are-small">
          <button class="button" id="show-column-settings"><span>Manage Columns</span></button>
          <button class="button" id="reset-to-defaults"><span>Reset to Defaults</span></button>
          <button class="button" id="show-all-override"><span>Show All</span></button>
        </div>
      </div>
    </div>`;
  }

  /**
   * Binds events specifically for column management buttons
   */
  bindColumnManagementEvents() {
    const $ = jQuery;

    // Column management buttons
    $('#show-column-settings')
      .off('click.columnManager')
      .on('click.columnManager', () =>
        this.app.columnManager.showColumnVisibilityModal()
      );
    $('#reset-to-defaults')
      .off('click.columnManager')
      .on('click.columnManager', () =>
        this.app.columnManager.resetToDefaultColumns()
      );
    $('#show-all-override')
      .off('click.columnManager')
      .on('click.columnManager', () =>
        this.app.columnManager.showAllColumnsOverride()
      );
  }

  /**
   * Creates the initial filter structure (only when it doesn't exist)
   * MODIFIED: Removed column visibility notification logic since it's now handled separately
   */
  createFilterStructure() {
    // Create the complete structure without column visibility notification
    // (that will be added separately by updateColumnVisibilityStatus)
    const fullStructure = `
      <div id="validation-controls-container" class="field"></div>
      <div class="field mb-2" id="active-filters-container" style="display: none;">
      <div class="field">
        <label class="label mb-1">Active Filters:</label>
      </div>
        <div class="field is-grouped is-grouped-multiline mb-2">
          <div class="control">
            <button class="button is-small is-danger is-outlined" id="clear-all-filters-btn">
              <span class="icon is-small"><i class="fas fa-times"></i></span>
              <span>Clear All Filters</span>
            </button>
          </div>
          <div class="control">
            <button class="button is-small is-outlined" id="clear-last-filter-btn">
              <span class="icon is-small"><i class="fas fa-undo"></i></span>
              <span>Remove Last</span>
            </button>
          </div>
          <div class="control">
            <button class="button is-small is-info is-outlined" id="manage-filters-btn">
              <span class="icon is-small"><i class="fas fa-cog"></i></span>
              <span>Manage All</span>
            </button>
          </div>
        </div>
        <div class="field is-grouped is-grouped-multiline" id="active-filters"></div>
        <div class="field mt-2" id="filter-stats">
          <small class="has-text-grey" id="filter-stats-text"></small>
        </div>
      </div>`;

    this.app.elements.filtersContainer.html(fullStructure);
    // Add column visibility notification after creating the base structure
    this.updateColumnVisibilityStatus();
  }

  /**
   * Shows the active filters container and updates its content
   */
  showActiveFilters() {
    const $ = jQuery;
    const containerWrapper = $('#active-filters-container');

    if (containerWrapper.length) {
      containerWrapper.show();
      this.activeFiltersVisible = true;
      this.updateActiveFiltersContent();
    }
  }

  /**
   * Hides the active filters container
   */
  hideActiveFilters() {
    const $ = jQuery;
    const containerWrapper = $('#active-filters-container');

    if (containerWrapper.length) {
      containerWrapper.hide();
      this.activeFiltersVisible = false;
    }
  }

  /**
   * Updates only the validation controls without affecting the rest of the structure
   */
  updateValidationControls() {
    const $ = jQuery;
    const validationContainer = $('#validation-controls-container');

    if (!validationContainer.length) return;

    const validationControls = this.buildValidationFilterControls();
    validationContainer.html(validationControls);

    // Re-bind validation filter events
    $('#filter-errors-btn')
      .off('click.validation')
      .on('click.validation', () => this.applyValidationFilter('errors'));

    // Add event for new validation report button
    $('#validate-report-btn')
      .off('click.validation')
      .on('click.validation', () => {
        if (this.app.validationManager) {
          this.app.validationManager.showValidationSummary();
        }
      });
  }

  /**
   * Builds validation filter controls
   * @return {string} HTML string for validation filter controls
   */
  buildValidationFilterControls() {
    if (!this.app.validationManager) {
      return '';
    }

    const stats = this.app.validationManager.getValidationRowStats();

    if (stats.errorCount === 0) {
      return ''; // No validation issues, don't show validation filters
    }

    // Create a notification with identical structure to the column visibility notification
    let controls = `
    <div class="field">
      <div class="notification is-danger is-light" style="min-height: 140px; display: flex; flex-direction: column;">
        <div>
          <label class="label has-text-danger">
            <span class="icon"><i class="fas fa-exclamation-triangle"></i></span>
            Validation Issues Found
          </label>
          <p class="mb-3">There are ${
            stats.errorCount
          } rows with validation errors that need attention.</p>
        </div>
        <div class="buttons mt-auto are-small">
          <button class="button is-danger" id="filter-errors-btn">
            <span class="icon"><i class="fas fa-exclamation-triangle"></i></span>
            <span><strong>Show All ${stats.errorCount} Error${
      stats.errorCount !== 1 ? 's' : ''
    }</strong></span>
          </button>
          <button class="button is-light" id="validate-report-btn">
            <span class="icon"><i class="fas fa-file-alt"></i></span>
            <span>View Validation Report</span>
          </button>
        </div>
      </div>
    </div>`;

    return controls;
  }

  /**
   * Binds events for the filter control area with enhanced removal functionality.
   */
  bindFilterEvents() {
    const $ = jQuery;

    // Clear all filters button
    $('#clear-all-filters-btn')
      .off('click.filterManager')
      .on('click.filterManager', () => this.clearAllFilters());

    // Clear last filter button
    $('#clear-last-filter-btn')
      .off('click.filterManager')
      .on('click.filterManager', () => this.clearLastFilter());

    // Manage filters button
    $('#manage-filters-btn')
      .off('click.filterManager')
      .on('click.filterManager', () => this.showFilterManagementModal());
  }

  /**
   * Shows the filter modal for a specific column.
   * @param {number} columnIndex - The index of the column to filter.
   */
  showColumnFilter(columnIndex) {
    this.app.utilities.showQuickLoader('Loading filter options...');
    setTimeout(() => {
      try {
        const header = this.app.data.filtered[0][columnIndex];
        const uniqueValues = this.getUniqueColumnValues(columnIndex);
        const modalHtml = this.buildFilterModalHtml(
          header,
          uniqueValues,
          columnIndex
        );
        jQuery('body').append(modalHtml);
        this.bindFilterModalEvents(
          jQuery('#filter-modal'),
          columnIndex,
          header
        );
        this.updateFilterSelectedCount(jQuery('#filter-modal'));
      } finally {
        this.app.utilities.hideQuickLoader();
      }
    }, 50);
  }

  /**
   * Builds the HTML string for the filter modal with enhanced UI.
   * @param {string} header - The column header to display.
   * @param {Array} uniqueValues - The unique values in the column.
   * @param {number} columnIndex - The index of the column being filtered.
   * @return {string} The HTML string for the filter modal.
   */
  buildFilterModalHtml(header, uniqueValues, columnIndex) {
    const checkboxOptionsHtml = uniqueValues
      .map((val) => {
        const value = val || '';
        const displayValue = value === '' ? '(empty)' : value;
        const isChecked = this.isValueSelectedInFilter(columnIndex, value);
        return `<div class="column is-half"><label class="checkbox filter-checkbox-item"><input type="checkbox" value="${this.app.utilities.escapeHtml(
          value
        )}" ${
          isChecked ? 'checked' : ''
        } class="filter-value-checkbox mr-1"><span class="filter-checkbox-label">${this.app.utilities.escapeHtml(
          displayValue
        )}</span></label></div>`;
      })
      .join('');

    return `<div class="modal is-active" id="filter-modal" style="display: flex !important; z-index: 99999;">
              <div class="modal-background"></div>
              <div class="modal-content">
                <div class="box">
                  <h3 class="title is-4"><span class="icon"><i class="fas fa-filter"></i></span> Filter: ${this.app.utilities.escapeHtml(
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
                  <div class="field is-grouped is-grouped-right buttons">
                    <div class="control"><button class="button is-warning" id="clear-column-filter">
                      <span class="icon is-small"><i class="fas fa-times"></i></span>
                      <span>Clear Filter</span>
                    </button></div>
                    <div class="control"><button class="button" id="cancel-filter">Cancel</button></div>
                    <div class="control"><button class="button is-primary" id="apply-filter">Apply Filter</button></div>
                  </div>
                </div>
              </div>
              <button class="modal-close is-large" aria-label="close"></button>
            </div>`;
  }

  /**
   * Binds events for the filter modal.
   * @param {jQuery} modal - The jQuery object for the filter modal.
   * @param {number} columnIndex - The index of the column being filtered.
   * @param {string} header - The column header to display.
   */
  // eslint-disable-next-line no-unused-vars
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
        jQuery(this).prop('checked', !jQuery(this).prop('checked'));
      });
      this.updateFilterSelectedCount(modal);
    });

    modal.find('#filter-search').on('input', (e) => {
      const searchTerm = jQuery(e.target).val().toLowerCase();
      modal.find('.filter-checkbox-item').each(function () {
        const label = jQuery(this)
          .find('.filter-checkbox-label')
          .text()
          .toLowerCase();
        jQuery(this).closest('.column').toggle(label.includes(searchTerm));
      });
    });

    modal.find('.filter-value-checkbox').on('change', () => {
      this.updateFilterSelectedCount(modal);
    });

    modal.find('#clear-column-filter').on('click', async () => {
      try {
        await this.removeFilter(columnIndex);
        modal.remove();
      } finally {
        this.app.utilities.hideQuickLoader();
      }
    });

    modal.find('#apply-filter').on('click', async () => {
      try {
        await this.applyFilterFromModal(modal, columnIndex);
        modal.remove();
      } finally {
        this.app.utilities.hideQuickLoader();
      }
    });
  }

  /**
   * Applies the filter settings from the filter modal.
   * @param {jQuery} modal - The jQuery object for the filter modal.
   * @param {number} columnIndex - The index of the column being filtered.
   * @return {Promise<void>} - A promise that resolves when the filter is applied.
   */
  async applyFilterFromModal(modal, columnIndex) {
    const selectedValues = [];
    modal.find('.filter-value-checkbox:checked').each(function () {
      selectedValues.push(jQuery(this).val());
    });

    if (selectedValues.length > 0) {
      this.app.state.currentFilters[columnIndex] = {
        type: 'quick',
        selected: selectedValues,
      };
    } else {
      delete this.app.state.currentFilters[columnIndex];
    }

    await this.applyFilters();
  }

  /**
   * Applies all active filters to the original data.
   * @return {Promise<void>} - A promise that resolves when the filters are applied.
   */
  async applyFilters() {
    const shouldShowLoader = this.app.data.original.length > 100;

    if (shouldShowLoader) {
      this.app.utilities.showProcessLoader('Applying filters...');
    } else {
      this.app.utilities.showQuickLoader('Filtering...');
    }

    try {
      await new Promise((resolve) => setTimeout(resolve, 50));

      if (Object.keys(this.app.state.currentFilters).length === 0) {
        this.app.data.filtered = this.app.utilities.deepClone(
          this.app.data.original
        );
      } else {
        this.app.data.filtered = [this.app.data.original[0]]; // Keep header

        // Check for validation filters
        const validationFilter = this.app.state.currentFilters['_validation'];

        for (let i = 1; i < this.app.data.original.length; i++) {
          const row = this.app.data.original[i];
          let includeRow = true;

          // Handle validation filter first (it's exclusive)
          if (validationFilter) {
            includeRow = validationFilter.targetRows.has(i);
          } else {
            // Apply regular filters
            includeRow = Object.entries(this.app.state.currentFilters).every(
              ([colIndex, filter]) => {
                if (filter.type === 'validation') return true; // Skip validation filters in regular processing
                return this.rowMatchesFilter(row[parseInt(colIndex)], filter);
              }
            );
          }

          if (includeRow) {
            this.app.data.filtered.push(row);
          }
        }
      }

      this.app.data.selected.clear();
      await this.app.uiRenderer.renderTable();
      this.app.dataManager.updateSelectionCount();

      // CRITICAL: Update filters display immediately after applying
      this.updateActiveFiltersDisplay();

      // Trigger validation after filters are applied and table is re-rendered
      if (this.app.validationManager) {
        setTimeout(() => {
          this.app.validationManager.validateExistingBarcodeFields();
        }, 150);
      }
    } finally {
      if (shouldShowLoader) {
        this.app.utilities.hideProcessLoader();
      } else {
        this.app.utilities.hideQuickLoader();
      }
    }
  }

  /**
   * Checks if a single cell's value matches a given filter.
   * @param {any} cellValue - The value of the cell to check.
   * @param {Object} filter - The filter object containing type and value.
   * @return {boolean} - True if the cell value matches the filter, false otherwise.
   */
  rowMatchesFilter(cellValue, filter) {
    // Handle validation filters specially
    if (filter.type === 'validation') {
      // For validation filters, we need the row index, which isn't passed to this method
      // So validation filtering is handled in applyFilters() directly
      return true;
    }

    const originalValue = String(cellValue || '').trim();
    const value = filter.caseSensitive
      ? originalValue
      : originalValue.toLowerCase();
    const filterValue = filter.caseSensitive
      ? String(filter.value || '').trim()
      : String(filter.value || '')
          .trim()
          .toLowerCase();

    switch (filter.type) {
      case 'quick':
        // eslint-disable-next-line no-case-declarations
        const trimmedSelected = filter.selected.map((val) =>
          String(val || '').trim()
        );
        return trimmedSelected.includes(String(cellValue || '').trim());

      case 'equals':
        return value === filterValue;

      case 'contains':
        return value.includes(filterValue);

      case 'starts':
        return value.startsWith(filterValue);

      case 'ends':
        return value.endsWith(filterValue);

      case 'not_equals':
        return value !== filterValue;

      case 'not_contains':
        return !value.includes(filterValue);

      case 'empty':
        return !cellValue || String(cellValue).trim() === '';

      case 'not_empty':
        return cellValue && String(cellValue).trim() !== '';

      default:
        return true;
    }
  }

  /**
   * Clears all active filters and re-renders the table.
   */
  clearAllFilters() {
    if (Object.keys(this.app.state.currentFilters).length === 0) {
      this.app.utilities.showMessage('No filters to clear', 'info', 2000);
      return;
    }

    // Show confirmation for multiple filters
    const filterCount = Object.keys(this.app.state.currentFilters).length;
    if (filterCount > 1) {
      if (!confirm(`Clear all ${filterCount} active filters?`)) {
        return;
      }
    }

    this.app.utilities.showQuickLoader('Clearing all filters...');

    setTimeout(async () => {
      try {
        this.app.state.currentFilters = {};
        await this.applyFilters();
        this.app.utilities.showMessage('All filters cleared', 'success');
      } finally {
        this.app.utilities.hideQuickLoader();
      }
    }, 50);
  }

  /**
   * Clears the most recently applied filter.
   */
  clearLastFilter() {
    const filterKeys = Object.keys(this.app.state.currentFilters);

    if (filterKeys.length === 0) {
      this.app.utilities.showMessage('No filters to remove', 'info', 2000);
      return;
    }

    this.app.utilities.showQuickLoader('Removing last filter...');

    setTimeout(async () => {
      try {
        // Remove the last filter (most recently added)
        const lastFilterKey = filterKeys[filterKeys.length - 1];
        const removedFilter = this.app.state.currentFilters[lastFilterKey];
        delete this.app.state.currentFilters[lastFilterKey];

        await this.applyFilters();

        // Show which filter was removed
        let filterName = 'Last filter';
        if (lastFilterKey !== '_validation') {
          const columnName = this.app.data.original[0][parseInt(lastFilterKey)];
          filterName = `Filter on "${columnName}"`;
        } else {
          filterName = 'Validation filter';
        }

        this.app.utilities.showMessage(`${filterName} removed`, 'success');
      } finally {
        this.app.utilities.hideQuickLoader();
      }
    }, 50);
  }

  /**
   * Removes a specific filter by column index.
   * @param {string|number} columnIndex - The column index of the filter to remove.
   */
  async removeFilter(columnIndex) {
    if (!this.app.state.currentFilters[columnIndex]) {
      return;
    }

    this.app.utilities.showQuickLoader('Removing filter...');

    try {
      delete this.app.state.currentFilters[columnIndex];
      await this.applyFilters();
    } finally {
      this.app.utilities.hideQuickLoader();
    }
  }

  /**
   * Gets unique values for a column to populate the filter modal.
   * @param {number} columnIndex - The index of the column to get unique values from.
   * @return {Array} - An array of unique values in the column, sorted.
   */
  getUniqueColumnValues(columnIndex) {
    const values = new Set();
    for (let i = 1; i < this.app.data.original.length; i++) {
      const rawValue = this.app.data.original[i][columnIndex];
      values.add(String(rawValue || '').trim());
    }
    return Array.from(values).sort();
  }

  /**
   * Checks if a value is currently selected in a column's filter.
   * @param {number} columnIndex - The index of the column to check.
   * @param {string} value - The value to check for selection.
   * @return {boolean} - True if the value is selected in the filter, false otherwise.
   */
  isValueSelectedInFilter(columnIndex, value) {
    if (!this.app.state.currentFilters[columnIndex]) {
      return true;
    }
    const filter = this.app.state.currentFilters[columnIndex];
    if (filter.type === 'quick' && filter.selected) {
      return filter.selected.includes(value);
    }
    return false;
  }

  /**
   * Updates the selected count display in the filter modal.
   * @param {jQuery} modal - The jQuery object for the filter modal.
   */
  updateFilterSelectedCount(modal) {
    const checkedBoxes = modal.find('.filter-value-checkbox:checked');
    modal.find('#selected-count').text(checkedBoxes.length);
  }

  /**
   * Updates the display of active filters - the core method that handles visibility
   */
  updateActiveFiltersDisplay() {
    const filterCount = Object.keys(this.app.state.currentFilters).length;

    if (filterCount === 0) {
      this.hideActiveFilters();
      return;
    }

    // Show the container and update content
    this.showActiveFilters();
  }

  /**
   * Updates the active filters content display
   */
  updateActiveFiltersContent() {
    const $ = jQuery;
    const activeFiltersContainer = $('#active-filters');

    if (!activeFiltersContainer.length) return;

    const filterEntries = Object.entries(this.app.state.currentFilters);

    if (filterEntries.length === 0) {
      activeFiltersContainer.html(
        '<p class="has-text-grey">No active filters</p>'
      );
      this.updateFilterStats(); // Correct function name
      return;
    }

    const tagsHtml = filterEntries
      .map(([columnIndex, filter]) => {
        // Safety check - ensure filter is valid before accessing properties
        if (!filter || typeof filter !== 'object') {
          return ''; // Skip this filter if it's invalid
        }

        const columnName =
          this.app.data.filtered[0][columnIndex] || `Column ${columnIndex}`;
        const filterType = filter.type || 'unknown';

        let filterDescription = '';

        // Safely handle different filter types
        switch (filterType) {
          case 'include':
            // Handle include filters safely
            if (Array.isArray(filter.values)) {
              const valueCount = filter.values.length;
              filterDescription = `${valueCount} value${
                valueCount !== 1 ? 's' : ''
              }`;
            } else {
              filterDescription = 'values';
            }
            break;
          case 'search':
            // Handle search filters safely
            filterDescription = `"${filter.searchTerm || ''}"`;
            break;
          case 'validation':
            // Handle validation filters
            filterDescription = filter.validationType || 'validation';
            break;
          default:
            filterDescription = 'custom filter';
        }

        return `<div class="control mb-1">
      <div class="tags has-addons">
        <span class="tag is-info is-light">${this.app.utilities.escapeHtml(
          columnName
        )} ${filterDescription}
        <button class="delete remove-filter" data-column="${columnIndex}"></button></span>
      </div>
    </div>`;
      })
      .join('');

    activeFiltersContainer.html(tagsHtml);
    this.updateFilterStats(); // Correct function name

    // Rebind remove filter events
    $('.remove-filter').on('click', (e) => {
      const columnIndex = $(e.currentTarget).data('column');
      this.removeFilter(columnIndex);
    });
  }

  /**
   * Updates the filter statistics display
   */
  updateFilterStats() {
    const $ = jQuery;
    const statsContainer = $('#filter-stats-text');

    if (!statsContainer.length) return;

    // Calculate filtered vs total row counts (skip header row)
    const totalRows = this.app.data.original.length - 1;
    const filteredRows = this.app.data.filtered.length - 1;
    const hiddenRows = totalRows - filteredRows;

    if (hiddenRows === 0) {
      statsContainer.html(`Showing all ${totalRows} rows`);
    } else {
      const percentVisible = Math.round((filteredRows / totalRows) * 100);
      statsContainer.html(
        `Showing ${filteredRows} of ${totalRows} rows (${percentVisible}%) • ${hiddenRows} rows filtered out`
      );
    }
  }

  /**
   * Gets a human-readable description of a filter.
   * @param {Object} filter - The filter object containing type and value.
   * @return {string} - A description of the filter.
   */
  getFilterDescription(filter) {
    switch (filter.type) {
      case 'validation':
        return filter.description;
      case 'quick':
        return `${filter.selected.length} selected`;
      case 'equals':
        return `= "${this.app.utilities.escapeHtml(filter.value)}"`;
      case 'contains':
        return `contains "${this.app.utilities.escapeHtml(filter.value)}"`;
      case 'starts':
        return `starts with "${this.app.utilities.escapeHtml(filter.value)}"`;
      case 'ends':
        return `ends with "${this.app.utilities.escapeHtml(filter.value)}"`;
      case 'not_equals':
        return `≠ "${this.app.utilities.escapeHtml(filter.value)}"`;
      case 'not_contains':
        return `doesn't contain "${this.app.utilities.escapeHtml(
          filter.value
        )}"`;
      case 'empty':
        return 'is empty';
      case 'not_empty':
        return 'is not empty';
      default:
        return 'unknown filter';
    }
  }

  /**
   * Gets statistics about the current filtered data.
   * @return {Object} - An object containing total rows, filtered rows, hidden rows, and percentages.
   */
  getFilterStats() {
    const totalRows = this.app.data.original.length - 1;
    const filteredRows = this.app.data.filtered.length - 1;
    const hiddenRows = totalRows - filteredRows;
    const hiddenPercentage =
      totalRows > 0 ? ((hiddenRows / totalRows) * 100).toFixed(1) : 0;

    return {
      totalRows,
      filteredRows,
      hiddenRows,
      hiddenPercentage,
      activeFilters: Object.keys(this.app.state.currentFilters).length,
      hasFilters: Object.keys(this.app.state.currentFilters).length > 0,
    };
  }

  /**
   * Applies validation-based filters (errors only)
   * @param {string} filterType - The type of validation filter to apply ('errors' for now).
   * @return {Promise<void>} - A promise that resolves when the filter is applied.
   */
  async applyValidationFilter(filterType) {
    if (!this.app.validationManager) {
      this.app.utilities.showMessage(
        'Validation system not available',
        'warning'
      );
      return;
    }

    this.app.utilities.showQuickLoader('Applying validation filter...');

    try {
      // Clear existing filters first
      this.app.state.currentFilters = {};

      const stats = this.app.validationManager.getValidationRowStats();
      let targetRows = [];
      let filterDescription = '';

      if (filterType === 'errors') {
        targetRows = stats.errorRows;
        filterDescription = `${stats.errorCount} rows with errors`;
      }

      if (targetRows.length === 0) {
        this.app.utilities.showMessage('No rows found with errors', 'info');
        return;
      }

      // Apply the validation filter
      this.app.state.currentFilters['_validation'] = {
        type: 'validation',
        filterType: filterType,
        targetRows: new Set(targetRows),
        description: filterDescription,
      };

      await this.applyFilters();
      this.app.utilities.showMessage(`Showing ${filterDescription}`, 'success');
    } catch (error) {
      this.app.utilities.handleError(
        'Failed to apply validation filter',
        error
      );
    } finally {
      this.app.utilities.hideQuickLoader();
    }
  }

  /**
   * Shows filter statistics to the user.
   */
  showFilterStats() {
    const stats = this.getFilterStats();
    const message = `Showing ${stats.filteredRows} of ${stats.totalRows} rows (${stats.hiddenPercentage}% hidden)`;
    this.app.utilities.showMessage(message, 'info', 3000);
  }

  /**
   * Shows a modal to manage all active filters
   */
  showFilterManagementModal() {
    const $ = jQuery;
    const filterEntries = Object.entries(this.app.state.currentFilters);

    if (filterEntries.length === 0) {
      this.app.utilities.showMessage('No active filters to manage', 'info');
      return;
    }

    const filtersList = filterEntries
      .map(([columnIndex, filter]) => {
        // Safety check - ensure filter is valid before accessing properties
        if (!filter || typeof filter !== 'object') {
          return ''; // Skip this filter if it's invalid
        }

        const columnName =
          this.app.data.filtered[0][columnIndex] || `Column ${columnIndex}`;
        const filterType = filter.type || 'unknown';

        let filterDetails = '';

        // Safely handle different filter types
        switch (filterType) {
          case 'include':
            // Only process if values is an array
            if (Array.isArray(filter.values)) {
              const valueList = filter.values
                .slice(0, 5)
                .map(
                  (v) =>
                    `<span class="tag is-info is-light mr-1 mb-1">${this.app.utilities.escapeHtml(
                      v || '(empty)'
                    )}</span>`
                )
                .join('');

              const moreValues =
                filter.values.length > 5
                  ? `<span class="tag is-light">+${
                      filter.values.length - 5
                    } more</span>`
                  : '';

              filterDetails = `<div class="filter-values mt-1">${valueList}${moreValues}</div>`;
            } else {
              filterDetails =
                '<div class="filter-values mt-1"><span class="tag is-light">Invalid filter values</span></div>';
            }
            break;
          case 'search':
            filterDetails = `<div class="filter-values mt-1"><span class="tag is-warning is-light">"${this.app.utilities.escapeHtml(
              filter.searchTerm || ''
            )}"</span></div>`;
            break;
          case 'validation':
            filterDetails = `<div class="filter-values mt-1"><span class="tag is-danger is-light">${
              filter.validationType || 'validation'
            }</span></div>`;
            break;
          default:
            filterDetails =
              '<div class="filter-values mt-1"><span class="tag is-light">Custom filter</span></div>';
        }

        return `<div class="box p-3 mb-3 filter-management-item">
      <div class="level is-mobile mb-2">
        <div class="level-left">
          <strong>${this.app.utilities.escapeHtml(columnName)}</strong>
        </div>
        <div class="level-right">
          <div class="buttons are-small">
            <button class="button is-small is-danger is-outlined remove-filter-btn" data-column="${columnIndex}">
              <span class="icon is-small"><i class="fas fa-times"></i></span>
              <span>Remove</span>
            </button>
            <button class="button is-small is-info is-outlined edit-filter-btn" data-column="${columnIndex}">
              <span class="icon is-small"><i class="fas fa-edit"></i></span>
              <span>Edit</span>
            </button>
          </div>
        </div>
      </div>
      ${filterDetails}
    </div>`;
      })
      .join('');

    const modalHtml = `<div class="modal is-active" id="filter-management-modal">
    <div class="modal-background"></div>
    <div class="modal-card">
      <header class="modal-card-head">
        <p class="modal-card-title"><span class="icon"><i class="fas fa-filter"></i></span> Manage Filters</p>
        <button class="delete" aria-label="close"></button>
      </header>
      <section class="modal-card-body">
        <div class="filter-management-list">
          ${filtersList}
        </div>
      </section>
      <footer class="modal-card-foot">
        <button class="button is-danger" id="clear-all-filters-modal-btn">Clear All Filters</button>
        <button class="button" id="close-filter-modal-btn">Close</button>
      </footer>
    </div>
  </div>`;

    $('body').append(modalHtml);

    // Bind events
    $('#filter-management-modal .delete, #close-filter-modal-btn').on(
      'click',
      () => {
        $('#filter-management-modal').remove();
      }
    );

    $('#clear-all-filters-modal-btn').on('click', () => {
      this.clearAllFilters();
      $('#filter-management-modal').remove();
    });

    $('.remove-filter-btn').on('click', (e) => {
      const columnIndex = $(e.currentTarget).data('column');
      this.removeFilter(columnIndex);
      $(e.currentTarget).closest('.filter-management-item').remove();

      // If no more filters, close the modal
      if (Object.keys(this.app.state.currentFilters).length === 0) {
        $('#filter-management-modal').remove();
      }
    });

    $('.edit-filter-btn').on('click', (e) => {
      const columnIndex = $(e.currentTarget).data('column');
      $('#filter-management-modal').remove();
      this.showColumnFilter(columnIndex);
    });
  }
}
