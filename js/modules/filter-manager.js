/**
 * @file
 * Excel Editor Filter Manager Module
 *
 * Handles all filtering logic, filter modals, and filter management.
 * This module adds filtering methods to the ExcelEditor class.
 */

/* eslint-disable no-console */
(function ($) {
  'use strict';

  /**
   * Filter Manager module for Excel Editor.
   * This function is called on ExcelEditor instances to add filtering methods.
   */
  window.ExcelEditorFilterManager = function () {
    // =========================================================================
    // FILTER SETUP & MANAGEMENT
    // =========================================================================

    /**
     * Sets up the filter controls area above the table.
     */
    this.setupFilters = function () {
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
    };

    /**
     * Binds events for the filter control area.
     */
    this.bindFilterEvents = function () {
      $('#clear-all-filters-btn').on('click', () => this.clearAllFilters());
      $('#show-column-settings').on('click', () =>
        this.showColumnVisibilityModal()
      );
      $('#reset-to-defaults').on('click', () => this.resetToDefaultColumns());
      $('#show-all-override').on('click', () => this.showAllColumnsOverride());
    };

    // =========================================================================
    // FILTER MODAL & INTERACTION
    // =========================================================================

    /**
     * Shows the filter modal for a specific column.
     * @param {number} columnIndex The index of the column to filter.
     */
    this.showColumnFilter = function (columnIndex) {
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
    };

    /**
     * [HELPER] Builds the HTML string for the filter modal.
     * @param {string} header The column header text.
     * @param {Array<string>} uniqueValues The unique values in the column.
     * @param {number} columnIndex The index of the column.
     * @returns {string} The complete HTML string for the modal.
     */
    this._buildFilterModalHtml = function (header, uniqueValues, columnIndex) {
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
    };

    /**
     * Binds events for the filter modal (buttons, search, etc.).
     * @param {jQuery} modal The jQuery object for the modal.
     * @param {number} columnIndex The column index being filtered.
     * @param {string} header The name of the column.
     */
    this.bindFilterModalEvents = function (modal, columnIndex, header) {
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
        this.showQuickLoader('Clearing column filter...');

        try {
          delete this.state.currentFilters[columnIndex];
          await this.applyFilters();
          this.updateActiveFiltersDisplay();
          modal.remove();
        } finally {
          this.hideQuickLoader();
        }
      });

      modal.find('#apply-filter').on('click', async () => {
        this.showQuickLoader('Applying filter...');
        try {
          await this.applyFilterFromModal(modal, columnIndex);
          modal.remove();
        } finally {
          this.hideQuickLoader();
        }
      });
    };

    /**
     * Shows an advanced filter modal with more options.
     * @param {number} columnIndex The column index to filter.
     */
    this.showAdvancedFilter = function (columnIndex) {
      const header = this.data.filtered[0][columnIndex];

      const modalHtml = `
        <div class="modal is-active" id="advanced-filter-modal" style="z-index: 99999;">
          <div class="modal-background"></div>
          <div class="modal-content">
            <div class="box">
              <h3 class="title is-4">
                <span class="icon"><i class="fas fa-filter"></i></span>
                Advanced Filter: ${this.escapeHtml(header)}
              </h3>

              <div class="field">
                <label class="label">Filter Type</label>
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
                <label class="label">Filter Value</label>
                <div class="control">
                  <input class="input" type="text" id="filter-value" placeholder="Enter filter value...">
                </div>
              </div>

              <div class="field">
                <label class="label">Case Sensitive</label>
                <div class="control">
                  <label class="checkbox">
                    <input type="checkbox" id="case-sensitive">
                    Match case exactly
                  </label>
                </div>
              </div>

              <div class="field is-grouped is-grouped-right">
                <div class="control">
                  <button class="button" id="cancel-advanced-filter">Cancel</button>
                </div>
                <div class="control">
                  <button class="button is-primary" id="apply-advanced-filter">Apply Filter</button>
                </div>
              </div>
            </div>
          </div>
          <button class="modal-close is-large" aria-label="close"></button>
        </div>`;

      const modal = $(modalHtml);
      $('body').append(modal);

      // Show/hide value field based on filter type
      modal.find('#filter-type').on('change', function () {
        const filterType = $(this).val();
        const valueField = modal.find('#filter-value-field');

        if (filterType === 'empty' || filterType === 'not_empty') {
          valueField.hide();
        } else {
          valueField.show();
        }
      });

      // Bind events
      modal
        .find('.modal-close, #cancel-advanced-filter, .modal-background')
        .on('click', () => modal.remove());

      modal.find('#apply-advanced-filter').on('click', async () => {
        const filterType = modal.find('#filter-type').val();
        const filterValue = modal.find('#filter-value').val();
        const caseSensitive = modal.find('#case-sensitive').is(':checked');

        this.state.currentFilters[columnIndex] = {
          type: filterType,
          value: filterValue,
          caseSensitive: caseSensitive,
        };

        this.showQuickLoader('Applying advanced filter...');
        try {
          await this.applyFilters();
          this.updateActiveFiltersDisplay();
          modal.remove();
        } finally {
          this.hideQuickLoader();
        }
      });
    };

    // =========================================================================
    // FILTER APPLICATION & LOGIC
    // =========================================================================

    /**
     * Applies the filter settings from the filter modal.
     * @param {jQuery} modal The modal element.
     * @param {number} columnIndex The column index being filtered.
     */
    this.applyFilterFromModal = async function (modal, columnIndex) {
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
    };

    /**
     * Applies all active filters to the original data to create the filtered dataset.
     * Enhanced with loading indicators.
     */
    this.applyFilters = async function () {
      // Show loader for larger datasets
      const shouldShowLoader = this.data.original.length > 100;

      if (shouldShowLoader) {
        this.showProcessLoader('Applying filters...');
      } else {
        this.showQuickLoader('Filtering...');
      }

      try {
        // Add a small delay to ensure loader is visible
        await new Promise((resolve) => setTimeout(resolve, 50));

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
      } finally {
        if (shouldShowLoader) {
          this.hideProcessLoader();
        } else {
          this.hideQuickLoader();
        }
      }
    };

    /**
     * Checks if a single cell's value matches a given filter.
     * @param {string} cellValue The value of the cell.
     * @param {object} filter The filter object.
     * @returns {boolean} True if the cell matches the filter.
     */
    this.rowMatchesFilter = function (cellValue, filter) {
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
    };

    /**
     * Clears all active filters and re-renders the table.
     * Enhanced with loading indicators.
     */
    this.clearAllFilters = function () {
      this.showQuickLoader('Clearing filters...');

      // Use setTimeout to ensure loader shows
      setTimeout(async () => {
        try {
          this.state.currentFilters = {};
          await this.applyFilters(); // This now has its own loader
          this.updateActiveFiltersDisplay();
          this.showMessage('All filters cleared', 'success');
        } finally {
          this.hideQuickLoader();
        }
      }, 50);
    };

    /**
     * Clears filters for a specific column.
     * @param {number} columnIndex The column index to clear.
     */
    this.clearColumnFilter = function (columnIndex) {
      delete this.state.currentFilters[columnIndex];
      this.applyFilters();
      this.updateActiveFiltersDisplay();

      const header = this.data.original[0][columnIndex];
      this.showMessage(`Cleared filter for "${header}"`, 'success');
    };

    // =========================================================================
    // FILTER UTILITIES & HELPERS
    // =========================================================================

    /**
     * Gets unique values for a column to populate the filter modal.
     * @param {number} columnIndex The index of the column.
     * @returns {Array<string>} A sorted array of unique values.
     */
    this.getUniqueColumnValues = function (columnIndex) {
      const values = new Set();
      for (let i = 1; i < this.data.original.length; i++) {
        const rawValue = this.data.original[i][columnIndex];
        values.add(String(rawValue || '').trim());
      }
      return Array.from(values).sort();
    };

    /**
     * Checks if a value is currently selected in a column's filter.
     * @param {number} columnIndex The index of the column.
     * @param {string} value The value to check.
     * @returns {boolean} True if the value is selected (or if there's no filter).
     */
    this.isValueSelectedInFilter = function (columnIndex, value) {
      if (!this.state.currentFilters[columnIndex]) {
        return true; // If no filter, all are considered selected
      }
      const filter = this.state.currentFilters[columnIndex];
      if (filter.type === 'quick' && filter.selected) {
        return filter.selected.includes(value);
      }
      return false;
    };

    /**
     * Updates the selected count display in the filter modal.
     * @param {jQuery} modal The jQuery object for the modal.
     */
    this.updateFilterSelectedCount = function (modal) {
      const checkedBoxes = modal.find('.filter-value-checkbox:checked');
      modal.find('#selected-count').text(checkedBoxes.length);
    };

    /**
     * Updates the display of active filters above the table.
     */
    this.updateActiveFiltersDisplay = function () {
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
    };

    /**
     * Gets a human-readable description of a filter.
     * @param {object} filter The filter object.
     * @returns {string} The filter description.
     */
    this.getFilterDescription = function (filter) {
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
    };

    // =========================================================================
    // FILTER PRESETS & SAVED FILTERS
    // =========================================================================

    /**
     * Saves the current filter state as a preset.
     * @param {string} name The name for the filter preset.
     */
    this.saveFilterPreset = function (name) {
      const presets = JSON.parse(
        localStorage.getItem('excel_editor_filter_presets') || '{}'
      );
      presets[name] = {
        filters: this.deepClone(this.state.currentFilters),
        created: new Date().toISOString(),
        description: this.getFilterSummary(),
      };
      localStorage.setItem(
        'excel_editor_filter_presets',
        JSON.stringify(presets)
      );
      this.showMessage(`Filter preset "${name}" saved`, 'success');
    };

    /**
     * Loads a filter preset.
     * @param {string} name The name of the preset to load.
     */
    this.loadFilterPreset = function (name) {
      const presets = JSON.parse(
        localStorage.getItem('excel_editor_filter_presets') || '{}'
      );
      if (presets[name]) {
        this.state.currentFilters = this.deepClone(presets[name].filters);
        this.applyFilters();
        this.updateActiveFiltersDisplay();
        this.showMessage(`Filter preset "${name}" loaded`, 'success');
      } else {
        this.showMessage(`Filter preset "${name}" not found`, 'warning');
      }
    };

    /**
     * Gets a summary of current filters.
     * @returns {string} A human-readable filter summary.
     */
    this.getFilterSummary = function () {
      const filterCount = Object.keys(this.state.currentFilters).length;
      if (filterCount === 0) return 'No filters';

      const descriptions = Object.entries(this.state.currentFilters).map(
        ([colIndex, filter]) => {
          const header = this.data.original[0][parseInt(colIndex)];
          return `${header}: ${this.getFilterDescription(filter)}`;
        }
      );

      return descriptions.join(', ');
    };

    // =========================================================================
    // FILTER STATISTICS & ANALYSIS
    // =========================================================================

    /**
     * Gets statistics about the current filtered data.
     * @returns {object} Filter statistics.
     */
    this.getFilterStats = function () {
      const totalRows = this.data.original.length - 1; // Exclude header
      const filteredRows = this.data.filtered.length - 1; // Exclude header
      const hiddenRows = totalRows - filteredRows;
      const hiddenPercentage =
        totalRows > 0 ? ((hiddenRows / totalRows) * 100).toFixed(1) : 0;

      return {
        totalRows,
        filteredRows,
        hiddenRows,
        hiddenPercentage,
        activeFilters: Object.keys(this.state.currentFilters).length,
        hasFilters: Object.keys(this.state.currentFilters).length > 0,
      };
    };

    /**
     * Shows filter statistics to the user.
     */
    this.showFilterStats = function () {
      const stats = this.getFilterStats();
      const message = `Showing ${stats.filteredRows} of ${stats.totalRows} rows (${stats.hiddenPercentage}% hidden)`;
      this.showMessage(message, 'info', 3000);
    };

    this.logDebug('ExcelEditorFilterManager module loaded');
  };
})(jQuery);
