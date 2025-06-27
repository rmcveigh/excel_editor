/**
 * @file
 * Excel Editor Filter Manager Module
 */

export class ExcelEditorFilterManager {
  constructor(app) {
    this.app = app;
  }

  /**
   * Sets up the filter controls area above the table.
   */
  setupFilters() {
    if (!this.app.data.filtered.length) return;

    let statusMessages = '';

    if (this.app.state.hiddenColumns.size > 0) {
      statusMessages += `<div class="field"><div class="notification is-info is-light"><span class="icon"><i class="fas fa-eye-slash"></i></span> ${
        this.app.state.hiddenColumns.size
      } column${
        this.app.state.hiddenColumns.size !== 1 ? 's' : ''
      } hidden. <button class="button is-small is-light ml-2" id="show-column-settings"><span>Manage Columns</span></button></div></div>`;
    }

    if (
      this.app.config.settings.hideBehavior === 'hide_others' &&
      this.app.config.settings.defaultVisibleColumns?.length > 0
    ) {
      statusMessages += `<div class="field"><div class="notification is-primary is-light"><span class="icon"><i class="fas fa-cog"></i></span> Default column visibility applied. <button class="button is-small is-light ml-2" id="reset-to-defaults"><span>Reset to Defaults</span></button> <button class="button is-small is-light ml-2" id="show-all-override"><span>Show All</span></button></div></div>`;
    }

    this.app.elements.filtersContainer.html(
      `${statusMessages} <div class="field mb-2" id="active-filters-container" style="display: none;"><label class="label">Active Filters:</label><div class="control" id="active-filters"></div><div class="control mt-2"><button class="button is-small is-light" id="clear-all-filters-btn"><span>Clear All Filters</span></button></div></div>`
    );

    this.bindFilterEvents();
  }

  /**
   * Binds events for the filter control area.
   */
  bindFilterEvents() {
    const $ = jQuery;
    $('#clear-all-filters-btn').on('click', () => this.clearAllFilters());
    $('#show-column-settings').on('click', () =>
      this.app.columnManager.showColumnVisibilityModal()
    );
    $('#reset-to-defaults').on('click', () =>
      this.app.columnManager.resetToDefaultColumns()
    );
    $('#show-all-override').on('click', () =>
      this.app.columnManager.showAllColumnsOverride()
    );
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
   * Builds the HTML string for the filter modal.
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
   * Binds events for the filter modal.
   * @param {jQuery} modal - The jQuery object for the filter modal.
   * @param {number} columnIndex - The index of the column being filtered.
   * @param {string} header - The column header to display.
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
      this.app.utilities.showQuickLoader('Clearing column filter...');
      try {
        delete this.app.state.currentFilters[columnIndex];
        await this.applyFilters();
        this.updateActiveFiltersDisplay();
        modal.remove();
      } finally {
        this.app.utilities.hideQuickLoader();
      }
    });

    modal.find('#apply-filter').on('click', async () => {
      this.app.utilities.showQuickLoader('Applying filter...');
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
    this.updateActiveFiltersDisplay();
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
        this.app.data.filtered = [this.app.data.original[0]];
        for (let i = 1; i < this.app.data.original.length; i++) {
          const row = this.app.data.original[i];
          if (
            Object.entries(this.app.state.currentFilters).every(
              ([colIndex, filter]) =>
                this.rowMatchesFilter(row[parseInt(colIndex)], filter)
            )
          ) {
            this.app.data.filtered.push(row);
          }
        }
      }

      this.app.data.selected.clear();
      await this.app.uiRenderer.renderTable();
      this.app.dataManager.updateSelectionCount();
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
    this.app.utilities.showQuickLoader('Clearing filters...');

    setTimeout(async () => {
      try {
        this.app.state.currentFilters = {};
        await this.applyFilters();
        this.updateActiveFiltersDisplay();
        this.app.utilities.showMessage('All filters cleared', 'success');
      } finally {
        this.app.utilities.hideQuickLoader();
      }
    }, 50);
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
   * Updates the display of active filters above the table.
   */
  updateActiveFiltersDisplay() {
    const $ = jQuery;
    const filtersContainer = $('#active-filters');
    const containerWrapper = $('#active-filters-container');

    if (Object.keys(this.app.state.currentFilters).length === 0) {
      containerWrapper.hide();
      return;
    }

    const filterTags = Object.entries(this.app.state.currentFilters)
      .map(([columnIndex, filter]) => {
        const header = this.app.data.original[0][parseInt(columnIndex)];
        const filterDescription = this.getFilterDescription(filter);
        return `<span class="tag is-info"><strong>${this.app.utilities.escapeHtml(
          header
        )}</strong>: ${filterDescription}<button class="delete is-small ml-1" data-column="${columnIndex}"></button></span>`;
      })
      .join(' ');

    filtersContainer.html(filterTags);
    containerWrapper.show();

    filtersContainer.find('.delete').on('click', (e) => {
      const columnIndex = $(e.target).data('column');
      delete this.app.state.currentFilters[columnIndex];
      this.applyFilters();
      this.updateActiveFiltersDisplay();
    });
  }

  /**
   * Gets a human-readable description of a filter.
   * @param {Object} filter - The filter object containing type and value.
   * @return {string} - A description of the filter.
   */
  getFilterDescription(filter) {
    switch (filter.type) {
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
   * Shows filter statistics to the user.
   */
  showFilterStats() {
    const stats = this.getFilterStats();
    const message = `Showing ${stats.filteredRows} of ${stats.totalRows} rows (${stats.hiddenPercentage}% hidden)`;
    this.app.utilities.showMessage(message, 'info', 3000);
  }
}
