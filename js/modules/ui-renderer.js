/**
 * @file
 * Excel Editor UI Renderer Module
 */

export class ExcelEditorUIRenderer {
  constructor(app) {
    this.app = app;
  }

  /**
   * Renders the main interface after data is loaded.
   */
  renderInterface() {
    this.app.elements.uploadArea.hide();
    this.app.elements.mainArea.show();
    this.renderTable();
    this.app.filterManager.setupFilters();
  }

  /**
   * Renders the main data table with enhanced loading indicators.
   * @returns {Promise<void>} A promise that resolves when the table is fully rendered.
   */
  async renderTable() {
    // Clear existing table
    this.app.elements.table.empty();

    if (!this.app.data.filtered || !this.app.data.filtered.length) {
      this.app.elements.table.html('<p>No data to display</p>');
      return;
    }

    const shouldShowLoader = this.app.data.filtered.length > 200;

    if (shouldShowLoader) {
      this.app.utilities.showProcessLoader('Rebuilding table...');
      await new Promise((resolve) => setTimeout(resolve, 100));
    }

    try {
      const fragment = document.createDocumentFragment();
      const table = document.createElement('table');
      table.className = 'excel-editor-table table is-fullwidth is-striped';
      table.id = 'excel-table';

      table.appendChild(this.createTableHeader());
      table.appendChild(await this.createTableBodyAsync());
      fragment.appendChild(table);

      this.app.elements.tableContainer.html(fragment);
      this.app.elements.table = jQuery('#excel-table');
      this.app.bindTableEvents();
      this.app.dataManager.applyRowStyling();

      // Auto-validate barcode fields after table is rendered
      if (this.app.validationManager) {
        // Small delay to ensure DOM is fully updated
        setTimeout(() => {
          this.app.validationManager.validateExistingBarcodeFields();
          // Refresh filter controls to update validation filter counts
          if (this.app.filterManager) {
            this.app.filterManager.setupFilters();
          }
        }, 100);
      }

      // After the table is rendered, fetch links for subject IDs
      this.fetchSubjectIdLinks();
    } catch (error) {
      this.app.utilities.handleError('Error rendering table', error);
    } finally {
      if (shouldShowLoader) {
        this.app.utilities.hideProcessLoader();
      }
    }
  }

  /**
   * Creates the table header (<thead>).
   * @return {HTMLTableSectionElement} The created <thead> element.
   */
  createTableHeader() {
    const thead = document.createElement('thead');
    const headerRow = document.createElement('tr');
    headerRow.innerHTML =
      '<th class="selection-column"><label class="checkbox"><input type="checkbox" id="select-all-checkbox" /></label></th>';

    this.app.data.filtered[0].forEach((header, index) => {
      if (!this.app.state.hiddenColumns.has(index)) {
        const th = document.createElement('th');
        th.dataset.column = index;
        th.innerHTML = `${this.app.utilities.escapeHtml(
          header
        )}<br><small><a href="#" class="filter-link" data-column="${index}">Filter</a></small>`;
        headerRow.appendChild(th);
      }
    });

    thead.appendChild(headerRow);
    return thead;
  }

  /**
   * Creates the table body (<tbody>) - legacy synchronous version.
   * @return {HTMLTableSectionElement} The created <tbody> element.
   */
  createTableBody() {
    const tbody = document.createElement('tbody');
    for (let i = 1; i < this.app.data.filtered.length; i++) {
      tbody.appendChild(this.createTableRow(i));
    }
    return tbody;
  }

  /**
   * Creates an async version of createTableBody for better performance
   * @return {Promise<HTMLTableSectionElement>} A promise that resolves to the created <tbody> element.
   */
  async createTableBodyAsync() {
    const tbody = document.createElement('tbody');
    const batchSize = 50;

    for (let i = 1; i < this.app.data.filtered.length; i += batchSize) {
      const endIndex = Math.min(i + batchSize, this.app.data.filtered.length);

      for (let j = i; j < endIndex; j++) {
        tbody.appendChild(this.createTableRow(j));
      }

      if (this.app.data.filtered.length > 500 && i % (batchSize * 4) === 1) {
        await new Promise((resolve) => setTimeout(resolve, 10));
      }
    }

    return tbody;
  }

  /**
   * Creates an individual table row (<tr>).
   * @param {number} rowIndex - The index of the row in the filtered data.
   * @return {HTMLTableRowElement} The created table row element.
   */
  createTableRow(rowIndex) {
    const row = document.createElement('tr');
    row.dataset.row = rowIndex;
    const rowData = this.app.data.filtered[rowIndex];
    const isSelected = this.app.data.selected.has(rowIndex);
    if (isSelected) row.classList.add('selected-row');

    row.innerHTML = `<td class="selection-column"><label class="checkbox"><input type="checkbox" class="row-checkbox" data-row="${rowIndex}" ${
      isSelected ? 'checked' : ''
    } /></label></td>`;

    rowData.forEach((cell, colIndex) => {
      if (!this.app.state.hiddenColumns.has(colIndex)) {
        row.appendChild(this.createTableCell(rowIndex, colIndex, cell));
      }
    });
    return row;
  }

  /**
   * Creates an individual table cell (<td>).
   * @param {number} rowIndex - The index of the row in the filtered data.
   * @param {number} colIndex - The index of the column in the filtered data.
   * @param {string} cellValue - The value to display in the cell.
   * @return {HTMLTableCellElement} The created table cell element.
   */
  createTableCell(rowIndex, colIndex, cellValue) {
    const td = document.createElement('td');
    const columnName = this.app.data.filtered[0][colIndex];
    const isEditable = this.app.config.editableColumns.includes(columnName);
    const isSubjectId =
      columnName && columnName.toLowerCase().includes('subject id');

    // Add column-specific class using safe CSS identifier cleaner
    const columnClass = this.createColumnClass(columnName);
    td.classList.add(columnClass);

    // Add data attribute for JavaScript targeting
    if (columnName) {
      td.setAttribute('data-column-name', columnName);
    }

    // Add specific class based on column name
    if (isSubjectId) {
      td.classList.add('subject-id-cell');
    }

    if (isEditable) {
      td.classList.add('editable-column');

      if (columnName === 'actions') {
        td.classList.add('actions-column');
        td.innerHTML = this.createActionsDropdown(
          rowIndex,
          colIndex,
          cellValue
        );
      } else if (columnName === 'notes') {
        td.classList.add('notes-column');
        td.innerHTML = this.createNotesTextarea(rowIndex, colIndex, cellValue);
      } else {
        // For new_barcode and other editable fields
        td.innerHTML = this.createTextInput(
          rowIndex,
          colIndex,
          cellValue,
          'Enter barcode...'
        );
      }
    } else {
      td.classList.add('readonly-cell');

      if (isSubjectId) {
        td.innerHTML = `<span class="excel-editor-readonly">${this.app.utilities.escapeHtml(
          cellValue || ''
        )}</span>`;
      } else {
        td.innerHTML = `<span class="excel-editor-readonly">${this.app.utilities.escapeHtml(
          cellValue || ''
        )}</span>`;
      }
    }

    return td;
  }

  /**
   * Creates a CSS-safe class name from a column name.
   * Mimics Drupal's Html::cleanCssIdentifier() PHP function.
   * @param {string} columnName - The column name to convert.
   * @return {string} A CSS-safe class name.
   */
  createColumnClass(columnName) {
    if (!columnName) return 'column-unknown';

    let cleanName = String(columnName)
      // Convert to lowercase
      .toLowerCase()
      // Replace any character that's not a-z, 0-9, hyphen, or underscore with hyphen
      .replace(/[^a-z0-9\-_]/g, '-')
      // Remove leading numbers (CSS identifiers can't start with numbers)
      .replace(/^[0-9]+/, '')
      // Collapse multiple consecutive hyphens into single hyphen
      .replace(/-+/g, '-')
      // Remove leading and trailing hyphens/underscores
      .replace(/^[-_]+|[-_]+$/g, '');

    // Ensure we have a valid identifier
    if (!cleanName || cleanName === '') {
      cleanName = 'column';
    }

    return 'column-' + cleanName;
  }

  /**
   * Creates an actions dropdown for a table cell.
   * @param {number} rowIndex - The index of the row in the filtered data.
   * @param {number} colIndex - The index of the column in the filtered data.
   * @param {string} value - The current value of the cell (if any).
   * @return {string} The HTML string for the dropdown.
   */
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

  /**
   * Creates a notes textarea for a table cell.
   * @param {number} rowIndex - The index of the row in the filtered data.
   * @param {number} colIndex - The index of the column in the filtered data.
   * @param {string} value - The current value of the cell (if any).
   * @return {string} The HTML string for the textarea.
   */
  createNotesTextarea(rowIndex, colIndex, value) {
    return `<textarea class="excel-editor-cell editable notes-textarea" data-row="${rowIndex}" data-col="${colIndex}" placeholder="${Drupal.t(
      'Add notes...'
    )}" rows="2">${this.app.utilities.escapeHtml(value || '')}</textarea>`;
  }

  /**
   * Creates a text input for a table cell.
   * @param {number} rowIndex - The index of the row in the filtered data.
   * @param {number} colIndex - The index of the column in the filtered data.
   * @param {string} value - The current value of the cell (if any).
   * @return {string} The HTML string for the input.
   */
  createTextInput(rowIndex, colIndex, value, placeholder) {
    return `<input type="text" class="excel-editor-cell editable" data-row="${rowIndex}" data-col="${colIndex}" value="${this.app.utilities.escapeHtml(
      value || ''
    )}" placeholder="${Drupal.t(placeholder)}" />`;
  }

  /**
   * Creates a basic modal structure.
   * @param {string} id - The ID for the modal.
   * @param {string} title - The title of the modal.
   * @param {string} content - The HTML content for the modal body.
   * @param {Object} options - Additional options for the modal.
   * @return {jQuery} The created modal element.
   */
  createModal(id, title, content, options = {}) {
    const defaults = {
      size: '',
      showCloseButton: true,
      showFooter: true,
      footerButtons: [],
      customClass: '',
    };

    const settings = { ...defaults, ...options };
    const sizeClass = settings.size ? `modal-${settings.size}` : '';

    const footerHtml = settings.showFooter
      ? `<footer class="modal-card-foot">${settings.footerButtons
          .map(
            (btn) =>
              `<button class="button ${btn.class || ''}" id="${btn.id || ''}">${
                btn.text
              }</button>`
          )
          .join('')}</footer>`
      : '';

    const modalHtml = `
      <div class="modal is-active ${settings.customClass}" id="${id}">
        <div class="modal-background"></div>
        <div class="modal-card ${sizeClass}">
          <header class="modal-card-head">
            <p class="modal-card-title">${this.app.utilities.escapeHtml(
              title
            )}</p>
            ${
              settings.showCloseButton
                ? '<button class="delete" aria-label="close"></button>'
                : ''
            }
          </header>
          <section class="modal-card-body">${content}</section>
          ${footerHtml}
        </div>
      </div>`;

    const modal = jQuery(modalHtml);
    jQuery('body').append(modal);

    if (settings.showCloseButton) {
      modal
        .find('.delete, .modal-background')
        .on('click', () => modal.remove());
    }

    return modal;
  }

  /**
   * Creates a confirmation dialog.
   * @param {string} title - The title of the dialog.
   * @param {string} message - The confirmation message.
   * @param {Function} onConfirm - Callback for confirmation action.
   * @param {Function} [onCancel] - Optional callback for cancellation action.
   * @return {jQuery} The created confirmation dialog element.
   */
  createConfirmDialog(title, message, onConfirm, onCancel = null) {
    const content = `<p>${this.app.utilities.escapeHtml(message)}</p>`;

    const modal = this.createModal('confirm-dialog', title, content, {
      footerButtons: [
        { id: 'confirm-cancel', text: Drupal.t('Cancel'), class: '' },
        { id: 'confirm-ok', text: Drupal.t('Confirm'), class: 'is-primary' },
      ],
    });

    modal.find('#confirm-cancel').on('click', () => {
      if (onCancel) onCancel();
      modal.remove();
    });

    modal.find('#confirm-ok').on('click', () => {
      if (onConfirm) onConfirm();
      modal.remove();
    });

    return modal;
  }

  /**
   * Creates a loading modal for long operations.
   * @param {string} message - The message to display in the modal.
   * @return {jQuery} The created loading modal element.
   */
  createLoadingModal(message) {
    const content = `
      <div class="has-text-centered">
        <div class="loader mb-4"></div>
        <p>${this.app.utilities.escapeHtml(message)}</p>
      </div>`;

    return this.createModal('loading-modal', Drupal.t('Please Wait'), content, {
      showCloseButton: false,
      showFooter: false,
      customClass: 'loading-modal',
    });
  }

  /**
   * Updates the table header with current column state.
   */
  updateTableHeader() {
    if (!this.app.elements.table || !this.app.elements.table.length) return;

    const thead = this.app.elements.table.find('thead');
    if (thead.length) {
      thead.replaceWith(this.createTableHeader());
    }
  }

  /**
   * Updates specific table cells without full re-render.
   * @param {Array<Object>} updates - Array of objects with row, col, and value properties.
   */
  updateTableCells(updates) {
    if (!this.app.elements.table || !this.app.elements.table.length) return;

    updates.forEach((update) => {
      const { row, col, value } = update;
      const cell = this.app.elements.table.find(
        `[data-row="${row}"][data-col="${col}"]`
      );
      if (cell.length) {
        if (cell.is('input, textarea, select')) {
          cell.val(value);
        } else {
          cell.text(value);
        }
      }
    });
  }

  /**
   * Highlights specific rows or cells.
   * @param {Array<Object>} targets - Array of objects with row, col, and class properties.
   * @param {number} [duration=2000] - Duration to keep the highlight (in milliseconds).
   */
  highlightElements(targets, duration = 2000) {
    targets.forEach((target) => {
      let selector;
      if (target.col !== undefined) {
        selector = `[data-row="${target.row}"][data-col="${target.col}"]`;
      } else {
        selector = `tr[data-row="${target.row}"]`;
      }

      const element = this.app.elements.table.find(selector);
      if (element.length) {
        element.addClass(target.class || 'highlight');
        setTimeout(() => {
          element.removeClass(target.class || 'highlight');
        }, duration);
      }
    });
  }

  /**
   * Scrolls the table to a specific row.
   * @param {number} rowIndex - The index of the row to scroll to.
   */
  scrollToRow(rowIndex) {
    const row = this.app.elements.table.find(`tr[data-row="${rowIndex}"]`);
    if (row.length) {
      this.app.elements.tableContainer.animate(
        {
          scrollTop:
            row.offset().top - this.app.elements.tableContainer.offset().top,
        },
        500
      );
    }
  }

  /**
   * Adjusts table for mobile view.
   */
  adjustForMobile() {
    const $ = jQuery;
    const isMobile = $(window).width() < 768;

    if (isMobile) {
      this.app.elements.tableContainer.addClass('mobile-view');
      const lessImportantColumns =
        this.app.data.filtered[0]
          ?.map((header, index) => {
            if (
              !this.app.config.editableColumns.includes(header) &&
              !['id', 'name', 'title'].some((key) =>
                header.toLowerCase().includes(key)
              )
            ) {
              return index;
            }
            return null;
          })
          .filter((index) => index !== null) || [];

      lessImportantColumns.slice(3).forEach((colIndex) => {
        this.app.state.hiddenColumns.add(colIndex);
      });
    } else {
      this.app.elements.tableContainer.removeClass('mobile-view');
    }
  }

  /**
   * Enhances table accessibility.
   */
  enhanceAccessibility() {
    if (!this.app.elements.table || !this.app.elements.table.length) return;

    this.app.elements.table.attr('role', 'grid');
    this.app.elements.table.find('th').attr('role', 'columnheader');
    this.app.elements.table.find('td').attr('role', 'gridcell');

    this.app.elements.table.on('keydown', 'input, textarea, select', (e) => {
      if (e.key === 'Tab') {
        // Custom tab navigation logic could be added here
      }
    });
  }

  /**
   * Debounced table update to prevent excessive re-renders.
   */
  getDebouncedTableUpdate() {
    if (!this._debouncedTableUpdate) {
      this._debouncedTableUpdate = this.app.utilities.debounce(() => {
        this.renderTable();
      }, 250);
    }
    return this._debouncedTableUpdate;
  }

  /**
   * Saves the current table state for restoration.
   * @return {Object} An object containing the current scroll position, selected rows, and hidden columns.
   */
  saveTableState() {
    return {
      scrollTop: this.app.elements.tableContainer.scrollTop(),
      scrollLeft: this.app.elements.tableContainer.scrollLeft(),
      selectedRows: Array.from(this.app.data.selected),
      hiddenColumns: Array.from(this.app.state.hiddenColumns),
    };
  }

  /**
   * Restores a previously saved table state.
   * @param {Object} state - The state object containing scroll position, selected rows, and hidden columns.
   */
  restoreTableState(state) {
    if (state.scrollTop !== undefined) {
      this.app.elements.tableContainer.scrollTop(state.scrollTop);
    }
    if (state.scrollLeft !== undefined) {
      this.app.elements.tableContainer.scrollLeft(state.scrollLeft);
    }
    if (state.selectedRows) {
      this.app.data.selected = new Set(state.selectedRows);
      this.app.dataManager.updateSelectionCount();
      this.app.dataManager.updateSelectAllCheckbox();
    }
    if (state.hiddenColumns) {
      this.app.state.hiddenColumns = new Set(state.hiddenColumns);
    }
  }

  /**
   * Fetches entity URLs for visible subject IDs in batch
   */
  fetchSubjectIdLinks() {
    const $ = jQuery;

    // Ensure cache is initialized
    if (!this.app.state.dogEntityUrlCache) {
      this.app.state.dogEntityUrlCache = {};
    }

    // Find all subject ID cells that need links
    const elementsToProcess = $('.subject-id-cell:not(.processed)');

    if (elementsToProcess.length === 0) return;

    // Collect unique IDs that aren't already in cache
    const uniqueIds = new Set();
    elementsToProcess.each((index, element) => {
      try {
        const $element = $(element);
        const grlsId = $element.text().trim();

        if (grlsId && !this.app.state.dogEntityUrlCache[grlsId]) {
          uniqueIds.add(grlsId);
          // Mark as processing to avoid reprocessing
          $element.addClass('processing');
        }
      } catch (error) {
        // Silently handle individual element processing errors
        this.app.utilities.logDebug(
          'Error processing subject ID element:',
          error
        );
      }
    });

    if (uniqueIds.size === 0) return;

    // Get all the URLs in a single batch request
    this.fetchDogEntityUrlsBatch(Array.from(uniqueIds))
      .then((result) => {
        if (result && result.success && result.urls) {
          // Process all the returned URLs
          Object.entries(result.urls).forEach(([grlsId, urlData]) => {
            try {
              // Store in cache (with safety check)
              if (urlData && urlData.url) {
                this.app.state.dogEntityUrlCache[grlsId] = urlData;

                // Update all matching elements with the link
                $(`.subject-id-cell.processing`).each((i, el) => {
                  const $el = $(el);
                  try {
                    // Only process cells that match this ID
                    if ($el.text().trim() === grlsId) {
                      // Create a link but keep the original text
                      $el.html(
                        `<a href="${this.app.utilities.escapeHtml(
                          urlData.url
                        )}" target="_blank">${this.app.utilities.escapeHtml(
                          $el.text().trim()
                        )}</a>`
                      );
                      // Change from 'processing' to 'processed' state
                      $el.removeClass('processing').addClass('processed');
                    }
                  } catch (linkError) {
                    // Silently handle link creation errors
                    this.app.utilities.logDebug(
                      `Error creating link for ${grlsId}:`,
                      linkError
                    );
                    $el.removeClass('processing').addClass('processed');
                  }
                });
              }
            } catch (cacheError) {
              // Silently handle cache storage errors
              this.app.utilities.logDebug(
                `Error caching URL for ${grlsId}:`,
                cacheError
              );
            }
          });
        }

        // Mark any remaining processing elements as processed (silently fail)
        $('.subject-id-cell.processing')
          .removeClass('processing')
          .addClass('processed');
      })
      .catch((error) => {
        // Silently handle batch fetch errors
        this.app.utilities.logDebug('Error fetching dog entity links:', error);

        // Mark failed elements as processed to avoid retrying
        $('.subject-id-cell.processing')
          .removeClass('processing')
          .addClass('processed');
      });
  }

  /**
   * Fetches multiple dog entity URLs in a single request
   * @param {Array} grlsIds - Array of GRLS IDs to fetch URLs for
   * @return {Promise<Object>} - Promise resolving to response with URLs
   */
  async fetchDogEntityUrlsBatch(grlsIds) {
    try {
      // Use the correct endpoint and method
      return await this.app.utilities.apiPost(
        this.app.config.endpoints.getDogEntityUrls,
        { grls_ids: grlsIds }
      );
    } catch (error) {
      // Return a safe empty response instead of throwing
      this.app.utilities.logDebug('Failed to fetch dog entity links:', error);
      return {
        success: false,
        urls: {},
        message: error.message,
      };
    }
  }
}
