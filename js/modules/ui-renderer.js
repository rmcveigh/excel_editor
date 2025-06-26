/**
 * @file
 * Excel Editor UI Renderer Module
 *
 * Handles table creation, modals, interface rendering, and UI updates.
 * This module adds UI rendering methods to the ExcelEditor class.
 */

/* eslint-disable no-console */
(function ($) {
  'use strict';

  /**
   * UI Renderer module for Excel Editor.
   * This function is called on ExcelEditor instances to add UI rendering methods.
   */
  window.ExcelEditorUIRenderer = function () {
    // =========================================================================
    // MAIN INTERFACE RENDERING
    // =========================================================================

    /**
     * Renders the main interface after data is loaded.
     */
    this.renderInterface = function () {
      this.elements.uploadArea.hide();
      this.elements.mainArea.show();
      this.renderTable();
      this.setupFilters();
    };

    /**
     * Renders the main data table with enhanced loading indicators.
     */
    this.renderTable = async function () {
      if (!this.data.filtered.length) {
        this.elements.tableContainer.html(
          '<p class="has-text-centered">No data available</p>'
        );
        return;
      }

      const shouldShowLoader = this.data.filtered.length > 200;

      if (shouldShowLoader) {
        this.showProcessLoader('Rebuilding table...');
        // Add delay to ensure loader shows
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

        this.elements.tableContainer.html(fragment);
        this.elements.table = $('#excel-table');
        this.bindTableEvents();
        this.applyRowStyling();
      } finally {
        if (shouldShowLoader) {
          this.hideProcessLoader();
        }
      }
    };

    // =========================================================================
    // TABLE STRUCTURE CREATION
    // =========================================================================

    /**
     * Creates the table header (<thead>).
     * @returns {HTMLTableSectionElement} The created thead element.
     */
    this.createTableHeader = function () {
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
    };

    /**
     * Creates the table body (<tbody>) - legacy synchronous version.
     * @returns {HTMLTableSectionElement} The created tbody element.
     */
    this.createTableBody = function () {
      const tbody = document.createElement('tbody');
      for (let i = 1; i < this.data.filtered.length; i++) {
        tbody.appendChild(this.createTableRow(i));
      }
      return tbody;
    };

    /**
     * Creates an async version of createTableBody for better performance
     */
    this.createTableBodyAsync = async function () {
      const tbody = document.createElement('tbody');
      const batchSize = 50; // Process rows in batches

      for (let i = 1; i < this.data.filtered.length; i += batchSize) {
        const endIndex = Math.min(i + batchSize, this.data.filtered.length);

        // Process batch
        for (let j = i; j < endIndex; j++) {
          tbody.appendChild(this.createTableRow(j));
        }

        // Yield control to prevent UI blocking for large datasets
        if (this.data.filtered.length > 500 && i % (batchSize * 4) === 1) {
          await new Promise((resolve) => setTimeout(resolve, 10));
        }
      }

      return tbody;
    };

    /**
     * Creates an individual table row (<tr>).
     * @param {number} rowIndex The index of the row.
     * @returns {HTMLTableRowElement} The created tr element.
     */
    this.createTableRow = function (rowIndex) {
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
    };

    /**
     * Creates an individual table cell (<td>).
     * @param {number} rowIndex The row index.
     * @param {number} colIndex The column index.
     * @param {string} cellValue The value of the cell.
     * @returns {HTMLTableCellElement} The created td element.
     */
    this.createTableCell = function (rowIndex, colIndex, cellValue) {
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
    };

    // =========================================================================
    // CELL INPUT CREATION
    // =========================================================================

    /**
     * Creates an actions dropdown for a table cell.
     * @param {number} rowIndex The row index.
     * @param {number} colIndex The column index.
     * @param {string} value The current value.
     * @returns {string} HTML string for the dropdown.
     */
    this.createActionsDropdown = function (rowIndex, colIndex, value) {
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
    };

    /**
     * Creates a notes textarea for a table cell.
     * @param {number} rowIndex The row index.
     * @param {number} colIndex The column index.
     * @param {string} value The current value.
     * @returns {string} HTML string for the textarea.
     */
    this.createNotesTextarea = function (rowIndex, colIndex, value) {
      return `<textarea class="excel-editor-cell editable notes-textarea" data-row="${rowIndex}" data-col="${colIndex}" placeholder="${Drupal.t(
        'Add notes...'
      )}" rows="2">${this.escapeHtml(value || '')}</textarea>`;
    };

    /**
     * Creates a text input for a table cell.
     * @param {number} rowIndex The row index.
     * @param {number} colIndex The column index.
     * @param {string} value The current value.
     * @param {string} placeholder The placeholder text.
     * @returns {string} HTML string for the input.
     */
    this.createTextInput = function (rowIndex, colIndex, value, placeholder) {
      return `<input type="text" class="excel-editor-cell editable" data-row="${rowIndex}" data-col="${colIndex}" value="${this.escapeHtml(
        value || ''
      )}" placeholder="${Drupal.t(placeholder)}" />`;
    };

    // =========================================================================
    // MODAL CREATION HELPERS
    // =========================================================================

    /**
     * Creates a basic modal structure.
     * @param {string} id The modal ID.
     * @param {string} title The modal title.
     * @param {string} content The modal content HTML.
     * @param {object} options Additional modal options.
     * @returns {jQuery} The modal element.
     */
    this.createModal = function (id, title, content, options = {}) {
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
        ? `
        <footer class="modal-card-foot">
          ${settings.footerButtons
            .map(
              (btn) =>
                `<button class="button ${btn.class || ''}" id="${
                  btn.id || ''
                }">${btn.text}</button>`
            )
            .join('')}
        </footer>`
        : '';

      const modalHtml = `
        <div class="modal is-active ${settings.customClass}" id="${id}">
          <div class="modal-background"></div>
          <div class="modal-card ${sizeClass}">
            <header class="modal-card-head">
              <p class="modal-card-title">${this.escapeHtml(title)}</p>
              ${
                settings.showCloseButton
                  ? '<button class="delete" aria-label="close"></button>'
                  : ''
              }
            </header>
            <section class="modal-card-body">
              ${content}
            </section>
            ${footerHtml}
          </div>
        </div>`;

      const modal = $(modalHtml);
      $('body').append(modal);

      // Bind close events
      if (settings.showCloseButton) {
        modal
          .find('.delete, .modal-background')
          .on('click', () => modal.remove());
      }

      return modal;
    };

    /**
     * Creates a confirmation dialog.
     * @param {string} title The dialog title.
     * @param {string} message The confirmation message.
     * @param {function} onConfirm Callback for confirmation.
     * @param {function} onCancel Callback for cancellation.
     * @returns {jQuery} The modal element.
     */
    this.createConfirmDialog = function (
      title,
      message,
      onConfirm,
      onCancel = null
    ) {
      const content = `<p>${this.escapeHtml(message)}</p>`;

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
    };

    /**
     * Creates a loading modal for long operations.
     * @param {string} message The loading message.
     * @returns {jQuery} The modal element.
     */
    this.createLoadingModal = function (message) {
      const content = `
        <div class="has-text-centered">
          <div class="loader mb-4"></div>
          <p>${this.escapeHtml(message)}</p>
        </div>`;

      return this.createModal(
        'loading-modal',
        Drupal.t('Please Wait'),
        content,
        {
          showCloseButton: false,
          showFooter: false,
          customClass: 'loading-modal',
        }
      );
    };

    // =========================================================================
    // UI UPDATE HELPERS
    // =========================================================================

    /**
     * Updates the table header with current column state.
     */
    this.updateTableHeader = function () {
      if (!this.elements.table || !this.elements.table.length) return;

      const thead = this.elements.table.find('thead');
      if (thead.length) {
        thead.replaceWith(this.createTableHeader());
      }
    };

    /**
     * Updates specific table cells without full re-render.
     * @param {Array} updates Array of {row, col, value} objects.
     */
    this.updateTableCells = function (updates) {
      if (!this.elements.table || !this.elements.table.length) return;

      updates.forEach((update) => {
        const { row, col, value } = update;
        const cell = this.elements.table.find(
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
    };

    /**
     * Highlights specific rows or cells.
     * @param {Array} targets Array of {row, col?, class} objects.
     * @param {number} duration Duration to show highlight (ms).
     */
    this.highlightElements = function (targets, duration = 2000) {
      targets.forEach((target) => {
        let selector;
        if (target.col !== undefined) {
          // Highlight specific cell
          selector = `[data-row="${target.row}"][data-col="${target.col}"]`;
        } else {
          // Highlight entire row
          selector = `tr[data-row="${target.row}"]`;
        }

        const element = this.elements.table.find(selector);
        if (element.length) {
          element.addClass(target.class || 'highlight');
          setTimeout(() => {
            element.removeClass(target.class || 'highlight');
          }, duration);
        }
      });
    };

    /**
     * Scrolls the table to a specific row.
     * @param {number} rowIndex The row index to scroll to.
     */
    this.scrollToRow = function (rowIndex) {
      const row = this.elements.table.find(`tr[data-row="${rowIndex}"]`);
      if (row.length) {
        this.elements.tableContainer.animate(
          {
            scrollTop:
              row.offset().top - this.elements.tableContainer.offset().top,
          },
          500
        );
      }
    };

    // =========================================================================
    // RESPONSIVE & ACCESSIBILITY HELPERS
    // =========================================================================

    /**
     * Adjusts table for mobile view.
     */
    this.adjustForMobile = function () {
      const isMobile = $(window).width() < 768;

      if (isMobile) {
        this.elements.tableContainer.addClass('mobile-view');
        // Hide less important columns on mobile
        const lessImportantColumns =
          this.data.filtered[0]
            ?.map((header, index) => {
              if (
                !this.config.editableColumns.includes(header) &&
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
          this.state.hiddenColumns.add(colIndex);
        });
      } else {
        this.elements.tableContainer.removeClass('mobile-view');
      }
    };

    /**
     * Enhances table accessibility.
     */
    this.enhanceAccessibility = function () {
      if (!this.elements.table || !this.elements.table.length) return;

      // Add ARIA labels
      this.elements.table.attr('role', 'grid');
      this.elements.table.find('th').attr('role', 'columnheader');
      this.elements.table.find('td').attr('role', 'gridcell');

      // Add keyboard navigation
      this.elements.table.on('keydown', 'input, textarea, select', (e) => {
        if (e.key === 'Tab') {
          // Custom tab navigation logic could be added here
        }
      });
    };

    // =========================================================================
    // PERFORMANCE OPTIMIZATION
    // =========================================================================

    /**
     * Virtualizes table rows for better performance with large datasets.
     * @param {number} startIndex Starting row index.
     * @param {number} endIndex Ending row index.
     */
    this.virtualizeTableRows = function (startIndex, endIndex) {
      // Implementation for virtual scrolling
      // This would be useful for very large datasets (10,000+ rows)
      this.logDebug(`Virtualizing rows ${startIndex} to ${endIndex}`);
    };

    /**
     * Debounced table update to prevent excessive re-renders.
     */
    this.debouncedTableUpdate = this.debounce(function () {
      this.renderTable();
    }, 250);

    // =========================================================================
    // TABLE STATE MANAGEMENT
    // =========================================================================

    /**
     * Saves the current table state for restoration.
     * @returns {object} The saved table state.
     */
    this.saveTableState = function () {
      return {
        scrollTop: this.elements.tableContainer.scrollTop(),
        scrollLeft: this.elements.tableContainer.scrollLeft(),
        selectedRows: Array.from(this.data.selected),
        hiddenColumns: Array.from(this.state.hiddenColumns),
      };
    };

    /**
     * Restores a previously saved table state.
     * @param {object} state The table state to restore.
     */
    this.restoreTableState = function (state) {
      if (state.scrollTop !== undefined) {
        this.elements.tableContainer.scrollTop(state.scrollTop);
      }
      if (state.scrollLeft !== undefined) {
        this.elements.tableContainer.scrollLeft(state.scrollLeft);
      }
      if (state.selectedRows) {
        this.data.selected = new Set(state.selectedRows);
        this.updateSelectionCount();
        this.updateSelectAllCheckbox();
      }
      if (state.hiddenColumns) {
        this.state.hiddenColumns = new Set(state.hiddenColumns);
      }
    };

    this.logDebug('ExcelEditorUIRenderer module loaded');
  };
})(jQuery);
