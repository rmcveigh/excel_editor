/**
 * @file
 * Excel Editor Column Manager Module
 *
 * Handles column visibility management in the Excel Editor.
 */

/* eslint-disable no-console */
(function ($) {
  'use strict';

  window.ExcelEditorColumnManager = function () {
    // =========================================================================
    // COLUMN VISIBILITY
    // =========================================================================

    /**
     * Shows the modal for managing column visibility.
     */
    this.showColumnVisibilityModal = function () {
      if (!this.data.filtered.length) {
        this.showMessage('No data loaded', 'warning');
        return;
      }

      const modal = $(this._buildColumnVisibilityModalHtml());
      $('body').append(modal);
      this.bindColumnModalEvents(modal);
    };

    /**
     * [HELPER] Builds the HTML string for the column visibility modal.
     * @returns {string} The complete HTML string for the modal.
     */
    this._buildColumnVisibilityModalHtml = function () {
      const headers = this.data.filtered[0] || [];

      const checkboxesHtml = headers
        .map((header, index) => {
          const isVisible = !this.state.hiddenColumns.has(index);
          const isEditable = this.config.editableColumns.includes(header);
          const editableTag = isEditable
            ? `<span class="tag is-small is-info ml-2">${Drupal.t(
                'Editable'
              )}</span>`
            : '';

          return `
          <div class="column is-half">
            <label class="checkbox">
              <input type="checkbox"
                     class="column-visibility-checkbox"
                     data-column-index="${index}"
                     ${isVisible ? 'checked' : ''}>
              <span class="column-name">${this.escapeHtml(header)}</span>
              ${editableTag}
            </label>
          </div>`;
        })
        .join('');

      return `
      <div class="modal is-active" id="column-visibility-modal">
        <div class="modal-background"></div>
        <div class="modal-content">
          <div class="box">
            <h3 class="title is-4">Manage Column Visibility</h3>
            <div class="field is-grouped">
              <button class="button is-small" id="show-all-columns">Show All</button>
              <button class="button is-small" id="show-only-editable">Show Only Editable</button>
            </div>
            <div class="column-checkboxes columns is-multiline">
              ${checkboxesHtml}
            </div>
            <div class="field is-grouped is-grouped-right">
              <button class="button" id="cancel-column-visibility">Cancel</button>
              <button class="button is-primary" id="apply-column-visibility">Apply</button>
            </div>
          </div>
        </div>
        <button class="modal-close is-large" aria-label="close"></button>
      </div>`;
    };

    /**
     * Binds events for the column visibility modal.
     * @param {jQuery} modal The jQuery object for the modal.
     */
    this.bindColumnModalEvents = function (modal) {
      // Close modal handlers
      const closeElements = modal.find(
        '.modal-close, #cancel-column-visibility, .modal-background'
      );
      closeElements.on('click', () => modal.remove());

      // Show all columns handler
      modal.find('#show-all-columns').on('click', () => {
        modal.find('.column-visibility-checkbox').prop('checked', true);
      });

      // Show only editable columns handler
      modal.find('#show-only-editable').on('click', () => {
        const checkboxes = modal.find('.column-visibility-checkbox');
        checkboxes.each((index, checkbox) => {
          const colIndex = parseInt($(checkbox).data('column-index'), 10);
          const header = this.data.filtered[0][colIndex];
          $(checkbox).prop(
            'checked',
            this.config.editableColumns.includes(header)
          );
        });
      });

      // Apply changes handler
      modal.find('#apply-column-visibility').on('click', () => {
        this.applyColumnVisibilityChanges(modal);
        modal.remove();
      });
    };

    /**
     * Applies the column visibility changes selected in the modal.
     * @param {jQuery} modal The jQuery object for the modal.
     */
    this.applyColumnVisibilityChanges = async function (modal) {
      const isLargeDataset = this.data.filtered.length > 200;

      try {
        if (isLargeDataset) {
          this.showProcessLoader('Updating column visibility...');
        } else {
          this.showQuickLoader('Updating columns...');
        }

        // Small delay to ensure loader is visible
        await new Promise((resolve) => setTimeout(resolve, 50));

        // Update hidden columns state
        this.state.hiddenColumns.clear();
        modal.find('.column-visibility-checkbox').each((index, checkbox) => {
          if (!$(checkbox).is(':checked')) {
            const columnIndex = parseInt($(checkbox).data('column-index'), 10);
            this.state.hiddenColumns.add(columnIndex);
          }
        });

        await this.renderTable();
        this.setupFilters();
      } catch (error) {
        console.error('Error applying column visibility changes:', error);
        this.showMessage('Failed to update column visibility.', 'error');
      } finally {
        if (isLargeDataset) {
          this.hideProcessLoader();
        } else {
          this.hideQuickLoader();
        }
      }
    };

    /**
     * Resets the visible columns to the defaults specified in the module settings.
     */
    this.resetToDefaultColumns = async function () {
      if (!this.data.filtered.length) {
        this.showMessage('No data loaded', 'warning');
        return;
      }

      try {
        this.showQuickLoader('Resetting columns...');
        this.applyDefaultColumnVisibility();
        await this.renderTable();
        this.setupFilters();
        this.showMessage('Columns reset to default visibility.', 'success');
      } catch (error) {
        console.error('Error resetting columns:', error);
        this.showMessage('Failed to reset columns.', 'error');
      } finally {
        this.hideQuickLoader();
      }
    };

    /**
     * Overrides any settings and makes all columns visible.
     */
    this.showAllColumnsOverride = async function () {
      try {
        this.showQuickLoader('Showing all columns...');
        this.state.hiddenColumns.clear();
        await this.renderTable();
        this.setupFilters();
        this.showMessage('All columns are now visible.', 'success');
      } catch (error) {
        console.error('Error showing all columns:', error);
        this.showMessage('Failed to show all columns.', 'error');
      } finally {
        this.hideQuickLoader();
      }
    };
  };
})(jQuery);
