/**
 * @file
 * Excel Editor Column Manager Module
 */

export class ExcelEditorColumnManager {
  constructor(app) {
    this.app = app;
  }

  /**
   * Shows the modal for managing column visibility.
   */
  showColumnVisibilityModal() {
    if (!this.app.data.filtered.length) {
      this.app.utilities.showMessage('No data loaded', 'warning');
      return;
    }

    const modal = jQuery(this.buildColumnVisibilityModalHtml());
    jQuery('body').append(modal);
    this.bindColumnModalEvents(modal);
  }

  /**
   * Builds the HTML string for the column visibility modal.
   * @returns {string} The HTML string for the modal.
   */
  buildColumnVisibilityModalHtml() {
    const headers = this.app.data.filtered[0] || [];

    const checkboxesHtml = headers
      .map((header, index) => {
        const isVisible = !this.app.state.hiddenColumns.has(index);
        const isEditable = this.app.config.editableColumns.includes(header);
        const editableTag = isEditable
          ? `<span class="tag is-small is-info ml-2">${Drupal.t('Editable')}</span>`
          : '';

        return `
        <div class="column is-half">
          <label class="checkbox">
            <input type="checkbox"
                   class="column-visibility-checkbox"
                   data-column-index="${index}"
                   ${isVisible ? 'checked' : ''}>
            <span class="column-name">${this.app.utilities.escapeHtml(header)}</span>
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
  }

  /**
   * Binds events for the column visibility modal.
   * @param {jQuery} modal The jQuery object representing the modal.
   */
  bindColumnModalEvents(modal) {
    // Close modal handlers
    const closeElements = modal.find('.modal-close, #cancel-column-visibility, .modal-background');
    closeElements.on('click', () => modal.remove());

    // Show all columns handler
    modal.find('#show-all-columns').on('click', () => {
      modal.find('.column-visibility-checkbox').prop('checked', true);
    });

    // Show only editable columns handler
    modal.find('#show-only-editable').on('click', () => {
      const checkboxes = modal.find('.column-visibility-checkbox');
      checkboxes.each((index, checkbox) => {
        const colIndex = parseInt(jQuery(checkbox).data('column-index'), 10);
        const header = this.app.data.filtered[0][colIndex];
        jQuery(checkbox).prop('checked', this.app.config.editableColumns.includes(header));
      });
    });

    // Apply changes handler
    modal.find('#apply-column-visibility').on('click', () => {
      this.applyColumnVisibilityChanges(modal);
      modal.remove();
    });
  }

  /**
   * Applies the column visibility changes selected in the modal.
   * @param {jQuery} modal The jQuery object representing the modal.
   */
  async applyColumnVisibilityChanges(modal) {
    const isLargeDataset = this.app.data.filtered.length > 200;

    try {
      if (isLargeDataset) {
        this.app.utilities.showProcessLoader('Updating column visibility...');
      } else {
        this.app.utilities.showQuickLoader('Updating columns...');
      }

      // Small delay to ensure loader is visible
      await new Promise((resolve) => setTimeout(resolve, 50));

      // Update hidden columns state
      this.app.state.hiddenColumns.clear();
      modal.find('.column-visibility-checkbox').each((index, checkbox) => {
        if (!jQuery(checkbox).is(':checked')) {
          const columnIndex = parseInt(jQuery(checkbox).data('column-index'), 10);
          this.app.state.hiddenColumns.add(columnIndex);
        }
      });

      await this.app.uiRenderer.renderTable();
      this.app.filterManager.setupFilters();
    } catch (error) {
      console.error('Error applying column visibility changes:', error);
      this.app.utilities.showMessage('Failed to update column visibility.', 'error');
    } finally {
      if (isLargeDataset) {
        this.app.utilities.hideProcessLoader();
      } else {
        this.app.utilities.hideQuickLoader();
      }
    }
  }

  /**
   * Resets the visible columns to the defaults specified in the module settings.
   */
  async resetToDefaultColumns() {
    if (!this.app.data.filtered.length) {
      this.app.utilities.showMessage('No data loaded', 'warning');
      return;
    }

    try {
      this.app.utilities.showQuickLoader('Resetting columns...');
      this.app.dataManager.applyDefaultColumnVisibility();
      await this.app.uiRenderer.renderTable();
      this.app.filterManager.setupFilters();
      this.app.utilities.showMessage('Columns reset to default visibility.', 'success');
    } catch (error) {
      console.error('Error resetting columns:', error);
      this.app.utilities.showMessage('Failed to reset columns.', 'error');
    } finally {
      this.app.utilities.hideQuickLoader();
    }
  }

  /**
   * Overrides any settings and makes all columns visible.
   */
  async showAllColumnsOverride() {
    try {
      this.app.utilities.showQuickLoader('Showing all columns...');
      this.app.state.hiddenColumns.clear();
      await this.app.uiRenderer.renderTable();
      this.app.filterManager.setupFilters();
      this.app.utilities.showMessage('All columns are now visible.', 'success');
    } catch (error) {
      console.error('Error showing all columns:', error);
      this.app.utilities.showMessage('Failed to show all columns.', 'error');
    } finally {
      this.app.utilities.hideQuickLoader();
    }
  }
}
