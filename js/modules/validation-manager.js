/**
 * @file
 * Excel Editor Validation Manager Module (ES Module)
 */

export class ExcelEditorValidationManager {
  constructor(app) {
    this.app = app;
  }

  /**
   * Validates a barcode field and applies visual feedback
   * @param {jQuery} $cell - The jQuery cell element containing the barcode input
   * @param {string} value - The barcode value to validate
   * @param {number} rowIndex - The index of the row being validated
   */
  validateBarcodeField($cell, value, rowIndex) {
    const validation = this.validateBarcode(value, rowIndex);

    // Remove existing validation classes
    $cell.removeClass('is-valid is-warning is-danger');

    // Remove existing validation messages
    $cell.siblings('.validation-message').remove();

    if (!value || value.trim() === '') {
      // Empty field - no validation styling needed
      return;
    }

    // Apply validation styling based on validation results
    if (!validation.isValid) {
      // Has errors - red border
      $cell.addClass('is-danger');
    }

    // Add validation message if there are issues
    if (validation.messages.length > 0) {
      const messageHtml = `<div class="validation-message has-text-danger is-size-7 mt-1">
        ${validation.messages
          .map((msg) => `<div>• ${this.app.utilities.escapeHtml(msg)}</div>`)
          .join('')}
      </div>`;
      $cell.after(messageHtml);
    }
  }

  /**
   * Validates all existing barcode fields in the rendered table
   */
  validateExistingBarcodeFields() {
    if (!this.app.data.filtered || this.app.data.filtered.length <= 1) {
      return;
    }

    const barcodeColumnIndex = this.app.data.filtered[0].indexOf('new_barcode');
    if (barcodeColumnIndex === -1) {
      return;
    }
    this.app.utilities.showLoading('Validating barcodes...');

    // Find all barcode input fields and validate them
    this.app.elements.tableContainer
      .find('input[data-col="' + barcodeColumnIndex + '"]')
      .each((index, input) => {
        const $input = jQuery(input);
        const rowIndex = parseInt($input.data('row'));
        const value = $input.val() || '';

        // Only validate non-empty fields for auto-validation
        if (value.trim()) {
          this.validateBarcodeField($input, value.trim(), rowIndex);
        }
      });

    this.app.utilities.logDebug(
      'Auto-validated barcode fields after table load'
    );
    this.app.utilities.hideLoading();
  }

  /**
   * Validates a barcode value according to business rules
   * @param {string} value - The barcode value to validate
   * @param {number} excludeRowIndex - Optional row index to exclude from duplicate checks
   * @return {Object} Validation result object with isValid, hasWarnings, and messages
   */
  validateBarcode(value, excludeRowIndex = -1) {
    const result = {
      isValid: true,
      hasWarnings: false,
      messages: [],
    };

    if (!value || value.trim() === '') {
      return result; // Empty values are not validated
    }

    const trimmedValue = value.trim();

    // Check length (16-17 characters)
    if (trimmedValue.length < 16 || trimmedValue.length > 17) {
      result.isValid = false;
      result.messages.push(
        `Length should be 16-17 characters (current: ${trimmedValue.length})`
      );
    }

    // Check for 'X' characters (now treated as ERROR, not warning)
    if (trimmedValue.includes('X')) {
      result.isValid = false; // Changed from hasWarnings = true to isValid = false
      const xCount = (trimmedValue.match(/X/g) || []).length;
      result.messages.push(
        `Contains ${xCount} 'X' character${
          xCount > 1 ? 's' : ''
        } (incomplete data)`
      );
    }

    // Check for duplicates on the page.
    const duplicateInfoOnPage = this.findBarcodeDuplicatesOnPage(
      trimmedValue,
      excludeRowIndex
    );
    if (duplicateInfoOnPage.isDuplicate) {
      result.isValid = false;
      result.messages.push(
        `Duplicate barcode found in ${duplicateInfoOnPage.count} other row${
          duplicateInfoOnPage.count > 1 ? 's' : ''
        } on this page.`
      );
    }

    return result;
  }

  /**
   * Finds duplicate barcodes in the dataset
   * @param {string} barcode - The barcode value to check for duplicates
   * @param {number} excludeRowIndex - Optional row index to exclude from duplicate checks
   * @return {Object} Object with isDuplicate flag and count of duplicates
   */
  findBarcodeDuplicatesOnPage(barcode, excludeRowIndex = -1) {
    if (!this.app.data.filtered || this.app.data.filtered.length <= 1) {
      return { isDuplicate: false, count: 0 };
    }

    const barcodeColumnIndex = this.app.data.filtered[0].indexOf('new_barcode');
    if (barcodeColumnIndex === -1) {
      return { isDuplicate: false, count: 0 };
    }

    let duplicateCount = 0;
    for (let i = 1; i < this.app.data.filtered.length; i++) {
      if (i === excludeRowIndex) continue; // Skip the row we're currently editing

      const rowBarcode = String(
        this.app.data.filtered[i][barcodeColumnIndex] || ''
      ).trim();
      if (rowBarcode === barcode && rowBarcode !== '') {
        duplicateCount++;
      }
    }

    return {
      isDuplicate: duplicateCount > 0,
      count: duplicateCount,
    };
  }

  /**
   * Validates all barcode fields in the current dataset
   * @returns {Object} Summary object with counts of valid, warning, error, and empty barcodes
   */
  validateAllBarcodes() {
    const summary = {
      totalBarcodes: 0,
      validBarcodes: 0,
      errorBarcodes: 0,
      emptyBarcodes: 0,
      issues: [],
    };

    if (!this.app.data.filtered || this.app.data.filtered.length <= 1) {
      return summary;
    }

    const barcodeColumnIndex = this.app.data.filtered[0].indexOf('new_barcode');
    if (barcodeColumnIndex === -1) {
      return summary;
    }

    for (let i = 1; i < this.app.data.filtered.length; i++) {
      const barcode = String(
        this.app.data.filtered[i][barcodeColumnIndex] || ''
      ).trim();

      if (!barcode) {
        summary.emptyBarcodes++;
        continue;
      }

      summary.totalBarcodes++;
      const validation = this.validateBarcode(barcode, i);

      if (validation.isValid) {
        summary.validBarcodes++;
      } else {
        summary.errorBarcodes++;
        summary.issues.push({
          row: i,
          barcode: barcode,
          type: 'error',
          messages: validation.messages,
        });
      }
    }

    return summary;
  }

  /**
   * Gets validation statistics for display
   * @return {Object} Object containing total fields, completion rate, error rate, and warning rate
   */
  getValidationStats() {
    const summary = this.validateAllBarcodes();
    const totalFields = summary.totalBarcodes + summary.emptyBarcodes;

    return {
      ...summary,
      totalFields,
      completionRate:
        totalFields > 0
          ? ((summary.totalBarcodes / totalFields) * 100).toFixed(1)
          : 0,
      errorRate:
        summary.totalBarcodes > 0
          ? ((summary.errorBarcodes / summary.totalBarcodes) * 100).toFixed(1)
          : 0,
    };
  }

  /**
   * Shows a validation summary modal
   */
  showValidationSummary() {
    const summary = this.validateAllBarcodes();
    const stats = this.getValidationStats();

    const modalHtml = `
      <div class="modal is-active" id="validation-summary-modal">
        <div class="modal-background"></div>
        <div class="modal-content">
          <div class="box">
            <h3 class="title is-4">
              <span class="icon"><i class="fas fa-check-circle"></i></span>
              Barcode Validation Summary
            </h3>

            <div class="columns">
              <div class="column is-half">
                <div class="notification is-light">
                  <h4 class="title is-6">Field Status</h4>
                  <strong>Total Fields:</strong> ${stats.totalFields}<br>
                  <strong>Completed:</strong> ${summary.totalBarcodes} (${
      stats.completionRate
    }%)<br>
                  <strong>Empty:</strong> ${summary.emptyBarcodes}<br>
                </div>
              </div>
              <div class="column is-half">
                <div class="notification is-light">
                  <h4 class="title is-6">Validation Results</h4>
                  <strong class="has-text-success">✓ Valid:</strong> ${
                    summary.validBarcodes
                  }<br>
                  <strong class="has-text-danger">✗ Errors:</strong> ${
                    summary.errorBarcodes
                  } (${stats.errorRate}%)<br>
                </div>
              </div>
            </div>

            ${this.renderValidationIssues(summary.issues)}

            <div class="field is-grouped is-grouped-right">
              <div class="control">
                <button class="button" id="close-validation-summary">Close</button>
              </div>
              ${
                summary.errorBarcodes > 0
                  ? `<div class="control">
                      <button class="button is-danger" id="highlight-errors">Highlight Errors</button>
                    </div>`
                  : ''
              }
            </div>
          </div>
        </div>
        <button class="modal-close is-large" aria-label="close"></button>
      </div>`;

    const modal = jQuery(modalHtml);
    jQuery('body').append(modal);

    // Bind close events
    modal
      .find('.modal-close, #close-validation-summary, .modal-background')
      .on('click', () => modal.remove());

    // Bind highlight errors button
    modal.find('#highlight-errors').on('click', () => {
      this.highlightValidationErrors(summary.issues);
      modal.remove();
    });
  }

  /**
   * Renders validation issues for the summary modal
   * @param {Array} issues - Array of validation issue objects
   * @return {string} HTML string for the issues section
   */
  renderValidationIssues(issues) {
    if (issues.length === 0) {
      return `
        <div class="notification is-success is-light">
          <span class="icon"><i class="fas fa-check"></i></span>
          <strong>All barcodes are valid!</strong> No issues found.
        </div>`;
    }

    const errorIssues = issues.filter((issue) => issue.type === 'error');

    let html = '<div class="field"><label class="label">Issues Found:</label>';

    if (errorIssues.length > 0) {
      html += '<h5 class="title is-6 has-text-danger">Errors (Must Fix):</h5>';
      html +=
        '<div style="max-height: 200px; overflow-y: auto; margin-bottom: 1rem;">';
      errorIssues.forEach((issue) => {
        html += `
          <div class="notification is-danger is-light mb-2">
            <strong>Row ${issue.row}:</strong> ${this.app.utilities.escapeHtml(
          issue.barcode
        )}<br>
            <small>${issue.messages.join(', ')}</small>
          </div>`;
      });
      html += '</div>';
    }

    html += '</div>';
    return html;
  }

  /**
   * Highlights validation errors in the table
   * @param {Array} issues - Array of validation issue objects
   */
  highlightValidationErrors(issues) {
    if (!this.app.data.filtered || this.app.data.filtered.length <= 1) {
      return;
    }

    const barcodeColumnIndex = this.app.data.filtered[0].indexOf('new_barcode');
    if (barcodeColumnIndex === -1) {
      return;
    }

    // First scroll to the first error
    const firstError = issues.find((issue) => issue.type === 'error');
    if (firstError) {
      const $errorCell = this.app.elements.tableContainer.find(
        `input[data-row="${firstError.row}"][data-col="${barcodeColumnIndex}"]`
      );
      if ($errorCell.length) {
        // Scroll to the cell
        this.app.elements.tableContainer.animate(
          {
            scrollTop:
              $errorCell.offset().top -
              this.app.elements.tableContainer.offset().top +
              this.app.elements.tableContainer.scrollTop() -
              100,
          },
          500
        );

        // Focus the cell
        setTimeout(() => {
          $errorCell.focus().select();
        }, 600);
      }
    }

    this.app.utilities.showMessage(
      `Highlighted ${
        issues.filter((i) => i.type === 'error').length
      } validation errors. Check the barcode fields with red borders.`,
      'warning',
      5000
    );
  }

  /**
   * Gets a list of row indices that have validation errors
   * @return {Array} Array of row indices with validation errors
   */
  getRowsWithErrors() {
    const errorRows = [];

    if (!this.app.data.filtered || this.app.data.filtered.length <= 1) {
      return errorRows;
    }

    const barcodeColumnIndex = this.app.data.filtered[0].indexOf('new_barcode');
    if (barcodeColumnIndex === -1) {
      return errorRows;
    }

    for (let i = 1; i < this.app.data.filtered.length; i++) {
      const barcode = String(
        this.app.data.filtered[i][barcodeColumnIndex] || ''
      ).trim();

      if (barcode) {
        // Only check non-empty barcodes
        const validation = this.validateBarcode(barcode, i);
        if (!validation.isValid) {
          errorRows.push(i);
        }
      }
    }

    return errorRows;
  }

  /**
   * Gets validation statistics including row counts
   * @return {Object} Object containing error row count, total rows, and error status
   */
  getValidationRowStats() {
    const errorRows = this.getRowsWithErrors();
    const totalRows = this.app.data.filtered.length - 1; // Exclude header

    return {
      errorRows,
      errorCount: errorRows.length,
      totalRows,
      hasErrors: errorRows.length > 0,
    };
  }

  /**
   * Checks if the dataset has any validation issues
   * @return {boolean} True if there are validation errors, false otherwise
   */
  hasValidationErrors() {
    const summary = this.validateAllBarcodes();
    return summary.errorBarcodes > 0;
  }

  /**
   * Gets a quick validation status for display
   * @return {Object} Object containing status, message, and CSS class for display
   */
  getQuickValidationStatus() {
    const summary = this.validateAllBarcodes();
    const totalWithData = summary.totalBarcodes;

    if (totalWithData === 0) {
      return {
        status: 'empty',
        message: 'No barcodes to validate',
        class: 'is-light',
      };
    }

    if (summary.errorBarcodes > 0) {
      return {
        status: 'error',
        message: `${summary.errorBarcodes} error${
          summary.errorBarcodes > 1 ? 's' : ''
        }`,
        class: 'is-danger',
      };
    }

    return {
      status: 'valid',
      message: `All ${summary.validBarcodes} barcodes valid`,
      class: 'is-success',
    };
  }

  /**
   * Validates barcodes before export
   * @return {boolean} True if validation passes, false if user cancels export
   */
  validateBeforeExport() {
    const summary = this.validateAllBarcodes();

    if (summary.errorBarcodes === 0) {
      return true; // No errors, proceed
    }

    const proceed = confirm(
      `Warning: ${summary.errorBarcodes} barcode${
        summary.errorBarcodes > 1 ? 's have' : ' has'
      } validation errors. ` +
        `Export anyway? Errors include duplicate barcodes, invalid lengths, and incomplete data ('X' characters).`
    );

    return proceed;
  }
}
