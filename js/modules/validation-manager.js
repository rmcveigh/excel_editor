/**
 * @file
 * Excel Editor Validation Manager Module
 *
 * Handles barcode validation, field validation, and validation feedback.
 * This module adds validation methods to the ExcelEditor class.
 */

/* eslint-disable no-console */
(function ($) {
  'use strict';

  /**
   * Validation Manager module for Excel Editor.
   * This function is called on ExcelEditor instances to add validation methods.
   */
  window.ExcelEditorValidationManager = function () {
    // =========================================================================
    // BARCODE FIELD VALIDATION
    // =========================================================================

    /**
     * Validates a barcode field and applies visual feedback
     * @param {jQuery} $cell - The input cell element
     * @param {string} value - The barcode value to validate
     * @param {number} rowIndex - The row index for context
     */
    this.validateBarcodeField = function ($cell, value, rowIndex) {
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
      } else if (validation.hasWarnings) {
        // Has warnings but is valid - yellow border
        $cell.addClass('is-warning');
      } else {
        // Completely valid - green border
        $cell.addClass('is-valid');
      }

      // Add validation message if there are issues
      if (validation.messages.length > 0) {
        const messageClass = validation.isValid && !validation.hasWarnings ? 'has-text-success' :
                            validation.hasWarnings && validation.isValid ? 'has-text-warning' : 'has-text-danger';
        const messageHtml = `<div class="validation-message ${messageClass} is-size-7 mt-1">
          ${validation.messages.map(msg => `<div>• ${this.escapeHtml(msg)}</div>`).join('')}
        </div>`;
        $cell.after(messageHtml);
      }
    };

    /**
     * Validates all existing barcode fields in the rendered table
     */
    this.validateExistingBarcodeFields = function () {
      if (!this.data.filtered || this.data.filtered.length <= 1) {
        return;
      }

      const barcodeColumnIndex = this.data.filtered[0].indexOf('new_barcode');
      if (barcodeColumnIndex === -1) {
        return;
      }

      // Find all barcode input fields and validate them
      this.elements.tableContainer.find('input[data-col="' + barcodeColumnIndex + '"]').each((index, input) => {
        const $input = $(input);
        const rowIndex = parseInt($input.data('row'));
        const value = $input.val() || '';

        // Validate all fields, including empty ones
        this.validateBarcodeField($input, value.trim(), rowIndex);
      });

      this.logDebug('Validated all existing barcode fields on load');
    };

    /**
     * Validates all barcode fields and updates their styling
     * This is called when data changes or filters are applied
     */
    this.revalidateAllVisibleFields = function () {
      setTimeout(() => {
        this.validateExistingBarcodeFields();
      }, 100); // Small delay to ensure DOM is updated
    };

    // =========================================================================
    // BARCODE VALIDATION LOGIC
    // =========================================================================

    /**
     * Validates a barcode value according to business rules
     * @param {string} value - The barcode value to validate
     * @param {number} excludeRowIndex - Row index to exclude from duplicate checking
     * @returns {object} Validation result with isValid, hasWarnings, and messages
     */
    this.validateBarcode = function (value, excludeRowIndex = -1) {
      const result = {
        isValid: true,
        hasWarnings: false,
        messages: []
      };

      if (!value || value.trim() === '') {
        return result; // Empty values are not validated
      }

      const trimmedValue = value.trim();

      // Check length (16-17 characters)
      if (trimmedValue.length < 16 || trimmedValue.length > 17) {
        result.isValid = false;
        result.messages.push(`Length should be 16-17 characters (current: ${trimmedValue.length})`);
      }

      // Check for 'X' characters (warning, not error)
      if (trimmedValue.includes('X')) {
        result.hasWarnings = true;
        const xCount = (trimmedValue.match(/X/g) || []).length;
        result.messages.push(`Contains ${xCount} 'X' character${xCount > 1 ? 's' : ''} (may indicate incomplete data)`);
      }

      // Check for duplicates
      const duplicateInfo = this.findBarcodeDuplicates(trimmedValue, excludeRowIndex);
      if (duplicateInfo.isDuplicate) {
        result.isValid = false;
        result.messages.push(`Duplicate barcode found in ${duplicateInfo.count} other row${duplicateInfo.count > 1 ? 's' : ''}`);
      }

      // If no errors but has warnings, it's still "valid" but with warnings
      if (!result.isValid) {
        result.hasWarnings = false; // Errors take precedence over warnings
      }

      return result;
    };

    /**
     * Finds duplicate barcodes in the dataset
     * @param {string} barcode - The barcode to search for
     * @param {number} excludeRowIndex - Row index to exclude from search
     * @returns {object} Object with isDuplicate boolean and count number
     */
    this.findBarcodeDuplicates = function (barcode, excludeRowIndex = -1) {
      if (!this.data.filtered || this.data.filtered.length <= 1) {
        return { isDuplicate: false, count: 0 };
      }

      const barcodeColumnIndex = this.data.filtered[0].indexOf('new_barcode');
      if (barcodeColumnIndex === -1) {
        return { isDuplicate: false, count: 0 };
      }

      let duplicateCount = 0;
      for (let i = 1; i < this.data.filtered.length; i++) {
        if (i === excludeRowIndex) continue; // Skip the row we're currently editing

        const rowBarcode = String(this.data.filtered[i][barcodeColumnIndex] || '').trim();
        if (rowBarcode === barcode && rowBarcode !== '') {
          duplicateCount++;
        }
      }

      return {
        isDuplicate: duplicateCount > 0,
        count: duplicateCount
      };
    };

    // =========================================================================
    // VALIDATION SUMMARY & ANALYSIS
    // =========================================================================

    /**
     * Validates all barcode fields in the current dataset
     * @returns {object} Summary of validation results
     */
    this.validateAllBarcodes = function () {
      const summary = {
        totalBarcodes: 0,
        validBarcodes: 0,
        warningBarcodes: 0,
        errorBarcodes: 0,
        emptyBarcodes: 0,
        issues: []
      };

      if (!this.data.filtered || this.data.filtered.length <= 1) {
        return summary;
      }

      const barcodeColumnIndex = this.data.filtered[0].indexOf('new_barcode');
      if (barcodeColumnIndex === -1) {
        return summary;
      }

      for (let i = 1; i < this.data.filtered.length; i++) {
        const barcode = String(this.data.filtered[i][barcodeColumnIndex] || '').trim();

        if (!barcode) {
          summary.emptyBarcodes++;
          continue;
        }

        summary.totalBarcodes++;
        const validation = this.validateBarcode(barcode, i);

        if (validation.isValid && !validation.hasWarnings) {
          summary.validBarcodes++;
        } else if (validation.hasWarnings && validation.isValid) {
          summary.warningBarcodes++;
          summary.issues.push({
            row: i,
            barcode: barcode,
            type: 'warning',
            messages: validation.messages
          });
        } else {
          summary.errorBarcodes++;
          summary.issues.push({
            row: i,
            barcode: barcode,
            type: 'error',
            messages: validation.messages
          });
        }
      }

      return summary;
    };

    /**
     * Gets validation statistics for display
     * @returns {object} Formatted validation statistics
     */
    this.getValidationStats = function () {
      const summary = this.validateAllBarcodes();
      const totalFields = summary.totalBarcodes + summary.emptyBarcodes;

      return {
        ...summary,
        totalFields,
        completionRate: totalFields > 0 ? ((summary.totalBarcodes / totalFields) * 100).toFixed(1) : 0,
        errorRate: summary.totalBarcodes > 0 ? ((summary.errorBarcodes / summary.totalBarcodes) * 100).toFixed(1) : 0,
        warningRate: summary.totalBarcodes > 0 ? ((summary.warningBarcodes / summary.totalBarcodes) * 100).toFixed(1) : 0
      };
    };

    // =========================================================================
    // VALIDATION UI & MODALS
    // =========================================================================

    /**
     * Shows a validation summary modal
     */
    this.showValidationSummary = function () {
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
                    <strong>Completed:</strong> ${summary.totalBarcodes} (${stats.completionRate}%)<br>
                    <strong>Empty:</strong> ${summary.emptyBarcodes}<br>
                  </div>
                </div>
                <div class="column is-half">
                  <div class="notification is-light">
                    <h4 class="title is-6">Validation Results</h4>
                    <strong class="has-text-success">✓ Valid:</strong> ${summary.validBarcodes}<br>
                    <strong class="has-text-warning">⚠ Warnings:</strong> ${summary.warningBarcodes} (${stats.warningRate}%)<br>
                    <strong class="has-text-danger">✗ Errors:</strong> ${summary.errorBarcodes} (${stats.errorRate}%)<br>
                  </div>
                </div>
              </div>

              ${this.renderValidationIssues(summary.issues)}

              <div class="field is-grouped is-grouped-right">
                <div class="control">
                  <button class="button" id="close-validation-summary">Close</button>
                </div>
                ${summary.errorBarcodes > 0 ? `
                <div class="control">
                  <button class="button is-warning" id="highlight-errors">Highlight Errors</button>
                </div>
                ` : ''}
              </div>
            </div>
          </div>
          <button class="modal-close is-large" aria-label="close"></button>
        </div>`;

      const modal = $(modalHtml);
      $('body').append(modal);

      // Bind close events
      modal.find('.modal-close, #close-validation-summary, .modal-background')
        .on('click', () => modal.remove());

      // Bind highlight errors button
      modal.find('#highlight-errors').on('click', () => {
        this.highlightValidationErrors(summary.issues);
        modal.remove();
      });
    };

    /**
     * Renders validation issues for the summary modal
     * @param {Array} issues - Array of validation issues
     * @returns {string} HTML string for issues display
     */
    this.renderValidationIssues = function (issues) {
      if (issues.length === 0) {
        return `
          <div class="notification is-success is-light">
            <span class="icon"><i class="fas fa-check"></i></span>
            <strong>All barcodes are valid!</strong> No issues found.
          </div>`;
      }

      const errorIssues = issues.filter(issue => issue.type === 'error');
      const warningIssues = issues.filter(issue => issue.type === 'warning');

      let html = '<div class="field"><label class="label">Issues Found:</label>';

      if (errorIssues.length > 0) {
        html += '<h5 class="title is-6 has-text-danger">Errors (Must Fix):</h5>';
        html += '<div style="max-height: 200px; overflow-y: auto; margin-bottom: 1rem;">';
        errorIssues.forEach(issue => {
          html += `
            <div class="notification is-danger is-light mb-2">
              <strong>Row ${issue.row}:</strong> ${this.escapeHtml(issue.barcode)}<br>
              <small>${issue.messages.join(', ')}</small>
            </div>`;
        });
        html += '</div>';
      }

      if (warningIssues.length > 0) {
        html += '<h5 class="title is-6 has-text-warning">Warnings (Review Recommended):</h5>';
        html += '<div style="max-height: 200px; overflow-y: auto;">';
        warningIssues.forEach(issue => {
          html += `
            <div class="notification is-warning is-light mb-2">
              <strong>Row ${issue.row}:</strong> ${this.escapeHtml(issue.barcode)}<br>
              <small>${issue.messages.join(', ')}</small>
            </div>`;
        });
        html += '</div>';
      }

      html += '</div>';
      return html;
    };

    /**
     * Highlights validation errors in the table
     * @param {Array} issues - Array of validation issues
     */
    this.highlightValidationErrors = function (issues) {
      if (!this.data.filtered || this.data.filtered.length <= 1) {
        return;
      }

      const barcodeColumnIndex = this.data.filtered[0].indexOf('new_barcode');
      if (barcodeColumnIndex === -1) {
        return;
      }

      // First scroll to the first error
      const firstError = issues.find(issue => issue.type === 'error');
      if (firstError) {
        const $errorCell = this.elements.tableContainer.find(`input[data-row="${firstError.row}"][data-col="${barcodeColumnIndex}"]`);
        if ($errorCell.length) {
          // Scroll to the cell
          this.elements.tableContainer.animate({
            scrollTop: $errorCell.offset().top - this.elements.tableContainer.offset().top + this.elements.tableContainer.scrollTop() - 100
          }, 500);

          // Focus the cell
          setTimeout(() => {
            $errorCell.focus().select();
          }, 600);
        }
      }

      this.showMessage(`Highlighted ${issues.filter(i => i.type === 'error').length} validation errors. Check the barcode fields with red borders.`, 'warning', 5000);
    };

    // =========================================================================
    // VALIDATION HELPERS
    // =========================================================================

    /**
     * Checks if the dataset has any validation issues
     * @returns {boolean} True if there are validation errors (not warnings)
     */
    this.hasValidationErrors = function () {
      const summary = this.validateAllBarcodes();
      return summary.errorBarcodes > 0;
    };

    /**
     * Gets a quick validation status for display
     * @returns {object} Quick status object
     */
    this.getQuickValidationStatus = function () {
      const summary = this.validateAllBarcodes();
      const totalWithData = summary.totalBarcodes;

      if (totalWithData === 0) {
        return {
          status: 'empty',
          message: 'No barcodes to validate',
          class: 'is-light'
        };
      }

      if (summary.errorBarcodes > 0) {
        return {
          status: 'error',
          message: `${summary.errorBarcodes} error${summary.errorBarcodes > 1 ? 's' : ''}`,
          class: 'is-danger'
        };
      }

      if (summary.warningBarcodes > 0) {
        return {
          status: 'warning',
          message: `${summary.warningBarcodes} warning${summary.warningBarcodes > 1 ? 's' : ''}`,
          class: 'is-warning'
        };
      }

      return {
        status: 'valid',
        message: `All ${summary.validBarcodes} barcodes valid`,
        class: 'is-success'
      };
    };

    /**
     * Validates barcodes before export
     * @returns {boolean} True if validation passes or user confirms export
     */
    this.validateBeforeExport = function () {
      const summary = this.validateAllBarcodes();

      if (summary.errorBarcodes === 0) {
        return true; // No errors, proceed
      }

      const proceed = confirm(
        `Warning: ${summary.errorBarcodes} barcode${summary.errorBarcodes > 1 ? 's have' : ' has'} validation errors. ` +
        `Export anyway? Errors include duplicate barcodes and invalid lengths.`
      );

      return proceed;
    };

    this.logDebug('ExcelEditorValidationManager module loaded');
  };
})(jQuery);
