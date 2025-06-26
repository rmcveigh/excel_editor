/**
 * @file
 * Excel Editor Barcode System Module
 *
 * Handles generic barcode formatting, column detection, and barcode management.
 * This module adds barcode-related methods to the ExcelEditor class.
 */

/* eslint-disable no-console */
(function ($) {
  'use strict';

  /**
   * Barcode System module for Excel Editor.
   * This function is called on ExcelEditor instances to add barcode methods.
   */
  window.ExcelEditorBarcodeSystem = function () {
    // =========================================================================
    // GENERIC BARCODE FORMATTING SYSTEM
    // =========================================================================

    /**
     * Generic barcode formatting function that can work with any source value
     * @param {string} sourceValue - The source value to format (Subject ID, Product Code, etc.)
     * @param {object} options - Formatting options
     * @param {string} contextValue - Optional context value for suffixes (health status, category, etc.)
     * @param {string} contextType - Type of context ('health', 'category', 'status', etc.)
     * @returns {string} Formatted barcode
     */
    this.formatBarcode = function (
      sourceValue,
      options = {},
      contextValue = null,
      contextType = 'health'
    ) {
      if (!sourceValue) return '';

      const defaults = {
        removeDashes: true,
        removeSpaces: true,
        removeUnderscores: true,
        removeDots: true,
        removeNonAlphanumeric: false,
        toUpperCase: true,
        maxLength: null,
        prefix: '',
        suffix: '',
        includeContext: true,
      };

      const settings = { ...defaults, ...options };
      let formatted = String(sourceValue).trim();

      // Apply cleaning rules
      if (settings.removeDashes) formatted = formatted.replace(/-/g, '');
      if (settings.removeSpaces) formatted = formatted.replace(/\s/g, '');
      if (settings.removeUnderscores) formatted = formatted.replace(/_/g, '');
      if (settings.removeDots) formatted = formatted.replace(/\./g, '');
      if (settings.removeNonAlphanumeric)
        formatted = formatted.replace(/[^a-zA-Z0-9]/g, '');
      if (settings.toUpperCase) formatted = formatted.toUpperCase();

      // Add prefix
      formatted = settings.prefix + formatted;

      // Add context-based suffix if enabled and context value provided
      let contextSuffix = '';
      if (settings.includeContext && contextValue !== null) {
        contextSuffix = this.getContextSuffix(contextValue, contextType);
      }

      // Add context suffix and regular suffix
      formatted = formatted + contextSuffix + settings.suffix;

      // Apply max length after all additions
      if (settings.maxLength)
        formatted = formatted.substring(0, settings.maxLength);

      this.logDebug(
        `Formatted barcode: "${sourceValue}" + ${contextType}:"${contextValue}" → "${formatted}"`,
        settings
      );

      return formatted;
    };

    /**
     * Generic context suffix determination based on value and type
     * @param {string} contextValue - The context value (health status, category, etc.)
     * @param {string} contextType - Type of context ('health', 'category', 'status', etc.)
     * @returns {string} Appropriate suffix or empty string
     */
    this.getContextSuffix = function (contextValue, contextType = 'health') {
      if (!contextValue) return '';

      const value = String(contextValue).trim().toLowerCase();

      switch (contextType) {
        case 'health':
          return this.getHealthStatusSuffix(value);
        case 'category':
          return this.getCategorySuffix(value);
        case 'priority':
          return this.getPrioritySuffix(value);
        case 'status':
          return this.getStatusSuffix(value);
        default:
          this.logDebug(`Unknown context type: ${contextType}`);
          return '';
      }
    };

    /**
     * Health status suffix (H for healthy, D for diseased)
     * @param {string} value - The health status value
     * @returns {string} 'H', 'D', or ''
     */
    this.getHealthStatusSuffix = function (value) {
      if (!value) return '';

      const normalizedValue = String(value).trim().toLowerCase();

      // Check for Healthy variants
      const healthyValues = ['h', 'healthy', 'normal', 'good'];
      if (healthyValues.includes(normalizedValue)) {
        return 'H';
      }

      // Check for Diseased variants
      const diseasedValues = [
        'd',
        'diseased',
        'disease',
        'abnormal',
        'pathological',
      ];
      if (diseasedValues.includes(normalizedValue)) {
        return 'D';
      }

      return '';
    };

    /**
     * Category suffix for future extensibility
     * @param {string} value - The category value
     * @returns {string} Category suffix or ''
     */
    this.getCategorySuffix = function (value) {
      if (!value) return '';

      const normalizedValue = String(value).trim().toLowerCase();

      const categoryMap = {
        primary: 'P',
        secondary: 'S',
        control: 'C',
        test: 'T',
        sample: 'S',
        reference: 'R',
      };

      return categoryMap[normalizedValue] || '';
    };

    /**
     * Priority suffix for future extensibility
     * @param {string} value - The priority value
     * @returns {string} Priority suffix or ''
     */
    this.getPrioritySuffix = function (value) {
      if (!value) return '';

      const normalizedValue = String(value).trim().toLowerCase();

      const priorityMap = {
        high: 'H',
        medium: 'M',
        low: 'L',
        urgent: 'U',
        critical: 'C',
      };

      return priorityMap[normalizedValue] || '';
    };

    /**
     * Status suffix for future extensibility
     * @param {string} value - The status value
     * @returns {string} Status suffix or ''
     */
    this.getStatusSuffix = function (value) {
      if (!value) return '';

      const normalizedValue = String(value).trim().toLowerCase();

      const statusMap = {
        active: 'A',
        inactive: 'I',
        pending: 'P',
        complete: 'C',
        failed: 'F',
        approved: 'A',
        rejected: 'R',
      };

      return statusMap[normalizedValue] || '';
    };

    // =========================================================================
    // GENERIC COLUMN DETECTION SYSTEM
    // =========================================================================

    /**
     * Generic column finder that can locate various types of columns
     * @param {Array} headerRow - The header row array
     * @param {string} columnType - Type of column to find ('subject_id', 'health_status', etc.)
     * @returns {number} Column index or -1 if not found
     */
    this.findColumnByType = function (headerRow, columnType) {
      switch (columnType) {
        case 'subject_id':
          return this.findSubjectIdColumn(headerRow);
        case 'health_status':
          return this.findHealthStatusColumn(headerRow);
        case 'barcode':
          return this.findBarcodeColumn(headerRow);
        case 'category':
          return this.findCategoryColumn(headerRow);
        case 'priority':
          return this.findPriorityColumn(headerRow);
        default:
          this.logDebug(`Unknown column type: ${columnType}`);
          return -1;
      }
    };

    /**
     * Enhanced Subject ID column finder
     * @param {Array} headerRow - The header row array
     * @returns {number} Column index or -1 if not found
     */
    this.findSubjectIdColumn = function (headerRow) {
      // Try exact matches first
      const exactMatches = [
        'Subject ID',
        'SubjectID',
        'subject_id',
        'SUBJECT_ID',
        'Subject_ID',
        'Patient ID',
        'PatientID',
        'patient_id',
        'PATIENT_ID',
        'Sample ID',
        'SampleID',
        'sample_id',
        'SAMPLE_ID',
      ];

      for (const exactMatch of exactMatches) {
        const index = headerRow.indexOf(exactMatch);
        if (index !== -1) {
          this.logDebug(
            `Found ID column (exact): "${exactMatch}" at index ${index}`
          );
          return index;
        }
      }

      // Try partial matches
      const flexibleMatches = [
        'subject id',
        'subjectid',
        'subject-id',
        'subject_id',
        'patient id',
        'patientid',
        'patient-id',
        'patient_id',
        'sample id',
        'sampleid',
        'sample-id',
        'sample_id',
        'specimen id',
        'specimenid',
        'specimen-id',
        'id',
        'identifier',
      ];

      for (let i = 0; i < headerRow.length; i++) {
        const header = String(headerRow[i]).trim().toLowerCase();

        for (const pattern of flexibleMatches) {
          if (header.includes(pattern)) {
            this.logDebug(
              `Found ID column (flexible): "${headerRow[i]}" at index ${i}`
            );
            return i;
          }
        }
      }

      this.logDebug('No ID column found');
      return -1;
    };

    /**
     * Enhanced Health Status column finder
     */
    this.findHealthStatusColumn = function (headerRow) {
      // Try exact matches first
      const exactMatches = [
        'Tissue Diseased or Healthy',
        'Tissue_Diseased_or_Healthy',
        'tissue diseased or healthy',
        'TISSUE DISEASED OR HEALTHY',
        'Health Status',
        'Disease Status',
      ];

      for (const exactMatch of exactMatches) {
        const index = headerRow.indexOf(exactMatch);
        if (index !== -1) {
          this.logDebug(
            `Found Health Status column (exact): "${exactMatch}" at index ${index}`
          );
          return index;
        }
      }

      // Try partial matches
      const flexibleMatches = [
        'tissue diseased',
        'diseased or healthy',
        'health status',
        'tissue health',
        'disease status',
        'healthy diseased',
      ];

      for (let i = 0; i < headerRow.length; i++) {
        const header = String(headerRow[i]).trim().toLowerCase();

        for (const pattern of flexibleMatches) {
          if (header.includes(pattern)) {
            this.logDebug(
              `Found Health Status column (flexible): "${headerRow[i]}" at index ${i}`
            );
            return i;
          }
        }
      }

      this.logDebug('No Health Status column found');
      return -1;
    };

    /**
     * Barcode column finder
     * @param {Array} headerRow - The header row array
     * @returns {number} Column index or -1 if not found
     */
    this.findBarcodeColumn = function (headerRow) {
      const exactMatches = [
        'new_barcode',
        'barcode',
        'Barcode',
        'BARCODE',
        'Bar Code',
        'bar_code',
        'BAR_CODE',
      ];

      for (const exactMatch of exactMatches) {
        const index = headerRow.indexOf(exactMatch);
        if (index !== -1) {
          this.logDebug(
            `Found Barcode column: "${exactMatch}" at index ${index}`
          );
          return index;
        }
      }

      return headerRow.indexOf('new_barcode'); // Default to our added column
    };

    /**
     * Category column finder for future extensibility
     * @param {Array} headerRow - The header row array
     * @returns {number} Column index or -1 if not found
     */
    this.findCategoryColumn = function (headerRow) {
      const patterns = [
        'category',
        'type',
        'classification',
        'group',
        'sample type',
        'specimen type',
        'tissue type',
      ];

      for (let i = 0; i < headerRow.length; i++) {
        const header = String(headerRow[i]).trim().toLowerCase();

        for (const pattern of patterns) {
          if (header.includes(pattern)) {
            this.logDebug(
              `Found Category column: "${headerRow[i]}" at index ${i}`
            );
            return i;
          }
        }
      }

      return -1;
    };

    /**
     * Priority column finder for future extensibility
     * @param {Array} headerRow - The header row array
     * @returns {number} Column index or -1 if not found
     */
    this.findPriorityColumn = function (headerRow) {
      const patterns = ['priority', 'urgency', 'importance'];

      for (let i = 0; i < headerRow.length; i++) {
        const header = String(headerRow[i]).trim().toLowerCase();

        for (const pattern of patterns) {
          if (header.includes(pattern)) {
            this.logDebug(
              `Found Priority column: "${headerRow[i]}" at index ${i}`
            );
            return i;
          }
        }
      }

      return -1;
    };

    // =========================================================================
    // BARCODE MANAGEMENT (RESET BARCODES FUNCTIONALITY)
    // =========================================================================

    /**
     * Resets/populates barcode values using generic barcode formatter
     */
    this.resetBarcodes = function () {
      if (!this.data.original.length) {
        this.showMessage('No data loaded', 'warning');
        return;
      }

      const headerRow = this.data.original[0];
      const barcodeIndex = this.findColumnByType(headerRow, 'barcode');
      const sourceColumnIndex = this.findColumnByType(headerRow, 'subject_id');

      if (barcodeIndex === -1) {
        this.showMessage('No barcode column found', 'warning');
        return;
      }

      if (sourceColumnIndex === -1) {
        this.showMessage('No source ID column found in this file', 'warning');
        return;
      }

      // Show configuration modal before proceeding
      this.showBarcodeResetModal(barcodeIndex, sourceColumnIndex);
    };

    /**
     * Shows a modal to configure barcode formatting options with context info
     */
    this.showBarcodeResetModal = function (barcodeIndex, sourceColumnIndex) {
      const headerRow = this.data.original[0];
      const sourceColumnName = headerRow[sourceColumnIndex];
      const contextColumnIndex = this.findColumnByType(
        headerRow,
        'health_status'
      );
      const contextColumnName =
        contextColumnIndex !== -1 ? headerRow[contextColumnIndex] : null;

      // Check if there are existing barcode values
      const hasExistingValues = this.data.original
        .slice(1)
        .some((row) => row[barcodeIndex]);
      const existingWarning = hasExistingValues
        ? '<div class="notification is-warning is-light mb-3"><strong>Warning:</strong> Some barcode values already exist and will be overwritten.</div>'
        : '';

      // Context info
      const contextInfo = contextColumnName
        ? `<div class="notification is-info is-light mb-3"><strong>Context Detection:</strong> Found "${contextColumnName}" column. Barcodes will automatically include context suffixes.</div>`
        : '<div class="notification is-light mb-3"><strong>Context Status:</strong> No context column found. Barcodes will not include context suffixes.</div>';

      // Get a sample value for preview
      const sampleRow = this.data.original
        .slice(1)
        .find((row) => row[sourceColumnIndex]);
      const sampleSourceValue = sampleRow?.[sourceColumnIndex] || 'ABC-123-XYZ';
      const sampleContextValue =
        contextColumnIndex !== -1 && sampleRow
          ? sampleRow[contextColumnIndex]
          : 'Healthy';

      const modalHtml = `
        <div class="modal is-active" id="reset-barcodes-modal">
          <div class="modal-background"></div>
          <div class="modal-content">
            <div class="box">
              <h3 class="title is-4">
                <span class="icon"><i class="fas fa-barcode"></i></span>
                Reset Barcodes from ${this.escapeHtml(sourceColumnName)}
              </h3>

              ${existingWarning}
              ${contextInfo}

              <div class="field">
                <label class="label">Formatting Options</label>
              </div>

              <div class="columns">
                <div class="column">
                  <div class="field">
                    <label class="checkbox">
                      <input type="checkbox" id="remove-dashes" checked>
                      Remove dashes (-)
                    </label>
                  </div>
                  <div class="field">
                    <label class="checkbox">
                      <input type="checkbox" id="remove-spaces" checked>
                      Remove spaces
                    </label>
                  </div>
                  <div class="field">
                    <label class="checkbox">
                      <input type="checkbox" id="remove-underscores" checked>
                      Remove underscores (_)
                    </label>
                  </div>
                  <div class="field">
                    <label class="checkbox">
                      <input type="checkbox" id="remove-dots" checked>
                      Remove dots (.)
                    </label>
                  </div>
                </div>
                <div class="column">
                  <div class="field">
                    <label class="checkbox">
                      <input type="checkbox" id="remove-nonalpha">
                      Remove all non-alphanumeric
                    </label>
                  </div>
                  <div class="field">
                    <label class="checkbox">
                      <input type="checkbox" id="to-uppercase" checked>
                      Convert to uppercase
                    </label>
                  </div>
                  <div class="field">
                    <label class="label">Max Length</label>
                    <div class="control">
                      <input class="input is-small" type="number" id="max-length" placeholder="No limit" min="1" max="50">
                    </div>
                  </div>
                </div>
              </div>

              <div class="columns">
                <div class="column">
                  <div class="field">
                    <label class="label">Prefix</label>
                    <div class="control">
                      <input class="input is-small" type="text" id="prefix" placeholder="Optional prefix">
                    </div>
                  </div>
                </div>
                <div class="column">
                  <div class="field">
                    <label class="label">Suffix (after context)</label>
                    <div class="control">
                      <input class="input is-small" type="text" id="suffix" placeholder="Optional suffix">
                    </div>
                  </div>
                </div>
              </div>

              <div class="field">
                <label class="label">Preview</label>
                <div class="control">
                  <div class="box has-background-light">
                    <strong>Sample Source:</strong> <code id="sample-input">${this.escapeHtml(
                      sampleSourceValue
                    )}</code><br>
                    ${
                      contextColumnName
                        ? `<strong>Sample Context:</strong> <code>${this.escapeHtml(
                            sampleContextValue
                          )}</code><br>`
                        : ''
                    }
                    <strong>Result:</strong> <code id="sample-output"></code>
                  </div>
                </div>
              </div>

              <div class="field is-grouped is-grouped-right">
                <div class="control">
                  <button class="button" id="cancel-reset-barcodes">Cancel</button>
                </div>
                <div class="control">
                  <button class="button is-primary" id="confirm-reset-barcodes">Reset Barcodes</button>
                </div>
              </div>
            </div>
          </div>
          <button class="modal-close is-large" aria-label="close"></button>
        </div>`;

      const modal = $(modalHtml);
      $('body').append(modal);

      // Update preview when options change
      const updatePreview = () => {
        const options = this.getBarcodeFormattingOptions(modal);
        const contextValue =
          contextColumnIndex !== -1 ? sampleContextValue : null;
        const preview = this.formatBarcode(
          sampleSourceValue,
          { ...options, includeContext: true },
          contextValue,
          'health'
        );
        modal.find('#sample-output').text(preview);
      };

      // Bind events
      modal
        .find(
          'input[type="checkbox"], input[type="number"], input[type="text"]'
        )
        .on('input change', updatePreview);
      modal
        .find('.modal-close, #cancel-reset-barcodes, .modal-background')
        .on('click', () => modal.remove());

      modal.find('#confirm-reset-barcodes').on('click', () => {
        const options = this.getBarcodeFormattingOptions(modal);
        this.executeResetBarcodes(barcodeIndex, sourceColumnIndex, options);
        modal.remove();
      });

      // Initial preview
      updatePreview();
    };

    /**
     * Extracts formatting options from the modal form
     */
    this.getBarcodeFormattingOptions = function (modal) {
      const maxLength = modal.find('#max-length').val();

      return {
        removeDashes: modal.find('#remove-dashes').is(':checked'),
        removeSpaces: modal.find('#remove-spaces').is(':checked'),
        removeUnderscores: modal.find('#remove-underscores').is(':checked'),
        removeDots: modal.find('#remove-dots').is(':checked'),
        removeNonAlphanumeric: modal.find('#remove-nonalpha').is(':checked'),
        toUpperCase: modal.find('#to-uppercase').is(':checked'),
        maxLength: maxLength ? parseInt(maxLength) : null,
        prefix: modal.find('#prefix').val().trim(),
        suffix: modal.find('#suffix').val().trim(),
      };
    };

    /**
     * Executes the actual barcode reset using generic formatting
     */
    this.executeResetBarcodes = function (
      barcodeIndex,
      sourceColumnIndex,
      options
    ) {
      this.showProcessLoader('Resetting barcodes...');

      setTimeout(() => {
        try {
          let updatedCount = 0;
          const errors = [];
          const headerRow = this.data.original[0];
          const contextColumnIndex = this.findColumnByType(
            headerRow,
            'health_status'
          );

          // Show info about context detection
          if (contextColumnIndex !== -1) {
            this.logDebug(
              `Context column found: "${headerRow[contextColumnIndex]}" at index ${contextColumnIndex}`
            );
          } else {
            this.logDebug(
              'No context column found - barcodes will not include context suffix'
            );
          }

          for (let i = 1; i < this.data.original.length; i++) {
            const row = this.data.original[i];

            if (row[sourceColumnIndex]) {
              try {
                const contextValue =
                  contextColumnIndex !== -1 ? row[contextColumnIndex] : null;
                const newBarcode = this.formatBarcode(
                  row[sourceColumnIndex],
                  { ...options, includeContext: true },
                  contextValue,
                  'health'
                );
                row[barcodeIndex] = newBarcode;
                updatedCount++;
              } catch (error) {
                errors.push(`Row ${i}: ${error.message}`);
              }
            }
          }

          if (updatedCount > 0) {
            this.data.dirty = true;
            this.data.filtered = this.deepClone(this.data.original);
            this.renderTable();

            let message = `Reset ${updatedCount} barcode values`;
            if (contextColumnIndex !== -1) {
              message += ' with context suffixes';
            }
            if (errors.length > 0) {
              message += ` (${errors.length} errors)`;
              this.logDebug('Barcode reset errors:', errors);
            }

            this.showMessage(
              message,
              errors.length > 0 ? 'warning' : 'success'
            );
          } else {
            this.showMessage('No source values found to convert', 'warning');
          }
        } catch (error) {
          this.handleError('Failed to reset barcodes', error);
        } finally {
          this.hideProcessLoader();
        }
      }, 100);
    };

    this.logDebug('ExcelEditorBarcodeSystem module loaded');
  };
})(jQuery);
