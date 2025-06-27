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
    // SPECIALIZED TISSUE RESEARCH BARCODE SYSTEM
    // =========================================================================

    /**
     * Specialized barcode formatter based on tissue research Excel formula
     * Implements: CONCAT(SubjectID, BiopsyType, TissueCode, BiopsyCode, HealthStatus, "R")
     * @param {string} subjectId - The subject ID value
     * @param {string} biopsyType - Biopsy or Necropsy value
     * @param {string} reqTissueType - Required tissue type
     * @param {string} vialTissueType - Vial label tissue type
     * @param {string} healthStatus - Tissue diseased or healthy status
     * @param {object} options - Additional formatting options
     * @returns {string} Formatted barcode
     */
    this.formatTissueResearchBarcode = function (
      subjectId,
      biopsyType,
      reqTissueType,
      vialTissueType,
      healthStatus,
      options = {}
    ) {
      if (!subjectId) return '';

      // Clean subject ID (remove dashes)
      let cleanSubjectId = String(subjectId).replace(/-/g, '');

      // Determine biopsy type code (B or N) - case insensitive
      const biopsyCode = this.getBiopsyTypeCode(biopsyType);

      // Get tissue type code (70-99 or XX)
      const tissueCode = this.getTissueTypeCode(reqTissueType, vialTissueType);

      // Get biopsy-specific code (XX for biopsy, 99 for necropsy)
      const biopsySpecificCode = this.getBiopsySpecificCode(biopsyType);

      // Get health status code (D or H) - case insensitive
      const healthCode = this.getHealthStatusCode(healthStatus);

      // Combine all parts
      const barcode = cleanSubjectId + biopsyCode + tissueCode + biopsySpecificCode + healthCode + 'R';

      this.logDebug(`Tissue Research Barcode: "${subjectId}" → "${barcode}"`, {
        subjectId: cleanSubjectId,
        biopsyCode,
        tissueCode,
        biopsySpecificCode,
        healthCode,
        result: barcode
      });

      return barcode;
    };

    /**
     * Gets the biopsy type code (B for biopsy, N for necropsy)
     * @param {string} biopsyType - The biopsy type value
     * @returns {string} 'B' or 'N'
     */
    this.getBiopsyTypeCode = function (biopsyType) {
      if (!biopsyType) return 'N';

      const normalized = String(biopsyType).trim().toLowerCase();
      return (normalized === 'b' || normalized === 'biopsy') ? 'B' : 'N';
    };

    /**
     * Gets the tissue type code based on req and vial tissue types
     * @param {string} reqTissueType - Required tissue type
     * @param {string} vialTissueType - Vial label tissue type
     * @returns {string} Tissue code (70-99) or 'XX'
     */
    this.getTissueTypeCode = function (reqTissueType, vialTissueType) {
      if (!reqTissueType || !vialTissueType) return 'XX';

      const reqNormalized = String(reqTissueType).trim().toUpperCase();
      const vialNormalized = String(vialTissueType).trim().toUpperCase();

      // Only return a code if both values match
      if (reqNormalized !== vialNormalized) return 'XX';

      // Tissue type mapping from your Excel formula
      const tissueMap = {
        'TISSUE_OTHER': '70',
        'TISSUE_ADRENAL_GLAND': '71',
        'TISSUE_BONE': '72',
        'TISSUE_BONE_MARROW': '73',
        'TISSUE_BRAIN': '74',
        'TISSUE_COLON': '75',
        'TISSUE_DUODENUM': '76',
        'TISSUE_ESOPHAGUS': '77',
        'TISSUE_EYE': '78',
        'TISSUE_GONADS': '79',
        'TISSUE_HEART': '80',
        'TISSUE_ILEOCECOCOLIC_JUNCTION': '81',
        'TISSUE_ILEUM': '82',
        'TISSUE_JEJUNUM': '83',
        'TISSUE_KIDNEY': '84',
        'TISSUE_LIVER': '85',
        'TISSUE_LUNG': '86',
        'TISSUE_LYMPH_NODE': '87',
        'TISSUE_ORAL_CAVITY': '88',
        'TISSUE_PANCREAS': '89',
        'TISSUE_PARATHYROID_GLAND': '90',
        'TISSUE_PROSTATE': '91',
        'TISSUE_RECTUM': '92',
        'TISSUE_SKELETAL_MUSCLE': '93',
        'TISSUE_SKIN': '94',
        'TISSUE_SPINAL_CORD': '95',
        'TISSUE_SPLEEN': '96',
        'TISSUE_STOMACH': '97',
        'TISSUE_THYROID': '98',
        'TISSUE_URINARY_BLADDER': '99'
      };

      return tissueMap[reqNormalized] || 'XX';
    };

    /**
     * Gets the biopsy-specific code (XX for biopsy, 99 for necropsy)
     * @param {string} biopsyType - The biopsy type value
     * @returns {string} 'XX' or '99'
     */
    this.getBiopsySpecificCode = function (biopsyType) {
      if (!biopsyType) return '99';

      const normalized = String(biopsyType).trim().toLowerCase();
      return (normalized === 'b' || normalized === 'biopsy') ? 'XX' : '99';
    };

    /**
     * Gets the health status code (D for diseased, H for healthy)
     * @param {string} healthStatus - The health status value
     * @returns {string} 'D' or 'H'
     */
    this.getHealthStatusCode = function (healthStatus) {
      if (!healthStatus) return 'H';

      const normalized = String(healthStatus).trim().toLowerCase();
      return (normalized === 'diseased' || normalized === 'd') ? 'D' : 'H';
    };

    // =========================================================================
    // ENHANCED COLUMN DETECTION FOR TISSUE RESEARCH
    // =========================================================================

    /**
     * Enhanced column finder for tissue research columns
     * @param {Array} headerRow - The header row array
     * @param {string} columnType - Type of column to find
     * @returns {number} Column index or -1 if not found
     */
    this.findTissueResearchColumn = function (headerRow, columnType) {
      switch (columnType) {
        case 'biopsy_necropsy':
          return this.findBiopsyNecropsyColumn(headerRow);
        case 'req_tissue_type':
          return this.findReqTissueTypeColumn(headerRow);
        case 'vial_tissue_type':
          return this.findVialTissueTypeColumn(headerRow);
        default:
          return this.findColumnByType(headerRow, columnType);
      }
    };

    /**
     * Finds the Biopsy or Necropsy column
     * @param {Array} headerRow - The header row array
     * @returns {number} Column index or -1 if not found
     */
    this.findBiopsyNecropsyColumn = function (headerRow) {
      const exactMatches = [
        'Biopsy or Necropsy',
        'Biopsy_or_Necropsy',
        'biopsy or necropsy',
        'BIOPSY OR NECROPSY',
        'BiopsyOrNecropsy'
      ];

      for (const exactMatch of exactMatches) {
        const index = headerRow.indexOf(exactMatch);
        if (index !== -1) {
          this.logDebug(`Found Biopsy/Necropsy column: "${exactMatch}" at index ${index}`);
          return index;
        }
      }

      // Try partial matches
      const flexibleMatches = ['biopsy', 'necropsy', 'procedure', 'type'];

      for (let i = 0; i < headerRow.length; i++) {
        const header = String(headerRow[i]).trim().toLowerCase();

        for (const pattern of flexibleMatches) {
          if (header.includes(pattern) && (header.includes('biopsy') || header.includes('necropsy'))) {
            this.logDebug(`Found Biopsy/Necropsy column (flexible): "${headerRow[i]}" at index ${i}`);
            return i;
          }
        }
      }

      return -1;
    };

    /**
     * Finds the Required Tissue Type column
     * @param {Array} headerRow - The header row array
     * @returns {number} Column index or -1 if not found
     */
    this.findReqTissueTypeColumn = function (headerRow) {
      const exactMatches = [
        'Req. Tissue Type',
        'Req Tissue Type',
        'Required Tissue Type',
        'Req_Tissue_Type',
        'REQ_TISSUE_TYPE'
      ];

      for (const exactMatch of exactMatches) {
        const index = headerRow.indexOf(exactMatch);
        if (index !== -1) {
          this.logDebug(`Found Req Tissue Type column: "${exactMatch}" at index ${index}`);
          return index;
        }
      }

      // Try partial matches
      for (let i = 0; i < headerRow.length; i++) {
        const header = String(headerRow[i]).trim().toLowerCase();

        if ((header.includes('req') || header.includes('required')) &&
            header.includes('tissue') && header.includes('type')) {
          this.logDebug(`Found Req Tissue Type column (flexible): "${headerRow[i]}" at index ${i}`);
          return i;
        }
      }

      return -1;
    };

    /**
     * Finds the Vial Label Tissue Type column
     * @param {Array} headerRow - The header row array
     * @returns {number} Column index or -1 if not found
     */
    this.findVialTissueTypeColumn = function (headerRow) {
      const exactMatches = [
        'Vial Label Tissue Type',
        'Vial_Label_Tissue_Type',
        'VialLabelTissueType',
        'VIAL_LABEL_TISSUE_TYPE',
        'Vial Tissue Type'
      ];

      for (const exactMatch of exactMatches) {
        const index = headerRow.indexOf(exactMatch);
        if (index !== -1) {
          this.logDebug(`Found Vial Tissue Type column: "${exactMatch}" at index ${index}`);
          return index;
        }
      }

      // Try partial matches
      for (let i = 0; i < headerRow.length; i++) {
        const header = String(headerRow[i]).trim().toLowerCase();

        if (header.includes('vial') && header.includes('tissue') && header.includes('type')) {
          this.logDebug(`Found Vial Tissue Type column (flexible): "${headerRow[i]}" at index ${i}`);
          return i;
        }
      }

      return -1;
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
    // BARCODE MANAGEMENT (RESET BARCODES FUNCTIONALITY)
    // =========================================================================

    /**
     * Resets/populates barcode values using either tissue research or generic formatter
     */
    this.resetBarcodes = function () {
      if (!this.data.original.length) {
        this.showMessage('No data loaded', 'warning');
        return;
      }

      const headerRow = this.data.original[0];
      const barcodeIndex = this.findColumnByType(headerRow, 'barcode');

      if (barcodeIndex === -1) {
        this.showMessage('No barcode column found', 'warning');
        return;
      }

      // Detect available barcode types
      const isTissueResearchFile = this.detectTissueResearchFile ?
        this.detectTissueResearchFile(headerRow) : false;
      const hasGenericColumns = this.findColumnByType(headerRow, 'subject_id') !== -1;

      if (!isTissueResearchFile && !hasGenericColumns) {
        this.showMessage('No source columns found for barcode generation', 'warning');
        return;
      }

      // Show configuration modal with appropriate options
      this.showBarcodeResetModalEnhanced(barcodeIndex, headerRow, isTissueResearchFile, hasGenericColumns);
    };

    /**
     * Shows enhanced barcode reset modal with tissue research and generic options
     */
    this.showBarcodeResetModalEnhanced = function (barcodeIndex, headerRow, isTissueResearchFile, hasGenericColumns) {
      // Check if there are existing barcode values
      const hasExistingValues = this.data.original
        .slice(1)
        .some((row) => row[barcodeIndex]);
      const existingWarning = hasExistingValues
        ? '<div class="notification is-warning is-light mb-3"><strong>Warning:</strong> Some barcode values already exist and will be overwritten.</div>'
        : '';

      // Build format options
      let formatOptions = '';

      if (isTissueResearchFile) {
        formatOptions += `
          <div class="field">
            <label class="radio">
              <input type="radio" name="barcode-format" value="tissue-research" checked>
              <strong>Tissue Research Format</strong> - Based on your Excel formula
              <br><small class="has-text-grey">Format: SubjectID + BiopsyType + TissueCode + BiopsyCode + HealthStatus + "R"</small>
            </label>
          </div>`;
      }

      if (hasGenericColumns) {
        formatOptions += `
          <div class="field">
            <label class="radio">
              <input type="radio" name="barcode-format" value="generic" ${!isTissueResearchFile ? 'checked' : ''}>
              <strong>Generic Format</strong> - Simple subject ID with health suffix
              <br><small class="has-text-grey">Format: SubjectID + HealthSuffix (H/D)</small>
            </label>
          </div>`;
      }

      // Get sample data for preview
      const sampleRow = this.data.original.slice(1).find((row) => row.length > 0);
      const previewContent = this.buildBarcodePreview(headerRow, sampleRow, isTissueResearchFile);

      const modalHtml = `
        <div class="modal is-active" id="reset-barcodes-modal">
          <div class="modal-background"></div>
          <div class="modal-content">
            <div class="box">
              <h3 class="title is-4">
                <span class="icon"><i class="fas fa-barcode"></i></span>
                Reset Barcodes
              </h3>

              ${existingWarning}

              <div class="field">
                <label class="label">Barcode Format</label>
                ${formatOptions}
              </div>

              <div id="format-specific-options">
                <!-- Format-specific options will be loaded here -->
              </div>

              <div class="field">
                <label class="label">Preview</label>
                <div class="control">
                  <div class="box has-background-light" id="barcode-preview">
                    ${previewContent}
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

      // Update preview and options when format changes
      const updatePreviewAndOptions = () => {
        const selectedFormat = modal.find('input[name="barcode-format"]:checked').val();
        this.updateBarcodeFormatOptions(modal, selectedFormat, headerRow);
        this.updateBarcodePreview(modal, selectedFormat, headerRow, sampleRow);
      };

      // Bind events
      modal.find('input[name="barcode-format"]').on('change', updatePreviewAndOptions);
      modal.find('.modal-close, #cancel-reset-barcodes, .modal-background').on('click', () => modal.remove());

      modal.find('#confirm-reset-barcodes').on('click', () => {
        const selectedFormat = modal.find('input[name="barcode-format"]:checked').val();
        const options = this.getBarcodeFormattingOptionsEnhanced(modal, selectedFormat);
        this.executeResetBarcodesEnhanced(barcodeIndex, selectedFormat, options);
        modal.remove();
      });

      // Initial setup
      updatePreviewAndOptions();
    };

    /**
     * Updates the format-specific options section
     */
    this.updateBarcodeFormatOptions = function (modal, format, headerRow) {
      const optionsContainer = modal.find('#format-specific-options');

      if (format === 'generic') {
        optionsContainer.html(`
          <div class="field">
            <label class="label">Generic Format Options</label>
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
                  <input type="checkbox" id="to-uppercase" checked>
                  Convert to uppercase
                </label>
              </div>
            </div>
            <div class="column">
              <div class="field">
                <label class="label">Prefix</label>
                <div class="control">
                  <input class="input is-small" type="text" id="prefix" placeholder="Optional prefix">
                </div>
              </div>
              <div class="field">
                <label class="label">Suffix</label>
                <div class="control">
                  <input class="input is-small" type="text" id="suffix" placeholder="Optional suffix">
                </div>
              </div>
            </div>
          </div>`);
      } else {
        optionsContainer.html(`
          <div class="notification is-info is-light">
            <strong>Tissue Research Format</strong><br>
            This format follows your Excel formula exactly and doesn't have configurable options.
            All components are automatically determined from your data columns.
          </div>`);
      }

      // Bind change events for preview updates
      optionsContainer.find('input').on('input change', () => {
        this.updateBarcodePreview(modal, format, headerRow, this.data.original.slice(1).find((row) => row.length > 0));
      });
    };

    /**
     * Updates the barcode preview
     */
    this.updateBarcodePreview = function (modal, format, headerRow, sampleRow) {
      if (!sampleRow) {
        modal.find('#barcode-preview').html('<em>No sample data available</em>');
        return;
      }

      let preview = '';

      if (format === 'tissue-research') {
        const columnIndices = {
          subjectId: this.findTissueResearchColumn(headerRow, 'subject_id'),
          biopsyType: this.findTissueResearchColumn(headerRow, 'biopsy_necropsy'),
          reqTissueType: this.findTissueResearchColumn(headerRow, 'req_tissue_type'),
          vialTissueType: this.findTissueResearchColumn(headerRow, 'vial_tissue_type'),
          healthStatus: this.findTissueResearchColumn(headerRow, 'health_status')
        };

        if (columnIndices.subjectId !== -1) {
          const result = this.formatTissueResearchBarcode(
            sampleRow[columnIndices.subjectId] || 'SAMPLE123',
            columnIndices.biopsyType !== -1 ? sampleRow[columnIndices.biopsyType] : 'Biopsy',
            columnIndices.reqTissueType !== -1 ? sampleRow[columnIndices.reqTissueType] : 'TISSUE_LIVER',
            columnIndices.vialTissueType !== -1 ? sampleRow[columnIndices.vialTissueType] : 'TISSUE_LIVER',
            columnIndices.healthStatus !== -1 ? sampleRow[columnIndices.healthStatus] : 'Healthy'
          );
          preview = `<strong>Result:</strong> <code>${this.escapeHtml(result)}</code>`;
        }
      } else {
        const options = this.getBarcodeFormattingOptionsEnhanced(modal, format);
        const sourceColumnIndex = this.findColumnByType(headerRow, 'subject_id');
        const healthStatusIndex = this.findColumnByType(headerRow, 'health_status');

        if (sourceColumnIndex !== -1) {
          const sourceValue = sampleRow[sourceColumnIndex] || 'SAMPLE123';
          const healthValue = healthStatusIndex !== -1 ? sampleRow[healthStatusIndex] : 'Healthy';
          const result = this.formatBarcode(sourceValue, options, healthValue, 'health');
          preview = `<strong>Result:</strong> <code>${this.escapeHtml(result)}</code>`;
        }
      }

      modal.find('#barcode-preview').html(preview || '<em>Unable to generate preview</em>');
    };

    /**
     * Builds initial barcode preview content
     */
    this.buildBarcodePreview = function (headerRow, sampleRow, isTissueResearchFile) {
      if (!sampleRow) return '<em>No sample data available</em>';

      if (isTissueResearchFile) {
        return '<strong>Format:</strong> SubjectID + BiopsyType + TissueCode + BiopsyCode + HealthStatus + "R"<br><em>Preview will update when you select options above.</em>';
      } else {
        return '<strong>Format:</strong> SubjectID + HealthSuffix<br><em>Preview will update when you select options above.</em>';
      }
    };

    /**
     * Gets enhanced barcode formatting options from the modal
     */
    this.getBarcodeFormattingOptionsEnhanced = function (modal, format) {
      if (format === 'tissue-research') {
        return { format: 'tissue-research' };
      }

      return {
        format: 'generic',
        removeDashes: modal.find('#remove-dashes').is(':checked'),
        removeSpaces: modal.find('#remove-spaces').is(':checked'),
        toUpperCase: modal.find('#to-uppercase').is(':checked'),
        prefix: modal.find('#prefix').val().trim(),
        suffix: modal.find('#suffix').val().trim(),
      };
    };

    /**
     * Executes the enhanced barcode reset
     */
    this.executeResetBarcodesEnhanced = function (barcodeIndex, format, options) {
      this.showProcessLoader('Resetting barcodes...');

      setTimeout(() => {
        try {
          let updatedCount = 0;
          const errors = [];
          const headerRow = this.data.original[0];

          if (format === 'tissue-research') {
            const result = this.executeTissueResearchReset(barcodeIndex, headerRow);
            updatedCount = result.updated;
            errors.push(...result.errors);
          } else {
            const result = this.executeGenericReset(barcodeIndex, headerRow, options);
            updatedCount = result.updated;
            errors.push(...result.errors);
          }

          if (updatedCount > 0) {
            this.data.dirty = true;
            this.data.filtered = this.deepClone(this.data.original);
            this.renderTable();

            let message = `Reset ${updatedCount} barcode values using ${format} format`;
            if (errors.length > 0) {
              message += ` (${errors.length} errors)`;
              this.logDebug('Barcode reset errors:', errors);
            }

            this.showMessage(message, errors.length > 0 ? 'warning' : 'success');
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

    /**
     * Executes tissue research barcode reset
     */
    this.executeTissueResearchReset = function (barcodeIndex, headerRow) {
      const columnIndices = {
        subjectId: this.findTissueResearchColumn(headerRow, 'subject_id'),
        biopsyType: this.findTissueResearchColumn(headerRow, 'biopsy_necropsy'),
        reqTissueType: this.findTissueResearchColumn(headerRow, 'req_tissue_type'),
        vialTissueType: this.findTissueResearchColumn(headerRow, 'vial_tissue_type'),
        healthStatus: this.findTissueResearchColumn(headerRow, 'health_status')
      };

      let updated = 0;
      const errors = [];

      for (let i = 1; i < this.data.original.length; i++) {
        const row = this.data.original[i];

        if (columnIndices.subjectId !== -1 && row[columnIndices.subjectId]) {
          try {
            const newBarcode = this.formatTissueResearchBarcode(
              row[columnIndices.subjectId],
              columnIndices.biopsyType !== -1 ? row[columnIndices.biopsyType] : '',
              columnIndices.reqTissueType !== -1 ? row[columnIndices.reqTissueType] : '',
              columnIndices.vialTissueType !== -1 ? row[columnIndices.vialTissueType] : '',
              columnIndices.healthStatus !== -1 ? row[columnIndices.healthStatus] : ''
            );
            row[barcodeIndex] = newBarcode;
            updated++;
          } catch (error) {
            errors.push(`Row ${i}: ${error.message}`);
          }
        }
      }

      return { updated, errors };
    };

    /**
     * Executes generic barcode reset
     */
    this.executeGenericReset = function (barcodeIndex, headerRow, options) {
      const sourceColumnIndex = this.findColumnByType(headerRow, 'subject_id');
      const contextColumnIndex = this.findColumnByType(headerRow, 'health_status');

      let updated = 0;
      const errors = [];

      for (let i = 1; i < this.data.original.length; i++) {
        const row = this.data.original[i];

        if (sourceColumnIndex !== -1 && row[sourceColumnIndex]) {
          try {
            const contextValue = contextColumnIndex !== -1 ? row[contextColumnIndex] : null;
            const newBarcode = this.formatBarcode(
              row[sourceColumnIndex],
              { ...options, includeContext: true },
              contextValue,
              'health'
            );
            row[barcodeIndex] = newBarcode;
            updated++;
          } catch (error) {
            errors.push(`Row ${i}: ${error.message}`);
          }
        }
      }

      return { updated, errors };
    };

    this.logDebug('ExcelEditorBarcodeSystem module loaded');
  };
})(jQuery);
