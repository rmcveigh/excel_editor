/**
 * @file
 * Excel Editor Barcode System Module
 */

export class ExcelEditorBarcodeSystem {
  constructor(app) {
    this.app = app;
  }

  // =========================================================================
  // SPECIALIZED TISSUE RESEARCH BARCODE SYSTEM
  // =========================================================================

  /**
   * Specialized barcode formatter based on tissue research Excel formula
   * @param {string} subjectId - The subject ID (e.g., "12345")
   * @param {string} biopsyType - The type of biopsy (e.g., "B" for biopsy, "N" for necropsy)
   * @param {string} reqTissueType - The required tissue type (e.g., "TISSUE_BRAIN")
   * @param {string} vialTissueType - The tissue type from the vial label (e.g., "TISSUE_BRAIN")
   * @param {string} healthStatus - The health status (e.g., "Healthy", "Diseased")
   * @return {string} The formatted tissue research barcode
   */
  formatTissueResearchBarcode(
    subjectId,
    biopsyType,
    reqTissueType,
    vialTissueType,
    healthStatus
  ) {
    if (!subjectId) return '';

    let cleanSubjectId = String(subjectId).replace(/-/g, '');
    const biopsyCode = this.getBiopsyTypeCode(biopsyType);
    const tissueCode = this.getTissueTypeCode(reqTissueType, vialTissueType);
    const biopsySpecificCode = this.getBiopsySpecificCode(biopsyType);
    const healthCode = this.getHealthStatusCode(healthStatus);

    const barcode =
      cleanSubjectId +
      biopsyCode +
      tissueCode +
      biopsySpecificCode +
      healthCode +
      'R';

    return barcode;
  }

  /**
   * Gets the biopsy type code (B for biopsy, N for necropsy)
   * @param {string} biopsyType - The type of biopsy (e.g., "B", "N", "biopsy", "necropsy")
   * @return {string} The biopsy type code ('B' for biopsy, 'N' for necropsy)
   */
  getBiopsyTypeCode(biopsyType) {
    if (!biopsyType) return 'N';
    const normalized = String(biopsyType).trim().toLowerCase();
    return normalized === 'b' || normalized === 'biopsy' ? 'B' : 'N';
  }

  /**
   * Gets the tissue type code based on req and vial tissue types
   * @param {string} reqTissueType - The required tissue type (e.g., "TISSUE_BRAIN")
   * @param {string} vialTissueType - The tissue type from the vial label (e.g., "TISSUE_BRAIN")
   * @return {string} The tissue type code (e.g., '74' for brain, 'XX' for mismatch)
   */
  getTissueTypeCode(reqTissueType, vialTissueType) {
    if (!reqTissueType || !vialTissueType) return 'XX';

    const reqNormalized = String(reqTissueType).trim().toUpperCase();
    const vialNormalized = String(vialTissueType).trim().toUpperCase();

    if (reqNormalized !== vialNormalized) return 'XX';

    const tissueMap = {
      TISSUE_OTHER: '70',
      TISSUE_ADRENAL_GLAND: '71',
      TISSUE_BONE: '72',
      TISSUE_BONE_MARROW: '73',
      TISSUE_BRAIN: '74',
      TISSUE_COLON: '75',
      TISSUE_DUODENUM: '76',
      TISSUE_ESOPHAGUS: '77',
      TISSUE_EYE: '78',
      TISSUE_GONADS: '79',
      TISSUE_HEART: '80',
      TISSUE_ILEOCECOCOLIC_JUNCTION: '81',
      TISSUE_ILEUM: '82',
      TISSUE_JEJUNUM: '83',
      TISSUE_KIDNEY: '84',
      TISSUE_LIVER: '85',
      TISSUE_LUNG: '86',
      TISSUE_LYMPH_NODE: '87',
      TISSUE_ORAL_CAVITY: '88',
      TISSUE_PANCREAS: '89',
      TISSUE_PARATHYROID_GLAND: '90',
      TISSUE_PROSTATE: '91',
      TISSUE_RECTUM: '92',
      TISSUE_SKELETAL_MUSCLE: '93',
      TISSUE_SKIN: '94',
      TISSUE_SPINAL_CORD: '95',
      TISSUE_SPLEEN: '96',
      TISSUE_STOMACH: '97',
      TISSUE_THYROID: '98',
      TISSUE_URINARY_BLADDER: '99',
    };

    return tissueMap[reqNormalized] || 'XX';
  }

  /**
   * Gets the biopsy-specific code (XX for biopsy, 99 for necropsy)
   * @param {string} biopsyType - The type of biopsy (e.g., "B", "N", "biopsy", "necropsy")
   * @return {string} The biopsy-specific code ('XX' for biopsy, '99' for necropsy)
   */
  getBiopsySpecificCode(biopsyType) {
    if (!biopsyType) return '99';
    const normalized = String(biopsyType).trim().toLowerCase();
    return normalized === 'b' || normalized === 'biopsy' ? 'XX' : '99';
  }

  /**
   * Gets the health status code (D for diseased, H for healthy)
   * @param {string} healthStatus - The health status (e.g., "Healthy", "Diseased")
   * @return {string} The health status code ('H' for healthy, 'D' for diseased)
   */
  getHealthStatusCode(healthStatus) {
    if (!healthStatus) return 'H';
    const normalized = String(healthStatus).trim().toLowerCase();
    return normalized === 'diseased' || normalized === 'd' ? 'D' : 'H';
  }

  // =========================================================================
  // ENHANCED COLUMN DETECTION FOR TISSUE RESEARCH
  // =========================================================================

  /**
   * Enhanced column finder for tissue research columns
   * @param {Array<string>} headerRow - The header row of the data
   * @param {string} columnType - The type of column to find (e.g., 'biopsy_necropsy', 'req_tissue_type', 'vial_tissue_type')
   * @return {number} The index of the found column, or -1 if not found
   */
  findTissueResearchColumn(headerRow, columnType) {
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
  }

  /**
   * Finds the Biopsy or Necropsy column
   * @param {Array<string>} headerRow - The header row of the data
   * @return {number} The index of the Biopsy or Necropsy column, or -1 if not found
   */
  findBiopsyNecropsyColumn(headerRow) {
    const exactMatches = [
      'Biopsy or Necropsy',
      'Biopsy_or_Necropsy',
      'biopsy or necropsy',
      'BIOPSY OR NECROPSY',
      'BiopsyOrNecropsy',
    ];

    for (const exactMatch of exactMatches) {
      const index = headerRow.indexOf(exactMatch);
      if (index !== -1) {
        return index;
      }
    }

    const flexibleMatches = ['biopsy', 'necropsy', 'procedure', 'type'];

    for (let i = 0; i < headerRow.length; i++) {
      const header = String(headerRow[i]).trim().toLowerCase();
      for (const pattern of flexibleMatches) {
        if (
          header.includes(pattern) &&
          (header.includes('biopsy') || header.includes('necropsy'))
        ) {
          return i;
        }
      }
    }

    return -1;
  }

  /**
   * Finds the Required Tissue Type column
   * @param {Array<string>} headerRow - The header row of the data
   * @return {number} The index of the Required Tissue Type column, or -1 if not found
   */
  findReqTissueTypeColumn(headerRow) {
    const exactMatches = [
      'Req. Tissue Type',
      'Req Tissue Type',
      'Required Tissue Type',
      'Req_Tissue_Type',
      'REQ_TISSUE_TYPE',
    ];

    for (const exactMatch of exactMatches) {
      const index = headerRow.indexOf(exactMatch);
      if (index !== -1) {
        return index;
      }
    }

    for (let i = 0; i < headerRow.length; i++) {
      const header = String(headerRow[i]).trim().toLowerCase();
      if (
        (header.includes('req') || header.includes('required')) &&
        header.includes('tissue') &&
        header.includes('type')
      ) {
        return i;
      }
    }

    return -1;
  }

  /**
   * Finds the Vial Label Tissue Type column
   * @param {Array<string>} headerRow - The header row of the data
   * @return {number} The index of the Vial Label Tissue Type column, or -1 if not found
   */
  findVialTissueTypeColumn(headerRow) {
    const exactMatches = [
      'Vial Label Tissue Type',
      'Vial_Label_Tissue_Type',
      'VialLabelTissueType',
      'VIAL_LABEL_TISSUE_TYPE',
      'Vial Tissue Type',
    ];

    for (const exactMatch of exactMatches) {
      const index = headerRow.indexOf(exactMatch);
      if (index !== -1) {
        return index;
      }
    }

    for (let i = 0; i < headerRow.length; i++) {
      const header = String(headerRow[i]).trim().toLowerCase();
      if (
        header.includes('vial') &&
        header.includes('tissue') &&
        header.includes('type')
      ) {
        return i;
      }
    }

    return -1;
  }

  // =========================================================================
  // GENERIC COLUMN DETECTION SYSTEM
  // =========================================================================

  /**
   * Generic column finder that can locate various types of columns
   * @param {Array<string>} headerRow - The header row of the data
   * @param {string} columnType - The type of column to find (e.g., 'subject_id', 'health_status', 'barcode', 'category', 'priority')
   * @return {number} The index of the found column, or -1 if not found
   */
  findColumnByType(headerRow, columnType) {
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
        this.app.utilities.logDebug(`Unknown column type: ${columnType}`);
        return -1;
    }
  }

  /**
   * Enhanced Subject ID column finder
   * @param {Array<string>} headerRow - The header row of the data
   * @return {number} The index of the Subject ID column, or -1 if not found
   */
  findSubjectIdColumn(headerRow) {
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
        return index;
      }
    }

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
          return i;
        }
      }
    }

    this.app.utilities.logDebug('No ID column found');
    return -1;
  }

  /**
   * Enhanced Health Status column finder
   * @param {Array<string>} headerRow - The header row of the data
   * @return {number} The index of the Health Status column, or -1 if not found
   */
  findHealthStatusColumn(headerRow) {
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
        return index;
      }
    }

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
          return i;
        }
      }
    }

    this.app.utilities.logDebug('No Health Status column found');
    return -1;
  }

  /**
   * Barcode column finder
   * @param {Array<string>} headerRow - The header row of the data
   * @return {number} The index of the Barcode column, or -1 if not found
   */
  findBarcodeColumn(headerRow) {
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
        return index;
      }
    }

    return headerRow.indexOf('new_barcode');
  }

  /**
   * Category column finder for future extensibility
   * @param {Array<string>} headerRow - The header row of the data
   * @return {number} The index of the Category column, or -1 if not found
   */
  findCategoryColumn(headerRow) {
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
          return i;
        }
      }
    }

    return -1;
  }

  /**
   * Priority column finder for future extensibility
   * @param {Array<string>} headerRow - The header row of the data
   * @return {number} The index of the Priority column, or -1 if not found
   */
  findPriorityColumn(headerRow) {
    const patterns = ['priority', 'urgency', 'importance'];

    for (let i = 0; i < headerRow.length; i++) {
      const header = String(headerRow[i]).trim().toLowerCase();
      for (const pattern of patterns) {
        if (header.includes(pattern)) {
          return i;
        }
      }
    }

    return -1;
  }

  // =========================================================================
  // GENERIC BARCODE FORMATTING SYSTEM
  // =========================================================================

  /**
   * Generic barcode formatting function that can work with any source value
   * @param {string} sourceValue - The source value to format (e.g., subject ID, barcode)
   * @param {Object} options - Formatting options
   * @param {string|null} contextValue - Optional context value (e.g., health status, category)
   * @param {string} contextType - The type of context (e.g., 'health', 'category', 'priority', 'status')
   * @return {string} The formatted barcode
   */
  formatBarcode(
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

    if (settings.removeDashes) formatted = formatted.replace(/-/g, '');
    if (settings.removeSpaces) formatted = formatted.replace(/\s/g, '');
    if (settings.removeUnderscores) formatted = formatted.replace(/_/g, '');
    if (settings.removeDots) formatted = formatted.replace(/\./g, '');
    if (settings.removeNonAlphanumeric)
      formatted = formatted.replace(/[^a-zA-Z0-9]/g, '');
    if (settings.toUpperCase) formatted = formatted.toUpperCase();

    formatted = settings.prefix + formatted;

    let contextSuffix = '';
    if (settings.includeContext && contextValue !== null) {
      contextSuffix = this.getContextSuffix(contextValue, contextType);
    }

    formatted = formatted + contextSuffix + settings.suffix;

    if (settings.maxLength)
      formatted = formatted.substring(0, settings.maxLength);

    return formatted;
  }

  /**
   * Generic context suffix determination based on value and type
   * @param {string|null} contextValue - The context value to determine the suffix for
   * @param {string} contextType - The type of context (e.g., 'health', 'category', 'priority', 'status')
   * @return {string} The context suffix (e.g., 'H', 'D', 'P', 'S', etc.)
   */
  getContextSuffix(contextValue, contextType = 'health') {
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
        this.app.utilities.logDebug(`Unknown context type: ${contextType}`);
        return '';
    }
  }

  /**
   * Health status suffix (H for healthy, D for diseased)
   * @param {string|null} value - The health status value to determine the suffix for
   * @return {string} The health status suffix ('H' for healthy, 'D' for diseased, or empty string)
   */
  getHealthStatusSuffix(value) {
    if (!value) return '';

    const normalizedValue = String(value).trim().toLowerCase();

    const healthyValues = ['h', 'healthy', 'normal', 'good'];
    if (healthyValues.includes(normalizedValue)) {
      return 'H';
    }

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
  }

  /**
   * Category suffix for future extensibility
   * @param {string|null} value - The category value to determine the suffix for
   * @return {string} The category suffix (e.g., 'P' for primary, 'S' for secondary, etc.)
   */
  getCategorySuffix(value) {
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
  }

  /**
   * Priority suffix for future extensibility
   * @param {string|null} value - The priority value to determine the suffix for
   * @return {string} The priority suffix (e.g., 'H' for high, 'M' for medium, 'L' for low, etc.)
   */
  getPrioritySuffix(value) {
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
  }

  /**
   * Status suffix for future extensibility
   * @param {string|null} value - The status value to determine the suffix for
   * @return {string} The status suffix (e.g., 'A' for active, 'I' for inactive, 'P' for pending, etc.)
   */
  getStatusSuffix(value) {
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
  }

  // =========================================================================
  // BARCODE MANAGEMENT (RESET BARCODES FUNCTIONALITY)
  // =========================================================================

  /**
   * Resets/populates barcode values using either tissue research or generic formatter
   */
  resetBarcodes() {
    if (!this.app.data.original.length) {
      this.app.utilities.showMessage('No data loaded', 'warning');
      return;
    }

    const headerRow = this.app.data.original[0];
    const barcodeIndex = this.findColumnByType(headerRow, 'barcode');

    if (barcodeIndex === -1) {
      this.app.utilities.showMessage('No barcode column found', 'warning');
      return;
    }

    const isTissueResearchFile = this.app.dataManager.detectTissueResearchFile
      ? this.app.dataManager.detectTissueResearchFile(headerRow)
      : false;
    const hasGenericColumns =
      this.findColumnByType(headerRow, 'subject_id') !== -1;

    if (!isTissueResearchFile && !hasGenericColumns) {
      this.app.utilities.showMessage(
        'No source columns found for barcode generation',
        'warning'
      );
      return;
    }

    this.showBarcodeResetModalEnhanced(
      barcodeIndex,
      headerRow,
      isTissueResearchFile,
      hasGenericColumns
    );
  }

  /**
   * Shows enhanced barcode reset modal with tissue research and generic options
   * @param {number} barcodeIndex - The index of the barcode column
   * @param {Array<string>} headerRow - The header row of the data
   * @param {boolean} isTissueResearchFile - Whether the file is a tissue research file
   * @param {boolean} hasGenericColumns - Whether generic columns are present
   */
  showBarcodeResetModalEnhanced(
    barcodeIndex,
    headerRow,
    isTissueResearchFile,
    hasGenericColumns
  ) {
    const $ = jQuery;

    const hasExistingValues = this.app.data.original
      .slice(1)
      .some((row) => row[barcodeIndex]);
    const existingWarning = hasExistingValues
      ? '<div class="notification is-warning is-light mb-3"><strong>Warning:</strong> Some barcode values already exist and will be overwritten.</div>'
      : '';

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
            <input type="radio" name="barcode-format" value="generic" ${
              !isTissueResearchFile ? 'checked' : ''
            }>
            <strong>Generic Format</strong> - Simple subject ID with health suffix
            <br><small class="has-text-grey">Format: SubjectID + HealthSuffix (H/D)</small>
          </label>
        </div>`;
    }

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

    modal
      .find('.modal-close, #cancel-reset-barcodes, .modal-background')
      .on('click', () => modal.remove());

    modal.find('#confirm-reset-barcodes').on('click', () => {
      const selectedFormat = modal
        .find('input[name="barcode-format"]:checked')
        .val();
      this.executeResetBarcodesEnhanced(barcodeIndex, selectedFormat);
      setTimeout(() => modal.remove(), 1000);
    });
  }

  /**
   * Executes the enhanced barcode reset
   * @param {number} barcodeIndex - The index of the barcode column
   * @param {string} format - The selected format for resetting barcodes ('tissue-research' or 'generic')
   */
  executeResetBarcodesEnhanced(barcodeIndex, format) {
    this.app.utilities.showProcessLoader('Resetting barcodes...');

    setTimeout(() => {
      try {
        let updatedCount = 0;
        const errors = [];
        const headerRow = this.app.data.original[0];

        if (format === 'tissue-research') {
          const result = this.executeTissueResearchReset(
            barcodeIndex,
            headerRow
          );
          updatedCount = result.updated;
          errors.push(...result.errors);
        } else {
          const result = this.executeGenericReset(barcodeIndex, headerRow);
          updatedCount = result.updated;
          errors.push(...result.errors);
        }

        if (updatedCount > 0) {
          this.app.data.dirty = true;
          this.app.data.filtered = this.app.utilities.deepClone(
            this.app.data.original
          );
          this.app.uiRenderer.renderTable();

          let message = `Reset ${updatedCount} barcode values using ${format} format`;
          if (errors.length > 0) {
            message += ` (${errors.length} errors)`;
            this.app.utilities.logDebug('Barcode reset errors:', errors);
          }

          this.app.utilities.showMessage(
            message,
            errors.length > 0 ? 'warning' : 'success'
          );

          // Trigger validation after barcode reset
          setTimeout(() => {
            if (this.app.validationManager) {
              this.app.validationManager.validateExistingBarcodeFields();
            }
          }, 200);
        } else {
          this.app.utilities.showMessage(
            'No source values found to convert',
            'warning'
          );
        }
      } catch (error) {
        this.app.utilities.handleError('Failed to reset barcodes', error);
      } finally {
        this.app.utilities.hideProcessLoader();
      }
    }, 100);
  }

  /**
   * Executes tissue research barcode reset
   * @param {number} barcodeIndex - The index of the barcode column
   * @param {Array<string>} headerRow - The header row of the data
   * @return {Object} An object containing the number of updated rows and any errors encountered
   */
  executeTissueResearchReset(barcodeIndex, headerRow) {
    const columnIndices = {
      subjectId: this.findTissueResearchColumn(headerRow, 'subject_id'),
      biopsyType: this.findTissueResearchColumn(headerRow, 'biopsy_necropsy'),
      reqTissueType: this.findTissueResearchColumn(
        headerRow,
        'req_tissue_type'
      ),
      vialTissueType: this.findTissueResearchColumn(
        headerRow,
        'vial_tissue_type'
      ),
      healthStatus: this.findTissueResearchColumn(headerRow, 'health_status'),
    };

    let updated = 0;
    const errors = [];

    for (let i = 1; i < this.app.data.original.length; i++) {
      const row = this.app.data.original[i];

      if (columnIndices.subjectId !== -1 && row[columnIndices.subjectId]) {
        try {
          const newBarcode = this.formatTissueResearchBarcode(
            row[columnIndices.subjectId],
            columnIndices.biopsyType !== -1
              ? row[columnIndices.biopsyType]
              : '',
            columnIndices.reqTissueType !== -1
              ? row[columnIndices.reqTissueType]
              : '',
            columnIndices.vialTissueType !== -1
              ? row[columnIndices.vialTissueType]
              : '',
            columnIndices.healthStatus !== -1
              ? row[columnIndices.healthStatus]
              : ''
          );
          row[barcodeIndex] = newBarcode;
          updated++;
        } catch (error) {
          errors.push(`Row ${i}: ${error.message}`);
        }
      }
    }

    return { updated, errors };
  }

  /**
   * Executes generic barcode reset
   * @param {number} barcodeIndex - The index of the barcode column
   * @param {Array<string>} headerRow - The header row of the data
   * @return {Object} An object containing the number of updated rows and any errors encountered
   */
  executeGenericReset(barcodeIndex, headerRow) {
    const sourceColumnIndex = this.findColumnByType(headerRow, 'subject_id');
    const contextColumnIndex = this.findColumnByType(
      headerRow,
      'health_status'
    );

    let updated = 0;
    const errors = [];

    for (let i = 1; i < this.app.data.original.length; i++) {
      const row = this.app.data.original[i];

      if (sourceColumnIndex !== -1 && row[sourceColumnIndex]) {
        try {
          const contextValue =
            contextColumnIndex !== -1 ? row[contextColumnIndex] : null;
          const newBarcode = this.formatBarcode(
            row[sourceColumnIndex],
            {
              removeDashes: true,
              removeSpaces: true,
              toUpperCase: true,
              includeContext: true,
            },
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
  }
}
