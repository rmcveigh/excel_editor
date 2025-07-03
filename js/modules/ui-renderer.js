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
   * Only fetch subject ID links once per dataset
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
        }, 1000);
      }

      // Only fetch subject ID and tube links if not cached for this dataset
      this.fetchSubjectIdLinksOptimized();
      this.fetchTubeLinksOptimized();
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
   * Better subject ID detection
   * @param {number} rowIndex - The index of the row in the filtered data.
   * @param {number} colIndex - The index of the column in the filtered data.
   * @param {string} cellValue - The value to display in the cell.
   * @return {HTMLTableCellElement} The created table cell element.
   */
  createTableCell(rowIndex, colIndex, cellValue) {
    const td = document.createElement('td');
    const columnName = this.app.data.filtered[0][colIndex];
    const isEditable = this.app.config.editableColumns.includes(columnName);

    // More flexible subject ID detection
    const isSubjectId = columnName && this.isSubjectIdColumn(columnName);

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

  // =========================================================================
  // OPTIMIZED SUBJECT ID AND TUBE LINK CACHING SYSTEM
  // =========================================================================

  /**
   * Fetches subject ID links only once per dataset
   */
  fetchSubjectIdLinksOptimized() {
    // Check if we've already fetched links for this dataset
    if (this.app.state.subjectLinksFullyFetched) {
      this.app.utilities.logDebug(
        'Subject ID links already fetched for this dataset, applying cached links'
      );
      this.applyCachedSubjectIdLinks();
      return;
    }

    // Check if we have subject ID data at all
    const allSubjectIds = this.getAllSubjectIdsFromDataset();

    if (allSubjectIds.length === 0) {
      this.app.utilities.logDebug('No subject IDs found in dataset');
      return;
    }

    // Filter out IDs we already have cached
    const uncachedIds = allSubjectIds.filter(
      (id) => id && !this.app.state.dogEntityUrlCache[id]
    );

    if (uncachedIds.length === 0) {
      this.app.utilities.logDebug('All subject IDs already cached');
      this.app.state.subjectLinksFullyFetched = true;
      this.applyCachedSubjectIdLinks();
      return;
    }

    // Fetch the missing links
    this.fetchDogEntityUrlsBatch(uncachedIds)
      .then((result) => {
        if (result && result.success && result.urls) {
          // Store all URLs in cache
          Object.entries(result.urls).forEach(([grlsId, urlData]) => {
            if (urlData && urlData.url) {
              this.app.state.dogEntityUrlCache[grlsId] = urlData;
            }
          });

          // Mark this dataset as fully fetched
          this.app.state.subjectLinksFullyFetched = true;

          // Apply all cached links to the current table
          this.applyCachedSubjectIdLinks();
        } else {
          this.app.utilities.logDebug('Failed to fetch subject ID links');
        }
      })
      .catch((error) => {
        this.app.utilities.logDebug('Error fetching subject ID links:', error);
      });
  }

  /**
   * Fetches tube links only once per dataset
   */
  fetchTubeLinksOptimized() {
    // Check if we've already fetched links for this dataset
    if (this.app.state.tubeLinksFullyFetched) {
      this.app.utilities.logDebug(
        'Tube links already fetched for this dataset, applying cached links'
      );
      this.applyCachedTubeLinks();
      return;
    }

    // Check if we have subject ID data at all
    const allAzentaIds = this.getAllAzentaIdsFromDataset();

    if (allAzentaIds.length === 0) {
      this.app.utilities.logDebug('No Azenta IDs found in dataset');
      return;
    }

    // Filter out IDs we already have cached
    const uncachedTubeIds = allAzentaIds.filter(
      (id) => id && !this.app.state.tubeEntityUrlCache[id]
    );

    if (uncachedTubeIds.length === 0) {
      this.app.utilities.logDebug('All Azenta IDs already cached');
      this.app.state.tubeLinksFullyFetched = true;
      this.applyCachedTubeLinks();
      return;
    }

    // Fetch the missing links
    this.fetchTubeEntityUrlsBatch(uncachedTubeIds)
      .then((result) => {
        if (result && result.success && result.urls) {
          // Store all URLs in cache
          Object.entries(result.urls).forEach(([azentaId, urlData]) => {
            if (urlData && urlData.url) {
              this.app.state.tubeEntityUrlCache[azentaId] = urlData;
            }
          });

          // Mark this dataset as fully fetched
          this.app.state.tubeLinksFullyFetched = true;

          // Apply all cached links to the current table
          this.applyCachedTubeLinks();
        } else {
          this.app.utilities.logDebug('Failed to fetch tube links');
        }
      })
      .catch((error) => {
        this.app.utilities.logDebug('Error fetching tube links:', error);
      });
  }

  /**
   * Gets all subject IDs from the current dataset
   */
  getAllSubjectIdsFromDataset() {
    const subjectIdColumnIndex = this.findSubjectIdColumnIndex();

    if (subjectIdColumnIndex === -1) {
      return [];
    }

    const subjectIds = [];

    // Get all subject IDs from the original dataset (not just filtered)
    for (let i = 1; i < this.app.data.original.length; i++) {
      const id = this.app.data.original[i][subjectIdColumnIndex];
      if (id && String(id).trim()) {
        subjectIds.push(String(id).trim());
      }
    }

    // Return unique IDs
    return [...new Set(subjectIds)];
  }

  /**
   * Gets all Azenta IDs from the current dataset
   */
  getAllAzentaIdsFromDataset() {
    const azentaIdColumnIndex = this.findAzentaIdColumnIndex();

    if (azentaIdColumnIndex === -1) {
      return [];
    }

    const azentaIds = [];

    // Get all subject IDs from the original dataset (not just filtered)
    for (let i = 1; i < this.app.data.original.length; i++) {
      const id = this.app.data.original[i][azentaIdColumnIndex];
      if (id && String(id).trim()) {
        azentaIds.push(String(id).trim().toUpperCase());
      }
    }

    // Return unique IDs
    return [...new Set(azentaIds)];
  }

  /**
   * Finds the subject ID column index
   */
  findSubjectIdColumnIndex() {
    const headerRow = this.app.data.original[0];

    for (let i = 0; i < headerRow.length; i++) {
      const columnName = headerRow[i];
      if (columnName && this.isSubjectIdColumn(columnName)) {
        return i;
      }
    }

    return -1;
  }

  /**
   * Finds the Azenta ID column index
   */
  findAzentaIdColumnIndex() {
    const headerRow = this.app.data.original[0];

    for (let i = 0; i < headerRow.length; i++) {
      const columnName = String(headerRow[i]).toLowerCase();
      if (columnName && columnName.includes('name')) {
        return i;
      }
    }

    return -1;
  }

  /**
   * Determines if a column name represents subject ID data
   */
  isSubjectIdColumn(columnName) {
    const name = String(columnName).toLowerCase();
    return (
      name.includes('subject id') ||
      name.includes('grls id') ||
      name.includes('dog id') ||
      name.includes('subject_id') ||
      name.includes('grls_id')
    );
  }

  /**
   * Applies cached tube links to all visible name cells
   */
  applyCachedTubeLinks() {
    const $ = jQuery;

    $('.name-cell:not(.processed)').each((index, element) => {
      const $element = $(element);
      const azentaId = $element.text().trim();

      if (azentaId && this.app.state.tubeEntityUrlCache[azentaId]) {
        const urlData = this.app.state.tubeEntityUrlCache[azentaId];

        // Create the link
        $element.html(
          `<a href="${this.app.utilities.escapeHtml(
            urlData.url
          )}" target="_blank">
            ${this.app.utilities.escapeHtml(azentaId)}
          </a>`
        );

        // Mark as processed
        $element.addClass('processed');
      } else {
        // No cached link available, mark as processed anyway
        $element.addClass('processed');
      }
    });
  }

  /**
   * Applies cached subject ID links to all visible subject ID cells
   */
  applyCachedSubjectIdLinks() {
    const $ = jQuery;

    $('.subject-id-cell:not(.processed)').each((index, element) => {
      const $element = $(element);
      const grlsId = $element.text().trim();

      if (grlsId && this.app.state.dogEntityUrlCache[grlsId]) {
        const urlData = this.app.state.dogEntityUrlCache[grlsId];

        // Create the link
        $element.html(
          `<a href="${this.app.utilities.escapeHtml(
            urlData.url
          )}" target="_blank">
            ${this.app.utilities.escapeHtml(grlsId)}
          </a>`
        );

        // Mark as processed
        $element.addClass('processed');
      } else {
        // No cached link available, mark as processed anyway
        $element.addClass('processed');
      }
    });
  }

  /**
   * Resets the tube link cache (call when loading new data)
   */
  resetTubeLinkCache() {
    this.app.state.tubeLinksFullyFetched = false;
    this.app.state.tubeEntityUrlCache = {};
  }

  /**
   * Resets the subject ID link cache (call when loading new data)
   */
  resetSubjectIdLinkCache() {
    this.app.state.subjectLinksFullyFetched = false;
    this.app.state.dogEntityUrlCache = {};
  }

  /**
   * Fetches dog entity URLs in batches
   * @param {Array} grlsIds - Array of GRLS IDs to fetch URLs
   * @returns {Promise<Object>} - A promise that resolves with the fetched URLs
   */
  async fetchDogEntityUrlsBatch(grlsIds) {
    try {
      if (!grlsIds || grlsIds.length === 0) {
        return { success: true, urls: {} };
      }

      return await this.app.utilities.apiPost(
        this.app.config.endpoints.getDogEntityUrls,
        { grls_ids: grlsIds }
      );
    } catch (error) {
      this.app.utilities.logDebug('Failed to fetch dog entity links:', error);
      return {
        success: false,
        urls: {},
        message: error.message,
      };
    }
  }

  /**
   * Fetches tube entity URLs in batches
   * @param {Array} azentaIds - Array of Azenta IDs to fetch URLs
   * @returns {Promise<Object>} - A promise that resolves with the fetched URLs
   */
  async fetchTubeEntityUrlsBatch(azentaIds) {
    try {
      if (!azentaIds || azentaIds.length === 0) {
        return { success: true, urls: {} };
      }

      return await this.app.utilities.apiPost(
        this.app.config.endpoints.getTubeEntityUrls,
        { azenta_ids: azentaIds }
      );
    } catch (error) {
      this.app.utilities.logDebug('Failed to fetch tube entity links:', error);
      return {
        success: false,
        urls: {},
        message: error.message,
      };
    }
  }
}
