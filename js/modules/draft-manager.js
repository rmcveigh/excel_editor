/**
 * @file
 * Excel Editor Draft Manager Module
 */

export class ExcelEditorDraftManager {
  constructor(app) {
    this.app = app;
  }

  /**
   * Shows a modal to ask the user for a draft name before saving manually.
   * @returns {Promise<void>} - Resolves when the draft is saved or canceled.
   */
  async saveDraft() {
    if (!this.app.data.original.length) {
      this.app.utilities.showMessage('No data to save', 'warning');
      return;
    }

    const draftName = await this.promptForDraftName();
    if (!draftName) {
      return; // User canceled
    }

    this.app.utilities.showLoading('Saving draft...');
    try {
      const draftData = {
        name: draftName,
        data: this.app.data.original,
        filters: this.app.state.currentFilters,
        hiddenColumns: Array.from(this.app.state.hiddenColumns),
        selected: Array.from(this.app.data.selected),
        timestamp: new Date().toISOString(),
      };

      const response = await this.app.utilities.apiCall(
        'POST',
        this.app.config.endpoints.saveDraft,
        draftData
      );

      if (response.success && response.draft_id) {
        this.app.data.dirty = false;
        this.app.state.currentDraftId = response.draft_id;
        this.app.state.currentDraftName = draftName;
        this.app.utilities.showMessage(
          `Draft "${this.app.utilities.escapeHtml(
            draftName
          )}" saved successfully`,
          'success'
        );
        this.loadDrafts();
      } else {
        throw new Error(response.message || 'Failed to save draft');
      }
    } catch (error) {
      this.app.utilities.handleError('Failed to save draft', error);
    } finally {
      this.app.utilities.hideLoading();
    }
  }

  /**
   * Loads the data from a specific draft into the editor.
   * @param {string} draftId - The ID of the draft to load.
   * @returns {Promise<void>} - Resolves when the draft is loaded or an error occurs.
   */
  async loadDraft(draftId) {
    if (!draftId) {
      this.app.utilities.showMessage('Invalid draft ID', 'error');
      return;
    }

    this.app.utilities.showLoading('Loading draft...');
    try {
      const response = await this.app.utilities.apiCall(
        'GET',
        `${this.app.config.endpoints.loadDraft}${draftId}`
      );

      if (response.success && response.data) {
        this.app.state.currentDraftId = response.id;
        this.app.state.currentDraftName = response.name;
        this.loadDraftData(response.data);
        this.app.utilities.showMessage(
          `Draft "${this.app.utilities.escapeHtml(
            response.name
          )}" loaded successfully`,
          'success'
        );
      } else {
        throw new Error(response.message || 'Failed to load draft');
      }
    } catch (error) {
      this.app.utilities.handleError('Failed to load draft', error);
    } finally {
      this.app.utilities.hideLoading();
    }
  }

  /**
   * Deletes a draft from the server.
   * @param {string} draftId - The ID of the draft to delete.
   * @returns {Promise<void>} - Resolves when the draft is deleted or an error occurs.
   */
  async deleteDraft(draftId) {
    if (!draftId) {
      this.app.utilities.showMessage('Invalid draft ID', 'error');
      return;
    }

    if (!confirm(Drupal.t('Are you sure you want to delete this draft?'))) {
      return;
    }

    this.app.utilities.showQuickLoader('Deleting draft...');
    try {
      const response = await this.app.utilities.apiCall(
        'DELETE',
        `${this.app.config.endpoints.deleteDraft}${draftId}`
      );

      if (response.success) {
        this.app.utilities.showMessage('Draft deleted successfully', 'success');

        // Reset current draft if we deleted the active one
        if (this.app.state.currentDraftId === draftId) {
          this.app.state.currentDraftId = null;
          this.app.state.currentDraftName = null;
        }

        this.loadDrafts();
      } else {
        throw new Error(response.message || 'Failed to delete draft');
      }
    } catch (error) {
      this.app.utilities.handleError('Failed to delete draft', error);
    } finally {
      this.app.utilities.hideQuickLoader();
    }
  }

  /**
   * Loads a list of the user's drafts from the server.
   * @returns {Promise<void>} - Resolves when the drafts are loaded or an error occurs.
   */
  async loadDrafts() {
    this.app.utilities.showQuickLoader('Loading drafts...');
    try {
      const response = await this.app.utilities.apiCall(
        'GET',
        this.app.config.endpoints.listDrafts
      );

      if (response.success && response.drafts) {
        this.renderDrafts(response.drafts);
      } else {
        this.app.elements.draftsContainer.html(
          `<p class="has-text-grey">${Drupal.t('No drafts available')}</p>`
        );
      }
    } catch (error) {
      this.app.utilities.logDebug('Failed to load drafts:', error);
      this.app.elements.draftsContainer.html(
        `<p class="has-text-danger">${Drupal.t('Error loading drafts')}</p>`
      );
    } finally {
      this.app.utilities.hideQuickLoader();
    }
  }

  /**
   * Performs the autosave operation.
   * @returns {Promise<void>} - Resolves when the draft is autosaved or an error occurs.
   */
  async autosaveDraft() {
    if (!this.app.data.dirty || !this.app.state.currentDraftId) {
      return;
    }

    this.app.utilities.logDebug(
      `Autosaving draft ID: ${this.app.state.currentDraftId}`
    );
    this.app.utilities.showQuickLoader('Autosaving...');

    try {
      const draftData = {
        draft_id: this.app.state.currentDraftId,
        name: this.app.state.currentDraftName,
        data: this.app.data.original,
        filters: this.app.state.currentFilters,
        hiddenColumns: Array.from(this.app.state.hiddenColumns),
        selected: Array.from(this.app.data.selected),
        timestamp: new Date().toISOString(),
      };

      const response = await this.app.utilities.apiCall(
        'POST',
        this.app.config.endpoints.saveDraft,
        draftData
      );

      if (response.success) {
        this.app.data.dirty = false;
        this.app.utilities.showQuickLoader('Draft autosaved', 'success');
        setTimeout(() => this.app.utilities.hideQuickLoader(), 2000);
        this.loadDrafts();
      } else {
        throw new Error(response.message || 'Autosave failed');
      }
    } catch (error) {
      this.app.utilities.handleError('Autosave failed', error);
    } finally {
      // Ensure the loader is hidden in case of an early return
      setTimeout(() => this.app.utilities.hideQuickLoader(), 3000);
    }
  }

  /**
   * Renders the list of drafts into the UI.
   * @param {Array} drafts - The list of drafts to render.
   */
  renderDrafts(drafts) {
    if (!drafts || drafts.length === 0) {
      this.app.elements.draftsContainer.html(
        `<p class="has-text-grey">${Drupal.t('No drafts found')}</p>`
      );
      return;
    }

    const draftsHtml = drafts
      .map((draft) => this.renderDraftItem(draft))
      .join('');
    this.app.elements.draftsContainer.html(draftsHtml);

    // Attach event handlers
    this.attachDraftEventHandlers();
  }

  /**
   * Renders a single draft item
   * @param {Object} draft - The draft object containing id, name, changed timestamp, etc.
   * @return {string} HTML string for the draft item
   */
  renderDraftItem(draft) {
    const draftName = this.app.utilities.escapeHtml(
      draft.name || 'Untitled Draft'
    );
    const timestamp = new Date(draft.changed * 1000).toLocaleString();
    const isCurrentDraft = this.app.state.currentDraftId === draft.id;
    const currentClass = isCurrentDraft ? 'is-current-draft' : '';

    return `
    <div class="excel-editor-draft-item ${currentClass}">
      <div>
        <strong>${draftName}</strong>
        ${
          isCurrentDraft
            ? '<span class="tag is-small is-info ml-2">Current</span>'
            : ''
        }
        <br>
        <small class="has-text-grey">${timestamp}</small>
      </div>
      <div class="field is-grouped">
        <div class="control">
          <button class="button is-small is-info load-draft-btn" data-draft-id="${
            draft.id
          }">
            <span>${Drupal.t('Load')}</span>
          </button>
        </div>
        <div class="control">
          <button class="button is-small is-danger delete-draft-btn" data-draft-id="${
            draft.id
          }">
            <span class="icon is-small">
              <i class="fas fa-trash"></i>
            </span>
          </button>
        </div>
      </div>
    </div>`;
  }

  /**
   * Attaches event handlers to draft buttons
   */
  attachDraftEventHandlers() {
    const $ = jQuery;

    // Load draft handlers
    this.app.elements.draftsContainer
      .find('.load-draft-btn')
      .on('click', (e) => {
        const draftId = $(e.currentTarget).data('draft-id');
        this.loadDraft(draftId);
      });

    // Delete draft handlers
    this.app.elements.draftsContainer
      .find('.delete-draft-btn')
      .on('click', (e) => {
        const draftId = $(e.currentTarget).data('draft-id');
        this.deleteDraft(draftId);
      });
  }

  /**
   * Loads draft data into the application state.
   * @param {Object} draftData - The draft data object containing data, filters, hidden columns, and selected rows.
   */
  loadDraftData(draftData) {
    if (!draftData) {
      this.app.utilities.showMessage('Invalid draft data', 'error');
      return;
    }

    this.app.data.original = draftData.data || [];
    this.app.data.filtered = this.app.utilities.deepClone(
      this.app.data.original
    );
    this.app.state.currentFilters = draftData.filters || {};
    this.app.state.hiddenColumns = new Set(draftData.hiddenColumns || []);
    this.app.data.selected = new Set(draftData.selected || []);
    this.app.data.dirty = false;

    // Update UI
    this.app.uiRenderer.renderInterface();
    this.app.filterManager.applyFilters();
    this.app.filterManager.updateActiveFiltersDisplay();
    this.app.dataManager.updateSelectionCount();
  }

  /**
   * Creates and displays a modal to prompt the user for a draft name.
   * @returns {Promise<string|null>} A promise that resolves with the draft name or null if canceled.
   */
  promptForDraftName() {
    return new Promise((resolve) => {
      const $ = jQuery;

      // Remove any existing modals
      $('.modal#save-draft-modal').remove();

      const defaultName =
        this.app.state.currentDraftName ||
        `Draft ${new Date().toLocaleString()}`;

      const modalHtml = `
      <div class="modal is-active" id="save-draft-modal">
        <div class="modal-background"></div>
        <div class="modal-content">
          <div class="box">
            <h3 class="title is-4">${Drupal.t('Save Draft')}</h3>
            <div class="field">
              <label class="label">${Drupal.t('Draft Name')}</label>
              <div class="control">
                <input class="input" id="draft-name-input" type="text"
                       value="${this.app.utilities.escapeHtml(defaultName)}">
              </div>
            </div>
            <div class="field is-grouped is-grouped-right">
              <div class="control">
                <button class="button" id="cancel-save-draft">${Drupal.t(
                  'Cancel'
                )}</button>
              </div>
              <div class="control">
                <button class="button is-primary" id="confirm-save-draft">${Drupal.t(
                  'Save'
                )}</button>
              </div>
            </div>
          </div>
        </div>
      </div>`;

      const modal = $(modalHtml);
      $('body').append(modal);

      const nameInput = modal.find('#draft-name-input').focus().select();

      // Handle enter key in the input field
      nameInput.on('keypress', (e) => {
        if (e.which === 13) {
          // Enter key
          const draftName = nameInput.val().trim();
          if (draftName) {
            resolve(draftName);
            modal.remove();
          } else {
            nameInput.addClass('is-danger');
          }
        }
      });

      // Handle button clicks
      modal.find('#confirm-save-draft').on('click', () => {
        const draftName = nameInput.val().trim();
        if (draftName) {
          resolve(draftName);
          modal.remove();
        } else {
          nameInput.addClass('is-danger');
        }
      });

      modal.find('#cancel-save-draft, .modal-background').on('click', () => {
        resolve(null);
        modal.remove();
      });
    });
  }
}
