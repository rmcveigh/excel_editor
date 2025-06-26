/**
 * @file
 * Excel Editor Draft Manager Module
 *
 * Handles generic draft management operations in the Excel Editor.
 */

/* eslint-disable no-console */
(function ($) {
  'use strict';

  window.ExcelEditorDraftManager = function () {
    /**
     * Shows a modal to ask the user for a draft name before saving manually.
     */
    this.saveDraft = async function () {
      if (!this.data.original.length) {
        this.showMessage('No data to save', 'warning');
        return;
      }

      const draftName = await this._promptForDraftName();
      if (!draftName) {
        return; // User canceled
      }

      this.showLoading('Saving draft...');
      try {
        const draftData = {
          name: draftName,
          data: this.data.original,
          filters: this.state.currentFilters,
          hiddenColumns: Array.from(this.state.hiddenColumns),
          selected: Array.from(this.data.selected),
          timestamp: new Date().toISOString(),
        };

        const response = await this.apiCall(
          'POST',
          this.config.endpoints.saveDraft,
          draftData
        );

        if (response.success && response.draft_id) {
          this.data.dirty = false;
          this.state.currentDraftId = response.draft_id;
          this.state.currentDraftName = draftName;
          this.showMessage(
            `Draft "${this.escapeHtml(draftName)}" saved successfully`,
            'success'
          );
          this.loadDrafts();
        } else {
          throw new Error(response.message || 'Failed to save draft');
        }
      } catch (error) {
        this.handleError('Failed to save draft', error);
      } finally {
        this.hideLoading();
      }
    };

    /**
     * Loads the data from a specific draft into the editor.
     * @param {number} draftId The ID of the draft to load.
     */
    this.loadDraft = async function (draftId) {
      if (!draftId) {
        this.showMessage('Invalid draft ID', 'error');
        return;
      }

      this.showLoading('Loading draft...');
      try {
        const response = await this.apiCall(
          'GET',
          `${this.config.endpoints.loadDraft}${draftId}`
        );

        if (response.success && response.data) {
          this.state.currentDraftId = response.id;
          this.state.currentDraftName = response.name;
          this.loadDraftData(response.data);
          this.showMessage(
            `Draft "${this.escapeHtml(response.name)}" loaded successfully`,
            'success'
          );
        } else {
          throw new Error(response.message || 'Failed to load draft');
        }
      } catch (error) {
        this.handleError('Failed to load draft', error);
      } finally {
        this.hideLoading();
      }
    };

    /**
     * Deletes a draft from the server.
     * @param {number} draftId The ID of the draft to delete.
     */
    this.deleteDraft = async function (draftId) {
      if (!draftId) {
        this.showMessage('Invalid draft ID', 'error');
        return;
      }

      if (!confirm(Drupal.t('Are you sure you want to delete this draft?'))) {
        return;
      }

      this.showQuickLoader('Deleting draft...');
      try {
        const response = await this.apiCall(
          'DELETE',
          `${this.config.endpoints.deleteDraft}${draftId}`
        );

        if (response.success) {
          this.showMessage('Draft deleted successfully', 'success');

          // Reset current draft if we deleted the active one
          if (this.state.currentDraftId === draftId) {
            this.state.currentDraftId = null;
            this.state.currentDraftName = null;
          }

          this.loadDrafts();
        } else {
          throw new Error(response.message || 'Failed to delete draft');
        }
      } catch (error) {
        this.handleError('Failed to delete draft', error);
      } finally {
        this.hideQuickLoader();
      }
    };

    /**
     * Loads a list of the user's drafts from the server.
     */
    this.loadDrafts = async function () {
      this.showQuickLoader('Loading drafts...');
      try {
        const response = await this.apiCall(
          'GET',
          this.config.endpoints.listDrafts
        );

        if (response.success && response.drafts) {
          this.renderDrafts(response.drafts);
        } else {
          this.elements.draftsContainer.html(
            `<p class="has-text-grey">${Drupal.t('No drafts available')}</p>`
          );
        }
      } catch (error) {
        this.logDebug('Failed to load drafts:', error);
        this.elements.draftsContainer.html(
          `<p class="has-text-danger">${Drupal.t('Error loading drafts')}</p>`
        );
      } finally {
        this.hideQuickLoader();
      }
    };

    /**
     * Performs the autosave operation.
     */
    this.autosaveDraft = async function () {
      if (!this.data.dirty || !this.state.currentDraftId) {
        return;
      }

      this.logDebug(`Autosaving draft ID: ${this.state.currentDraftId}`);
      this.showQuickLoader('Autosaving...');

      try {
        const draftData = {
          draft_id: this.state.currentDraftId,
          name: this.state.currentDraftName,
          data: this.data.original,
          filters: this.state.currentFilters,
          hiddenColumns: Array.from(this.state.hiddenColumns),
          selected: Array.from(this.data.selected),
          timestamp: new Date().toISOString(),
        };

        const response = await this.apiCall(
          'POST',
          this.config.endpoints.saveDraft,
          draftData
        );

        if (response.success) {
          this.data.dirty = false;
          this.showQuickLoader('Draft autosaved', 'success');
          setTimeout(() => this.hideQuickLoader(), 2000);
          this.loadDrafts();
        } else {
          throw new Error(response.message || 'Autosave failed');
        }
      } catch (error) {
        this.handleError('Autosave failed', error);
      } finally {
        // Ensure the loader is hidden in case of an early return
        setTimeout(() => this.hideQuickLoader(), 3000);
      }
    };

    /**
     * Renders the list of drafts into the UI.
     * @param {Array} drafts The array of draft objects from the server.
     */
    this.renderDrafts = function (drafts) {
      if (!drafts || drafts.length === 0) {
        this.elements.draftsContainer.html(
          `<p class="has-text-grey">${Drupal.t('No drafts found')}</p>`
        );
        return;
      }

      const draftsHtml = drafts
        .map((draft) => this._renderDraftItem(draft))
        .join('');
      this.elements.draftsContainer.html(draftsHtml);

      // Attach event handlers
      this._attachDraftEventHandlers();
    };

    /**
     * [HELPER] Renders a single draft item
     * @param {Object} draft The draft object
     * @returns {string} HTML for the draft item
     */
    this._renderDraftItem = function (draft) {
      const draftName = this.escapeHtml(draft.name || 'Untitled Draft');
      const timestamp = new Date(draft.changed * 1000).toLocaleString();
      const isCurrentDraft = this.state.currentDraftId === draft.id;
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
    };

    /**
     * [HELPER] Attaches event handlers to draft buttons
     */
    this._attachDraftEventHandlers = function () {
      // Load draft handlers
      this.elements.draftsContainer.find('.load-draft-btn').on('click', (e) => {
        const draftId = $(e.currentTarget).data('draft-id');
        this.loadDraft(draftId);
      });

      // Delete draft handlers
      this.elements.draftsContainer
        .find('.delete-draft-btn')
        .on('click', (e) => {
          const draftId = $(e.currentTarget).data('draft-id');
          this.deleteDraft(draftId);
        });
    };

    /**
     * Loads draft data into the application state.
     * @param {object} draftData The draft data object from the server.
     */
    this.loadDraftData = function (draftData) {
      if (!draftData) {
        this.showMessage('Invalid draft data', 'error');
        return;
      }

      this.data.original = draftData.data || [];
      this.data.filtered = this.deepClone(this.data.original);
      this.state.currentFilters = draftData.filters || {};
      this.state.hiddenColumns = new Set(draftData.hiddenColumns || []);
      this.data.selected = new Set(draftData.selected || []);
      this.data.dirty = false;

      // Update UI
      this.renderInterface();
      this.applyFilters();
      this.updateActiveFiltersDisplay();
      this.updateSelectionCount();
    };

    /**
     * [HELPER] Creates and displays a modal to prompt the user for a draft name.
     * @returns {Promise<string|null>} A promise that resolves with the draft name, or null if canceled.
     */
    this._promptForDraftName = function () {
      return new Promise((resolve) => {
        // Remove any existing modals
        $('.modal#save-draft-modal').remove();

        const defaultName =
          this.state.currentDraftName || `Draft ${new Date().toLocaleString()}`;

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
                         value="${this.escapeHtml(defaultName)}">
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
    };
  };
})(jQuery);
