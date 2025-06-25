<?php

namespace Drupal\excel_editor\Controller;

use Drupal\Core\Controller\ControllerBase;
use Drupal\excel_editor\DraftManager;
use Symfony\Component\HttpFoundation\JsonResponse;
use Symfony\Component\HttpFoundation\Request;

/**
 * Returns responses for Excel Editor routes.
 */
class ExcelEditorController extends ControllerBase {

  /**
   * The controller constructor.
   */
  public function __construct(
    private readonly DraftManager $draftManager,
  ) {}

  /**
   * Main Excel Editor page.
   */
  public function page() {
    $config = $this->config('excel_editor.settings');
    $defaultVisibleColumns = $config->get('default_visible_columns') ?: [];
    $hideBehavior = $config->get('hide_behavior') ?: 'show_all';
    $maxVisibleColumns = $config->get('max_visible_columns') ?: 50;

    // This logic to handle string/array from config is good.
    if (is_string($defaultVisibleColumns)) {
      $defaultVisibleColumns = array_filter(array_map('trim', explode("\n", $defaultVisibleColumns)));
    }

    return [
      '#theme' => 'excel_editor_page',
      '#attached' => [
        'library' => [
          'excel_editor/excel_editor',
          'excel_editor/bulma',
          'excel_editor/fontawesome',
        ],
        'drupalSettings' => [
          'excelEditor' => [
            'endpoints' => [
              'saveDraft' => '/excel-editor/save-draft',
              'loadDraft' => '/excel-editor/load-draft/',
              'listDrafts' => '/excel-editor/drafts',
              'deleteDraft' => '/excel-editor/delete-draft/',
            ],
            'settings' => [
              'defaultVisibleColumns' => $defaultVisibleColumns,
              'hideBehavior' => $hideBehavior,
              'maxVisibleColumns' => (int) $maxVisibleColumns,
            ],
          ],
        ],
      ],
    ];
  }

  /**
   * Save draft endpoint.
   */
  public function saveDraft(Request $request) {
    try {
      $data = json_decode($request->getContent(), TRUE);
      $draftName = $data['name'] ?? 'Untitled Draft';
      $draftData = $data['data'] ?? [];

      if (empty($draftData)) {
        return new JsonResponse(['success' => FALSE, 'message' => 'Invalid data'], 400);
      }

      $draft_id = $this->draftManager->saveDraft($draftName, $draftData);

      return new JsonResponse([
        'success' => TRUE,
        'draft_id' => $draft_id,
        'message' => 'Draft saved successfully',
      ]);
    }
    catch (\Exception $e) {
      $this->getLogger('excel_editor')->error('Error saving draft: @error', ['@error' => $e->getMessage()]);
      return new JsonResponse(['success' => FALSE, 'message' => $e->getMessage()], 500);
    }
  }

  /**
   * Load draft endpoint.
   */
  public function loadDraft($draft_id) {
    try {
      $draft = $this->draftManager->loadDraft($draft_id);
      if ($draft) {
        return new JsonResponse(['success' => TRUE, 'data' => $draft->draft_data]);
      }
      return new JsonResponse(['success' => FALSE, 'message' => 'Draft not found'], 404);
    }
    catch (\Exception $e) {
      $this->getLogger('excel_editor')->error('Error loading draft: @error', ['@error' => $e->getMessage()]);
      return new JsonResponse(['success' => FALSE, 'message' => $e->getMessage()], 500);
    }
  }

  /**
   * Delete draft endpoint.
   */
  public function deleteDraft($draft_id) {
    try {
      $this->draftManager->deleteDraft($draft_id);
      return new JsonResponse(['success' => TRUE, 'message' => 'Draft deleted successfully']);
    }
    catch (\Exception $e) {
      $this->getLogger('excel_editor')->error('Error deleting draft: @error', ['@error' => $e->getMessage()]);
      return new JsonResponse(['success' => FALSE, 'message' => $e->getMessage()], 500);
    }
  }

  /**
   * List user's drafts endpoint.
   */
  public function listDrafts() {
    try {
      $drafts = $this->draftManager->listDrafts();
      // Format the drafts to match the expected JS format.
      $formatted_drafts = array_map(function ($draft) {
        return [
          'id' => $draft->id,
          'name' => $draft->name,
          'created' => date('Y-m-d H:i:s', $draft->changed),
          // You could store row count or calculate it if needed.
          'rows' => 'N/A',
        ];
      }, $drafts);

      return new JsonResponse(['success' => TRUE, 'drafts' => $formatted_drafts]);
    }
    catch (\Exception $e) {
      $this->getLogger('excel_editor')->error('Error listing drafts: @error', ['@error' => $e->getMessage()]);
      return new JsonResponse(['success' => FALSE, 'message' => $e->getMessage()], 500);
    }
  }

  /**
   * Get main interface HTML template.
   */
  private function getMainInterfaceTemplate() {
    return '
    <div class="excel-editor-upload" id="excel-upload-area">
      <div class="has-text-centered">
        <p class="title is-4">Upload Excel File</p>
        <p class="subtitle is-6">Drag and drop an Excel file here or click to browse</p>
        <input type="file" id="excel-file-input" accept=".xlsx,.xls" style="display: none;">
        <button class="button is-primary" onclick="document.getElementById(\'excel-file-input\').click()">
          <span class="icon"><i class="fas fa-upload"></i></span>
          <span>Choose File</span>
        </button>
      </div>
    </div>

    <div class="excel-editor-loading">
      <div class="has-text-centered">
        <div class="loader"></div>
        <p>Processing Excel file...</p>
      </div>
    </div>

    <div id="excel-editor-main" style="display: none;">
      <div class="excel-editor-toolbar">
        <div class="field is-grouped">
          <div class="control">
            <button class="button is-success" id="save-draft-btn">
              <span class="icon"><i class="fas fa-save"></i></span>
              <span>Save Draft</span>
            </button>
          </div>
          <div class="control">
            <button class="button is-info" id="export-btn">
              <span class="icon"><i class="fas fa-download"></i></span>
              <span>Export Selected</span>
            </button>
          </div>
          <div class="control">
            <button class="button is-link" id="export-all-btn">
              <span class="icon"><i class="fas fa-download"></i></span>
              <span>Export All</span>
            </button>
          </div>
          <div class="control">
            <button class="button is-light" id="toggle-columns-btn">
              <span class="icon"><i class="fas fa-eye"></i></span>
              <span>Show/Hide Columns</span>
            </button>
          </div>
        </div>

        <div class="field is-grouped selection-controls">
          <div class="control">
            <button class="button is-small is-outlined" id="select-all-visible-btn">
              <span class="icon is-small"><i class="fas fa-check-square"></i></span>
              <span>Select All Visible</span>
            </button>
          </div>
          <div class="control">
            <button class="button is-small is-outlined" id="deselect-all-btn">
              <span class="icon is-small"><i class="fas fa-square"></i></span>
              <span>Deselect All</span>
            </button>
          </div>
          <div class="control">
            <span class="tag is-info" id="selection-count">0 rows selected</span>
          </div>
        </div>
      </div>

      <div class="excel-editor-filters field is-grouped is-grouped-multiline" id="filter-controls">
        <!-- Dynamic filters will be added here -->
      </div>

      <div class="excel-editor-table-container">
        <table class="excel-editor-table table is-fullwidth is-striped" id="excel-table">
          <!-- Table content will be dynamically generated -->
        </table>
      </div>
    </div>

    <div class="excel-editor-drafts">
      <div class="box">
        <h3 class="title is-5">Your Drafts</h3>
        <div id="drafts-list">
          <!-- Drafts will be loaded here -->
        </div>
      </div>
    </div>
  ';
  }

  /**
   * Get action dropdown HTML template.
   */
  private function getActionDropdownTemplate() {
    return '
      <div class="select is-small is-fullwidth">
        <select class="excel-editor-cell editable actions-dropdown" data-row="{ROW}" data-col="{COL}">
          <option value="" {SELECTED_EMPTY}>-- Select Action --</option>
          <option value="relabel" {SELECTED_RELABEL}>Relabel</option>
          <option value="pending" {SELECTED_PENDING}>Pending</option>
          <option value="discard" {SELECTED_DISCARD}>Discard</option>
        </select>
      </div>
    ';
  }

  /**
   * Get filter modal HTML template.
   */
  private function getFilterModalTemplate() {
    return '
      <div class="modal is-active">
        <div class="modal-background"></div>
        <div class="modal-card">
          <header class="modal-card-head">
            <p class="modal-card-title">
              <span class="icon"><i class="fas fa-filter"></i></span>
              Filter: {COLUMN_NAME}
            </p>
            <button class="delete" aria-label="close"></button>
          </header>
          <section class="modal-card-body">
            <div class="tabs is-boxed">
              <ul>
                <li class="is-active" data-tab="quick">
                  <a><span class="icon is-small"><i class="fas fa-bolt"></i></span><span>Quick Filter</span></a>
                </li>
                <li data-tab="advanced">
                  <a><span class="icon is-small"><i class="fas fa-cog"></i></span><span>Advanced</span></a>
                </li>
              </ul>
            </div>

            <div id="quick-filter-tab" class="tab-content">
              <div class="field">
                <label class="label">Select Values to Show:</label>
                <div class="control">
                  <select id="quick-filter-select" multiple="multiple" data-placeholder="Choose values to filter by..." style="width: 100%;">
                    {OPTIONS}
                  </select>
                </div>
                <p class="help">Select one or more values to filter by. Leave empty to show all.</p>
              </div>
            </div>

            <div id="advanced-filter-tab" class="tab-content" style="display: none;">
              <div class="field">
                <label class="label">Filter Type:</label>
                <div class="control">
                  <div class="select is-fullwidth">
                    <select id="filter-type">
                      <option value="equals">Equals</option>
                      <option value="contains">Contains</option>
                      <option value="starts">Starts with</option>
                      <option value="ends">Ends with</option>
                      <option value="not_equals">Does not equal</option>
                      <option value="not_contains">Does not contain</option>
                      <option value="empty">Is empty</option>
                      <option value="not_empty">Is not empty</option>
                    </select>
                  </div>
                </div>
              </div>
              <div class="field" id="filter-value-field">
                <label class="label">Value:</label>
                <div class="control">
                  <input class="input" type="text" id="filter-value" placeholder="Enter filter value">
                </div>
              </div>
            </div>
          </section>
          <footer class="modal-card-foot">
            <button class="button is-success" id="apply-filter">
              <span class="icon"><i class="fas fa-check"></i></span>
              <span>Apply Filter</span>
            </button>
            <button class="button" id="cancel-filter">Cancel</button>
            <button class="button is-warning" id="clear-column-filter" style="margin-left: auto;">
              <span class="icon"><i class="fas fa-times"></i></span>
              <span>Clear This Filter</span>
            </button>
          </footer>
        </div>
      </div>
    ';
  }

}
