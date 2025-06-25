<?php

namespace Drupal\excel_editor\Controller;

use Drupal\Core\Controller\ControllerBase;
use Drupal\excel_editor\DraftManager;
use Symfony\Component\DependencyInjection\ContainerInterface;
use Symfony\Component\HttpFoundation\JsonResponse;
use Symfony\Component\HttpFoundation\Request;

/**
 * Returns responses for Excel Editor routes.
 */
class ExcelEditorController extends ControllerBase {

  /**
   * The draft manager service.
   *
   * @var \Drupal\excel_editor\DraftManager
   */
  protected $draftManager;

  /**
   * Constructs a new ExcelEditorController object.
   *
   * @param \Drupal\excel_editor\DraftManager $draft_manager
   *   The draft manager service.
   */
  public function __construct(DraftManager $draft_manager) {
    $this->draftManager = $draft_manager;
  }

  /**
   * {@inheritdoc}
   */
  public static function create(ContainerInterface $container) {
    return new static(
      $container->get('draft_manager')
    );
  }

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

}
