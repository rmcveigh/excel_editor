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
  protected DraftManager $draftManager;

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

    // Handle string/array from config.
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
              'debug' => $this->currentUser()->hasPermission('administer excel editor'),
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
    // Ensure proper JSON response headers.
    $response_data = ['success' => FALSE, 'message' => 'Unknown error'];

    try {
      // Log the request for debugging.
      $this->getLogger('excel_editor')->info('Save draft request received from user @uid', [
        '@uid' => $this->currentUser()->id(),
      ]);

      // Validate request method.
      if (!$request->isMethod('POST')) {
        $response_data['message'] = 'Only POST requests are allowed';
        return new JsonResponse($response_data, 405);
      }

      // Get and validate JSON data.
      $content = $request->getContent();
      if (empty($content)) {
        $response_data['message'] = 'No data received';
        return new JsonResponse($response_data, 400);
      }

      $data = json_decode($content, TRUE);
      if (json_last_error() !== JSON_ERROR_NONE) {
        $response_data['message'] = 'Invalid JSON data: ' . json_last_error_msg();
        return new JsonResponse($response_data, 400);
      }

      // Validate required fields.
      if (!isset($data['name']) || !isset($data['data'])) {
        $response_data['message'] = 'Missing required fields: name and data';
        return new JsonResponse($response_data, 400);
      }

      $draftName = trim($data['name']);
      $draftData = $data['data'];

      if (empty($draftName)) {
        $draftName = 'Untitled Draft ' . date('Y-m-d H:i:s');
      }

      if (empty($draftData) || !is_array($draftData)) {
        $response_data['message'] = 'Invalid or empty data';
        return new JsonResponse($response_data, 400);
      }

      // Prepare the complete draft data.
      $completeDraftData = [
        'data' => $draftData,
        'filters' => $data['filters'] ?? [],
        'hiddenColumns' => $data['hiddenColumns'] ?? [],
        'selected' => $data['selected'] ?? [],
        'timestamp' => $data['timestamp'] ?? date('c'),
      ];

      // Save the draft.
      $draft_id = $this->draftManager->saveDraft($draftName, $completeDraftData);

      if ($draft_id) {
        $response_data = [
          'success' => TRUE,
          'draft_id' => $draft_id,
          'message' => 'Draft saved successfully',
        ];

        $this->getLogger('excel_editor')->info('Draft saved successfully with ID @id for user @uid', [
          '@id' => $draft_id,
          '@uid' => $this->currentUser()->id(),
        ]);

        return new JsonResponse($response_data, 200);
      }
      else {
        throw new \Exception('Failed to save draft to database');
      }
    }
    catch (\Exception $e) {
      $this->getLogger('excel_editor')->error('Error saving draft: @error', [
        '@error' => $e->getMessage(),
      ]);

      $response_data['message'] = 'Failed to save draft: ' . $e->getMessage();
      return new JsonResponse($response_data, 500);
    }
  }

  /**
   * Load draft endpoint.
   */
  public function loadDraft($draft_id, Request $request) {
    try {
      // Validate draft ID.
      if (!is_numeric($draft_id)) {
        return new JsonResponse(['success' => FALSE, 'message' => 'Invalid draft ID'], 400);
      }

      $draft = $this->draftManager->loadDraft((int) $draft_id);

      if ($draft) {
        $this->getLogger('excel_editor')->info('Draft @id loaded for user @uid', [
          '@id' => $draft_id,
          '@uid' => $this->currentUser()->id(),
        ]);

        return new JsonResponse([
          'success' => TRUE,
          'data' => $draft->draft_data,
          'name' => $draft->name,
          'created' => $draft->created,
          'changed' => $draft->changed,
        ]);
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
  public function deleteDraft($draft_id, Request $request) {
    try {
      // Validate draft ID.
      if (!is_numeric($draft_id)) {
        return new JsonResponse(['success' => FALSE, 'message' => 'Invalid draft ID'], 400);
      }

      $deleted = $this->draftManager->deleteDraft((int) $draft_id);

      if ($deleted) {
        $this->getLogger('excel_editor')->info('Draft @id deleted for user @uid', [
          '@id' => $draft_id,
          '@uid' => $this->currentUser()->id(),
        ]);

        return new JsonResponse(['success' => TRUE, 'message' => 'Draft deleted successfully']);
      }

      return new JsonResponse(['success' => FALSE, 'message' => 'Draft not found or could not be deleted'], 404);
    }
    catch (\Exception $e) {
      $this->getLogger('excel_editor')->error('Error deleting draft: @error', ['@error' => $e->getMessage()]);
      return new JsonResponse(['success' => FALSE, 'message' => $e->getMessage()], 500);
    }
  }

  /**
   * List user's drafts endpoint.
   */
  public function listDrafts(Request $request) {
    try {
      $drafts = $this->draftManager->listDrafts();

      // Format the drafts to match the expected JS format.
      $formatted_drafts = array_map(function ($draft) {
        // Try to get row count from draft data if available.
        $row_count = 'N/A';
        try {
          // The draft data is already decoded by DraftManager.
          if (isset($draft->draft_data) && is_array($draft->draft_data)) {
            $data = $draft->draft_data;
            if (isset($data['data']) && is_array($data['data'])) {
              // Subtract 1 for header row.
              $row_count = max(0, count($data['data']) - 1);
            }
          }
        }
        catch (\Exception $e) {
          // If we can't get row count, just use N/A.
        }

        return [
          'id' => $draft->id,
          'name' => $draft->name,
        // ISO format.
          'created' => date('c', $draft->created),
        // ISO format.
          'changed' => date('c', $draft->changed),
          'rows' => $row_count,
        ];
      }, $drafts);

      return new JsonResponse([
        'success' => TRUE,
        'drafts' => $formatted_drafts,
        'count' => count($formatted_drafts),
      ]);
    }
    catch (\Exception $e) {
      $this->getLogger('excel_editor')->error('Error listing drafts: @error', ['@error' => $e->getMessage()]);
      return new JsonResponse(['success' => FALSE, 'message' => $e->getMessage()], 500);
    }
  }

}
