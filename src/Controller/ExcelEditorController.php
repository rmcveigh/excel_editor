<?php

namespace Drupal\excel_editor\Controller;

use Drupal\Core\Controller\ControllerBase;
use Drupal\Core\Database\Connection;
use Drupal\Core\Extension\ModuleExtensionList;
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
   * The database connection.
   *
   * @var \Drupal\Core\Database\Connection
   */
  protected Connection $connection;

  /**
   * Module extension list service.
   *
   * @var \Drupal\Core\Extension\ModuleExtensionList
   */
  protected ModuleExtensionList $moduleExtensionList;

  /**
   * Constructs a new ExcelEditorController object.
   *
   * @param \Drupal\excel_editor\DraftManager $draft_manager
   *   The draft manager service.
   * @param \Drupal\Core\Database\Connection $connection
   *   The database connection.
   * @param \Drupal\Core\Extension\ModuleExtensionList $module_extension_list
   *   The module extension list service.
   */
  public function __construct(DraftManager $draft_manager, Connection $connection, ModuleExtensionList $module_extension_list) {
    $this->draftManager = $draft_manager;
    $this->connection = $connection;
    $this->moduleExtensionList = $module_extension_list;
  }

  /**
   * {@inheritdoc}
   */
  public static function create(ContainerInterface $container) {
    return new static(
      $container->get('draft_manager'),
      $container->get('database'),
      $container->get('extension.list.module')
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
    $autosaveEnabled = $config->get('autosave_enabled');
    // Get the module path.
    $module_path = $this->moduleExtensionList->getPath('excel_editor');

    // Generate the proper path to the worker file.
    $worker_path = '/' . $module_path . '/js/workers/excel-worker.js';

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
              'getDogEntityUrls' => '/excel-editor/dog-entity-urls',
              'getTubeEntityUrls' => '/excel-editor/tube-entity-urls',
            ],
            'settings' => [
              'autosave_enabled' => (bool) $autosaveEnabled,
              'defaultVisibleColumns' => $defaultVisibleColumns,
              'hideBehavior' => $hideBehavior,
              'maxVisibleColumns' => (int) $maxVisibleColumns,
              'debug' => $this->currentUser()->hasPermission('administer excel editor'),
              'worker_path' => $worker_path,
              'base_path' => base_path(),
              'module_path' => '/' . $module_path,
            ],
          ],
        ],
      ],
    ];
  }

  /**
   * Save draft endpoint.
   *
   * @param \Symfony\Component\HttpFoundation\Request $request
   *   The request object containing the draft data.
   *
   * @return \Symfony\Component\HttpFoundation\JsonResponse
   *   Returns a JSON response indicating success or failure.
   */

  /**
   * Save draft endpoint.
   */
  public function saveDraft(Request $request) {
    try {
      if (!$request->isMethod('POST')) {
        return new JsonResponse(['success' => FALSE, 'message' => 'Only POST requests are allowed'], 405);
      }

      $content = $request->getContent();
      $data = json_decode($content, TRUE);

      if (json_last_error() !== JSON_ERROR_NONE || !isset($data['name']) || !isset($data['data'])) {
        return new JsonResponse(['success' => FALSE, 'message' => 'Invalid or missing data'], 400);
      }

      $draftName = trim($data['name']);
      if (empty($draftName)) {
        $draftName = 'Untitled Draft ' . date('Y-m-d H:i:s');
      }

      $draftId = $data['draft_id'] ?? NULL;

      if ($this->draftManager->isDraftNameTaken($draftName, $draftId ? (int) $draftId : NULL)) {
        return new JsonResponse([
          'success' => FALSE,
          'message' => 'This draft name is already in use. Please choose another.',
        ], 409);
      }

      $completeDraftData = [
        'data' => $data['data'],
        'filters' => $data['filters'] ?? [],
        'hiddenColumns' => $data['hiddenColumns'] ?? [],
        'selected' => $data['selected'] ?? [],
        'timestamp' => $data['timestamp'] ?? date('c'),
      ];

      if ($draftId) {
        $this->draftManager->updateDraft((int) $draftId, $draftName, $completeDraftData);
        $savedId = (int) $draftId;
      }
      else {
        $savedId = $this->draftManager->saveDraft($draftName, $completeDraftData);
      }

      if ($savedId) {
        return new JsonResponse(['success' => TRUE, 'draft_id' => $savedId, 'message' => 'Draft saved successfully']);
      }

      throw new \Exception('Failed to save draft to database');
    }
    catch (\Exception $e) {
      $this->getLogger('excel_editor')->error('Error saving draft: @error', ['@error' => $e->getMessage()]);
      return new JsonResponse(['success' => FALSE, 'message' => 'Failed to save draft: ' . $e->getMessage()], 500);
    }
  }

  /**
   * Load draft endpoint.
   *
   * @param mixed $draft_id
   *   The ID of the draft to load.
   * @param \Symfony\Component\HttpFoundation\Request $request
   *   The request object.
   *
   * @return \Symfony\Component\HttpFoundation\JsonResponse
   *   Returns a JSON response with the draft data or an error message.
   */
  public function loadDraft($draft_id, Request $request) {
    try {
      if (!is_numeric($draft_id)) {
        return new JsonResponse(['success' => FALSE, 'message' => 'Invalid draft ID'], 400);
      }

      $draft = $this->draftManager->loadDraft((int) $draft_id);

      if ($draft) {
        return new JsonResponse([
          'success' => TRUE,
          'id' => $draft->id,
          'name' => $draft->name,
          'data' => $draft->draft_data,
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
   *
   * @param mixed $draft_id
   *   The ID of the draft to delete.
   * @param \Symfony\Component\HttpFoundation\Request $request
   *   The request object.
   *
   * @return \Symfony\Component\HttpFoundation\JsonResponse
   *   Returns a JSON response indicating success or failure.
   */
  public function deleteDraft($draft_id, Request $request) {
    try {
      if (!is_numeric($draft_id)) {
        return new JsonResponse(['success' => FALSE, 'message' => 'Invalid draft ID'], 400);
      }

      if ($this->draftManager->deleteDraft((int) $draft_id)) {
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
   *
   * @param \Symfony\Component\HttpFoundation\Request $request
   *   The request object.
   *
   * @return \Symfony\Component\HttpFoundation\JsonResponse
   *   Returns a JSON response with the list of drafts or an error message.
   */
  public function listDrafts(Request $request) {
    try {
      $drafts = $this->draftManager->listDrafts();
      return new JsonResponse(['success' => TRUE, 'drafts' => $drafts]);
    }
    catch (\Exception $e) {
      $this->getLogger('excel_editor')->error('Error listing drafts: @error', ['@error' => $e->getMessage()]);
      return new JsonResponse(['success' => FALSE, 'message' => $e->getMessage()], 500);
    }
  }

  /**
   * Returns URLs for multiple grls_dog entities based on grls_ids.
   *
   * @param \Symfony\Component\HttpFoundation\Request $request
   *   The request object containing the grls_ids array.
   *
   * @return \Symfony\Component\HttpFoundation\JsonResponse
   *   Returns a JSON response with the URLs and labels of the grls_dogs.
   */
  public function getDogEntityUrls(Request $request) {
    try {
      $content = $request->getContent();
      $data = json_decode($content, TRUE);

      if (json_last_error() !== JSON_ERROR_NONE || empty($data['grls_ids']) || !is_array($data['grls_ids'])) {
        return new JsonResponse(['success' => FALSE, 'message' => 'Invalid or missing grls_ids array'], 400);
      }

      $grls_ids = array_unique(array_filter($data['grls_ids']));

      if (empty($grls_ids)) {
        return new JsonResponse(['success' => TRUE, 'urls' => []], 200);
      }

      // Use the database query directly for better performance with batches.
      $query = $this->connection->select('grls_dog', 'gd');
      $query->fields('gd', ['id', 'grls_id', 'name']);
      $query->condition('grls_id', $grls_ids, 'IN');
      $result = $query->execute()->fetchAllAssoc('grls_id');

      $urls = [];
      $dogStorage = $this->entityTypeManager()->getStorage('grls_dog');

      foreach ($grls_ids as $grls_id) {
        if (isset($result[$grls_id])) {
          $dog_id = $result[$grls_id]->id;
          $label = $result[$grls_id]->name;

          try {
            $dog = $dogStorage->load($dog_id);

            if ($dog) {
              $urls[$grls_id] = [
                'url' => $dog->toUrl('canonical')->toString(),
                'label' => $label,
                'entity_id' => $dog_id,
              ];
            }
          }
          catch (\Exception $entityError) {
            // Log the error but continue processing other entities.
            $this->getLogger('excel_editor')->warning('Failed to load dog entity @id: @error', [
              '@id' => $dog_id,
              '@error' => $entityError->getMessage(),
            ]);
            // Skip this entity but continue with others.
            continue;
          }
        }
      }

      return new JsonResponse([
        'success' => TRUE,
        'urls' => $urls,
      ]);
    }
    catch (\Exception $e) {
      $this->getLogger('excel_editor')->error('Error getting dog entity URLs: @error', ['@error' => $e->getMessage()]);
      return new JsonResponse([
        'success' => FALSE,
        'message' => 'Error getting dog entity URLs: ' . $e->getMessage(),
      ], 500);
    }
  }

  /**
   * Returns URLs for multiple grlstube entities based on azenta_ids.
   *
   * @param \Symfony\Component\HttpFoundation\Request $request
   *   The request object containing the azenta_id array.
   *
   * @return \Symfony\Component\HttpFoundation\JsonResponse
   *   Returns a JSON response with the URLs and labels of the grls_dogs.
   */
  public function getTubeEntityUrls(Request $request) {
    try {
      $content = $request->getContent();
      $data = json_decode($content, TRUE);

      if (json_last_error() !== JSON_ERROR_NONE || empty($data['azenta_ids']) || !is_array($data['azenta_ids'])) {
        return new JsonResponse(['success' => FALSE, 'message' => 'Invalid or missing azenta_ids array'], 400);
      }

      $azenta_ids = array_unique(array_filter($data['azenta_ids']));

      if (empty($azenta_ids)) {
        return new JsonResponse(['success' => TRUE, 'urls' => []], 200);
      }

      // Use the database query directly for better performance with batches.
      $query = $this->connection->select('grlstube', 'gt');
      $query->fields('gt', ['id', 'azenta_id', 'barcode']);
      $query->condition('azenta_id', $azenta_ids, 'IN');
      $result = $query->execute()->fetchAllAssoc('azenta_id');

      $urls = [];
      $tubeStorage = $this->entityTypeManager()->getStorage('grlstube');

      foreach ($azenta_ids as $azenta_id) {
        if (isset($result[$azenta_id])) {
          $tube_id = $result[$azenta_id]->id || NULL;
          $barcode = $result[$azenta_id]->barcode || NULL;

          try {
            $tube = $tubeStorage->load($tube_id);

            if ($tube) {
              $urls[$azenta_id] = [
                'url' => $tube->toUrl('canonical')->toString(),
                'barcode' => $barcode,
                'azenta_id' => $azenta_id,
                'entity_id' => $tube_id,
              ];
            }
          }
          catch (\Exception $entityError) {
            // Log the error but continue processing other entities.
            $this->getLogger('excel_editor')->warning('Failed to load tube entity @id: @error', [
              '@id' => $tube_id,
              '@error' => $entityError->getMessage(),
            ]);
            // Skip this entity but continue with others.
            continue;
          }
        }
      }

      return new JsonResponse([
        'success' => TRUE,
        'urls' => $urls,
      ]);
    }
    catch (\Exception $e) {
      $this->getLogger('excel_editor')->error('Error getting dog entity URLs: @error', ['@error' => $e->getMessage()]);
      return new JsonResponse([
        'success' => FALSE,
        'message' => 'Error getting dog entity URLs: ' . $e->getMessage(),
      ], 500);
    }
  }

}
