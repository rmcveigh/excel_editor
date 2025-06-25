<?php

declare(strict_types=1);

namespace Drupal\excel_editor;

use Drupal\Core\Database\Connection;
use Drupal\Core\Session\AccountProxyInterface;

/**
 * Draft Manager service for Excel Editor module.
 */
final class DraftManager {

  /**
   * Constructs a DraftManager object.
   */
  public function __construct(
    private readonly Connection $connection,
    private readonly AccountProxyInterface $currentUser,
  ) {}

  /**
   * Saves a draft for the current user.
   *
   * @param string $name
   *   The name of the draft.
   * @param array $data
   *   The draft data to save.
   *
   * @return int
   *   The ID of the saved draft.
   */
  public function saveDraft(string $name, array $data) {
    $fields = [
      'uid' => $this->currentUser->id(),
      'name' => $name,
      'draft_data' => json_encode($data),
      'created' => time(),
      'changed' => time(),
    ];

    $query = $this->connection->insert('excel_editor_drafts')
      ->fields($fields);

    return $query->execute();
  }

  /**
   * Loads a specific draft for the current user.
   *
   * @param int $draft_id
   *   The ID of the draft to load.
   *
   * @return object|null
   *   The draft object or NULL if not found.
   */
  public function loadDraft(int $draft_id) {
    $query = $this->connection->select('excel_editor_drafts', 'd')
      ->fields('d')
      ->condition('d.id', $draft_id)
      ->condition('d.uid', $this->currentUser->id());

    $result = $query->execute()->fetchObject();

    if ($result && $result->draft_data) {
      // Decode the JSON data
      $result->draft_data = json_decode($result->draft_data, TRUE);
    }

    return $result;
  }

  /**
   * Deletes a specific draft for the current user.
   *
   * @param int $draft_id
   *   The ID of the draft to delete.
   *
   * @return int
   *   The number of rows deleted.
   */
  public function deleteDraft(int $draft_id) {
    $query = $this->connection->delete('excel_editor_drafts')
      ->condition('id', $draft_id)
      ->condition('uid', $this->currentUser->id());

    return $query->execute();
  }

  /**
   * Lists all drafts for the current user.
   *
   * @return array
   *   An array of draft objects.
   */
  public function listDrafts() {
    $query = $this->connection->select('excel_editor_drafts', 'd')
      ->fields('d', ['id', 'name', 'changed'])
      ->condition('d.uid', $this->currentUser->id())
      ->orderBy('d.changed', 'DESC');

    return $query->execute()->fetchAll();
  }

}
