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
   * Checks if a draft name is already taken by the current user.
   *
   * @param string $name
   *   The draft name to check.
   * @param int|null $excludeDraftId
   *   A draft ID to exclude from the check (used when updating a draft).
   *
   * @return bool
   *   TRUE if the name is taken, FALSE otherwise.
   *
   * @throws \Exception
   */
  public function isDraftNameTaken(string $name, ?int $excludeDraftId = NULL): bool {
    $query = $this->connection->select('excel_editor_drafts', 'd')
      ->fields('d', ['id'])
      ->condition('d.name', $name)
      ->condition('d.uid', $this->currentUser->id());

    if ($excludeDraftId !== NULL) {
      $query->condition('d.id', $excludeDraftId, '<>');
    }

    $result = $query->range(0, 1)->execute()->fetchField();

    return (bool) $result;
  }

  /**
   * Saves a new draft for the current user.
   *
   * @param string $name
   *   The name of the draft.
   * @param array $data
   *   The draft data to save.
   *
   * @return int|null
   *   The ID of the saved draft, or null on failure.
   *
   * @throws \Exception
   */
  public function saveDraft(string $name, array $data): ?int {
    $fields = [
      'uid' => $this->currentUser->id(),
      'name' => $name,
      'draft_data' => json_encode($data),
      'created' => time(),
      'changed' => time(),
    ];

    return (int) $this->connection->insert('excel_editor_drafts')
      ->fields($fields)
      ->execute();
  }

  /**
   * Updates an existing draft for the current user.
   *
   * @param int $draft_id
   *   The ID of the draft to update.
   * @param string $name
   *   The new name for the draft.
   * @param array $data
   *   The new draft data.
   *
   * @return int
   *   The number of affected rows.
   */
  public function updateDraft(int $draft_id, string $name, array $data): int {
    $fields = [
      'name' => $name,
      'draft_data' => json_encode($data),
      'changed' => time(),
    ];

    return $this->connection->update('excel_editor_drafts')
      ->fields($fields)
      ->condition('id', $draft_id)
      ->condition('uid', $this->currentUser->id())
      ->execute();
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
  public function loadDraft(int $draft_id): ?object {
    $result = $this->connection->select('excel_editor_drafts', 'd')
      ->fields('d')
      ->condition('d.id', $draft_id)
      ->condition('d.uid', $this->currentUser->id())
      ->execute()->fetchObject();

    if ($result && !empty($result->draft_data)) {
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
  public function deleteDraft(int $draft_id): int {
    return $this->connection->delete('excel_editor_drafts')
      ->condition('id', $draft_id)
      ->condition('uid', $this->currentUser->id())
      ->execute();
  }

  /**
   * Lists all drafts for the current user.
   *
   * @return array
   *   An array of draft objects.
   */
  public function listDrafts(): array {
    $query = $this->connection->select('excel_editor_drafts', 'd')
      ->fields('d', ['id', 'name', 'created', 'changed'])
      ->condition('d.uid', $this->currentUser->id())
      ->orderBy('d.changed', 'DESC');

    return $query->execute()->fetchAll(\PDO::FETCH_ASSOC);
  }

}
