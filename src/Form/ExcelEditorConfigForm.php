<?php

namespace Drupal\excel_editor\Form;

use Drupal\Core\Form\ConfigFormBase;
use Drupal\Core\Form\FormStateInterface;

/**
 * Configure Excel Editor settings.
 */
class ExcelEditorConfigForm extends ConfigFormBase {

  /**
   * {@inheritdoc}
   */
  protected function getEditableConfigNames() {
    return ['excel_editor.settings'];
  }

  /**
   * {@inheritdoc}
   */
  public function getFormId() {
    return 'excel_editor_admin_settings';
  }

  /**
   * {@inheritdoc}
   */
  public function buildForm(array $form, FormStateInterface $form_state) {
    $config = $this->config('excel_editor.settings');

    $form['column_visibility'] = [
      '#type' => 'fieldset',
      '#title' => $this->t('Default Column Visibility'),
      '#description' => $this->t('Configure which columns should be visible by default when users upload Excel files.'),
    ];

    $form['column_visibility']['default_visible_columns'] = [
      '#type' => 'textarea',
      '#title' => $this->t('Default Visible Columns'),
      '#description' => $this->t('Enter column names (one per line) that should be visible by default. Leave empty to show all columns. Common examples:<br>• new_barcode<br>• notes<br>• actions<br>• Product Name<br>• SKU<br>• Price'),
      '#default_value' => $this->formatColumnsForDisplay($config->get('default_visible_columns')),
      '#rows' => 8,
    ];

    $form['column_visibility']['hide_behavior'] = [
      '#type' => 'radios',
      '#title' => $this->t('Column Hiding Behavior'),
      '#description' => $this->t('Choose how the system should handle columns not in the default visible list.'),
      '#options' => [
        'hide_others' => $this->t('Hide all other columns (only show specified columns)'),
        'show_all' => $this->t('Show all columns (ignore the default visible setting)'),
      ],
      '#default_value' => $config->get('hide_behavior') ?: 'hide_others',
    ];

    $form['column_visibility']['always_visible'] = [
      '#type' => 'checkboxes',
      '#title' => $this->t('Always Visible Columns'),
      '#description' => $this->t('These columns will always be visible regardless of the default settings.'),
      '#options' => [
        'new_barcode' => $this->t('New Barcode Column'),
        'notes' => $this->t('Notes Column'),
        'actions' => $this->t('Actions Column'),
      ],
      '#default_value' => $config->get('always_visible') ?: ['new_barcode', 'notes', 'actions'],
    ];

    $form['performance'] = [
      '#type' => 'fieldset',
      '#title' => $this->t('Performance Settings'),
    ];

    $form['performance']['max_visible_columns'] = [
      '#type' => 'number',
      '#title' => $this->t('Maximum Visible Columns'),
      '#description' => $this->t('Limit the number of columns that can be visible at once to improve performance. Set to 0 for no limit.'),
      '#default_value' => $config->get('max_visible_columns') ?: 20,
      '#min' => 0,
      '#max' => 100,
    ];

    return parent::buildForm($form, $form_state);
  }

  /**
   * {@inheritdoc}
   */
  public function validateForm(array &$form, FormStateInterface $form_state) {
    $columns = $form_state->getValue('default_visible_columns');
    if (!empty($columns)) {
      $columnArray = array_filter(array_map('trim', explode("\n", $columns)));
      if (count($columnArray) > 50) {
        $form_state->setErrorByName('default_visible_columns', $this->t('Too many columns specified. Please limit to 50 or fewer.'));
      }
    }
    parent::validateForm($form, $form_state);
  }

  /**
   * {@inheritdoc}
   */
  public function submitForm(array &$form, FormStateInterface $form_state) {
    $columns = $form_state->getValue('default_visible_columns');
    $columnArray = [];

    if (!empty($columns)) {
      $columnArray = array_filter(array_map('trim', explode("\n", $columns)));
    }

    $alwaysVisible = array_filter($form_state->getValue('always_visible'));

    $this->config('excel_editor.settings')
      ->set('default_visible_columns', $columnArray)
      ->set('hide_behavior', $form_state->getValue('hide_behavior'))
      ->set('always_visible', $alwaysVisible)
      ->set('max_visible_columns', $form_state->getValue('max_visible_columns'))
      ->save();

    parent::submitForm($form, $form_state);
  }

  /**
   * Format columns array for textarea display.
   */
  private function formatColumnsForDisplay($columns) {
    if (empty($columns)) {
      return '';
    }
    if (is_array($columns)) {
      return implode("\n", $columns);
    }
    return $columns;
  }
}
