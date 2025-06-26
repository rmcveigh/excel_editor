/**
 * @file
 * Excel Editor Utilities Module
 *
 * Handles loading indicators, messages, error handling, and helper functions.
 * This module adds utility methods to the ExcelEditor class.
 */

/* eslint-disable no-console */
(function ($) {
  'use strict';

  /**
   * Utilities module for Excel Editor.
   * This function is called on ExcelEditor instances to add utility methods.
   */
  window.ExcelEditorUtilities = function () {
    // =========================================================================
    // LOADING INDICATORS
    // =========================================================================

    /**
     * Shows the main loading indicator.
     * @param {string} message The message to display.
     */
    this.showLoading = function (message = 'Loading...') {
      this.state.isLoading = true;
      this.elements.loadingArea.find('p').text(message);
      this.elements.loadingArea.addClass('active').show();
    };

    /**
     * Hides the main loading indicator.
     */
    this.hideLoading = function () {
      this.state.isLoading = false;
      if (this.elements.loadingArea && this.elements.loadingArea.length > 0) {
        this.elements.loadingArea.removeClass('active').hide();
      }
      $('.excel-editor-loading').removeClass('active').hide();
    };

    /**
     * Shows a temporary process loader for intensive operations.
     * @param {string} message The message to display.
     */
    this.showProcessLoader = function (message = 'Processing...') {
      this.hideProcessLoader();
      const loaderId = 'process-loader-' + Date.now();
      const loader = $(
        `<div class="excel-editor-overlay-loader" id="${loaderId}">
          <div class="loading-content">
            <div class="excel-editor-spinner"></div>
            <p><strong>${this.escapeHtml(message)}</strong></p>
          </div>
        </div>`
      );
      $('body').append(loader);
      this.state.currentProcessLoader = loaderId;
    };

    /**
     * Hides the process loader.
     */
    this.hideProcessLoader = function () {
      if (this.state.currentProcessLoader) {
        $(`#${this.state.currentProcessLoader}`).remove();
        this.state.currentProcessLoader = null;
      }
      $('.excel-editor-overlay-loader').remove();
    };

    /**
     * Shows a small, temporary loader in the corner of the screen.
     * @param {string} message The message to display.
     */
    this.showQuickLoader = function (message = 'Working...') {
      this.hideQuickLoader();
      const loader = $(
        `<div class="excel-editor-quick-loader" id="quick-loader">
          <div class="spinner"></div>
          <span>${this.escapeHtml(message)}</span>
        </div>`
      );
      $('body').append(loader);
      setTimeout(() => loader.addClass('slide-in'), 10);
    };

    /**
     * Hides the quick loader.
     */
    this.hideQuickLoader = function () {
      const loader = $('#quick-loader');
      if (loader.length) {
        loader.addClass('slide-out');
        setTimeout(() => loader.remove(), 300);
      }
    };

    // =========================================================================
    // MESSAGING SYSTEM
    // =========================================================================

    /**
     * Displays a notification message to the user.
     * @param {string} message The message content.
     * @param {string} type The type of message (success, error, warning, info).
     * @param {number} duration How long to display the message in ms.
     */
    this.showMessage = function (message, type = 'info', duration = 5000) {
      const alertClass =
        {
          success: 'is-success',
          error: 'is-danger',
          warning: 'is-warning',
          info: 'is-info',
        }[type] || 'is-info';

      const messageElement = $(
        `<div class="notification ${alertClass} excel-editor-message">
          <button class="delete"></button>
          ${this.escapeHtml(message)}
        </div>`
      );

      this.elements.container.prepend(messageElement);

      messageElement.find('.delete').on('click', () => {
        messageElement.fadeOut(() => messageElement.remove());
      });

      if (duration > 0) {
        setTimeout(() => {
          messageElement.fadeOut(() => messageElement.remove());
        }, duration);
      }
    };

    /**
     * Centralized error handler. Logs to console and shows a user message.
     * @param {string} message The user-facing message.
     * @param {Error|null} error The caught error object.
     */
    this.handleError = function (message, error = null) {
      this.logDebug(message, error);
      let userMessage = message;
      if (error && error.message) {
        userMessage += `: ${error.message}`;
      }
      this.showMessage(userMessage, 'error');
    };

    // =========================================================================
    // HELPER FUNCTIONS
    // =========================================================================

    /**
     * Centralized debug logger. Only logs if debug mode is enabled.
     * @param {string} message The debug message.
     * @param {*} [data=null] Additional data to log.
     */
    this.logDebug = function (message, data = null) {
      const isDebugMode =
        this.config.settings.debug ||
        window.location.search.includes('debug=1') ||
        localStorage.getItem('excel_editor_debug') === 'true';

      if (isDebugMode) {
        console.log(`[Excel Editor] ${message}`, data);
      }
    };

    /**
     * Escapes HTML to prevent XSS vulnerabilities.
     * @param {string} text The text to escape.
     * @returns {string} The escaped HTML string.
     */
    this.escapeHtml = function (text) {
      const div = document.createElement('div');
      div.textContent = text;
      return div.innerHTML;
    };

    /**
     * Creates a deep clone of an object or array.
     * @param {*} obj The object or array to clone.
     * @returns {*} A deep copy.
     */
    this.deepClone = function (obj) {
      return JSON.parse(JSON.stringify(obj));
    };

    /**
     * Validates a file based on size and format.
     * @param {File} file The file to validate.
     * @returns {boolean} True if the file is valid.
     */
    this.validateFile = function (file) {
      if (file.size > this.config.maxFileSize) {
        this.showMessage(
          `File too large. Maximum size is ${
            this.config.maxFileSize / (1024 * 1024)
          }MB`,
          'error'
        );
        return false;
      }

      const extension = '.' + file.name.split('.').pop().toLowerCase();
      if (!this.config.supportedFormats.includes(extension)) {
        this.showMessage(
          `Unsupported file format. Supported formats: ${this.config.supportedFormats.join(
            ', '
          )}`,
          'error'
        );
        return false;
      }

      return true;
    };

    /**
     * Reads a file into an ArrayBuffer.
     * @param {File} file The file to read.
     * @returns {Promise<ArrayBuffer>}
     */
    this.readFile = function (file) {
      return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => resolve(e.target.result);
        reader.onerror = () => reject(new Error('Failed to read file'));
        reader.readAsArrayBuffer(file);
      });
    };

    // =========================================================================
    // API HELPERS
    // =========================================================================

    /**
     * Fetches the CSRF token from Drupal's session endpoint.
     */
    this.getCsrfToken = async function () {
      try {
        const response = await fetch('/session/token');
        if (response.ok) {
          this.csrfToken = await response.text();
          this.logDebug('CSRF token obtained:', this.csrfToken ? 'Yes' : 'No');
        } else {
          console.warn('Failed to get CSRF token:', response.status);
        }
      } catch (error) {
        console.warn('Error getting CSRF token:', error);
      }
    };

    /**
     * Centralized helper function for making API calls to the Drupal backend.
     * @param {string} method The HTTP method (GET, POST, DELETE).
     * @param {string} url The API endpoint URL.
     * @param {Object|null} data The data to send in the request body.
     * @returns {Promise<Object>} The JSON response from the server.
     */
    this.apiCall = async function (method, url, data = null) {
      if (
        (method === 'POST' || method === 'PUT' || method === 'DELETE') &&
        !this.csrfToken
      ) {
        await this.getCsrfToken();
      }

      const options = {
        method: method,
        headers: {
          'Content-Type': 'application/json',
          'X-Requested-With': 'XMLHttpRequest',
        },
        credentials: 'same-origin',
      };

      if (
        (method === 'POST' || method === 'PUT' || method === 'DELETE') &&
        this.csrfToken
      ) {
        options.headers['X-CSRF-Token'] = this.csrfToken;
      }

      if (data) {
        options.body = JSON.stringify(data);
      }

      const response = await fetch(url, options);
      if (!response.ok) {
        let errorMessage = `HTTP ${response.status}: ${response.statusText}`;
        try {
          const errorData = await response.json();
          if (errorData.message) errorMessage = errorData.message;
        } catch (e) {
          // Ignore if response is not JSON
        }
        throw new Error(errorMessage);
      }

      return response.json();
    };

    // =========================================================================
    // PERFORMANCE HELPERS
    // =========================================================================

    /**
     * Determines if intensive loading should be shown based on data size.
     * @param {string} operationType The type of operation.
     * @returns {boolean} Whether to show intensive loader.
     */
    this.shouldShowIntensiveLoader = function (operationType = 'default') {
      const thresholds = {
        filter: 100,
        render: 200,
        export: 500,
        default: 150,
      };

      const threshold = thresholds[operationType] || thresholds.default;
      return this.data.original.length > threshold;
    };

    /**
     * Creates a debounced function wrapper.
     * @param {Function} func The function to debounce.
     * @param {number} wait The delay in milliseconds.
     * @returns {Function} The debounced function.
     */
    this.debounce = function (func, wait) {
      let timeout;
      return function executedFunction(...args) {
        const later = () => {
          clearTimeout(timeout);
          func.apply(this, args);
        };
        clearTimeout(timeout);
        timeout = setTimeout(later, wait);
      };
    };

    // =========================================================================
    // INITIALIZATION HELPERS
    // =========================================================================

    /**
     * Starts the autosave timer.
     */
    this.startAutosave = function () {
      this.stopAutosave(); // Clear any existing timer
      this.autosaveTimer = setInterval(() => {
        this.autosaveDraft();
      }, this.config.autosaveInterval);
      this.logDebug(
        `Autosave started with interval: ${this.config.autosaveInterval}ms`
      );
    };

    /**
     * Stops the autosave timer.
     */
    this.stopAutosave = function () {
      if (this.autosaveTimer) {
        clearInterval(this.autosaveTimer);
        this.autosaveTimer = null;
        this.logDebug('Autosave stopped.');
      }
    };

    this.logDebug('ExcelEditorUtilities module loaded');
  };
})(jQuery);
