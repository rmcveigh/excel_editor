/**
 * @file
 * Excel Editor Utilities Module
 */

export class ExcelEditorUtilities {
  constructor(app) {
    this.app = app;
  }

  /**
   * Shows the main loading indicator.
   * @param {string} message - The message to display in the loading indicator.
   */
  showLoading(message = 'Loading...') {
    this.app.state.isLoading = true;
    this.app.elements.loadingArea.find('p').text(message);
    this.app.elements.loadingArea.addClass('active').show();
  }

  /**
   * Hides the main loading indicator.
   */
  hideLoading() {
    this.app.state.isLoading = false;
    if (
      this.app.elements.loadingArea &&
      this.app.elements.loadingArea.length > 0
    ) {
      this.app.elements.loadingArea.removeClass('active').hide();
    }
    jQuery('.excel-editor-loading').removeClass('active').hide();
  }

  /**
   * Shows a temporary process loader for intensive operations.
   * @param {string} message - The message to display in the loader.
   */
  showProcessLoader(message = 'Processing...') {
    this.hideProcessLoader();
    const loaderId = 'process-loader-' + Date.now();
    const loader = jQuery(
      `<div class="excel-editor-overlay-loader" id="${loaderId}">
        <div class="loading-content">
          <div class="excel-editor-spinner"></div>
          <p><strong>${this.escapeHtml(message)}</strong></p>
        </div>
      </div>`
    );
    jQuery('body').append(loader);
    this.app.state.currentProcessLoader = loaderId;
  }

  /**
   * Hides the process loader.
   */
  hideProcessLoader() {
    if (this.app.state.currentProcessLoader) {
      jQuery(`#${this.app.state.currentProcessLoader}`).remove();
      this.app.state.currentProcessLoader = null;
    }
    jQuery('.excel-editor-overlay-loader').remove();
  }

  /**
   * Shows a small, temporary loader in the corner of the screen.
   * @param {string} message - The message to display in the quick loader.
   */
  showQuickLoader(message = 'Working...') {
    this.hideQuickLoader();
    const loader = jQuery(
      `<div class="excel-editor-quick-loader" id="quick-loader">
        <div class="spinner"></div>
        <span>${this.escapeHtml(message)}</span>
      </div>`
    );
    jQuery('body').append(loader);
    setTimeout(() => loader.addClass('slide-in'), 10);
  }

  /**
   * Hides the quick loader.
   */
  hideQuickLoader() {
    const loader = jQuery('#quick-loader');
    if (loader.length) {
      loader.addClass('slide-out');
      setTimeout(() => loader.remove(), 300);
    }
  }

  /**
   * Displays a notification message to the user.
   * @param {string} message - The message to display.
   * @param {string} [type='info'] - The type of message ('success', 'error', 'warning', 'info').
   * @param {number} [duration=5000] - Duration in milliseconds to show the message.
   */
  showMessage(message, type = 'info', duration = 5000) {
    const alertClass =
      {
        success: 'is-success',
        error: 'is-danger',
        warning: 'is-warning',
        info: 'is-info',
      }[type] || 'is-info';

    const messageElement = jQuery(
      `<div class="notification ${alertClass} excel-editor-message">
        <button class="delete"></button>
        ${this.escapeHtml(message)}
      </div>`
    );

    this.app.elements.container.prepend(messageElement);

    messageElement.find('.delete').on('click', () => {
      messageElement.fadeOut(() => messageElement.remove());
    });

    if (duration > 0) {
      setTimeout(() => {
        messageElement.fadeOut(() => messageElement.remove());
      }, duration);
    }
  }

  /**
   * Centralized error handler.
   * @param {string} message - The error message to display.
   * @param {Error|null} [error=null] - Optional error object for additional details.
   */
  handleError(message, error = null) {
    this.logDebug(message, error);
    let userMessage = message;
    if (error && error.message) {
      userMessage += `: ${error.message}`;
    }
    this.showMessage(userMessage, 'error');
  }

  /**
   * Centralized debug logger.
   * @param {string} message - The debug message to log.
   * @param {any} [data=null] - Optional data to log alongside the message.
   */
  logDebug(message, data = null) {
    const isDebugMode =
      this.app.config.settings.debug ||
      window.location.search.includes('debug=1') ||
      localStorage.getItem('excel_editor_debug') === 'true';

    if (isDebugMode) {
      // eslint-disable-next-line no-console
      console.log(`[Excel Editor] ${message}`, data);
    }
  }

  /**
   * Escapes HTML to prevent XSS vulnerabilities.
   * @param {string} text - The text to escape.
   * @return {string} - The escaped HTML string.
   */
  escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
  }

  /**
   * Creates a deep clone of an object or array.
   * @param {Object|Array} obj - The object or array to clone.
   */
  deepClone(obj) {
    return JSON.parse(JSON.stringify(obj));
  }

  /**
   * Validates a file based on size and format.
   * @param {File} file - The file to validate.
   * @return {boolean} - True if the file is valid, false otherwise.
   */
  validateFile(file) {
    if (file.size > this.app.config.maxFileSize) {
      this.showMessage(
        `File too large. Maximum size is ${
          this.app.config.maxFileSize / (1024 * 1024)
        }MB`,
        'error'
      );
      return false;
    }

    const extension = '.' + file.name.split('.').pop().toLowerCase();
    if (!this.app.config.supportedFormats.includes(extension)) {
      this.showMessage(
        `Unsupported file format. Supported formats: ${this.app.config.supportedFormats.join(
          ', '
        )}`,
        'error'
      );
      return false;
    }

    return true;
  }

  /**
   * Reads a file into an ArrayBuffer.
   * @param {File} file - The file to read.
   * @return {Promise<ArrayBuffer>} - A promise that resolves with the file data.
   */
  readFile(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => resolve(e.target.result);
      reader.onerror = () => reject(new Error('Failed to read file'));
      reader.readAsArrayBuffer(file);
    });
  }

  /**
   * Fetches the CSRF token from Drupal's session endpoint.
   * @returns {Promise<void>} - A promise that resolves when the CSRF token is obtained.
   */
  async getCsrfToken() {
    try {
      const response = await fetch('/session/token');
      if (response.ok) {
        this.app.csrfToken = await response.text();
        this.logDebug(
          'CSRF token obtained:',
          this.app.csrfToken ? 'Yes' : 'No'
        );
      } else {
        console.warn('Failed to get CSRF token:', response.status);
      }
    } catch (error) {
      console.warn('Error getting CSRF token:', error);
    }
  }

  /**
   * Centralized helper function for making API calls.
   * @param {string} method - The HTTP method (GET, POST, PUT, DELETE).
   * @param {string} url - The API endpoint URL.
   * @param {Object|null} data - The data to send with the request (for POST/PUT).
   * @return {Promise<Object>} - A promise that resolves with the response data.
   */
  async apiCall(method, url, data = null) {
    if (
      (method === 'POST' || method === 'PUT' || method === 'DELETE') &&
      !this.app.csrfToken
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
      this.app.csrfToken
    ) {
      options.headers['X-CSRF-Token'] = this.app.csrfToken;
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
  }

  /**
   * Determines if intensive loading should be shown based on data size.
   * @param {string} operationType - The type of operation ('filter', 'render', 'export', or 'default').
   * @return {boolean} - True if intensive loader should be shown, false otherwise.
   */
  shouldShowIntensiveLoader(operationType = 'default') {
    const thresholds = {
      filter: 100,
      render: 200,
      export: 500,
      default: 150,
    };

    const threshold = thresholds[operationType] || thresholds.default;
    return this.app.data.original.length > threshold;
  }

  /**
   * Creates a debounced function wrapper.
   * @param {Function} func - The function to debounce.
   * @param {number} wait - The number of milliseconds to wait before invoking the function.
   * @return {Function} - A debounced version of the function.
   */
  debounce(func, wait) {
    let timeout;
    return function executedFunction(...args) {
      const later = () => {
        clearTimeout(timeout);
        func.apply(this, args);
      };
      clearTimeout(timeout);
      timeout = setTimeout(later, wait);
    };
  }

  /**
   * Starts the autosave timer.
   */
  startAutosave() {
    this.stopAutosave();
    this.app.autosaveTimer = setInterval(() => {
      this.app.draftManager.autosaveDraft();
    }, this.app.config.autosaveInterval);
    this.logDebug(
      `Autosave started with interval: ${this.app.config.autosaveInterval}ms`
    );
  }

  /**
   * Stops the autosave timer.
   */
  stopAutosave() {
    if (this.app.autosaveTimer) {
      clearInterval(this.app.autosaveTimer);
      this.app.autosaveTimer = null;
      this.logDebug('Autosave stopped.');
    }
  }

  /**
   * Convenience method for GET requests.
   * @param {string} url - The API endpoint URL.
   * @return {Promise<Object>} - A promise that resolves with the response data.
   */
  async apiGet(url) {
    return this.apiCall('GET', url);
  }

  /**
   * Convenience method for POST requests.
   * @param {string} url - The API endpoint URL.
   * @param {Object} data - The data to send with the request.
   * @return {Promise<Object>} - A promise that resolves with the response data.
   */
  async apiPost(url, data) {
    return this.apiCall('POST', url, data);
  }

  /**
   * Convenience method for PUT requests.
   * @param {string} url - The API endpoint URL.
   * @param {Object} data - The data to send with the request.
   * @return {Promise<Object>} - A promise that resolves with the response data.
   */
  async apiPut(url, data) {
    return this.apiCall('PUT', url, data);
  }

  /**
   * Convenience method for DELETE requests.
   * @param {string} url - The API endpoint URL.
   * @return {Promise<Object>} - A promise that resolves with the response data.
   */
  async apiDelete(url) {
    return this.apiCall('DELETE', url);
  }
}
