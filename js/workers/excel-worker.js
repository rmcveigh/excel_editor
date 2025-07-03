/**
 * @file
 * Excel Worker - Handles Excel parsing and export operations in a Web Worker
 * This is a self-contained worker with no external module dependencies.
 */

// Use the reliable importScripts to load the SheetJS library
importScripts(
  'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js'
);

// --- Start of Embedded Shared Logic ---

/**
 * Parses Excel data from an ArrayBuffer using SheetJS.
 * @param {ArrayBuffer} data - The file data.
 * @returns {Array} - The parsed data.
 */
function parseExcelData(data) {
  const workbook = XLSX.read(data, { type: 'array' });
  if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
    throw new Error('No worksheets found in Excel file.');
  }

  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  const jsonData = XLSX.utils.sheet_to_json(worksheet, {
    header: 1,
    raw: false,
    defval: '',
  });

  const trimmedData = jsonData.map((row) =>
    Array.isArray(row) ? row.map((cell) => String(cell || '').trim()) : row
  );

  const filteredData = trimmedData.filter(
    (row) =>
      Array.isArray(row) &&
      row.some(
        (cell) =>
          cell !== null && cell !== undefined && String(cell).trim() !== ''
      )
  );

  if (filteredData.length <= 1) {
    throw new Error('Excel file contains no data rows.');
  }

  return filteredData;
}

/**
 * Parses CSV data from a string.
 * @param {string} text - The CSV data.
 * @returns {Array} - The parsed data.
 */
function parseCSVData(text) {
  const lines = text.split('\n').filter((line) => line.trim());
  return lines.map((line) => {
    return line.split(',').map((cell) => {
      return cell.trim().replace(/^["']|["']$/g, '');
    });
  });
}

/**
 * Creates an XLSX file from an array of arrays.
 * @param {Array<Array>} data - The data to export.
 * @returns {Blob} - The generated XLSX file as a Blob.
 */
function createXLSX(data) {
  const ws = XLSX.utils.aoa_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Export');
  const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  return new Blob([wbout], { type: 'application/octet-stream' });
}

// --- End of Embedded Shared Logic ---

/**
 * Handles messages sent from the main thread.
 */
self.onmessage = function (event) {
  const { type, data, taskId } = event.data;

  try {
    let result;
    switch (type) {
      case 'parse_excel':
        result = parseExcelData(data);
        self.postMessage({
          type: 'parse_complete',
          taskId,
          success: true,
          data: result,
        });
        break;

      case 'parse_csv':
        result = parseCSVData(data);
        self.postMessage({
          type: 'parse_complete',
          taskId,
          success: true,
          data: result,
        });
        break;

      case 'export_excel':
        // eslint-disable-next-line no-case-declarations
        const blob = createXLSX(data.data);
        self.postMessage({
          type: 'export_complete',
          taskId,
          success: true,
          data: blob,
          filename: data.filename,
        });
        break;

      default:
        throw new Error(`Unknown operation type: ${type}`);
    }
  } catch (error) {
    self.postMessage({
      type: 'error',
      taskId,
      error: error.message,
    });
  }
};
