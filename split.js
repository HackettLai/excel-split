/**
 * ============================================
 * DATA SPLIT TOOL - Main JavaScript Module
 * ============================================
 *
 * Purpose: Process Excel/CSV files and split cells containing multiple values
 * into separate rows while preserving all other data.
 *
 * Author: Hackett Lai
 * Version: 1.2.1
 *
 * Key Features:
 * - Excel/CSV file parsing using SheetJS library
 * - Multiple separator detection (comma, semicolon, slash, ampersand, newline)
 * - Support for both Western and Chinese separators
 * - Special term protection (N/A variants)
 * - Header row detection option
 * - Column selection by name or index
 * - Visual feedback with highlighted split rows
 * - Excel export with formatting
 */

// ============================================
// STATE MANAGEMENT
// ============================================

/**
 * Application state variables
 * These hold the current state of the application data
 */
let originalData = []; // Raw data rows from uploaded file (excluding headers)
let processedData = []; // Data after splitting operations (includes split metadata)
let headers = []; // Column headers (from file or auto-generated)
let currentFile = null; // Reference to currently loaded file object

// ============================================
// DOM ELEMENT REFERENCES
// ============================================

/**
 * Cache all DOM elements at initialization for better performance
 * and code readability. These elements are accessed frequently
 * throughout the application lifecycle.
 */

// File upload elements
const fileInput = document.getElementById('fileInput');
const fileName = document.getElementById('fileName');

// Column selection section
const columnSection = document.getElementById('columnSection');
const columnInput = document.getElementById('columnInput');
const confirmBtn = document.getElementById('confirmBtn');
const inputHint = document.getElementById('inputHint');
const inputTypeRadios = document.querySelectorAll('input[name="inputType"]');
const hasHeaderCheckbox = document.getElementById('hasHeaderCheckbox');
const colNameList = document.getElementById('colNameList');
const columnListWrapper = document.getElementById('columnListWrapper');

// Feedback and status elements
const loadingIndicator = document.getElementById('loadingIndicator');
const errorMessage = document.getElementById('errorMessage');

// Results section
const resultSection = document.getElementById('resultSection');
const resultTable = document.getElementById('resultTable');
const downloadBtn = document.getElementById('downloadBtn');

// ============================================
// EVENT LISTENERS SETUP
// ============================================

/**
 * Input type radio buttons change handler
 * Toggles between "Column Name" and "Column Index" input modes
 * Updates UI hints and placeholders accordingly
 */
inputTypeRadios.forEach((radio) => {
  radio.addEventListener('change', (e) => {
    if (e.target.value === 'name') {
      // Column Name mode - user enters exact header text
      columnInput.placeholder = 'Enter column name (e.g., Email)';
      inputHint.textContent = 'Enter the exact column header name';
    } else {
      // Column Index mode - user enters position number
      columnInput.placeholder = 'Enter column number (e.g., 7 for 7th column)';
      inputHint.textContent = 'Enter the column position (1 = first column, 2 = second column, etc.)';
    }
    columnInput.value = ''; // Clear previous input
  });
});

/**
 * Header checkbox change handler
 * Controls whether first row is treated as headers or data
 *
 * Logic:
 * - When unchecked: Forces "Column Index" mode (can't use names without headers)
 * - When checked: Re-enables "Column Name" mode
 * - Reprocesses file if already loaded to update column display
 */
hasHeaderCheckbox.addEventListener('change', (e) => {
  console.log('Checkbox changed to:', e.target.checked);

  if (!e.target.checked) {
    // No headers mode - force Column Index selection
    document.querySelector('input[name="inputType"][value="index"]').checked = true;
    document.querySelector('input[name="inputType"][value="name"]').disabled = true;
    columnInput.placeholder = 'Enter column number (e.g., 1 for first column)';
    inputHint.textContent = 'Enter the column position (1 = first column, 2 = second column, etc.)';
  } else {
    // Has headers mode - enable Column Name selection
    document.querySelector('input[name="inputType"][value="name"]').disabled = false;
    document.querySelector('input[name="inputType"][value="name"]').checked = true;
    columnInput.placeholder = 'Enter column name (e.g., Email)';
    inputHint.textContent = 'Enter the exact column header name';
  }

  columnInput.value = ''; // Clear input when mode changes

  // Reprocess file with new header interpretation
  if (currentFile) {
    console.log('Reprocessing file...');
    processFile(currentFile);
  } else {
    console.log('No file loaded yet');
  }
});

/**
 * Primary event listeners for core functionality
 */
fileInput.addEventListener('change', handleFileSelect);
confirmBtn.addEventListener('click', processData);
downloadBtn.addEventListener('click', downloadExcel);

// ============================================
// FILE HANDLING FUNCTIONS
// ============================================

/**
 * Handles file selection from input element
 *
 * @param {Event} e - File input change event
 *
 * Workflow:
 * 1. Extract file from event
 * 2. Store reference for potential reprocessing
 * 3. Update UI to show selected filename
 * 4. Trigger file processing
 */
function handleFileSelect(e) {
  const file = e.target.files[0];
  if (!file) return; // User cancelled file selection

  currentFile = file;
  fileName.textContent = `Selected: ${file.name}`;
  processFile(file);
}

/**
 * Processes uploaded Excel/CSV file
 *
 * @param {File} file - File object to process
 *
 * Process flow:
 * 1. Read file as ArrayBuffer using FileReader API
 * 2. Parse using SheetJS (XLSX) library
 * 3. Extract first sheet and convert to JSON array
 * 4. Determine headers based on checkbox setting
 * 5. Display column list for user selection
 * 6. Show success feedback
 *
 * Configuration:
 * - raw: false - Parse dates and numbers properly
 * - codepage: 65001 - UTF-8 encoding for international characters
 * - defval: '' - Use empty string for blank cells
 */
function processFile(file) {
  const reader = new FileReader();

  reader.onload = function (e) {
    try {
      // Convert file to Uint8Array for SheetJS
      const data = new Uint8Array(e.target.result);

      // Parse Excel/CSV file
      const workbook = XLSX.read(data, {
        type: 'array',
        raw: false, // Parse values properly
        codepage: 65001, // UTF-8 encoding
      });

      // Get first sheet
      const firstSheet = workbook.Sheets[workbook.SheetNames[0]];

      // Convert sheet to 2D array
      const jsonData = XLSX.utils.sheet_to_json(firstSheet, {
        header: 1, // Return array of arrays (not objects)
        defval: '', // Default value for empty cells
        raw: false, // Format values as strings
      });

      // Validate data exists
      if (jsonData.length < 1) {
        showError('File is empty or contains no data');
        return;
      }

      const hasHeader = hasHeaderCheckbox.checked;

      if (hasHeader) {
        // First row contains headers
        if (jsonData.length < 2) {
          showError('File must contain at least a header row and one data row');
          return;
        }

        // Extract headers from first row
        headers = jsonData[0].map((h) => String(h));

        // Data starts from second row
        originalData = jsonData.slice(1);

        showSuccess(`✓ Loaded with headers: ${headers.join(', ')}`);
      } else {
        // No headers - generate generic column names
        const numColumns = jsonData[0]?.length || 0;
        headers = Array.from({ length: numColumns }, (_, i) => `Column ${i + 1}`);

        // All rows are data
        originalData = jsonData;

        showSuccess(`✓ Loaded without headers (${originalData.length} rows, ${numColumns} columns)`);
      }

      // Display clickable column badges
      displayColumnList();

      // Show column selection section
      columnSection.style.display = 'block';
      resultSection.style.display = 'none';

      // Debug logging
      console.log('Has Header:', hasHeader);
      console.log('Headers:', headers);
      console.log('Data rows:', originalData.length);
      console.log('First data row:', originalData[0]);
    } catch (error) {
      showError('Error reading file: ' + error.message);
    }
  };

  // Read file as binary array
  reader.readAsArrayBuffer(file);
}

// ============================================
// UI DISPLAY FUNCTIONS
// ============================================

/**
 * Displays clickable column badges for easy selection
 *
 * Purpose: Provides visual, interactive way for users to select columns
 * instead of typing names or numbers manually
 *
 * Display Logic:
 * - With headers: Shows actual column names (e.g., "Email", "Name")
 * - Without headers: Shows numbered labels (e.g., "1. Column 1", "2. Column 2")
 *
 * Interaction:
 * - Clicking a badge auto-fills the column input field
 * - Respects current input mode (name vs index)
 */
function displayColumnList() {
  if (headers.length === 0) {
    columnListWrapper.style.display = 'none';
    return;
  }

  let html = '';

  // Build badge HTML for each column
  headers.forEach((header, index) => {
    const displayText = hasHeaderCheckbox.checked
      ? header // Show header name
      : `${index + 1}. ${header}`; // Show index + name

    html += `<span class="column-badge clickable" 
                      data-column="${header}" 
                      data-index="${index + 1}">
                    ${escapeHtml(displayText)}
                 </span>`;
  });

  colNameList.innerHTML = html;
  columnListWrapper.style.display = 'block';

  // Attach click handlers to all badges
  document.querySelectorAll('.column-badge').forEach((badge) => {
    badge.addEventListener('click', (e) => {
      const columnName = e.target.getAttribute('data-column');
      const columnIndex = e.target.getAttribute('data-index');

      // Fill input based on current mode
      if (hasHeaderCheckbox.checked && document.querySelector('input[name="inputType"][value="name"]').checked) {
        columnInput.value = columnName; // Use column name
      } else {
        columnInput.value = columnIndex; // Use column index
      }

      columnInput.focus(); // Focus input for immediate editing if needed
    });
  });
}

// ============================================
// DATA PROCESSING FUNCTIONS
// ============================================

/**
 * Main data processing trigger function
 * Validates user input and initiates column splitting
 *
 * Process Flow:
 * 1. Validate input exists
 * 2. Determine target column index (from name or number)
 * 3. Show loading indicator
 * 4. Execute split operation (async-like via setTimeout)
 * 5. Display results
 *
 * Error Handling:
 * - Empty input validation
 * - Column name not found
 * - Invalid column number (out of range)
 */
function processData() {
  const inputType = document.querySelector('input[name="inputType"]:checked').value;
  const inputValue = columnInput.value.trim();

  // Validate input exists
  if (!inputValue) {
    showError('Please enter a column ' + (inputType === 'name' ? 'name' : 'number'));
    return;
  }

  let columnIndex = -1;

  if (inputType === 'name') {
    // Find column by name (case-insensitive)
    columnIndex = headers.findIndex((h) => String(h).toLowerCase() === inputValue.toLowerCase());

    if (columnIndex === -1) {
      showError(`Column "${inputValue}" not found. Available columns: ${headers.join(', ')}`);
      return;
    }
  } else if (inputType === 'index') {
    // Parse and validate column number
    const colNum = parseInt(inputValue);

    if (isNaN(colNum) || colNum < 1 || colNum > headers.length) {
      showError(`Invalid column number. Please enter a number between 1 and ${headers.length}`);
      return;
    }

    // Convert to 0-based index
    columnIndex = colNum - 1;
  }

  // Show loading state
  loadingIndicator.style.display = 'block';
  hideError();

  // Use setTimeout to allow UI to update before heavy processing
  setTimeout(() => {
    try {
      processedData = splitColumn(columnIndex);
      displayResults();
      loadingIndicator.style.display = 'none';
      resultSection.style.display = 'block';
    } catch (error) {
      loadingIndicator.style.display = 'none';
      showError('Error processing data: ' + error.message);
    }
  }, 100); // 100ms delay allows loading spinner to render
}

/**
 * Core splitting algorithm (MULTI-SEPARATOR VERSION)
 * Splits cell values containing multiple items into separate rows
 * NOW SUPPORTS MULTIPLE DIFFERENT SEPARATORS IN THE SAME CELL!
 *
 * @param {number} columnIndex - Zero-based index of column to split
 * @returns {Array} Array of objects with {data: row[], isSplit: boolean}
 *
 * NEW APPROACH:
 * - Replaces ALL separators with a consistent delimiter (||)
 * - Then splits on that single delimiter
 * - This allows mixing separators like "email1 ; email2 , email3 / email4"
 */
function splitColumn(columnIndex) {
  const results = [];

  // Define all supported separators (Western + Chinese)
  const separators = [';', '；', ',', '，', '/', '／', '&', '\n'];

  // Terms that should NOT be split despite containing separators
  const specialTerms = ['n/a', 'N/A', 'n\\a', 'N\\A'];

  originalData.forEach((row, rowIndex) => {
    // Ensure row has enough columns (pad with empty strings)
    const fullRow = [...row];
    while (fullRow.length < headers.length) {
      fullRow.push('');
    }

    const cellValue = String(fullRow[columnIndex] || '').trim();

    // Handle empty cells - keep row as-is
    if (!cellValue) {
      results.push({
        data: fullRow,
        isSplit: false,
      });
      return;
    }

    // Check if cell contains a special term (case-insensitive)
    const isSpecialTerm = specialTerms.some((term) => cellValue.toLowerCase() === term.toLowerCase());

    if (isSpecialTerm) {
      // Don't split special terms (e.g., "N/A" contains "/" but shouldn't split)
      results.push({
        data: fullRow,
        isSplit: false,
      });
      return;
    }

    // NEW APPROACH: Replace ALL separators with a consistent delimiter
    let normalizedValue = cellValue;
    const TEMP_DELIMITER = '||SPLIT_HERE||'; // Use unlikely string to avoid conflicts

    // Replace each separator type with our temporary delimiter
    separators.forEach((sep) => {
      // Use regex with global flag to replace ALL occurrences
      const regex = new RegExp(escapeRegex(sep), 'g');
      normalizedValue = normalizedValue.replace(regex, TEMP_DELIMITER);
    });

    // Now split on the temporary delimiter
    let splitValues = normalizedValue
      .split(TEMP_DELIMITER)
      .map((v) => v.trim()) // Trim whitespace from each value
      .filter((v) => v); // Remove empty values

    // Debug logging
    if (splitValues.length > 1) {
      console.log(`Row ${rowIndex}: "${cellValue}" → ${splitValues.length} values:`, splitValues);
    }

    // Create result rows
    if (splitValues.length > 1) {
      // Multiple values found - create separate row for each
      splitValues.forEach((value) => {
        const newRow = [...fullRow];
        newRow[columnIndex] = value;
        results.push({
          data: newRow,
          isSplit: true, // Mark as split for highlighting
        });
      });
    } else {
      // Single value - keep original row
      results.push({
        data: fullRow,
        isSplit: false,
      });
    }
  });

  console.log(`Total processed: ${results.length} rows from ${originalData.length} original rows`);
  return results;
}

/**
 * Helper function to escape special regex characters
 * Needed because some separators like "/" and "&" have special meaning in regex
 *
 * @param {string} str - String to escape
 * @returns {string} Regex-safe string
 */
function escapeRegex(str) {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * Renders processed data in HTML table
 *
 * Features:
 * - Builds complete table HTML from scratch
 * - Applies highlighting to split rows
 * - Escapes HTML to prevent XSS attacks
 * - Handles missing cell values
 *
 * Table Structure:
 * - thead: Column headers
 * - tbody: Data rows with conditional highlighting
 */
function displayResults() {
  let tableHTML = '<thead><tr>';

  // Build header row
  headers.forEach((header) => {
    tableHTML += `<th>${escapeHtml(String(header))}</th>`;
  });
  tableHTML += '</tr></thead><tbody>';

  // Build data rows
  processedData.forEach((item) => {
    // Apply CSS class for highlighted rows
    const rowClass = item.isSplit ? 'highlighted' : '';
    tableHTML += `<tr class="${rowClass}">`;

    // Build cells for this row
    for (let i = 0; i < headers.length; i++) {
      const cellValue = item.data[i] !== undefined ? item.data[i] : '';
      tableHTML += `<td>${escapeHtml(String(cellValue))}</td>`;
    }

    tableHTML += '</tr>';
  });

  tableHTML += '</tbody>';
  resultTable.innerHTML = tableHTML;
}

// ============================================
// EXCEL EXPORT FUNCTION
// ============================================

/**
 * Exports processed data to Excel file with formatting
 *
 * Features:
 * - Adds "[SPLIT ROW]" indicator column
 * - Applies yellow background to split rows
 * - Auto-sizes columns based on content
 * - Uses SheetJS for Excel generation
 *
 * File Structure:
 * - Column 1: Split indicator (✓ SPLIT or empty)
 * - Remaining columns: Original data
 *
 * Formatting:
 * - Split rows: Yellow background (#FFF9E6)
 * - Column widths: Auto-calculated (min 10, max 50 chars)
 */
function downloadExcel() {
  // Add split indicator column to headers
  const excelHeaders = ['[SPLIT ROW]', ...headers];

  // Build data rows with split indicator
  const excelData = processedData.map((item) => {
    const fullRow = [...item.data];

    // Pad row if needed
    while (fullRow.length < headers.length) {
      fullRow.push('');
    }

    // Prepend split indicator
    return [item.isSplit ? '✓ SPLIT' : '', ...fullRow];
  });

  // Create worksheet from array of arrays
  const ws = XLSX.utils.aoa_to_sheet([excelHeaders, ...excelData]);

  // Calculate column widths based on content
  const colWidths = excelHeaders.map((_, i) => {
    const maxLength = Math.max(String(excelHeaders[i]).length, ...excelData.map((row) => String(row[i] || '')).map((v) => v.length));
    // Min width: 10, Max width: 50, Add padding: +2
    return { wch: Math.min(Math.max(maxLength + 2, 10), 50) };
  });
  ws['!cols'] = colWidths;

  // Apply yellow background to split rows
  const range = XLSX.utils.decode_range(ws['!ref']);
  for (let R = 1; R <= range.e.r; R++) {
    // Start at 1 to skip header
    if (processedData[R - 1]?.isSplit) {
      // Apply style to all cells in this row
      for (let C = range.s.c; C <= range.e.c; C++) {
        const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
        if (!ws[cellAddress]) continue;

        // Initialize style object if needed
        if (!ws[cellAddress].s) ws[cellAddress].s = {};

        // Set yellow background
        ws[cellAddress].s.fill = {
          fgColor: { rgb: 'FFF9E6' },
        };
      }
    }
  }

  // Create workbook and add worksheet
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Split Data');

  // Trigger download
  XLSX.writeFile(wb, 'split_data.xlsx');
}

// ============================================
// UTILITY FUNCTIONS
// ============================================

/**
 * Escapes HTML special characters to prevent XSS attacks
 *
 * @param {string} text - Text to escape
 * @returns {string} HTML-safe text
 *
 * Method: Uses browser's built-in escaping via textContent
 * Safer than manual replacement of &, <, >, ", '
 */
function escapeHtml(text) {
  const div = document.createElement('div');
  div.textContent = text; // Browser auto-escapes
  return div.innerHTML; // Return escaped HTML
}

/**
 * Displays error message to user
 *
 * @param {string} message - Error message to display
 *
 * Used for:
 * - File validation errors
 * - Column selection errors
 * - Processing errors
 */
function showError(message) {
  errorMessage.textContent = message;
  errorMessage.style.display = 'block';
}

/**
 * Hides error message
 * Called before showing success or when error is resolved
 */
function hideError() {
  errorMessage.style.display = 'none';
}

/**
 * Shows temporary success message
 *
 * @param {string} message - Success message to display
 *
 * Behavior:
 * - Displays in green for 3 seconds
 * - Uses error message element with custom styling
 * - Auto-hides after timeout
 * - Resets styling for future error messages
 */
function showSuccess(message) {
  hideError();

  // Temporarily repurpose error message element for success
  errorMessage.textContent = message;
  errorMessage.style.background = '#d4edda';
  errorMessage.style.color = '#155724';
  errorMessage.style.borderLeft = '4px solid #28a745';
  errorMessage.style.display = 'block';

  // Reset after 3 seconds
  setTimeout(() => {
    errorMessage.style.background = '';
    errorMessage.style.color = '';
    errorMessage.style.borderLeft = '';
    hideError();
  }, 3000);
}

/**
 * ============================================
 * END OF FILE
 * ============================================
 *
 * Dependencies:
 * - SheetJS (XLSX) library must be loaded via CDN
 * - Requires modern browser with ES6+ support
 *
 * Browser Compatibility:
 * - Chrome/Edge: ✓
 * - Firefox: ✓
 * - Safari: ✓
 * - IE11: ✗ (requires transpilation)
 */
