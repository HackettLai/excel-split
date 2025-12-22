let originalData = null;
let processedData = null;
let headers = null;
let highlightedRows = new Set(); // Track which rows were split/modified

const fileInput = document.getElementById('fileInput');
const fileName = document.getElementById('fileName');
const columnSection = document.getElementById('columnSection');
const columnInput = document.getElementById('columnInput');
const confirmBtn = document.getElementById('confirmBtn');
const loadingIndicator = document.getElementById('loadingIndicator');
const errorMessage = document.getElementById('errorMessage');
const resultSection = document.getElementById('resultSection');
const resultTable = document.getElementById('resultTable');
const downloadBtn = document.getElementById('downloadBtn');
const inputHint = document.getElementById('inputHint');

// Radio buttons
const radioButtons = document.querySelectorAll('input[name="inputType"]');

// File upload handler
fileInput.addEventListener('change', handleFileUpload);

// Radio button change handler
radioButtons.forEach(radio => {
    radio.addEventListener('change', updateInputPlaceholder);
});

function updateInputPlaceholder() {
    const selectedType = document.querySelector('input[name="inputType"]:checked').value;
    
    if (selectedType === 'name') {
        columnInput.placeholder = 'Enter column name (e.g., Email)';
        inputHint.textContent = 'Enter the exact column header name';
    } else {
        columnInput.placeholder = 'Enter column number (e.g., 3)';
        inputHint.textContent = 'Enter the column position (1 = first column, 2 = second column, etc.)';
    }
    
    columnInput.value = '';
    columnInput.focus();
}

function handleFileUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    fileName.textContent = `Selected: ${file.name}`;
    hideError();
    resultSection.style.display = 'none';

    const reader = new FileReader();
    
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });

            if (jsonData.length < 1) {
                showError('⚠️ No data found in the file!');
                return;
            }

            originalData = jsonData;
            headers = jsonData[0];
            columnSection.style.display = 'block';
            columnInput.focus();
        } catch (error) {
            showError('❌ Error reading file: ' + error.message);
        }
    };

    if (file.name.endsWith('.csv')) {
        reader.readAsText(file);
        reader.onload = function(e) {
            try {
                const csvData = e.target.result;
                const jsonData = parseCSV(csvData);
                
                if (jsonData.length < 1) {
                    showError('⚠️ No data found in the file!');
                    return;
                }

                originalData = jsonData;
                headers = jsonData[0];
                columnSection.style.display = 'block';
                columnInput.focus();
            } catch (error) {
                showError('❌ Error reading CSV: ' + error.message);
            }
        };
    } else {
        reader.readAsArrayBuffer(file);
    }
}

function parseCSV(csv) {
    const lines = csv.split('\n');
    return lines.map(line => {
        const values = [];
        let current = '';
        let inQuotes = false;
        
        for (let i = 0; i < line.length; i++) {
            const char = line[i];
            
            if (char === '"') {
                inQuotes = !inQuotes;
            } else if (char === ',' && !inQuotes) {
                values.push(current.trim());
                current = '';
            } else {
                current += char;
            }
        }
        values.push(current.trim());
        return values;
    }).filter(row => row.some(cell => cell));
}

// Confirm button handler
confirmBtn.addEventListener('click', processSplit);

// Allow Enter key to trigger split
columnInput.addEventListener('keypress', function(e) {
    if (e.key === 'Enter') {
        processSplit();
    }
});

function processSplit() {
    const userInput = columnInput.value.trim();
    const inputType = document.querySelector('input[name="inputType"]:checked').value;
    
    if (!userInput) {
        showError(`❌ Please enter a column ${inputType === 'name' ? 'name' : 'index'}.`);
        return;
    }

    if (!originalData) {
        showError('❌ Please upload a file first.');
        return;
    }

    hideError();
    loadingIndicator.style.display = 'block';
    resultSection.style.display = 'none';

    // Process in next tick to allow UI to update
    setTimeout(() => {
        try {
            let emailColIndex = -1;
            
            if (inputType === 'index') {
                // User selected column index
                const columnIndex = parseInt(userInput);
                
                if (isNaN(columnIndex)) {
                    showError('❌ Please enter a valid number for column index.');
                    loadingIndicator.style.display = 'none';
                    return;
                }
                
                if (columnIndex < 1) {
                    showError('❌ Column index must be 1 or greater.');
                    loadingIndicator.style.display = 'none';
                    return;
                }
                
                emailColIndex = columnIndex - 1; // Convert to 0-based index
                
                if (emailColIndex >= originalData[0].length) {
                    showError(`❌ Column index ${columnIndex} is out of range. Maximum is ${originalData[0].length}.`);
                    loadingIndicator.style.display = 'none';
                    return;
                }
            } else {
                // User selected column name
                const userColumnName = userInput.toLowerCase();
                emailColIndex = headers.findIndex(
                    h => String(h).trim().toLowerCase() === userColumnName
                );

                if (emailColIndex === -1) {
                    showError(`❌ Column "${userInput}" not found. Please check the spelling.`);
                    loadingIndicator.style.display = 'none';
                    return;
                }
            }

            let newData = [headers];
            highlightedRows = new Set();
            let currentRowIndex = 1; // Track position in new data array

            // Process each data row
            for (let i = 1; i < originalData.length; i++) {
                const row = originalData[i];
                const emailCell = row[emailColIndex];

                // Skip or keep if N/A or blank
                if (!emailCell || String(emailCell).trim().toUpperCase() === 'N/A') {
                    newData.push(row);
                    currentRowIndex++;
                    continue;
                }

                // Split emails by multiple delimiters
                const emails = String(emailCell)
                    .split(/[\n;,/\\]+/g)
                    .map(e => e.trim())
                    .filter(e => e && e.toUpperCase() !== 'N/A');

                // If no valid email after cleaning, keep as is
                if (emails.length === 0) {
                    newData.push(row);
                    currentRowIndex++;
                    continue;
                }

                // If only one email and it matches original, no split occurred
                if (emails.length === 1 && emails[0] === String(emailCell).trim()) {
                    newData.push(row);
                    currentRowIndex++;
                    continue;
                }

                // Create separate rows for each email and mark them as highlighted
                emails.forEach(email => {
                    const newRow = [...row];
                    newRow[emailColIndex] = email;
                    newData.push(newRow);
                    highlightedRows.add(currentRowIndex);
                    currentRowIndex++;
                });
            }

            processedData = newData;
            displayResults(newData);
            loadingIndicator.style.display = 'none';
            resultSection.style.display = 'block';
        } catch (error) {
            showError('❌ Error processing data: ' + error.message);
            loadingIndicator.style.display = 'none';
        }
    }, 100);
}

function displayResults(data) {
    let tableHTML = '<thead><tr>';
    
    // Add headers
    data[0].forEach(header => {
        tableHTML += `<th>${escapeHtml(String(header))}</th>`;
    });
    tableHTML += '</tr></thead><tbody>';

    // Add data rows
    for (let i = 1; i < data.length; i++) {
        const isHighlighted = highlightedRows.has(i);
        tableHTML += `<tr${isHighlighted ? ' class="highlighted"' : ''}>`;
        data[i].forEach(cell => {
            tableHTML += `<td>${escapeHtml(String(cell || ''))}</td>`;
        });
        tableHTML += '</tr>';
    }
    tableHTML += '</tbody>';

    resultTable.innerHTML = tableHTML;
}

function escapeHtml(text) {
    const map = {
        '&': '&amp;',
        '<': '&lt;',
        '>': '&gt;',
        '"': '&quot;',
        "'": '&#039;'
    };
    return text.replace(/[&<>"']/g, m => map[m]);
}

// Download button handler
downloadBtn.addEventListener('click', downloadExcel);

function downloadExcel() {
    if (!processedData) {
        showError('❌ No data to download.');
        return;
    }

    // Create a copy of processed data with marker column
    const excelData = processedData.map((row, index) => {
        if (index === 0) {
            // Add header for marker column
            return ['[SPLIT ROW]', ...row];
        } else {
            // Add marker for highlighted rows
            const marker = highlightedRows.has(index) ? '✓ SPLIT' : '';
            return [marker, ...row];
        }
    });

    // Create worksheet
    const ws = XLSX.utils.aoa_to_sheet(excelData);
    
    // Set column widths
    const range = XLSX.utils.decode_range(ws['!ref']);
    const colWidths = [];
    
    for (let C = range.s.c; C <= range.e.c; ++C) {
        let maxWidth = 10;
        for (let R = range.s.r; R <= range.e.r; ++R) {
            const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
            if (ws[cellAddress] && ws[cellAddress].v) {
                const cellLength = String(ws[cellAddress].v).length;
                maxWidth = Math.max(maxWidth, cellLength);
            }
        }
        colWidths.push({ wch: Math.min(maxWidth + 2, 50) });
    }
    ws['!cols'] = colWidths;
    
    // Create workbook
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Data_Split');
    XLSX.writeFile(wb, 'Data_Split.xlsx');
}

function showError(message) {
    errorMessage.textContent = message;
    errorMessage.style.display = 'block';
}

function hideError() {
    errorMessage.style.display = 'none';
}