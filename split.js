let originalData = [];
let processedData = [];
let headers = [];
let currentFile = null;

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
const inputTypeRadios = document.querySelectorAll('input[name="inputType"]');
const hasHeaderCheckbox = document.getElementById('hasHeaderCheckbox');
const colNameList = document.getElementById('colNameList');
const columnListWrapper = document.getElementById('columnListWrapper');

// Update placeholder and hint based on input type
inputTypeRadios.forEach(radio => {
    radio.addEventListener('change', (e) => {
        if (e.target.value === 'name') {
            columnInput.placeholder = 'Enter column name (e.g., Email)';
            inputHint.textContent = 'Enter the exact column header name';
        } else {
            columnInput.placeholder = 'Enter column number (e.g., 7 for 7th column)';
            inputHint.textContent = 'Enter the column position (1 = first column, 2 = second column, etc.)';
        }
        columnInput.value = '';
    });
});

// Update UI when header checkbox changes
hasHeaderCheckbox.addEventListener('change', (e) => {
    console.log('Checkbox changed to:', e.target.checked);

    if (!e.target.checked) {
        // If no headers, force Column Index mode
        document.querySelector('input[name="inputType"][value="index"]').checked = true;
        document.querySelector('input[name="inputType"][value="name"]').disabled = true;
        columnInput.placeholder = 'Enter column number (e.g., 1 for first column)';
        inputHint.textContent = 'Enter the column position (1 = first column, 2 = second column, etc.)';
    } else {
        // Re-enable Column Name mode
        document.querySelector('input[name="inputType"][value="name"]').disabled = false;
        document.querySelector('input[name="inputType"][value="name"]').checked = true;
        columnInput.placeholder = 'Enter column name (e.g., Email)';
        inputHint.textContent = 'Enter the exact column header name';
    }
    columnInput.value = '';

    // Reload the file with new header setting
    if (currentFile) {
        console.log('Reprocessing file...');
        processFile(currentFile);
    } else {
        console.log('No file loaded yet');
    }
});

fileInput.addEventListener('change', handleFileSelect);
confirmBtn.addEventListener('click', processData);
downloadBtn.addEventListener('click', downloadExcel);

function handleFileSelect(e) {
    const file = e.target.files[0];
    if (!file) return;

    currentFile = file;
    fileName.textContent = `Selected: ${file.name}`;
    processFile(file);
}

function processFile(file) {
    const reader = new FileReader();
    reader.onload = function (e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, {
                type: 'array',
                raw: false,
                codepage: 65001
            });
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            const jsonData = XLSX.utils.sheet_to_json(firstSheet, {
                header: 1,
                defval: '',
                raw: false
            });

            if (jsonData.length < 1) {
                showError('File is empty or contains no data');
                return;
            }

            const hasHeader = hasHeaderCheckbox.checked;

            if (hasHeader) {
                if (jsonData.length < 2) {
                    showError('File must contain at least a header row and one data row');
                    return;
                }
                headers = jsonData[0].map(h => String(h));
                originalData = jsonData.slice(1);

                // Show feedback
                showSuccess(`✓ Loaded with headers: ${headers.join(', ')}`);
            } else {
                // Generate generic column names
                const numColumns = jsonData[0]?.length || 0;
                headers = Array.from({ length: numColumns }, (_, i) => `Column ${i + 1}`);
                originalData = jsonData;

                // Show feedback
                showSuccess(`✓ Loaded without headers (${originalData.length} rows, ${numColumns} columns)`);
            }

            // Display column list
            displayColumnList();

            columnSection.style.display = 'block';
            resultSection.style.display = 'none';

            console.log('Has Header:', hasHeader);
            console.log('Headers:', headers);
            console.log('Data rows:', originalData.length);
            console.log('First data row:', originalData[0]);
        } catch (error) {
            showError('Error reading file: ' + error.message);
        }
    };
    reader.readAsArrayBuffer(file);
}

// Add this new function
function displayColumnList() {
    if (headers.length === 0) {
        columnListWrapper.style.display = 'none';
        return;
    }

    let html = '';
    headers.forEach((header, index) => {
        const displayText = hasHeaderCheckbox.checked ? header : `${index + 1}. ${header}`;
        html += `<span class="column-badge clickable" data-column="${header}" data-index="${index + 1}">${escapeHtml(displayText)}</span>`;
    });

    colNameList.innerHTML = html;
    columnListWrapper.style.display = 'block';

    // Add click listeners to badges
    document.querySelectorAll('.column-badge').forEach(badge => {
        badge.addEventListener('click', (e) => {
            const columnName = e.target.getAttribute('data-column');
            const columnIndex = e.target.getAttribute('data-index');

            if (hasHeaderCheckbox.checked && document.querySelector('input[name="inputType"][value="name"]').checked) {
                columnInput.value = columnName;
            } else {
                columnInput.value = columnIndex;
            }

            columnInput.focus();
        });
    });
}


function processData() {
    const inputType = document.querySelector('input[name="inputType"]:checked').value;
    const inputValue = columnInput.value.trim();

    if (!inputValue) {
        showError('Please enter a column ' + (inputType === 'name' ? 'name' : 'number'));
        return;
    }

    let columnIndex = -1;

    if (inputType === 'name') {
        columnIndex = headers.findIndex(h => String(h).toLowerCase() === inputValue.toLowerCase());
        if (columnIndex === -1) {
            showError(`Column "${inputValue}" not found. Available columns: ${headers.join(', ')}`);
            return;
        }
    } else if (inputType === 'index') {
        const colNum = parseInt(inputValue);
        if (isNaN(colNum) || colNum < 1 || colNum > headers.length) {
            showError(`Invalid column number. Please enter a number between 1 and ${headers.length}`);
            return;
        }
        columnIndex = colNum - 1;
    }

    loadingIndicator.style.display = 'block';
    hideError();

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
    }, 100);
}

function splitColumn(columnIndex) {
    const results = [];
    const separators = [',', ';', '\n', '，', '；'];

    originalData.forEach((row) => {
        const fullRow = [...row];
        while (fullRow.length < headers.length) {
            fullRow.push('');
        }

        const cellValue = String(fullRow[columnIndex] || '').trim();

        if (!cellValue) {
            results.push({
                data: fullRow,
                isSplit: false
            });
            return;
        }

        let splitValues = [cellValue];

        for (const sep of separators) {
            if (cellValue.includes(sep)) {
                splitValues = cellValue.split(sep).map(v => v.trim()).filter(v => v);
                break;
            }
        }

        if (splitValues.length > 1) {
            splitValues.forEach(value => {
                const newRow = [...fullRow];
                newRow[columnIndex] = value;
                results.push({
                    data: newRow,
                    isSplit: true
                });
            });
        } else {
            results.push({
                data: fullRow,
                isSplit: false
            });
        }
    });

    return results;
}

function displayResults() {
    let tableHTML = '<thead><tr>';
    headers.forEach(header => {
        tableHTML += `<th>${escapeHtml(String(header))}</th>`;
    });
    tableHTML += '</tr></thead><tbody>';

    processedData.forEach(item => {
        const rowClass = item.isSplit ? 'highlighted' : '';
        tableHTML += `<tr class="${rowClass}">`;

        for (let i = 0; i < headers.length; i++) {
            const cellValue = item.data[i] !== undefined ? item.data[i] : '';
            tableHTML += `<td>${escapeHtml(String(cellValue))}</td>`;
        }

        tableHTML += '</tr>';
    });

    tableHTML += '</tbody>';
    resultTable.innerHTML = tableHTML;
}

function downloadExcel() {
    const excelHeaders = ['[SPLIT ROW]', ...headers];

    const excelData = processedData.map(item => {
        const fullRow = [...item.data];
        while (fullRow.length < headers.length) {
            fullRow.push('');
        }
        return [item.isSplit ? '✓ SPLIT' : '', ...fullRow];
    });

    const ws = XLSX.utils.aoa_to_sheet([excelHeaders, ...excelData]);

    const colWidths = excelHeaders.map((_, i) => {
        const maxLength = Math.max(
            String(excelHeaders[i]).length,
            ...excelData.map(row => String(row[i] || '')).map(v => v.length)
        );
        return { wch: Math.min(Math.max(maxLength + 2, 10), 50) };
    });
    ws['!cols'] = colWidths;

    const range = XLSX.utils.decode_range(ws['!ref']);
    for (let R = 1; R <= range.e.r; R++) {
        if (processedData[R - 1]?.isSplit) {
            for (let C = range.s.c; C <= range.e.c; C++) {
                const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
                if (!ws[cellAddress]) continue;

                if (!ws[cellAddress].s) ws[cellAddress].s = {};
                ws[cellAddress].s.fill = {
                    fgColor: { rgb: "FFF9E6" }
                };
            }
        }
    }

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Split Data');
    XLSX.writeFile(wb, 'split_data.xlsx');
}

function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

function showError(message) {
    errorMessage.textContent = message;
    errorMessage.style.display = 'block';
}

function hideError() {
    errorMessage.style.display = 'none';
}

function showSuccess(message) {
    hideError();
    // Temporarily show success in the error box with green styling
    errorMessage.textContent = message;
    errorMessage.style.background = '#d4edda';
    errorMessage.style.color = '#155724';
    errorMessage.style.borderLeft = '4px solid #28a745';
    errorMessage.style.display = 'block';

    setTimeout(() => {
        errorMessage.style.background = '';
        errorMessage.style.color = '';
        errorMessage.style.borderLeft = '';
        hideError();
    }, 3000);
}