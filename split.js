let originalData = [];
let processedData = [];
let headers = [];

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

fileInput.addEventListener('change', handleFileSelect);
confirmBtn.addEventListener('click', processData);
downloadBtn.addEventListener('click', downloadExcel);

function handleFileSelect(e) {
    const file = e.target.files[0];
    if (!file) return;

    fileName.textContent = `Selected: ${file.name}`;
    
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1, defval: '' });
            
            if (jsonData.length < 2) {
                showError('File must contain at least a header row and one data row');
                return;
            }

            headers = jsonData[0];
            originalData = jsonData.slice(1);
            
            columnSection.style.display = 'block';
            resultSection.style.display = 'none';
            hideError();
        } catch (error) {
            showError('Error reading file: ' + error.message);
        }
    };
    reader.readAsArrayBuffer(file);
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
        columnIndex = headers.findIndex(h => h.toLowerCase() === inputValue.toLowerCase());
        if (columnIndex === -1) {
            showError(`Column "${inputValue}" not found. Available columns: ${headers.join(', ')}`);
            return;
        }
    } else {
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
    const separators = [',', ';', '\n'];

    originalData.forEach((row) => {
        // Ensure row has all columns (fill missing columns with empty strings)
        const fullRow = [...row];
        while (fullRow.length < headers.length) {
            fullRow.push('');
        }

        const cellValue = String(fullRow[columnIndex] || '').trim();
        
        if (!cellValue) {
            // If cell is empty, keep the row as is
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
        tableHTML += `<th>${escapeHtml(header)}</th>`;
    });
    tableHTML += '</tr></thead><tbody>';

    processedData.forEach(item => {
        const rowClass = item.isSplit ? 'highlighted' : '';
        tableHTML += `<tr class="${rowClass}">`;
        
        // Ensure we display all columns, even if empty
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
    // Add [SPLIT ROW] column as FIRST column
    const excelHeaders = ['[SPLIT ROW]', ...headers];
    
    // Prepare data with split status as first column
    const excelData = processedData.map(item => {
        // Ensure all columns are present
        const fullRow = [...item.data];
        while (fullRow.length < headers.length) {
            fullRow.push('');
        }
        return [item.isSplit ? '✓ SPLIT' : '', ...fullRow];
    });

    // Create worksheet
    const ws = XLSX.utils.aoa_to_sheet([excelHeaders, ...excelData]);

    // Set column widths
    const colWidths = excelHeaders.map((_, i) => {
        const maxLength = Math.max(
            String(excelHeaders[i]).length,
            ...excelData.map(row => String(row[i] || '')).map(v => v.length)
        );
        return { wch: Math.min(Math.max(maxLength + 2, 10), 50) };
    });
    ws['!cols'] = colWidths;

    // Apply yellow background to split rows (shift column index by 1 due to SPLIT STATUS column)
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

    // Create workbook and download
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