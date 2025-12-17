let originalData = null;
let processedData = null;
let headers = null;

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

// File upload handler
fileInput.addEventListener('change', handleFileUpload);

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

            if (jsonData.length < 2) {
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
                
                if (jsonData.length < 2) {
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
    const userColumnName = columnInput.value.trim().toLowerCase();
    
    if (!userColumnName) {
        showError('❌ Column name cannot be empty.');
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
            const emailColIndex = headers.findIndex(
                h => String(h).trim().toLowerCase() === userColumnName
            );

            if (emailColIndex === -1) {
                showError(`❌ Column "${columnInput.value.trim()}" not found.`);
                loadingIndicator.style.display = 'none';
                return;
            }

            let newData = [headers];

            // Process each data row
            for (let i = 1; i < originalData.length; i++) {
                const row = originalData[i];
                const emailCell = row[emailColIndex];

                // Skip or keep if N/A or blank
                if (!emailCell || String(emailCell).trim().toUpperCase() === 'N/A') {
                    newData.push(row);
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
                    continue;
                }

                // Create separate rows for each email
                emails.forEach(email => {
                    const newRow = [...row];
                    newRow[emailColIndex] = email;
                    newData.push(newRow);
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
        tableHTML += '<tr>';
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

    const ws = XLSX.utils.aoa_to_sheet(processedData);
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