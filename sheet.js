let excelData = []; // Placeholder for Excel data
let currentSheetName = ''; // Placeholder for the current sheet name
let selectedRows = [];
let selectedCols = [];

// Load the Google Sheets file when the page loads
document.addEventListener('DOMContentLoaded', async () => {
    const urlParams = new URLSearchParams(window.location.search);
    const fileUrl = urlParams.get('fileUrl');

    if (fileUrl) {
        await loadExcelData(fileUrl);
    }
});

// Function to load Excel data
async function loadExcelData(url) {
    const response = await fetch(url);
    const data = await response.arrayBuffer();
    const workbook = XLSX.read(data);
    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
    excelData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
    displayData(excelData);
}

// Function to display data in the table
function displayData(data) {
    const sheetContent = document.getElementById('sheet-content');
    sheetContent.innerHTML = '';

    const table = document.createElement('table');
    data.forEach((row, rowIndex) => {
        const tr = document.createElement('tr');
        row.forEach((cell, cellIndex) => {
            const td = document.createElement('td');
            td.textContent = cell;

            // Set data attributes for row and column
            td.dataset.row = rowIndex;
            td.dataset.col = cellIndex;

            // Add event listener to each cell for selection
            td.addEventListener('click', () => {
                toggleCellSelection(td);
            });

            tr.appendChild(td);
        });
        table.appendChild(tr);
    });
    sheetContent.appendChild(table);
}

// Function to toggle cell selection and highlight
function toggleCellSelection(cell) {
    const rowIndex = cell.dataset.row;
    const colIndex = cell.dataset.col;

    // Check if the cell is already selected
    if (selectedRows.includes(rowIndex) && selectedCols.includes(colIndex)) {
        // Deselect
        selectedRows = selectedRows.filter(row => row !== rowIndex);
        selectedCols = selectedCols.filter(col => col !== colIndex);
        cell.style.backgroundColor = ''; // Reset color
    } else {
        // Select
        if (!selectedRows.includes(rowIndex)) {
            selectedRows.push(rowIndex);
        }
        if (!selectedCols.includes(colIndex)) {
            selectedCols.push(colIndex);
        }
        cell.style.backgroundColor = '#d1e7dd'; // Highlight color
    }

    // Highlight the entire row and column based on selection
    highlightData();
}

// Highlight data based on selections
function highlightData() {
    const table = document.querySelector('table');
    if (!table) return;

    const rows = table.querySelectorAll('tr');

    rows.forEach((row, rowIndex) => {
        // Highlight the selected row
        if (selectedRows.includes(rowIndex.toString())) {
            row.style.backgroundColor = '#d1e7dd'; // Highlight color
        } else {
            row.style.backgroundColor = ''; // Reset color
        }

        // Highlight selected columns
        row.querySelectorAll('td').forEach((cell, cellIndex) => {
            if (selectedCols.includes(cellIndex.toString())) {
                cell.style.backgroundColor = '#d1e7dd'; // Highlight color
            } else if (!selectedRows.includes(rowIndex.toString())) {
                cell.style.backgroundColor = ''; // Reset color
            }
        });
    });
}

// Check the operation condition
function checkOperation(rowIndex, primaryCell, operationColumns, operation, operationType) {
    const primaryValue = primaryCell.textContent.trim();

    if (operationType === 'and') {
        return operationColumns.every(col => {
            const colCell = primaryCell.parentNode.cells[col.charCodeAt(0) - 65]; // Get cell for operation
            const colValue = colCell.textContent.trim();
            return operation === 'null' ? !colValue : colValue !== '';
        });
    } else if (operationType === 'or') {
        return operationColumns.some(col => {
            const colCell = primaryCell.parentNode.cells[col.charCodeAt(0) - 65]; // Get cell for operation
            const colValue = colCell.textContent.trim();
            return operation === 'null' ? !colValue : colValue !== '';
        });
    }
    return false;
}

// Download functionality
document.getElementById('download-button').addEventListener('click', () => {
    document.getElementById('download-modal').style.display = 'flex';
});

// Confirm download button
document.getElementById('confirm-download').addEventListener('click', () => {
    const filename = document.getElementById('filename').value || 'downloaded_file';
    const format = document.getElementById('file-format').value;

    // Download logic based on format
    if (format === 'xlsx') {
        const ws = XLSX.utils.aoa_to_sheet(excelData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, currentSheetName);
        XLSX.writeFile(wb, `${filename}.xlsx`);
    } else if (format === 'csv') {
        const csvContent = XLSX.utils.sheet_to_csv(XLSX.utils.aoa_to_sheet(excelData));
        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.setAttribute('href', url);
        link.setAttribute('download', `${filename}.csv`);
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    }

    document.getElementById('download-modal').style.display = 'none';
});

// Close modal
document.getElementById('close-modal').addEventListener('click', () => {
    document.getElementById('download-modal').style.display = 'none';
});

// Load Excel Sheet
async function loadExcelSheet(fileUrl) {
    try {
        const response = await fetch(fileUrl);
        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];

        data = XLSX.utils.sheet_to_json(sheet, { defval: null });
        filteredData = [...data];
        displaySheet(filteredData);
    } catch (error) {
        console.error("Error loading Excel sheet:", error);
    }
}

// Display Sheet
function displaySheet(sheetData) {
    const sheetContentDiv = document.getElementById('sheet-content');
    sheetContentDiv.innerHTML = '';

    if (sheetData.length === 0) {
        sheetContentDiv.innerHTML = '<p>No data available</p>';
        return;
    }

    const table = document.createElement('table');
    const headerRow = document.createElement('tr');
    Object.keys(sheetData[0]).forEach(header => {
        const th = document.createElement('th');
        th.textContent = header;
        headerRow.appendChild(th);
    });
    table.appendChild(headerRow);

    sheetData.forEach(row => {
        const tr = document.createElement('tr');
        Object.values(row).forEach(cell => {
            const td = document.createElement('td');
            td.textContent = cell === null || cell === "" ? 'NULL' : cell;
            tr.appendChild(td);
        });
        table.appendChild(tr);
    });

    sheetContentDiv.appendChild(table);
}

// Apply Operation
function applyOperation() {
    const primaryColumn = document.getElementById('primary-column').value.trim();
    const operationColumnsInput = document.getElementById('operation-columns').value.trim();
    const operationType = document.getElementById('operation-type').value;
    const operation = document.getElementById('operation').value;

    const rowRangeFrom = parseInt(document.getElementById('row-range-from').value, 10);
    const rowRangeTo = parseInt(document.getElementById('row-range-to').value, 10);

    if (!primaryColumn || !operationColumnsInput) {
        alert('Please enter the primary column and columns to operate on.');
        return;
    }

    const operationColumns = operationColumnsInput.split(',').map(col => col.trim());
    filteredData = data.filter((row, index) => {
        // Check if the current row index is within the specified range
        if (index < rowRangeFrom - 1 || index > rowRangeTo - 1) return false;

        const isPrimaryNull = row[primaryColumn] === null || row[primaryColumn] === "";
        const columnChecks = operationColumns.map(col => operation === 'null' ? row[col] === null || row[col] === "" : row[col] !== null && row[col] !== "");

        return operationType === 'and' ? !isPrimaryNull && columnChecks.every(Boolean) : !isPrimaryNull && columnChecks.some(Boolean);
    });

    filteredData = filteredData.map(row => {
        const filteredRow = {};
        filteredRow[primaryColumn] = row[primaryColumn];
        operationColumns.forEach(col => filteredRow[col] = row[col] === null || row[col] === "" ? 'NULL' : row[col]);
        return filteredRow;
    });

    displaySheet(filteredData);
}
