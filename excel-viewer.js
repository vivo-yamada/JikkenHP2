let workbook = null;
let currentData = [];
let filteredData = [];
let currentPage = 1;
const rowsPerPage = 50;

// Initialize
document.addEventListener('DOMContentLoaded', function() {
    setupFileHandlers();
    setupSearchHandler();
    setupSheetSelector();
    setupExportButton();
});

// File handling setup
function setupFileHandlers() {
    const fileInput = document.getElementById('fileInput');
    const uploadArea = document.getElementById('uploadArea');
    
    // File input change
    fileInput.addEventListener('change', function(e) {
        const file = e.target.files[0];
        if (file) {
            processFile(file);
        }
    });
    
    // Drag and drop
    uploadArea.addEventListener('dragover', function(e) {
        e.preventDefault();
        uploadArea.classList.add('drag-over');
    });
    
    uploadArea.addEventListener('dragleave', function(e) {
        e.preventDefault();
        uploadArea.classList.remove('drag-over');
    });
    
    uploadArea.addEventListener('drop', function(e) {
        e.preventDefault();
        uploadArea.classList.remove('drag-over');
        
        const file = e.dataTransfer.files[0];
        if (file && (file.name.endsWith('.xlsx') || file.name.endsWith('.xls'))) {
            processFile(file);
        } else {
            alert('Excel ファイル (.xlsx または .xls) を選択してください。');
        }
    });
}

// Process Excel file
function processFile(file) {
    const loading = document.getElementById('loading');
    loading.style.display = 'flex';
    
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            workbook = XLSX.read(data, { type: 'array' });
            
            // Populate sheet selector
            populateSheetSelector();
            
            // Load first sheet
            loadSheet(workbook.SheetNames[0]);
            
            // Show controls and data display
            document.getElementById('dataControls').style.display = 'block';
            document.getElementById('dataDisplay').style.display = 'block';
            
            // Update upload area to show file name
            updateUploadArea(file.name);
            
        } catch (error) {
            alert('ファイルの読み込みに失敗しました: ' + error.message);
        } finally {
            loading.style.display = 'none';
        }
    };
    
    reader.readAsArrayBuffer(file);
}

// Populate sheet selector
function populateSheetSelector() {
    const sheetSelect = document.getElementById('sheetSelect');
    sheetSelect.innerHTML = '';
    
    workbook.SheetNames.forEach((sheetName, index) => {
        const option = document.createElement('option');
        option.value = sheetName;
        option.textContent = sheetName;
        if (index === 0) option.selected = true;
        sheetSelect.appendChild(option);
    });
}

// Load sheet data
function loadSheet(sheetName) {
    const worksheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
    
    if (jsonData.length === 0) {
        currentData = [];
        filteredData = [];
        displayData([]);
        return;
    }
    
    // Convert to array of objects
    const headers = jsonData[0];
    currentData = jsonData.slice(1).map(row => {
        const obj = {};
        headers.forEach((header, index) => {
            obj[header || `Column${index + 1}`] = row[index] || '';
        });
        return obj;
    });
    
    filteredData = [...currentData];
    currentPage = 1;
    displayData(filteredData);
    updateStats();
}

// Display data in table
function displayData(data) {
    const tableHead = document.getElementById('tableHead');
    const tableBody = document.getElementById('tableBody');
    
    // Clear existing content
    tableHead.innerHTML = '';
    tableBody.innerHTML = '';
    
    if (data.length === 0) {
        tableBody.innerHTML = '<tr><td colspan="100%" style="text-align: center;">データがありません</td></tr>';
        return;
    }
    
    // Create headers
    const headers = Object.keys(data[0]);
    const headerRow = document.createElement('tr');
    headers.forEach(header => {
        const th = document.createElement('th');
        th.textContent = header;
        th.onclick = () => sortData(header);
        th.style.cursor = 'pointer';
        th.title = 'クリックでソート';
        headerRow.appendChild(th);
    });
    tableHead.appendChild(headerRow);
    
    // Paginate data
    const startIndex = (currentPage - 1) * rowsPerPage;
    const endIndex = startIndex + rowsPerPage;
    const pageData = data.slice(startIndex, endIndex);
    
    // Create rows
    pageData.forEach(row => {
        const tr = document.createElement('tr');
        headers.forEach(header => {
            const td = document.createElement('td');
            const value = row[header];
            td.textContent = value;
            td.title = value; // Show full text on hover
            tr.appendChild(td);
        });
        tableBody.appendChild(tr);
    });
    
    // Update pagination
    updatePagination(data.length);
}

// Sort data
let sortColumn = null;
let sortDirection = 'asc';

function sortData(column) {
    if (sortColumn === column) {
        sortDirection = sortDirection === 'asc' ? 'desc' : 'asc';
    } else {
        sortColumn = column;
        sortDirection = 'asc';
    }
    
    filteredData.sort((a, b) => {
        const aVal = a[column];
        const bVal = b[column];
        
        // Try numeric comparison first
        const aNum = parseFloat(aVal);
        const bNum = parseFloat(bVal);
        
        if (!isNaN(aNum) && !isNaN(bNum)) {
            return sortDirection === 'asc' ? aNum - bNum : bNum - aNum;
        }
        
        // String comparison
        const comparison = String(aVal).localeCompare(String(bVal), 'ja');
        return sortDirection === 'asc' ? comparison : -comparison;
    });
    
    currentPage = 1;
    displayData(filteredData);
}

// Search functionality
function setupSearchHandler() {
    const searchInput = document.getElementById('searchInput');
    let searchTimeout;
    
    searchInput.addEventListener('input', function() {
        clearTimeout(searchTimeout);
        searchTimeout = setTimeout(() => {
            const searchTerm = this.value.toLowerCase();
            
            if (searchTerm === '') {
                filteredData = [...currentData];
            } else {
                filteredData = currentData.filter(row => {
                    return Object.values(row).some(value => 
                        String(value).toLowerCase().includes(searchTerm)
                    );
                });
            }
            
            currentPage = 1;
            displayData(filteredData);
            updateStats();
        }, 300);
    });
}

// Sheet selector handler
function setupSheetSelector() {
    const sheetSelect = document.getElementById('sheetSelect');
    sheetSelect.addEventListener('change', function() {
        loadSheet(this.value);
    });
}

// Pagination
function updatePagination(totalRows) {
    const pagination = document.getElementById('pagination');
    pagination.innerHTML = '';
    
    const totalPages = Math.ceil(totalRows / rowsPerPage);
    
    if (totalPages <= 1) return;
    
    // Previous button
    const prevBtn = document.createElement('button');
    prevBtn.textContent = '前へ';
    prevBtn.disabled = currentPage === 1;
    prevBtn.onclick = () => {
        if (currentPage > 1) {
            currentPage--;
            displayData(filteredData);
        }
    };
    pagination.appendChild(prevBtn);
    
    // Page numbers
    const pageInfo = document.createElement('span');
    pageInfo.className = 'page-info';
    pageInfo.textContent = `${currentPage} / ${totalPages} ページ`;
    pagination.appendChild(pageInfo);
    
    // Next button
    const nextBtn = document.createElement('button');
    nextBtn.textContent = '次へ';
    nextBtn.disabled = currentPage === totalPages;
    nextBtn.onclick = () => {
        if (currentPage < totalPages) {
            currentPage++;
            displayData(filteredData);
        }
    };
    pagination.appendChild(nextBtn);
}

// Update statistics
function updateStats() {
    document.getElementById('totalRows').textContent = `総行数: ${currentData.length}`;
    document.getElementById('filteredRows').textContent = `表示行数: ${filteredData.length}`;
}

// Export to CSV
function setupExportButton() {
    const exportBtn = document.getElementById('exportBtn');
    exportBtn.addEventListener('click', function() {
        if (filteredData.length === 0) {
            alert('エクスポートするデータがありません。');
            return;
        }
        
        const csv = convertToCSV(filteredData);
        downloadCSV(csv, 'export.csv');
    });
}

function convertToCSV(data) {
    const headers = Object.keys(data[0]);
    const csvHeaders = headers.join(',');
    
    const csvRows = data.map(row => {
        return headers.map(header => {
            const value = row[header];
            // Escape quotes and wrap in quotes if contains comma
            const escaped = String(value).replace(/"/g, '""');
            return escaped.includes(',') ? `"${escaped}"` : escaped;
        }).join(',');
    });
    
    return [csvHeaders, ...csvRows].join('\n');
}

function downloadCSV(csv, filename) {
    const blob = new Blob(['\ufeff' + csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
}

// Update upload area after file selection
function updateUploadArea(filename) {
    const uploadContent = document.querySelector('.upload-content');
    uploadContent.innerHTML = `
        <svg class="upload-icon success" width="64" height="64" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"></path>
            <polyline points="22 4 12 14.01 9 11.01"></polyline>
        </svg>
        <p class="file-loaded">${filename}</p>
        <button class="btn-upload" onclick="document.getElementById('fileInput').click()">別のファイルを選択</button>
    `;
}