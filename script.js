let excelData = null;
let currentPage = 1;
const rowsPerPage = 20;

const uploadArea = document.getElementById('uploadArea');
const fileInput = document.getElementById('fileInput');

uploadArea.addEventListener('dragover', (e) => {
    e.preventDefault();
    uploadArea.classList.add('dragover');
});

uploadArea.addEventListener('dragleave', () => {
    uploadArea.classList.remove('dragover');
});

uploadArea.addEventListener('drop', (e) => {
    e.preventDefault();
    uploadArea.classList.remove('dragover');
    const files = e.dataTransfer.files;
    if (files.length > 0) {
        handleFile(files[0]);
    }
});

uploadArea.addEventListener('click', () => {
    fileInput.click();
});

fileInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) {
        handleFile(e.target.files[0]);
    }
});

function showLoading() {
    document.getElementById('loading').style.display = 'block';
    document.getElementById('errorMsg').style.display = 'none';
    document.getElementById('successMsg').style.display = 'none';
}

function hideLoading() {
    document.getElementById('loading').style.display = 'none';
}

function showError(message) {
    const errorMsg = document.getElementById('errorMsg');
    errorMsg.textContent = message;
    errorMsg.style.display = 'block';
    document.getElementById('successMsg').style.display = 'none';
}

function showSuccess(message) {
    const successMsg = document.getElementById('successMsg');
    successMsg.textContent = message;
    successMsg.style.display = 'block';
    document.getElementById('errorMsg').style.display = 'none';
}

function handleFile(file) {
    if (!file.name.match(/\.(xlsx|xls)$/)) {
        showError('Mohon upload file Excel (.xlsx atau .xls)');
        return;
    }

    showLoading();

    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });

            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];

            excelData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

            hideLoading();
            showSuccess(`File "${file.name}" berhasil diupload! Menampilkan ${excelData.length - 1} baris data.`);

            currentPage = 1;
            generateWebContent();
        } catch (error) {
            hideLoading();
            showError('Error membaca file Excel: ' + error.message);
        }
    };

    reader.readAsArrayBuffer(file);
}

function generateWebContent() {
    if (!excelData || excelData.length === 0) return;

    const contentSection = document.getElementById('contentSection');
    document.getElementById('statsGrid').innerHTML = '';
    document.getElementById('dataGrid').innerHTML = '';
    document.getElementById('tableContainer').innerHTML = '';

    generateStats();
    generateDataCards();
    generateTable();

    contentSection.style.display = 'block';
    contentSection.scrollIntoView({ behavior: 'smooth' });
}

function generateStats() {
    const statsGrid = document.getElementById('statsGrid');
    const totalRows = excelData.length - 1;
    const totalColumns = excelData[0] ? excelData[0].length : 0;

    let filledCells = 0;
    for (let i = 1; i < excelData.length; i++) {
        for (let j = 0; j < excelData[i].length; j++) {
            if (excelData[i][j] !== undefined && excelData[i][j] !== null && excelData[i][j] !== '') {
                filledCells++;
            }
        }
    }

    statsGrid.innerHTML = `
        <div class="stat-card">
            <div class="stat-number">${totalRows}</div>
            <div class="stat-label">Total Baris</div>
        </div>
        <div class="stat-card">
            <div class="stat-number">${totalColumns}</div>
            <div class="stat-label">Total Kolom</div>
        </div>
        <div class="stat-card">
            <div class="stat-number">${filledCells}</div>
            <div class="stat-label">Data Terisi</div>
        </div>
        <div class="stat-card">
            <div class="stat-number">${Math.round((filledCells / (totalRows * totalColumns)) * 100)}%</div>
            <div class="stat-label">Kelengkapan</div>
        </div>
    `;
}

function generateDataCards() {
    const dataGrid = document.getElementById('dataGrid');
    if (excelData.length < 2) return;

    const headers = excelData[0];
    let cardsHTML = '';

    headers.forEach((header, index) => {
        if (!header) return;

        const columnData = [];
        for (let i = 1; i < excelData.length; i++) {
            if (excelData[i][index] !== undefined && excelData[i][index] !== null && excelData[i][index] !== '') {
                columnData.push(excelData[i][index]);
            }
        }

        let cardContent = '';
        if (columnData.length > 0) {
            const isNumeric = columnData.every(val => !isNaN(val) && val !== '');

            if (isNumeric) {
                const numbers = columnData.map(val => parseFloat(val));
                const sum = numbers.reduce((a, b) => a + b, 0);
                const avg = sum / numbers.length;
                const min = Math.min(...numbers);
                const max = Math.max(...numbers);

                cardContent = `
                    <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 15px; margin-top: 15px;">
                        <div><strong>Total:</strong> ${sum.toLocaleString()}</div>
                        <div><strong>Rata-rata:</strong> ${avg.toFixed(2)}</div>
                        <div><strong>Minimum:</strong> ${min}</div>
                        <div><strong>Maksimum:</strong> ${max}</div>
                    </div>
                `;
            } else {
                const uniqueValues = [...new Set(columnData)].slice(0, 5);
                cardContent = `
                    <div style="margin-top: 15px;">
                        <strong>Sampel data:</strong><br>
                        ${uniqueValues.map(val => `• ${val}`).join('<br>')}
                        ${columnData.length > 5 ? `<br>... dan ${columnData.length - 5} lainnya` : ''}
                    </div>
                `;
            }
        }

        cardsHTML += `
            <div class="data-card">
                <div class="card-header">📋 ${header}</div>
                <div class="card-content">
                    <strong>Total data:</strong> ${columnData.length} item
                    ${cardContent}
                </div>
            </div>
        `;
    });

    dataGrid.innerHTML = cardsHTML;
}

function generateTable() {
    const tableContainer = document.getElementById('tableContainer');
    if (excelData.length === 0) return;

    const totalRows = excelData.length - 1;
    const totalPages = Math.ceil(totalRows / rowsPerPage);
    const startIndex = (currentPage - 1) * rowsPerPage + 1;
    const endIndex = Math.min(startIndex + rowsPerPage - 1, excelData.length - 1);

    let tableHTML = '<table><thead><tr>';

    if (excelData[0]) {
        excelData[0].forEach(header => {
            tableHTML += `<th>${header || 'Kolom'}</th>`;
        });
    }

    tableHTML += '</tr></thead><tbody>';

    for (let i = startIndex; i <= endIndex; i++) {
        tableHTML += '<tr>';
        for (let j = 0; j < excelData[0].length; j++) {
            const cellValue = excelData[i][j] || '';
            tableHTML += `<td>${cellValue}</td>`;
        }
        tableHTML += '</tr>';
    }

    tableHTML += '</tbody></table>';

    // Pagination controls
    tableHTML += `
    <div class="pagination">
        <button class="pagination-button" onclick="prevPage()" ${currentPage === 1 ? 'disabled' : ''}>
            ⬅️ <span>Sebelumnya</span>
        </button>
        <span style="font-weight: 600;">Halaman ${currentPage} dari ${totalPages}</span>
        <button class="pagination-button" onclick="nextPage()" ${currentPage === totalPages ? 'disabled' : ''}>
            <span>Berikutnya</span> ➡️
        </button>
    </div>
`;

    tableContainer.innerHTML = tableHTML;
}

function nextPage() {
    const totalPages = Math.ceil((excelData.length - 1) / rowsPerPage);
    if (currentPage < totalPages) {
        currentPage++;
        generateTable();
    }
}

function prevPage() {
    if (currentPage > 1) {
        currentPage--;
        generateTable();
    }
}


