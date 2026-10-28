// Struktur data global
let globalData = [];
let diagnosaChartInstance = null;
let ageRangeChartInstance = null;
let filteredData = [];
let currentPage = 1;
const pageSize = 100;
const databaseName = 'DashboardISPA';
const databaseVersion = 1;
const dataStoreName = 'dataPasien';

// Membuka penyimpanan lokal browser untuk mempertahankan data setelah refresh.
function openDataDatabase() {
    return new Promise((resolve, reject) => {
        const request = indexedDB.open(databaseName, databaseVersion);
        request.onupgradeneeded = () => {
            if (!request.result.objectStoreNames.contains(dataStoreName)) {
                request.result.createObjectStore(dataStoreName);
            }
        };
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
    });
}

async function saveDataToBrowser(data) {
    const database = await openDataDatabase();
    return new Promise((resolve, reject) => {
        const transaction = database.transaction(dataStoreName, 'readwrite');
        transaction.objectStore(dataStoreName).put(data, 'current');
        transaction.oncomplete = () => {
            database.close();
            resolve();
        };
        transaction.onerror = () => {
            database.close();
            reject(transaction.error);
        };
    });
}

async function loadDataFromBrowser() {
    const database = await openDataDatabase();
    return new Promise((resolve, reject) => {
        const transaction = database.transaction(dataStoreName, 'readonly');
        const request = transaction.objectStore(dataStoreName).get('current');
        request.onsuccess = () => {
            database.close();
            resolve(request.result || []);
        };
        request.onerror = () => {
            database.close();
            reject(request.error);
        };
    });
}

async function clearStoredData() {
    const database = await openDataDatabase();
    return new Promise((resolve, reject) => {
        const transaction = database.transaction(dataStoreName, 'readwrite');
        transaction.objectStore(dataStoreName).delete('current');
        transaction.oncomplete = () => {
            database.close();
            resolve();
        };
        transaction.onerror = () => {
            database.close();
            reject(transaction.error);
        };
    });
}

async function showSavedDashboard(data) {
    globalData = data.map(item => ({
        ...item,
        rentangUsia: item.rentangUsia || item.umur || '-'
    }));
    populatePuskesmasFilter();
    populateDynamicFilters();
    document.getElementById('upload-section').classList.add('is-hidden');
    document.getElementById('dashboard-section').classList.remove('is-hidden');
    applyFilters();
}

// Mengubah tanggal Excel (nomor serial) menjadi string
function formatExcelDate(serial) {
    if (!serial) return '-';
    if (serial instanceof Date) {
        return serial.toLocaleDateString('id-ID', { day: '2-digit', month: 'short', year: 'numeric' });
    }
    if (typeof serial === 'string') return serial;
    const utc_days = Math.floor(serial - 25569);
    const utc_value = utc_days * 86400;
    const date_info = new Date(utc_value * 1000);
    return date_info.toLocaleDateString('id-ID', { day: '2-digit', month: 'short', year: 'numeric' });
}

function toDateInputValue(value) {
    let date;
    if (value instanceof Date) {
        date = value;
    } else if (typeof value === 'number') {
        date = new Date((value - 25569) * 86400 * 1000);
    } else {
        const text = String(value || '').trim();
        const parts = text.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/);
        date = parts
            ? new Date(Number(parts[3]), Number(parts[2]) - 1, Number(parts[1]))
            : new Date(text);
    }
    if (!date || Number.isNaN(date.getTime())) return '';
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
}

function getValueFromKeys(obj, possibleKeys) {
    const normalizedKeys = possibleKeys.map(key => normalizeHeader(key));
    for (let key of possibleKeys) {
        const foundKey = Object.keys(obj).find(k => normalizeHeader(k) === normalizeHeader(key));
        if (foundKey && obj[foundKey] !== undefined) {
            return obj[foundKey];
        }
    }
    const partialKey = Object.keys(obj).find(k => normalizedKeys.some(key => {
        const normalizedHeader = normalizeHeader(k);
        return normalizedHeader.includes(key) || key.includes(normalizedHeader);
    }));
    if (partialKey && obj[partialKey] !== undefined) return obj[partialKey];
    return '';
}

function normalizeHeader(value) {
    return String(value || '')
        .toLowerCase()
        .replace(/[\r\n]+/g, ' ')
        .replace(/[^a-z0-9]+/g, ' ')
        .trim()
        .replace(/\s+/g, ' ');
}

function normalizePuskesmasName(value) {
    const name = String(value || '').trim().replace(/\s+/g, ' ');
    if (normalizeHeader(name) === 'bunga jadi') return 'Bunga Jadi';
    return name;
}

// Memproses file Excel menggunakan SheetJS
function parseExcelFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array', cellDates: true, dense: true });
                const parsedData = [];

                for (const sheetName of workbook.SheetNames) {
                    const worksheet = workbook.Sheets[sheetName];
                    const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: true, blankrows: false, defval: '' });
                    if (rows.length < 2) continue;

                    let headerIndex = -1;
                    for (let i = 0; i < rows.length; i++) {
                        const rowStr = (rows[i] || []).map(normalizeHeader).join(' ');
                        const headerScore = ['nama', 'pasien', 'puskesmas', 'diagnosa', 'diagnosis', 'tanggal', 'jenis kelamin', 'umur', 'usia']
                            .filter(keyword => rowStr.includes(keyword)).length;
                        if (headerScore >= 2) {
                            headerIndex = i;
                            break;
                        }
                    }
                    if (headerIndex === -1) continue;

                    const headers = rows[headerIndex];
                    for (let i = headerIndex + 1; i < rows.length; i++) {
                        const row = rows[i];
                        if (!row || row.length === 0) continue;

                        const obj = {};
                        headers.forEach((header, index) => {
                            if (header !== undefined && header !== '') obj[String(header).trim()] = row[index];
                        });

                        const puskesmas = getValueFromKeys(obj, ['Nama Puskesmas', 'Puskesmas', 'Fasilitas Kesehatan']);
                        const nama = getValueFromKeys(obj, ['Nama Pasien', 'Nama Lengkap', 'Nama']);
                        const diagnosa = getValueFromKeys(obj, ['Diagnosis Utama', 'Diagnosa Utama', 'Diagnosis 1', 'Diagnosis', 'Diagnosa 1', 'Diagnosa']);
                        const tarikh = getValueFromKeys(obj, ['Tanggal Kunjungan', 'Tanggal', 'Tanggal Pelayanan', 'Date']);
                        const umur = getValueFromKeys(obj, ['Umur', 'Usia']);
                        const rentangUsia = getValueFromKeys(obj, ['Rentang Usia', 'Kelompok Usia', 'Range Usia', 'Age Range']) || umur;
                        let jantina = getValueFromKeys(obj, ['Jenis Kelamin', 'JK', 'Jantina']);
                        const alergi = getValueFromKeys(obj, ['Alergi', 'Riwayat Alergi', 'Alergi Obat', 'Alergi Makanan']);
                        const skalaNyeri = getValueFromKeys(obj, ['Skala Nyeri', 'Skala Sakit', 'Nyeri', 'Pain Scale']);

                        jantina = String(jantina).toUpperCase().startsWith('L') ? 'L' :
                                  String(jantina).toUpperCase().startsWith('P') ? 'P' : '-';

                        if (nama || diagnosa) {
                            parsedData.push({
                                puskesmas: normalizePuskesmasName(puskesmas),
                                nama: String(nama).trim(),
                                diagnosa: String(diagnosa).trim(),
                                tarikh: formatExcelDate(tarikh),
                                tanggalFilter: toDateInputValue(tarikh),
                                umur: umur || '-',
                                rentangUsia: String(rentangUsia || '-').trim(),
                                jantina: jantina,
                                alergi: String(alergi || '-').trim(),
                                skalaNyeri: String(skalaNyeri || '-').trim()
                            });
                        }
                    }
                }
                resolve(parsedData);
            } catch (error) {
                reject(new Error(`File "${file.name}" tidak dapat dibaca sebagai workbook Excel. ${error.message || ''}`));
            }
        };
        reader.onerror = function() {
            reject(new Error(`File "${file.name}" gagal dibaca oleh browser.`));
        };
        reader.readAsArrayBuffer(file);
    });
}

// Menangani unggahan file
document.getElementById('file-upload').addEventListener('change', async (event) => {
    const files = event.target.files;
    if (files.length === 0) return;

    document.getElementById('loading-indicator').classList.remove('is-hidden');
    globalData = [];

    try {
        for (let i = 0; i < files.length; i++) {
            const data = await parseExcelFile(files[i]);
            if (data.length === 0) {
                throw new Error(`Tidak ada data pasien yang dikenali dari file "${files[i].name}". Pastikan baris header memuat kolom nama pasien atau diagnosis.`);
            }
            globalData = globalData.concat(data);
        }

            try {
                await saveDataToBrowser(globalData);
            } catch (storageError) {
                console.warn('Data tidak dapat disimpan untuk pemulihan setelah refresh:', storageError);
            }
        populatePuskesmasFilter();
        populateDynamicFilters();
        document.getElementById('upload-section').classList.add('is-hidden');
        document.getElementById('dashboard-section').classList.remove('is-hidden');
        applyFilters();
    } catch (error) {
        console.error(error);
        alert(error.message || 'Terjadi kesalahan saat membaca file Excel. Pastikan formatnya benar.');
        document.getElementById('upload-section').classList.remove('is-hidden');
        document.getElementById('dashboard-section').classList.add('is-hidden');
    } finally {
        document.getElementById('loading-indicator').classList.add('is-hidden');
    }
});

// Memulihkan data terakhir secara otomatis setelah halaman direfresh.
document.getElementById('loading-indicator').classList.remove('is-hidden');
loadDataFromBrowser()
    .then(savedData => {
        if (savedData.length > 0) {
            showSavedDashboard(savedData);
        }
    })
    .catch(error => console.warn('Data tersimpan tidak dapat dipulihkan:', error))
    .finally(() => {
        // sembunyikan loading indicator
        document.getElementById('loading-indicator').classList.add('is-hidden');
    });

const filterPuskesmas = document.getElementById('filter-puskesmas');
const filterDiagnosis = document.getElementById('filter-diagnosis');
const filterGender = document.getElementById('filter-gender');
const filterAgeRange = document.getElementById('filter-age-range');
const filterDate = document.getElementById('filter-date');
const filterPainScale = document.getElementById('filter-pain-scale');
const searchInput = document.getElementById('search-input');
const previousPageButton = document.getElementById('previous-page');
const nextPageButton = document.getElementById('next-page');
const clearStoredDataButton = document.getElementById('btn-clear-stored-data');

clearStoredDataButton.addEventListener('click', async () => {
    if (!confirm('Hapus data Excel yang tersimpan di browser?')) return;

    try {
        await clearStoredData();
        globalData = [];
        filteredData = [];
        if (diagnosaChartInstance) {
            diagnosaChartInstance.destroy();
            diagnosaChartInstance = null;
        }
        if (ageRangeChartInstance) {
            ageRangeChartInstance.destroy();
            ageRangeChartInstance = null;
        }
        document.getElementById('dashboard-section').classList.add('is-hidden');
        document.getElementById('upload-section').classList.remove('is-hidden');
        document.getElementById('file-upload').value = '';
    } catch (error) {
        console.error(error);
        alert('Data tersimpan tidak dapat dihapus.');
    }
});

function debounce(func, wait) {
    let timeout;
    return function executedFunction(...args) {
        const later = () => {
            clearTimeout(timeout);
            func(...args);
        };
        clearTimeout(timeout);
        timeout = setTimeout(later, wait);
    };
}

filterPuskesmas.addEventListener('change', applyFilters);
filterDiagnosis.addEventListener('change', applyFilters);
filterGender.addEventListener('change', applyFilters);
filterAgeRange.addEventListener('change', applyFilters);
filterDate.addEventListener('change', applyFilters);
filterPainScale.addEventListener('change', applyFilters);
searchInput.addEventListener('input', debounce(applyFilters, 300));
previousPageButton.addEventListener('click', () => {
    if (currentPage > 1) {
        currentPage--;
        updateTable(filteredData);
    }
});
nextPageButton.addEventListener('click', () => {
    const totalPages = Math.ceil(filteredData.length / pageSize);
    if (currentPage < totalPages) {
        currentPage++;
        updateTable(filteredData);
    }
});

function populatePuskesmasFilter() {
    const currentValue = filterPuskesmas.value;
    const puskesmasMap = new Map();
    globalData.forEach(item => {
        const displayName = normalizePuskesmasName(item.puskesmas);
        const filterValue = normalizeHeader(displayName);
        if (filterValue && !puskesmasMap.has(filterValue)) {
            puskesmasMap.set(filterValue, displayName);
        }
    });
    const puskesmasList = [...puskesmasMap.entries()]
        .map(([value, label]) => ({ value, label }))
        .sort((first, second) => first.label.localeCompare(second.label, 'id-ID'));

    filterPuskesmas.innerHTML = '<option value="semua">Semua Puskesmas</option>';
    puskesmasList.forEach(({ value, label }) => {
        const option = document.createElement('option');
        option.value = value;
        option.textContent = label;
        filterPuskesmas.appendChild(option);
    });

    if ([...filterPuskesmas.options].some(option => option.value === currentValue)) {
        filterPuskesmas.value = currentValue;
    }
}

function populateDynamicFilters() {
    const diagnosisSet = new Set();
    const genderSet = new Set();
    const painScaleSet = new Set();

    // Hanya melakukan SATU kali looping pada seluruh data
    globalData.forEach(item => {
        if (item.diagnosa) diagnosisSet.add(item.diagnosa);
        if (item.jantina) genderSet.add(item.jantina);
        if (item.skalaNyeri) painScaleSet.add(item.skalaNyeri);
    });

    populateSelect(filterDiagnosis, 'Semua Diagnosis', Array.from(diagnosisSet));
    populateSelect(filterGender, 'Semua Jenis Kelamin', Array.from(genderSet));
    populateSelect(filterAgeRange, 'Semua Rentang Usia', globalData.map(item => item.rentangUsia));
    populateSelect(filterPainScale, 'Semua Skala Nyeri', Array.from(painScaleSet));
}

function populateSelect(select, defaultLabel, values) {
    const currentValue = select.value;
    const uniqueValues = [...new Set(values.map(value => String(value || '-').trim()).filter(Boolean))]
        .sort((first, second) => first.localeCompare(second, 'id-ID', { numeric: true }));

    select.innerHTML = '';
    const defaultOption = document.createElement('option');
    defaultOption.value = 'semua';
    defaultOption.textContent = defaultLabel;
    select.appendChild(defaultOption);

    uniqueValues.forEach(value => {
        const option = document.createElement('option');
        option.value = value.toLowerCase();
        option.textContent = value;
        select.appendChild(option);
    });

    if ([...select.options].some(option => option.value === currentValue)) {
        select.value = currentValue;
    }
}

function applyFilters() {
    const puskesmasVal = normalizeHeader(filterPuskesmas.value);
    const diagnosisVal = filterDiagnosis.value.toLowerCase();
    const genderVal = filterGender.value.toLowerCase();
    const ageRangeVal = filterAgeRange.value.toLowerCase();
    const dateVal = filterDate.value;
    const painScaleVal = filterPainScale.value.toLowerCase();
    const searchVal = searchInput.value.toLowerCase();

    filteredData = globalData.filter(item => {
        const p = normalizeHeader(item.puskesmas);
        const n = (item.nama || '').toLowerCase();
        const d = (item.diagnosa || '').toLowerCase();
        const g = (item.jantina || '').toLowerCase();
        const ageRange = (item.rentangUsia || '-').toLowerCase();
        const date = item.tanggalFilter || '';
        const pain = (item.skalaNyeri || '-').toLowerCase();

        const matchPuskesmas = puskesmasVal === 'semua' || p === puskesmasVal;
        const matchDiagnosis = diagnosisVal === 'semua' || d === diagnosisVal;
        const matchGender = genderVal === 'semua' || g === genderVal;
        const matchAgeRange = ageRangeVal === 'semua' || ageRange === ageRangeVal;
        const matchDate = dateVal === '' || date === dateVal;
        const matchPainScale = painScaleVal === 'semua' || pain === painScaleVal;
        const matchSearch = searchVal === '' || n.includes(searchVal) || d.includes(searchVal) || p.includes(searchVal);

        return matchPuskesmas && matchDiagnosis && matchGender && matchAgeRange && matchDate && matchPainScale && matchSearch;
    });

    currentPage = 1;
    updateCards(filteredData);
    updateTable(filteredData);
    updateChart(filteredData);
}

document.getElementById('btn-reset-filter').addEventListener('click', () => {
    filterPuskesmas.value = 'semua';
    filterDiagnosis.value = 'semua';
    filterGender.value = 'semua';
    filterAgeRange.value = 'semua';
    filterDate.value = '';
    filterPainScale.value = 'semua';
    searchInput.value = '';
    applyFilters();
});

function updateCards(data) {
    let total = data.length;
    let maleCount = 0;
    let femaleCount = 0;

    data.forEach(item => {
        if (item.jantina === 'L') maleCount++;
        if (item.jantina === 'P') femaleCount++;
    });

    document.getElementById('card-total').innerText = total.toLocaleString('id-ID');
    document.getElementById('card-male').innerText = maleCount.toLocaleString('id-ID');
    document.getElementById('card-female').innerText = femaleCount.toLocaleString('id-ID');
    document.getElementById('table-count').innerText = `${total.toLocaleString('id-ID')} Data`;
}

function updateTable(data) {
    const tbody = document.getElementById('table-body');
    tbody.innerHTML = '';

    if (data.length === 0) {
        tbody.innerHTML = `<tr><td colspan="9" class="empty-table-cell">Data tidak ditemukan.</td></tr>`;
        document.getElementById('page-info').innerText = 'Halaman 0 dari 0';
        document.getElementById('previous-page').disabled = true;
        document.getElementById('next-page').disabled = true;
        return;
    }

    const totalPages = Math.ceil(data.length / pageSize);
    currentPage = Math.min(currentPage, Math.max(totalPages, 1));
    const startIndex = (currentPage - 1) * pageSize;
    const dataToDisplay = data.slice(startIndex, startIndex + pageSize);

    dataToDisplay.forEach((row, index) => {
        const tr = document.createElement('tr');
        tr.className = 'data-row';
        const genderClass = row.jantina === 'L' ? 'gender-badge gender-badge-male' :
                            row.jantina === 'P' ? 'gender-badge gender-badge-female' : 'gender-badge gender-badge-unknown';

        tr.innerHTML = `
            <td class="table-cell-muted">${startIndex + index + 1}</td>
            <td class="table-cell-primary">${row.puskesmas}</td>
            <td>${row.tarikh}</td>
            <td class="table-cell-primary">${row.nama}</td>
            <td class="table-cell-center">${row.umur}</td>
            <td class="table-cell-center">
                <span class="${genderClass}">${row.jantina}</span>
            </td>
            <td class="diagnosis-cell" title="${row.diagnosa}">${row.diagnosa}</td>
            <td class="table-cell-center">${row.skalaNyeri}</td>
            <td class="diagnosis-cell" title="${row.alergi}">${row.alergi}</td>
        `;
        tbody.appendChild(tr);
    });

    document.getElementById('page-info').innerText = `Halaman ${totalPages === 0 ? 0 : currentPage} dari ${totalPages}`;
    document.getElementById('previous-page').disabled = currentPage <= 1 || totalPages === 0;
    document.getElementById('next-page').disabled = currentPage >= totalPages || totalPages === 0;
}

document.getElementById('btn-export').addEventListener('click', () => {
    if (filteredData.length === 0) {
        alert("Tidak ada data untuk diekspor.");
        return;
    }

    // Membersihkan data sebelum di-export agar rapi di Excel
    const dataToExport = filteredData.map((item, index) => ({
        'No': index + 1,
        'Puskesmas': item.puskesmas,
        'Tanggal': item.tarikh,
        'Nama Pasien': item.nama,
        'Usia': item.umur,
        'Jenis Kelamin': item.jantina,
        'Diagnosis': item.diagnosa,
        'Skala Nyeri': item.skalaNyeri,
        'Alergi': item.alergi
    }));

    const ws = XLSX.utils.json_to_sheet(dataToExport);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Filter Data ISPA");
    XLSX.writeFile(wb, "Laporan_ISPA_Tersaring.xlsx");
});

function updateChart(data) {
    const ctx = document.getElementById('diagnosaChart').getContext('2d');
    const diagCounts = {};
    data.forEach(item => {
        const d = item.diagnosa || 'Tidak Diketahui';
        diagCounts[d] = (diagCounts[d] || 0) + 1;
    });

    const sortedDiags = Object.entries(diagCounts)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 5);
    const labels = sortedDiags.map(item => item[0].length > 25 ? item[0].substring(0, 25) + '...' : item[0]);
    const counts = sortedDiags.map(item => item[1]);

    if (diagnosaChartInstance) diagnosaChartInstance.destroy();

    diagnosaChartInstance = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: [{
                label: 'Jumlah Kasus',
                data: counts,
                backgroundColor: [
                    'rgba(59, 130, 246, 0.8)',
                    'rgba(16, 185, 129, 0.8)',
                    'rgba(245, 158, 11, 0.8)',
                    'rgba(239, 68, 68, 0.8)',
                    'rgba(139, 92, 246, 0.8)'
                ],
                borderRadius: 6,
                borderSkipped: false
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: false },
                tooltip: {
                    callbacks: {
                        title: function(context) {
                            return sortedDiags[context[0].dataIndex][0];
                        }
                    }
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    grid: { color: '#f1f5f9' },
                    border: { display: false }
                },
                x: {
                    grid: { display: false },
                    border: { display: false },
                    ticks: { font: { size: 11 } }
                }
            }
        }
    });

    updateAgeRangeChart(data);
}

function updateAgeRangeChart(data) {
    const ctx = document.getElementById('ageRangeChart').getContext('2d');
    const ageCounts = {};

    data.forEach(item => {
        const ageRange = item.rentangUsia || item.umur || 'Tidak Diketahui';
        ageCounts[ageRange] = (ageCounts[ageRange] || 0) + 1;
    });

    const sortedAges = Object.entries(ageCounts)
        .sort((first, second) => second[1] - first[1]);

    if (ageRangeChartInstance) ageRangeChartInstance.destroy();

    ageRangeChartInstance = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: sortedAges.map(item => item[0]),
            datasets: [{
                label: 'Jumlah Pasien',
                data: sortedAges.map(item => item[1]),
                backgroundColor: 'rgba(14, 165, 233, 0.8)',
                borderRadius: 6,
                borderSkipped: false
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: false }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    grid: { color: '#f1f5f9' },
                    border: { display: false }
                },
                x: {
                    grid: { display: false },
                    border: { display: false },
                    ticks: { font: { size: 11 } }
                }
            }
        }
    });
}
