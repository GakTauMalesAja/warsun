// Global Variables
let inventory = [];
let invoices = [];
let currentInvoiceItems = [];

// CDN Script Loader
function loadScript(src, opts = {}) {
    return new Promise((resolve, reject) => {
        const script = document.createElement('script');
        script.src = src;
        script.async = true;
        if (opts.integrity) script.integrity = opts.integrity;
        if (opts.crossOrigin) script.crossOrigin = opts.crossOrigin;
        script.onload = () => resolve(script);
        script.onerror = () => reject(new Error(`Gagal load script ${src}`));
        document.head.appendChild(script);
    });
}

async function initLibraries() {
    try {
        await loadScript('https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js', {
            crossOrigin: 'anonymous'
        });
    } catch (e1) {
        try {
            await loadScript('https://unpkg.com/xlsx@0.18.5/dist/xlsx.full.min.js');
        } catch (e2) {
            console.error('Gagal memuat SheetJS (XLSX) dari CDN utama & fallback', e1, e2);
            alert('Gagal memuat library Excel. Silakan gunakan koneksi internet atau file lokal.');
        }
    }

    try {
        await loadScript('https://cdn.jsdelivr.net/npm/chart.js@3.9.1/dist/chart.min.js', {
            crossOrigin: 'anonymous'
        });
    } catch (e1) {
        try {
            await loadScript('https://unpkg.com/chart.js@3.9.1/dist/chart.min.js');
        } catch (e2) {
            console.error('Gagal memuat Chart.js', e1, e2);
            alert('Gagal memuat library Chart.js. Grafik tidak akan muncul.');
        }
    }

    if (window.initializeApp) {
        window.initializeApp();
    }
}

document.addEventListener('DOMContentLoaded', initLibraries);

// Application Initialization
function initializeApp() {
    loadInventory();
    loadInvoices();
    populateProductDropdowns();
    loadDashboard();
    initInvoiceForm();

    document.querySelectorAll('.nav-btn').forEach(btn => {
        btn.addEventListener('click', (e) => {
            const sectionId = e.target.dataset.section;
            switchSection(sectionId);
        });
    });

    document.getElementById('addItemForm').addEventListener('submit', handleAddItem);
    document.getElementById('editItemForm').addEventListener('submit', handleEditItem);
    document.getElementById('searchInput').addEventListener('input', handleSearch);

    const invoiceProductSearch = document.getElementById('invoiceProductSearch');
    if (invoiceProductSearch) {
        invoiceProductSearch.addEventListener('input', filterInvoiceProducts);
    }

    window.addEventListener('click', (e) => {
        const addItemModal = document.getElementById('addItemModal');
        const editItemModal = document.getElementById('editItemModal');
        const invoiceDetailModal = document.getElementById('invoiceDetailModal');
        if (e.target === addItemModal) closeAddItemModal();
        if (e.target === editItemModal) closeEditItemModal();
        if (e.target === invoiceDetailModal) closeInvoiceDetailModal();
    });
}

// Data Management Functions
function loadInventory() {
    const data = localStorage.getItem('inventory');
    inventory = data ? JSON.parse(data) : [];
    inventory = inventory.map(item => ({
        ...item,
        originalQty: item.originalQty !== undefined && item.originalQty !== null ? item.originalQty : item.qty
    }));
    renderInventoryTable();
    updateDashboard();
}

function saveInventory() {
    localStorage.setItem('inventory', JSON.stringify(inventory));
    updateDashboard();
}

function loadInvoices() {
    const data = localStorage.getItem('invoices');
    invoices = data ? JSON.parse(data) : [];
    renderInvoiceHistoryTable();
    updateDashboard();
}

function saveInvoices() {
    localStorage.setItem('invoices', JSON.stringify(invoices));
    populateCustomerList();
    updateDashboard();
}

function loadTransactions() {
    const data = localStorage.getItem('transactions');
    transactions = data ? JSON.parse(data) : [];
    updateDashboard();
}

function saveTransactions() {
    localStorage.setItem('transactions', JSON.stringify(transactions));
    updateDashboard();
}

// Navigation
function switchSection(sectionId) {
    document.querySelectorAll('.section').forEach(s => s.classList.remove('active'));
    document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));

    document.getElementById(sectionId).classList.add('active');
    event.target.classList.add('active');

    if (sectionId === 'dashboard') loadDashboard();
    if (sectionId === 'inventory') renderInventoryTable();
    if (sectionId === 'transactions') {
        initInvoiceForm();
        renderInvoiceHistoryTable();
    }
}

// Dashboard Functions
function loadDashboard() {
    updateDashboardStats();
    renderLowStockTable();
    renderTransactionChart();
}

function updateDashboardStats() {
    const today = new Date().toDateString();
    const yesterday = new Date(Date.now() - 86400000).toDateString();

    const todayInvoices = invoices.filter(inv => new Date(inv.tanggal).toDateString() === today);
    const yesterdayInvoices = invoices.filter(inv => new Date(inv.tanggal).toDateString() === yesterday);

    console.log('Today:', today, 'Today Invoices:', todayInvoices.length, 'Yesterday Invoices:', yesterdayInvoices.length);

    const todayValue = todayInvoices.reduce((sum, inv) => sum + (inv.totalNilai || 0), 0);
    const yesterdayValue = yesterdayInvoices.reduce((sum, inv) => sum + (inv.totalNilai || 0), 0);

    // Calculate profit: sales - cost of goods sold
    const todayProfit = todayInvoices.reduce((sum, inv) => {
        return sum + inv.items.reduce((itemSum, item) => {
            const product = inventory.find(p => p.nama.toLowerCase() === item.nama.toLowerCase());
            const cost = product ? product.hargaBeli * item.qty : 0;
            return itemSum + (item.hargaSatuan * item.qty - cost);
        }, 0);
    }, 0);

    const yesterdayProfit = yesterdayInvoices.reduce((sum, inv) => {
        return sum + inv.items.reduce((itemSum, item) => {
            const product = inventory.find(p => p.nama.toLowerCase() === item.nama.toLowerCase());
            const cost = product ? product.hargaBeli * item.qty : 0;
            return itemSum + (item.hargaSatuan * item.qty - cost);
        }, 0);
    }, 0);

    // Calculate total value of incoming stock purchases today
    const todayIncomingValue = inventory.reduce((sum, item) => {
        if (item.incomingDate && new Date(item.incomingDate).toDateString() === today) {
            sum += item.incomingValue || 0;
        }
        if (item.lastIncomingDate && new Date(item.lastIncomingDate).toDateString() === today) {
            sum += item.lastIncomingValue || 0;
        }
        if (!item.lastIncomingDate && item.lastAdjustmentDate && new Date(item.lastAdjustmentDate).toDateString() === today && item.lastAdjustmentValue > 0) {
            sum += item.lastAdjustmentValue || 0;
        }
        return sum;
    }, 0);

    const lowStockCount = inventory.filter(item => item.qty < 5).length;

    document.getElementById('todayValue').textContent = formatCurrency(todayValue);
    document.getElementById('yesterdayValue').textContent = formatCurrency(yesterdayValue);
    document.getElementById('totalProducts').textContent = inventory.length;
    document.getElementById('lowStockProducts').textContent = lowStockCount;
    document.getElementById('todayIncomingValue').textContent = formatCurrency(todayIncomingValue);
    document.getElementById('todayProfit').textContent = formatCurrency(todayProfit);
    document.getElementById('yesterdayProfit').textContent = formatCurrency(yesterdayProfit);

    console.log('Today Profit:', todayProfit, 'Yesterday Profit:', yesterdayProfit);
}

function renderLowStockTable() {
    const lowStock = inventory.filter(item => {
        if (item.originalQty !== undefined && item.originalQty !== null) {
            return item.qty < item.originalQty;
        }
        return item.qty < 5;
    }).sort((a, b) => a.qty - b.qty);
    const tbody = document.getElementById('lowStockTable');

    if (lowStock.length === 0) {
        tbody.innerHTML = '<tr><td colspan="7" class="text-center no-data">Tidak ada produk yang perlu dibeli kembali</td></tr>';
        return;
    }

    tbody.innerHTML = lowStock.map((item, idx) => {
        const originalQty = item.originalQty !== undefined && item.originalQty !== null ? item.originalQty : item.qty;
        const restockQty = Math.max(0, originalQty - item.qty);
        const restockLabel = restockQty > 0 ? `<span class="badge badge-warning">Beli ${restockQty}</span>` : '-';

        return `
        <tr class="low-stock">
            <td>${idx + 1}</td>
            <td>${item.nama}</td>
            <td>${item.kategori}</td>
            <td><span class="badge badge-danger">${item.qty}</span></td>
            <td>${restockLabel}</td>
            <td>${formatCurrency(item.hargaBeli)}</td>
            <td>${formatCurrency(restockQty * item.hargaBeli)}</td>
        </tr>
    `;
    }).join('');
}

function renderTransactionChart() {
    const ctx = document.getElementById('transactionChart');
    if (!ctx) return;

    const last7Days = getLast7Days();
    
    // Calculate daily invoice values (sales)
    const invoiceData = last7Days.map(date => {
        const dayInvoices = invoices.filter(inv => new Date(inv.tanggal).toDateString() === date.toDateString());
        return dayInvoices.reduce((sum, inv) => sum + (inv.totalNilai || 0), 0);
    });

    // Calculate daily incoming values (purchases)
    const incomingData = last7Days.map(date => {
        const today = date.toDateString();
        let dailyIncoming = 0;
        
        // New products added on this day
        dailyIncoming += inventory
            .filter(item => item.dateAdded && new Date(item.dateAdded).toDateString() === today)
            .reduce((sum, item) => sum + (item.hargaBeli * item.qty), 0);
        
        // Quantity adjustments on this day
        dailyIncoming += inventory
            .filter(item => item.lastAdjustmentDate && new Date(item.lastAdjustmentDate).toDateString() === today && item.lastAdjustmentValue > 0)
            .reduce((sum, item) => sum + (item.lastAdjustmentValue || 0), 0);
        
        return dailyIncoming;
    });

    const labels = last7Days.map(d => d.toLocaleDateString('id-ID', { month: 'short', day: 'numeric' }));

    if (window.transactionChartInstance) {
        window.transactionChartInstance.destroy();
    }

    window.transactionChartInstance = new Chart(ctx, {
        type: 'line',
        data: {
            labels: labels,
            datasets: [
                {
                    label: 'Nilai Penjualan (Invoice)',
                    data: invoiceData,
                    borderColor: '#3fb950',
                    backgroundColor: 'rgba(63, 185, 80, 0.1)',
                    borderWidth: 3,
                    fill: true,
                    tension: 0.4,
                    pointRadius: 5,
                    pointBackgroundColor: '#3fb950',
                    pointBorderColor: '#fff',
                    pointBorderWidth: 2
                },
                {
                    label: 'Nilai Masuk (Pembelian)',
                    data: incomingData,
                    borderColor: '#f85149',
                    backgroundColor: 'rgba(248, 81, 73, 0.1)',
                    borderWidth: 3,
                    fill: true,
                    tension: 0.4,
                    pointRadius: 5,
                    pointBackgroundColor: '#f85149',
                    pointBorderColor: '#fff',
                    pointBorderWidth: 2
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            interaction: {
                mode: 'index',
                intersect: false
            },
            plugins: {
                legend: {
                    display: true,
                    labels: { 
                        font: { size: 12 },
                        padding: 15,
                        usePointStyle: true
                    }
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    min: 0,
                    max: 3000000,
                    ticks: {
                        callback: function(value) {
                            return 'Rp ' + (value / 1000000).toFixed(1) + 'M';
                        },
                        stepSize: 500000
                    },
                    grid: {
                        color: 'rgba(255, 255, 255, 0.1)'
                    }
                },
                x: {
                    grid: {
                        display: false
                    }
                }
            }
        }
    });
}

function getLast7Days() {
    const days = [];
    for (let i = 6; i >= 0; i--) {
        const date = new Date(Date.now() - i * 86400000);
        days.push(date);
    }
    return days;
}

// Inventory Functions
function renderInventoryTable(itemsToRender = null) {
    const items = itemsToRender || inventory;
    const tbody = document.getElementById('inventoryTable');

    if (items.length === 0) {
        tbody.innerHTML = '<tr><td colspan="8" class="text-center no-data">Tidak ada data inventori. Silakan upload file Excel terlebih dahulu.</td></tr>';
        return;
    }

    tbody.innerHTML = items.map((item, idx) => {
        const badges = [];
        if (item.isNew) badges.push('<span class="badge badge-warning">BARU</span>');
        if (item.lastMovement === 'masuk') badges.push('<span class="badge badge-success">MASUK</span>');
        if (item.lastMovement === 'keluar') badges.push('<span class="badge badge-danger">KELUAR</span>');
        if (item.qty < 5) badges.push('<span class="badge badge-warning">KURANG</span>');

        return `
            <tr>
                <td>${item.nomor}</td>
                <td>${item.nama} ${badges.length ? '<div>' + badges.join(' ') + '</div>' : ''}</td>
                <td>${item.kategori}</td>
                <td>${item.qty}</td>
                <td>${formatCurrency(item.hargaBeli)}</td>
                <td>${formatCurrency(item.hargaJual)}</td>
                <td>${formatCurrency(item.total)}</td>
                <td>
                    <div class="btn-group">
                        <button class="btn btn-sm btn-primary" onclick="openEditItemModal(${item.nomor})">Edit</button>
                        <button class="btn btn-sm btn-danger" onclick="deleteItem(${item.nomor})">Hapus</button>
                        ${item.priceChangeIndicator === 'up' ? '<span class="badge badge-danger">⬆️</span>' : ''}
                        ${item.priceChangeIndicator === 'down' ? '<span class="badge badge-success">⬇️</span>' : ''}
                    </div>
                </td>
            </tr>
        `;
    }).join('');
}

function updateDashboard() {
    updateDashboardStats();
    renderLowStockTable();
    renderTransactionChart();
}

function openAddItemModal() {
    document.getElementById('addItemModal').classList.add('show');
    // Set default value to 'umum'
    document.getElementById('newItemCategory').value = 'umum';
    populateCategoryDropdown('newItemCategory');
}

function closeAddItemModal() {
    document.getElementById('addItemModal').classList.remove('show');
    document.getElementById('addItemForm').reset();
}

function handleAddItem(e) {
    e.preventDefault();

    const newNomor = inventory.length > 0 ? Math.max(...inventory.map(i => i.nomor)) + 1 : 1;
    const kategoriValue = document.getElementById('newItemCategory').value.trim() || 'umum';
    const newItemName = document.getElementById('newItemName').value.trim();

    // Check for duplicate product name
    if (inventory.some(item => item.nama.toLowerCase() === newItemName.toLowerCase())) {
        showAlert('inventoryAlert', 'error', 'Produk dengan nama ini sudah ada! Silakan gunakan nama yang berbeda.');
        return;
    }

    const qty = parseInt(document.getElementById('newItemQty').value);
    const hargaBeli = parseFloat(document.getElementById('newItemBuyPrice').value);
    const now = new Date().toISOString();

    const newItem = {
        nomor: newNomor,
        nama: newItemName,
        kategori: kategoriValue,
        hargaBeli: hargaBeli,
        hargaJual: parseFloat(document.getElementById('newItemSellPrice').value),
        qty: qty,
        originalQty: qty,
        incomingValue: hargaBeli * qty,
        incomingDate: now,
        isNew: true,
        lastMovement: 'masuk',
        dateAdded: now
    };

    newItem.total = newItem.hargaBeli * newItem.qty;

    inventory.push(newItem);
    saveInventory();
    renderInventoryTable();
    populateProductDropdowns();
    populateCategoryDropdown('newItemCategory');
    closeAddItemModal();
    showAlert('inventoryAlert', 'success', 'Produk baru berhasil ditambahkan!');
}

function openEditItemModal(nomor) {
    const item = inventory.find(i => i.nomor === nomor);
    if (!item) return;

    document.getElementById('editItemId').value = nomor;
    document.getElementById('editItemName').value = item.nama;
    document.getElementById('editItemCategory').value = item.kategori;
    document.getElementById('editItemBuyPrice').value = item.hargaBeli;
    document.getElementById('editItemSellPrice').value = item.hargaJual;
    document.getElementById('editItemQty').value = item.qty;

    document.getElementById('editItemModal').classList.add('show');
}

function closeEditItemModal() {
    document.getElementById('editItemModal').classList.remove('show');
    document.getElementById('editItemForm').reset();
}

function handleEditItem(e) {
    e.preventDefault();

    const nomor = parseInt(document.getElementById('editItemId').value);
    const item = inventory.find(i => i.nomor === nomor);

    if (item) {
        const oldQty = item.qty;
        const oldHargaBeli = item.hargaBeli;
        const newHargaBeli = parseFloat(document.getElementById('editItemBuyPrice').value);
        const newQty = parseInt(document.getElementById('editItemQty').value);
        
        item.nama = document.getElementById('editItemName').value;
        item.hargaBeli = newHargaBeli;
        item.hargaJual = parseFloat(document.getElementById('editItemSellPrice').value);
        item.qty = newQty;
        item.total = item.hargaBeli * item.qty;
        
        const qtyDiff = newQty - oldQty;
        if (qtyDiff > 0) {
            item.lastIncomingDate = new Date().toISOString();
            item.lastIncomingValue = qtyDiff * newHargaBeli;
        }

        if (newQty !== oldQty) {
            item.originalQty = newQty;
        }

        // Track price change indicator
        if (newHargaBeli > oldHargaBeli) {
            item.priceChangeIndicator = 'up';
        } else if (newHargaBeli < oldHargaBeli) {
            item.priceChangeIndicator = 'down';
        } else {
            item.priceChangeIndicator = '';
        }

        // Track adjustment value if qty or price changed
        if (qtyDiff !== 0 || newHargaBeli !== oldHargaBeli) {
            item.lastAdjustmentDate = new Date().toISOString();
            // Adjustment value = qty increase × new price
            item.lastAdjustmentValue = Math.max(0, qtyDiff) * newHargaBeli;
        }

        saveInventory();
        renderInventoryTable();
        closeEditItemModal();
        showAlert('inventoryAlert', 'success', 'Produk berhasil diperbarui!');
    }
}

function deleteItem(nomor) {
    if (confirm('Apakah Anda yakin ingin menghapus produk ini?')) {
        inventory = inventory.filter(i => i.nomor !== nomor);
        saveInventory();
        renderInventoryTable();
        showAlert('inventoryAlert', 'success', 'Produk berhasil dihapus!');
    }
}

function sortInventory(order) {
    const sorted = [...inventory].sort((a, b) => {
        if (order === 'asc') {
            return a.kategori.localeCompare(b.kategori);
        } else {
            return b.kategori.localeCompare(a.kategori);
        }
    });
    renderInventoryTable(sorted);
}

function handleSearch() {
    const query = document.getElementById('searchInput').value.toLowerCase();
    const filtered = inventory.filter(item => item.nama.toLowerCase().includes(query));
    renderInventoryTable(filtered);
}

// Excel Upload Functions
function handleExcelUpload() {
    const file = document.getElementById('excelFile').files[0];

    if (!file) {
        showAlert('uploadAlert', 'error', 'Silakan pilih file terlebih dahulu!');
        return;
    }

    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = e.target.result;
            const workbook = XLSX.read(data, { type: 'binary' });
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];
            const json = XLSX.utils.sheet_to_json(worksheet);

            console.log('Parsed JSON:', json);

            if (json.length === 0) {
                showAlert('uploadAlert', 'error', 'File Excel kosong atau tidak memiliki data!');
                return;
            }

            if (!json[0]) {
                showAlert('uploadAlert', 'error', 'File Excel tidak memiliki baris data!');
                return;
            }

            const headers = Object.keys(json[0]).map(key => key.trim().toLowerCase());
            console.log('Headers found:', headers);

            const requiredColumns = ['nomor', 'nama barang', 'kategori', 'harga beli', 'harga jual', 'qty', 'total'];
            const missingColumns = requiredColumns.filter(col => !headers.includes(col.toLowerCase()));

            if (missingColumns.length > 0) {
                showAlert('uploadAlert', 'error', `Kolom yang hilang: ${missingColumns.join(', ')}. Kolom yang diperlukan: ${requiredColumns.join(', ')}`);
                return;
            }

            inventory = json.map(row => {
                const getValue = (possibleKeys) => {
                    for (let key of possibleKeys) {
                        if (row[key] !== undefined) return row[key];
                    }
                    return '';
                };

                const qty = parseInt(getValue(['Qty', 'qty', 'QTY', 'Quantity', 'quantity'])) || 0;
                const hargaBeli = parseFloat(getValue(['Harga Beli', 'harga beli', 'HargaBeli', 'hargabeli'])) || 0;
                const dateAdded = new Date().toISOString();

                return {
                    nomor: parseInt(getValue(['Nomor', 'nomor', 'NO', 'no'])) || 0,
                    nama: String(getValue(['Nama Barang', 'nama barang', 'Nama', 'nama'])).trim() || '',
                    kategori: String(getValue(['Kategori', 'kategori', 'Category', 'category'])).trim() || '',
                    hargaBeli: hargaBeli,
                    hargaJual: parseFloat(getValue(['Harga Jual', 'harga jual', 'HargaJual', 'hargajual'])) || 0,
                    qty: qty,
                    originalQty: qty,
                    incomingValue: hargaBeli * qty,
                    incomingDate: dateAdded,
                    total: parseFloat(getValue(['Total', 'total', 'TOTAL'])) || 0,
                    lastMovement: 'masuk',
                    dateAdded: dateAdded
                };
            });

            console.log('Mapped inventory:', inventory);

            saveInventory();
            renderInventoryTable();
            populateProductDropdowns();
            updateDashboard();
            document.getElementById('excelFile').value = '';
            showAlert('uploadAlert', 'success', `${json.length} produk berhasil diimpor!`);
        } catch (error) {
            console.error('Upload error:', error);
            showAlert('uploadAlert', 'error', `Error membaca file: ${error.message}. Pastikan file Excel valid dan memiliki kolom yang benar.`);
        }
    };

    reader.readAsBinaryString(file);
}

function exportTemplate() {
    try {
        const data = [
            {
                'Nomor': 1,
                'Nama Barang': 'Monitor 24"',
                'Kategori': 'Elektronik',
                'Harga Beli': 2000000,
                'Harga Jual': 2500000,
                'Qty': 5,
                'Total': 10000000
            },
            {
                'Nomor': 2,
                'Nama Barang': 'Keyboard Mekanik',
                'Kategori': 'Aksesori',
                'Harga Beli': 500000,
                'Harga Jual': 750000,
                'Qty': 10,
                'Total': 5000000
            },
            {
                'Nomor': 3,
                'Nama Barang': 'Mouse Wireless',
                'Kategori': 'Aksesori',
                'Harga Beli': 150000,
                'Harga Jual': 200000,
                'Qty': 15,
                'Total': 2250000
            }
        ];

        const ws = XLSX.utils.json_to_sheet(data);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Template');

        XLSX.writeFile(wb, 'template_inventori.xlsx');
        showAlert('uploadAlert', 'success', 'Template berhasil didownload!');
    } catch (error) {
        console.error('Template export error:', error);
        showAlert('uploadAlert', 'error', 'Gagal mendownload template. Pastikan browser mendukung download file.');
        
        const csvData = 'Nomor,Nama Barang,Kategori,Harga Beli,Harga Jual,Qty,Total\n' +
            '1,"Monitor 24""",Elektronik,2000000,2500000,5,10000000\n' +
            '2,"Keyboard Mekanik",Aksesori,500000,750000,10,5000000\n' +
            '3,"Mouse Wireless",Aksesori,150000,200000,15,2250000';
        
        const blob = new Blob([csvData], { type: 'text/csv' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'template_inventori.csv';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
        
        showAlert('uploadAlert', 'warning', 'Template didownload sebagai CSV. Buka dengan Excel dan simpan sebagai .xlsx');
    }
}

function exportToExcel() {
    if (inventory.length === 0) {
        showAlert('uploadAlert', 'warning', 'Tidak ada data inventori untuk diekspor!');
        return;
    }

    const data = inventory.map(item => ({
        'Nomor': item.nomor,
        'Nama Barang': item.nama,
        'Kategori': item.kategori,
        'Harga Beli': item.hargaBeli,
        'Harga Jual': item.hargaJual,
        'Qty': item.qty,
        'Total': item.total
    }));

    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Inventori');
    XLSX.writeFile(wb, `inventori_${new Date().toLocaleDateString('id-ID')}.xlsx`);
    showAlert('uploadAlert', 'success', 'Data inventori berhasil diekspor!');
}

function exportTransactionsToExcel() {
    if (invoices.length === 0) {
        showAlert('uploadAlert', 'warning', 'Tidak ada invoice untuk diekspor!');
        return;
    }

    const data = invoices.map(inv => ({
        'No. Faktur': inv.invoiceNumber,
        'Tanggal': new Date(inv.tanggal).toLocaleDateString('id-ID'),
        'Customer': inv.customerName,
        'Item': inv.items.length,
        'Total Qty': inv.totalQty,
        'Total Nilai': inv.totalNilai
    }));

    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Invoice');
    XLSX.writeFile(wb, `invoice_${new Date().toLocaleDateString('id-ID')}.xlsx`);
    showAlert('uploadAlert', 'success', 'Data invoice berhasil diekspor!');
}

function checkFileHeaders() {
    const file = document.getElementById('excelFile').files[0];
    const resultEl = document.getElementById('headerCheckResult');

    if (!file) {
        resultEl.textContent = 'Pilih file terlebih dahulu';
        resultEl.style.color = 'red';
        return;
    }

    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = e.target.result;
            const workbook = XLSX.read(data, { type: 'binary' });
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];
            const json = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

            if (json.length === 0) {
                resultEl.textContent = 'File kosong';
                resultEl.style.color = 'red';
                return;
            }

            const headers = json[0].map(h => String(h).trim());
            resultEl.textContent = `Header ditemukan: ${headers.join(', ')}`;
            resultEl.style.color = headers.length > 0 ? 'green' : 'red';
        } catch (error) {
            resultEl.textContent = `Error: ${error.message}`;
            resultEl.style.color = 'red';
        }
    };

    reader.readAsBinaryString(file);
}

// Populate Customer List from Invoice History
function populateCustomerList() {
    // Get unique customers from invoice history, always include 'umum'
    const customers = [...new Set(invoices.map(inv => inv.customerName).filter(c => c && c !== 'umum'))].sort();
    
    // Ensure 'umum' is always first
    const allCustomers = ['umum', ...customers];
    
    // Update the datalist for customer input
    const inputList = document.getElementById('customerList');
    if (inputList) {
        inputList.innerHTML = allCustomers.map(cust => `<option value="${cust}"></option>`).join('');
    }
}

// Populate Dropdowns
function populateProductDropdowns() {
    const legacySelect = document.getElementById('transactionProduct');
    if (legacySelect) {
        const filter = document.getElementById('transactionProductSearch')?.value.toLowerCase() || '';
        const products = inventory.filter(item => item.nama.toLowerCase().includes(filter));
        legacySelect.innerHTML = '<option value="">-- Pilih Produk --</option>' + 
            products.map(item => `<option value="${item.nomor}">${item.nama} (${item.kategori})</option>`).join('');
    }

    populateCategoryFilter();
}

function populateCategoryDropdown(elementId) {
    // Get unique categories from inventory, always include 'umum'
    const categories = [...new Set(inventory.map(i => i.kategori).filter(k => k && k !== 'umum'))].sort();
    
    // Ensure 'umum' is always first
    const allCategories = ['umum', ...categories];
    
    // Update the datalist for new item category input
    const inputList = document.getElementById('categoryList');
    if (inputList) {
        inputList.innerHTML = allCategories.map(cat => `<option value="${cat}"></option>`).join('');
    }

    // For edit modal category (if exists)
    const select = document.getElementById(elementId);
    if (select && select.tagName === 'SELECT') {
        select.innerHTML = '<option value="">-- Pilih Kategori --</option>' + 
            allCategories.map(cat => `<option value="${cat}">${cat}</option>`).join('');
    }

    populateCategoryFilter();
}

function populateCategoryFilter() {
    const categoryEl = document.getElementById('categoryFilter');
    const categories = [...new Set(inventory.map(i => i.kategori))].sort();
    if (!categoryEl) return;
    categoryEl.innerHTML = '<option value="all">Semua</option>' +
        categories.map(cat => `<option value="${cat}">${cat}</option>`).join('');
}

function applyCategoryFilter() {
    const selected = document.getElementById('categoryFilter').value;
    const filtered = selected === 'all' ? inventory : inventory.filter(item => item.kategori === selected);
    renderInventoryTable(filtered);
}

// Invoice Functions
function generateInvoiceNumber() {
    const lastInvoiceNum = parseInt(localStorage.getItem('lastInvoiceNumber') || '0');
    const newNum = lastInvoiceNum + 1;
    localStorage.setItem('lastInvoiceNumber', newNum.toString());
    return 'INV-' + String(newNum).padStart(4, '0');
}

function initInvoiceForm() {
    const invoiceNumber = generateInvoiceNumber();
    document.getElementById('invoiceNumber').value = invoiceNumber;

    const now = new Date();
    const year = now.getFullYear();
    const month = String(now.getMonth() + 1).padStart(2, '0');
    const date = String(now.getDate()).padStart(2, '0');
    const hours = String(now.getHours()).padStart(2, '0');
    const minutes = String(now.getMinutes()).padStart(2, '0');
    document.getElementById('invoiceDate').value = `${year}-${month}-${date}T${hours}:${minutes}`;

    // Set default customer to 'umum'
    document.getElementById('customerName').value = 'umum';
    document.getElementById('invoiceProductSearch').value = '';
    document.getElementById('invoiceProduct').value = '';
    document.getElementById('invoiceQty').value = '1';

    currentInvoiceItems = [];
    renderInvoiceItemsTable();
    updateInvoiceTotalDisplay();

    populateInvoiceProductDropdown();
    populateCustomerList();
}

function populateInvoiceProductDropdown() {
    const select = document.getElementById('invoiceProduct');
    const filter = document.getElementById('invoiceProductSearch').value.toLowerCase();
    const products = inventory.filter(item => item.nama.toLowerCase().includes(filter));

    // Determine placeholder text based on search and results
    let placeholderText = '-- Pilih atau Cari Produk --';
    if (filter && products.length === 0) {
        placeholderText = '-- Produk tidak ditemukan --';
    } else if (filter && products.length > 0) {
        placeholderText = `-- ${products.length} produk ditemukan --`;
    }

    select.innerHTML = `<option value="">${placeholderText}</option>` + 
        products.map(item => `<option value="${item.nomor}" data-nama="${item.nama}" data-kategori="${item.kategori}" data-harga="${item.hargaJual}" data-qty="${item.qty}">${item.nama} - (${item.kategori}) - Stok: ${item.qty}</option>`).join('');
    
    // Auto-open dropdown if search has results
    if (filter && products.length > 0) {
        select.size = Math.min(products.length + 1, 8);
    } else {
        select.size = 1;
    }
}

function filterInvoiceProducts() {
    const searchInput = document.getElementById('invoiceProductSearch');
    // Add visual feedback for empty search
    if (!searchInput.value.trim()) {
        searchInput.style.borderColor = '';
    }
    populateInvoiceProductDropdown();
}

function addItemToInvoice() {
    const productSelect = document.getElementById('invoiceProduct');
    const qtyInput = document.getElementById('invoiceQty');
    
    const productId = parseInt(productSelect.value);
    const qty = parseInt(qtyInput.value);

    if (!productId || isNaN(productId)) {
        showAlert('transactionAlert', 'error', 'Pilih produk terlebih dahulu!');
        return;
    }

    if (!qty || qty <= 0) {
        showAlert('transactionAlert', 'error', 'Masukkan jumlah yang valid!');
        return;
    }

    const product = inventory.find(i => i.nomor === productId);
    if (!product) {
        showAlert('transactionAlert', 'error', 'Produk tidak ditemukan!');
        return;
    }

    if (product.qty < qty) {
        showAlert('transactionAlert', 'error', `Stok tidak cukup! Stok saat ini: ${product.qty}`);
        return;
    }

    const existingItem = currentInvoiceItems.find(i => i.nomorBarang === productId);
    if (existingItem) {
        if (product.qty < existingItem.qty + qty) {
            showAlert('transactionAlert', 'error', `Stok tidak cukup untuk penambahan! Tersedia: ${product.qty - existingItem.qty}`);
            return;
        }
        existingItem.qty += qty;
        existingItem.subtotal = existingItem.qty * existingItem.hargaSatuan;
    } else {
        currentInvoiceItems.push({
            nomorBarang: productId,
            nama: product.nama,
            kategori: product.kategori,
            qty: qty,
            hargaSatuan: product.hargaJual,
            subtotal: qty * product.hargaJual
        });
    }

    renderInvoiceItemsTable();
    updateInvoiceTotalDisplay();
    
    document.getElementById('invoiceProductSearch').value = '';
    document.getElementById('invoiceProduct').value = '';
    document.getElementById('invoiceQty').value = '1';
    populateInvoiceProductDropdown();

    showAlert('transactionAlert', 'success', 'Item berhasil ditambahkan!');
}

function removeItemFromInvoice(index) {
    currentInvoiceItems.splice(index, 1);
    renderInvoiceItemsTable();
    updateInvoiceTotalDisplay();
    showAlert('transactionAlert', 'success', 'Item berhasil dihapus!');
}

function renderInvoiceItemsTable() {
    const tbody = document.getElementById('invoiceItemsTable');

    if (currentInvoiceItems.length === 0) {
        tbody.innerHTML = '<tr><td colspan="6" class="text-center no-data">Belum ada item</td></tr>';
        return;
    }

    tbody.innerHTML = currentInvoiceItems.map((item, idx) => `
        <tr>
            <td>${item.nama}</td>
            <td>${item.kategori}</td>
            <td>${item.qty}</td>
            <td>${formatCurrency(item.hargaSatuan)}</td>
            <td>${formatCurrency(item.subtotal)}</td>
            <td>
                <button type="button" class="btn btn-sm btn-danger" onclick="removeItemFromInvoice(${idx})">Hapus</button>
            </td>
        </tr>
    `).join('');
}

function updateInvoiceTotalDisplay() {
    const total = currentInvoiceItems.reduce((sum, item) => sum + item.subtotal, 0);
    document.getElementById('invoiceTotalDisplay').textContent = formatCurrency(total);
}

function submitInvoice(event) {
    event.preventDefault();

    const invoiceNumber = document.getElementById('invoiceNumber').value;
    const customerNameInput = document.getElementById('customerName').value.trim();
    const customerName = customerNameInput || 'umum';
    const invoiceDate = document.getElementById('invoiceDate').value;

    if (currentInvoiceItems.length === 0) {
        showAlert('transactionAlert', 'error', 'Tambahkan minimal 1 item untuk membuat faktur!');
        return;
    }

    const totalQty = currentInvoiceItems.reduce((sum, item) => sum + item.qty, 0);
    const totalNilai = currentInvoiceItems.reduce((sum, item) => sum + item.subtotal, 0);

    currentInvoiceItems.forEach(item => {
        const product = inventory.find(i => i.nomor === item.nomorBarang);
        if (product) {
            product.qty -= item.qty;
            product.lastMovement = 'keluar';
            product.isNew = false;
            product.total = product.hargaBeli * product.qty;
        }
    });

    const newInvoice = {
        invoiceNumber: invoiceNumber,
        tanggal: invoiceDate,
        customerName: customerName,
        items: JSON.parse(JSON.stringify(currentInvoiceItems)),
        totalQty: totalQty,
        totalNilai: totalNilai
    };

    invoices.push(newInvoice);

    saveInventory();
    saveInvoices();

    currentInvoiceItems = [];
    renderInventoryTable();
    renderInvoiceHistoryTable();
    initInvoiceForm();

    showAlert('transactionAlert', 'success', `Faktur ${invoiceNumber} berhasil disimpan!`);
}

function resetInvoiceForm() {
    if (confirm('Bersihkan formulir dan mulai faktur baru?')) {
        currentInvoiceItems = [];
        initInvoiceForm();
        showAlert('transactionAlert', 'success', 'Formulir berhasil dibersihkan!');
    }
}

function renderInvoiceHistoryTable() {
    const tbody = document.getElementById('invoiceHistoryTable');

    if (invoices.length === 0) {
        tbody.innerHTML = '<tr><td colspan="7" class="text-center no-data">Tidak ada faktur</td></tr>';
        return;
    }

    const sorted = [...invoices].sort((a, b) => new Date(b.tanggal) - new Date(a.tanggal));

    tbody.innerHTML = sorted.map((invoice, idx) => `
        <tr>
            <td><strong style="cursor: pointer; color: #667eea; text-decoration: underline;" onclick="openInvoiceDetailModal('${invoice.invoiceNumber}')">${invoice.invoiceNumber}</strong></td>
            <td>${new Date(invoice.tanggal).toLocaleDateString('id-ID')}</td>
            <td>${invoice.customerName}</td>
            <td>${invoice.items.length} item</td>
            <td>${invoice.totalQty}</td>
            <td>${formatCurrency(invoice.totalNilai)}</td>
            <td>
                <button type="button" class="btn btn-sm btn-danger" onclick="deleteInvoice('${invoice.invoiceNumber}')">Hapus</button>
            </td>
        </tr>
    `).join('');
}

function searchInvoices() {
    const query = document.getElementById('invoiceSearchInput').value.toLowerCase();
    const filtered = invoices.filter(invoice => 
        invoice.invoiceNumber.toLowerCase().includes(query) || 
        invoice.customerName.toLowerCase().includes(query)
    );

    const tbody = document.getElementById('invoiceHistoryTable');

    if (filtered.length === 0) {
        tbody.innerHTML = '<tr><td colspan="7" class="text-center no-data">Tidak ada faktur yang cocok</td></tr>';
        return;
    }

    const sorted = [...filtered].sort((a, b) => new Date(b.tanggal) - new Date(a.tanggal));

    tbody.innerHTML = sorted.map((invoice, idx) => `
        <tr>
            <td><strong style="cursor: pointer; color: #667eea; text-decoration: underline;" onclick="openInvoiceDetailModal('${invoice.invoiceNumber}')">${invoice.invoiceNumber}</strong></td>
            <td>${new Date(invoice.tanggal).toLocaleDateString('id-ID')}</td>
            <td>${invoice.customerName}</td>
            <td>${invoice.items.length} item</td>
            <td>${invoice.totalQty}</td>
            <td>${formatCurrency(invoice.totalNilai)}</td>
            <td>
                <button type="button" class="btn btn-sm btn-danger" onclick="deleteInvoice('${invoice.invoiceNumber}')">Hapus</button>
            </td>
        </tr>
    `).join('');
}

function deleteInvoice(invoiceNumber) {
    if (confirm('Apakah Anda yakin ingin menghapus faktur ini? Stok akan dikembalikan.')) {
        const invoice = invoices.find(inv => inv.invoiceNumber === invoiceNumber);

        if (invoice) {
            invoice.items.forEach(item => {
                const product = inventory.find(i => i.nomor === item.nomorBarang);
                if (product) {
                    product.qty += item.qty;
                    product.total = product.hargaBeli * product.qty;
                }
            });

            invoices = invoices.filter(inv => inv.invoiceNumber !== invoiceNumber);

            saveInventory();
            saveInvoices();

            renderInventoryTable();
            renderInvoiceHistoryTable();
            updateDashboardStats();

            showAlert('transactionAlert', 'success', 'Faktur berhasil dihapus dan stok dikembalikan!');
        }
    }
}

function openInvoiceDetailModal(invoiceNumber) {
    const invoice = invoices.find(inv => inv.invoiceNumber === invoiceNumber);
    if (!invoice) {
        showAlert('transactionAlert', 'error', 'Faktur tidak ditemukan!');
        return;
    }

    document.getElementById('detailInvoiceNumber').textContent = invoice.invoiceNumber;
    document.getElementById('detailInvoiceDate').textContent = new Date(invoice.tanggal).toLocaleDateString('id-ID', { 
        weekday: 'long',
        year: 'numeric',
        month: 'numeric',
        day: 'numeric',
        hour: '2-digit',
        minute: '2-digit'
    });
    document.getElementById('detailCustomerName').textContent = invoice.customerName;
    document.getElementById('detailTotalItems').textContent = invoice.items.length;

    const itemsTableBody = document.getElementById('detailInvoiceItemsTable');
    itemsTableBody.innerHTML = invoice.items.map((item, idx) => `
        <tr>
            <td>${item.nama}</td>
            <td>${item.kategori}</td>
            <td>${item.qty}</td>
            <td>${formatCurrency(item.hargaSatuan)}</td>
            <td>${formatCurrency(item.subtotal)}</td>
        </tr>
    `).join('');

    document.getElementById('detailInvoiceTotal').textContent = formatCurrency(invoice.totalNilai);

    document.getElementById('invoiceDetailModal').style.display = 'block';
}

function closeInvoiceDetailModal() {
    document.getElementById('invoiceDetailModal').style.display = 'none';
}

function printThermalReceipt() {
    // Collect invoice data
    const invoiceNumber = document.getElementById('detailInvoiceNumber').textContent;
    const invoiceDate = document.getElementById('detailInvoiceDate').textContent;
    const customerName = document.getElementById('detailCustomerName').textContent;
    const totalAmount = document.getElementById('detailInvoiceTotal').textContent;
    
    // Get invoice items
    const itemsTable = document.getElementById('detailInvoiceItemsTable');
    const rows = itemsTable.querySelectorAll('tr');
    
    // Create thermal receipt HTML
    let receiptHTML = `
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="UTF-8">
            <meta http-equiv="X-UA-Compatible" content="IE=edge">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Cetak Faktur</title>
            <style>
                * {
                    margin: 0;
                    padding: 0;
                    box-sizing: border-box;
                }
                
                @media print {
                    @page {
                        size: 58mm auto;
                        margin: 0;
                        padding: 0;
                    }
                    
                    body {
                        margin: 0;
                        padding: 0;
                        width: 58mm;
                        background: white;
                    }
                }
                
                body {
                    width: 58mm;
                    margin: 0 auto;
                    padding: 0;
                    background: white;
                    font-family: 'Courier New', monospace;
                    color: #000;
                }
                
                .receipt {
                    width: 58mm;
                    padding: 5px;
                    font-size: 11px;
                    line-height: 1.2;
                }
                
                .header {
                    text-align: center;
                    font-weight: bold;
                    font-size: 13px;
                    margin-bottom: 5px;
                    border-bottom: 1px dashed #000;
                    padding-bottom: 5px;
                }
                
                .address {
                    text-align: center;
                    font-size: 9px;
                    margin-bottom: 5px;
                    border-bottom: 1px dashed #000;
                    padding-bottom: 5px;
                }
                
                .info-line {
                    display: flex;
                    justify-content: space-between;
                    font-size: 10px;
                    margin: 2px 0;
                    word-break: break-all;
                }
                
                .label {
                    font-weight: bold;
                    min-width: 60px;
                }
                
                .separator {
                    border-bottom: 1px dashed #000;
                    margin: 5px 0;
                }
                
                table {
                    width: 100%;
                    border-collapse: collapse;
                    margin: 5px 0;
                    font-size: 10px;
                }
                
                th, td {
                    padding: 2px 1px;
                    text-align: left;
                    border: none;
                }
                
                th {
                    border-bottom: 1px solid #000;
                    font-weight: bold;
                    font-size: 9px;
                }
                
                .item-name {
                    max-width: 25mm;
                    word-break: break-word;
                }
                
                .qty { text-align: center; width: 20px; }
                .price { text-align: right; width: 35px; }
                .subtotal { text-align: right; width: 35px; }
                
                .total-section {
                    text-align: right;
                    font-size: 12px;
                    font-weight: bold;
                    margin: 5px 0;
                    border-top: 1px dashed #000;
                    border-bottom: 1px dashed #000;
                    padding: 3px 0;
                }
                
                .footer {
                    text-align: center;
                    font-size: 9px;
                    margin-top: 5px;
                    color: #000;
                }
            </style>
        </head>
        <body>
            <div class="receipt">
                <div class="header">WARUNG SHINTA</div>
                <div class="address">jl. Gondrong No.1, RT.006/RW.006<br>Kenanga, Cipondoh, Kota Tangerang</div>
                
                <div class="separator"></div>
                
                <div class="info-line">
                    <span class="label">No. Faktur:</span>
                    <span>${invoiceNumber}</span>
                </div>
                <div class="info-line">
                    <span class="label">Tanggal:</span>
                    <span>${invoiceDate}</span>
                </div>
                <div class="info-line">
                    <span class="label">Customer:</span>
                    <span>${customerName}</span>
                </div>
                
                <div class="separator"></div>
                
                <table>
                    <thead>
                        <tr>
                            <th style="width: 25mm;">PRODUK</th>
                            <th class="qty">QTY</th>
                            <th class="price">HRG</th>
                            <th class="subtotal">TOTAL</th>
                        </tr>
                    </thead>
                    <tbody>
    `;
    
    // Add items to receipt
    rows.forEach(row => {
        const cells = row.querySelectorAll('td');
        if (cells.length >= 5) {
            const produk = cells[0].textContent.substring(0, 15); // Limit to 15 chars
            const qty = cells[2].textContent.trim();
            const harga = cells[3].textContent.trim();
            const subtotal = cells[4].textContent.trim();
            
            receiptHTML += `
                        <tr>
                            <td class="item-name">${produk}</td>
                            <td class="qty">${qty}</td>
                            <td class="price" style="font-size: 9px;">${harga.replace('Rp ', '')}</td>
                            <td class="subtotal" style="font-size: 9px;">${subtotal.replace('Rp ', '')}</td>
                        </tr>
            `;
        }
    });
    
    receiptHTML += `
                    </tbody>
                </table>
                
                <div class="separator"></div>
                
                <div class="total-section">
                    TOTAL: ${totalAmount}
                </div>
                
                <div class="footer">
                    Terima Kasih<br>
                    Atas Pembelian Anda
                </div>
            </div>
        </body>
        </html>
    `;
    
    // Open print dialog
    const printWindow = window.open('', '_blank');
    printWindow.document.write(receiptHTML);
    printWindow.document.close();
    
    // Auto print after content loads
    printWindow.onload = function() {
        setTimeout(() => {
            printWindow.print();
        }, 250);
    };
}

// Utility Functions
function formatCurrency(value) {
    if (value === null || value === undefined || isNaN(value)) return 'Rp 0';
    return new Intl.NumberFormat('id-ID', {
        style: 'currency',
        currency: 'IDR',
        minimumFractionDigits: 0,
        maximumFractionDigits: 0
    }).format(value);
}

function showAlert(elementId, type, message) {
    const alert = document.getElementById(elementId);
    alert.className = `alert ${type}`;
    alert.textContent = message;
    alert.style.display = 'block';

    setTimeout(() => {
        alert.style.display = 'none';
    }, 4000);
}
