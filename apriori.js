// --- UNIVERSAL APRIORI ADVANCED ENGINE ---
const fileInput = document.getElementById('fileInput');
const fileNameLabel = document.getElementById('fileNameLabel');
const rowCountStatus = document.getElementById('rowCountStatus');
const detectedColumnsInfo = document.getElementById('detectedColumnsInfo');
const runAprioriBtn = document.getElementById('runAprioriBtn');
const resultsSection = document.getElementById('resultsSection');
const previewSection = document.getElementById('previewSection');
const previewTableBody = document.getElementById('previewTableBody');
const downloadPdfBtn = document.getElementById('downloadPdfBtn');
const downloadExcelBtn = document.getElementById('downloadExcelBtn');
const filterSeason = document.getElementById('filterSeason');
const searchRuleInput = document.getElementById('searchRuleInput');
const resetDataBtn = document.getElementById('resetDataBtn');

const btnIcon = document.getElementById('btnIcon');
const btnText = document.getElementById('btnText');
const toastNotification = document.getElementById('toastNotification');
const toastTitle = document.getElementById('toastTitle');
const toastMsg = document.getElementById('toastMsg');

const inputSupport = document.getElementById('inputSupport');
const valSupport = document.getElementById('valSupport');
const inputConfidence = document.getElementById('inputConfidence');
const valConfidence = document.getElementById('valConfidence');
const inputLift = document.getElementById('inputLift');

if(inputSupport) inputSupport.oninput = () => valSupport.innerText = inputSupport.value + '%';
if(inputConfidence) inputConfidence.oninput = () => valConfidence.innerText = inputConfidence.value + '%';

let rawDataStore = [];
let processedTransactions = [];
let globalValidRules = [];
let globalItemCounts = {};
let chartTopProdInstance = null;
let chartStrengthInstance = null;
let chartFilterInstance = null;

function showToast(title, message, isError = false) {
    toastTitle.innerText = title;
    toastMsg.innerText = message;
    toastNotification.classList.remove('translate-y-20', 'opacity-0');
    if(isError) {
        toastNotification.querySelector('#toastIcon').innerHTML = `<i class="fa-solid fa-circle-exclamation text-rose-400"></i>`;
    } else {
        toastNotification.querySelector('#toastIcon').innerHTML = `<i class="fa-solid fa-circle-check text-emerald-400"></i>`;
    }
    setTimeout(() => { toastNotification.classList.add('translate-y-20', 'opacity-0'); }, 3500);
}

if(fileInput) {
    fileInput.onchange = (e) => {
        const file = e.target.files[0];
        if (!file) return;
        fileNameLabel.innerText = file.name;

        const reader = new FileReader();
        reader.onload = (event) => {
            try {
                const data = new Uint8Array(event.target.result);
                const workbook = XLSX.read(data, {type: 'array'});
                const rows = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], {header: 1});
                
                if (rows.length === 0) { showToast("Kesalahan", "File Excel kosong.", true); return; }

                let headerRowIndex = 0;
                let idColIdx = -1;
                let itemColIdx = -1;
                let dateColIdx = -1;
                let rawHeaders = rows[0].map(c => String(c).trim());

                for (let i = 0; i < Math.min(rows.length, 15); i++) {
                    let rowStr = rows[i].map(cell => String(cell).toLowerCase().trim());
                    if (idColIdx === -1) idColIdx = rowStr.findIndex(cell => /invoiceno|invoice|transaksi|id|nota|trx|nomor|no|faktur|kode_transaksi/i.test(cell));
                    if (itemColIdx === -1) itemColIdx = rowStr.findIndex(cell => /description|desc|produk|barang|item|nama|menu|product|deskripsi|nama_barang/i.test(cell));
                    if (dateColIdx === -1) dateColIdx = rowStr.findIndex(cell => /date|tanggal|waktu|time/i.test(cell));

                    if (idColIdx !== -1 && itemColIdx !== -1) {
                        headerRowIndex = i;
                        break;
                    }
                }

                if (idColIdx === -1) idColIdx = 0;
                if (itemColIdx === -1) itemColIdx = 1;

                let idColName = rawHeaders[idColIdx] || `Kolom ke-${idColIdx + 1}`;
                let itemColName = rawHeaders[itemColIdx] || `Kolom ke-${itemColIdx + 1}`;
                detectedColumnsInfo.innerHTML = `<i class="fa-solid fa-circle-check text-emerald-600 mr-1"></i> ID Transaksi: <strong>${idColName}</strong> | Produk: <strong>${itemColName}</strong>`;

                rawDataStore = [];
                for (let i = headerRowIndex + 1; i < rows.length; i++) {
                    let r = rows[i];
                    if (!r || r.length === 0) continue;

                    let trxId = String(r[idColIdx] || '').trim();
                    let itemName = String(r[itemColIdx] || '').trim();
                    let dateVal = dateColIdx !== -1 ? String(r[dateColIdx] || '').trim() : '';

                    if (trxId !== '' && itemName !== '' && itemName !== 'undefined') {
                        rawDataStore.push({ id: trxId, item: itemName, date: dateVal });
                    }
                }

                if (rawDataStore.length === 0) {
                    showToast("Kesalahan", "Tidak dapat menemukan data transaksi valid.", true);
                    return;
                }

                renderPreview(rawDataStore.slice(0, 5));
                rowCountStatus.innerText = `Total Baris Data: ${rawDataStore.length}`;
                document.getElementById('badgeDataset').innerHTML = `<i class="fa-solid fa-database mr-1"></i> ${file.name}`;
                runAprioriBtn.disabled = false;
                resetDataBtn.classList.remove('hidden');

                showToast("Berhasil", "File transaksi berhasil diunggah & diproses!");
            } catch (err) {
                console.error(err);
                showToast("Kesalahan", "Terjadi kesalahan saat membaca file.", true);
            }
        };
        reader.readAsArrayBuffer(file);
    };
}

if(resetDataBtn) {
    resetDataBtn.onclick = () => {
        fileInput.value = "";
        fileNameLabel.innerText = "Klik atau seret file ke sini (.xlsx, .csv)";
        rowCountStatus.innerText = "Total Baris: 0";
        detectedColumnsInfo.innerHTML = "Kolom Terdeteksi: -";
        previewSection.classList.add('hidden');
        resultsSection.classList.add('hidden');
        runAprioriBtn.disabled = true;
        resetDataBtn.classList.add('hidden');
        rawDataStore = [];
        document.getElementById('badgeDataset').innerHTML = `<i class="fa-solid fa-database mr-1"></i> Belum ada dataset`;
        showToast("Reset", "Dataset berhasil dihapus. Silakan unggah file baru.");
    };
}

function renderPreview(previewData) {
    if(!previewTableBody) return;
    previewSection.classList.remove('hidden');
    previewTableBody.innerHTML = previewData.map((d, idx) => `
        <tr class="hover:bg-slate-50">
            <td class="p-2 text-slate-500">${idx + 1}</td>
            <td class="p-2 font-medium text-slate-800">${d.id}</td>
            <td class="p-2 font-medium text-indigo-600">${d.item}</td>
            <td class="p-2 text-slate-500">${d.date || '-'}</td>
        </tr>
    `).join('');
}

if(runAprioriBtn) {
    runAprioriBtn.onclick = () => {
        runAprioriBtn.disabled = true;
        btnIcon.innerHTML = `<i class="fa-solid fa-spinner fa-spin"></i>`;
        btnText.innerText = "Sedang Memproses Algoritma...";

        setTimeout(() => {
            try {
                const seasonFilter = filterSeason.value;
                
                let filteredData = rawDataStore.filter(row => {
                    if (seasonFilter === 'all' || !row.date) return true;
                    let d = new Date(row.date);
                    if (isNaN(d.getTime())) return true;
                    let day = d.getDay();
                    let isWeekend = (day === 0 || day === 6);
                    if (seasonFilter === 'weekend') return isWeekend;
                    if (seasonFilter === 'weekday') return !isWeekend;
                    return true;
                });

                let grouped = {};
                filteredData.forEach(row => {
                    if (!grouped[row.id]) grouped[row.id] = new Set();
                    grouped[row.id].add(row.item);
                });

                processedTransactions = Object.keys(grouped).map(id => ({
                    transactionId: id,
                    items: Array.from(grouped[id])
                }));

                const minSupPct = parseFloat(inputSupport.value) / 100;
                const minConfPct = parseFloat(inputConfidence.value) / 100;
                const minLiftVal = parseFloat(inputLift.value);
                const totalTrans = processedTransactions.length;

                if (totalTrans === 0) {
                    showToast("Peringatan", "Tidak ada transaksi sesuai filter waktu.", true);
                    runAprioriBtn.disabled = false;
                    btnIcon.innerHTML = `<i class="fa-solid fa-play"></i>`;
                    btnText.innerText = "Jalankan Analisis Lanjutan";
                    return;
                }

                let itemCounts = {};
                processedTransactions.forEach(t => {
                    t.items.forEach(item => {
                        itemCounts[item] = (itemCounts[item] || 0) + 1;
                    });
                });
                globalItemCounts = itemCounts;

                let frequent1 = {};
                Object.keys(itemCounts).forEach(item => {
                    let sup = itemCounts[item] / totalTrans;
                    if (sup >= minSupPct) frequent1[item] = itemCounts[item];
                });

                let frequentItemsList = Object.keys(frequent1);
                let pairCounts = {};
                
                processedTransactions.forEach(t => {
                    let items = t.items.filter(i => frequent1[i] !== undefined);
                    for (let i = 0; i < items.length; i++) {
                        for (let j = i + 1; j < items.length; j++) {
                            let pair = [items[i], items[j]].sort().join(' | ');
                            pairCounts[pair] = (pairCounts[pair] || 0) + 1;
                        }
                    }
                });

                let rules = [];
                let droppedRulesCount = 0;
                let ruleIdCounter = 1;

                Object.keys(pairCounts).forEach(pairStr => {
                    let count = pairCounts[pairStr];
                    let supportAB = count / totalTrans;
                    let parts = pairStr.split(' | ');
                    let itemA = parts[0], itemB = parts[1];

                    let supportA = frequent1[itemA] / totalTrans;
                    let supportB = frequent1[itemB] / totalTrans;
                    
                    let confAtoB = supportAB / supportA;
                    let liftAtoB = confAtoB / supportB;

                    if (supportAB >= minSupPct) {
                        if (confAtoB >= minConfPct && liftAtoB >= minLiftVal) {
                            rules.push({
                                id: `R${String(ruleIdCounter++).padStart(2, '0')}`,
                                antecedent: itemA, consequent: itemB,
                                support: (supportAB * 100).toFixed(1) + '%',
                                confidence: (confAtoB * 100).toFixed(1) + '%',
                                lift: liftAtoB.toFixed(2),
                                strength: liftAtoB > 1.5 ? 'Sangat Kuat' : 'Kuat'
                            });
                        } else { droppedRulesCount++; }
                    } else { droppedRulesCount++; }

                    let confBtoA = supportAB / supportB;
                    let liftBtoA = confBtoA / supportA;

                    if (supportAB >= minSupPct) {
                        if (confBtoA >= minConfPct && liftBtoA >= minLiftVal) {
                            rules.push({
                                id: `R${String(ruleIdCounter++).padStart(2, '0')}`,
                                antecedent: itemB, consequent: itemA,
                                support: (supportAB * 100).toFixed(1) + '%',
                                confidence: (confBtoA * 100).toFixed(1) + '%',
                                lift: liftBtoA.toFixed(2),
                                strength: liftBtoA > 1.5 ? 'Sangat Kuat' : 'Kuat'
                            });
                        } else { droppedRulesCount++; }
                    }
                });

                globalValidRules = rules;

                document.getElementById('statTransactions').innerText = totalTrans;
                document.getElementById('statProducts').innerText = Object.keys(itemCounts).length;
                document.getElementById('statItemsets').innerText = frequentItemsList.length + Object.keys(pairCounts).length;
                document.getElementById('statRules').innerText = rules.length;

                renderRulesTable(rules);
                renderSmartPricingCards(rules);
                renderFrequentItemsTable(itemCounts, totalTrans);
                renderCharts(itemCounts, rules, rules.length, droppedRulesCount);
                renderMatrixTable(rules, itemCounts);

                resultsSection.classList.remove('hidden');
                resultsSection.scrollIntoView({ behavior: 'smooth' });
                showToast("Sukses", "Analisis Algoritma Apriori berhasil dijalankan!");
            } catch (err) {
                console.error(err);
                showToast("Kesalahan", "Terjadi kesalahan saat pemrosesan.", true);
            } finally {
                runAprioriBtn.disabled = false;
                btnIcon.innerHTML = `<i class="fa-solid fa-play"></i>`;
                btnText.innerText = "Jalankan Analisis Lanjutan";
            }
        }, 300);
    };
}

if(searchRuleInput) {
    searchRuleInput.oninput = (e) => {
        let keyword = e.target.value.toLowerCase();
        let filteredRules = globalValidRules.filter(r => 
            r.antecedent.toLowerCase().includes(keyword) || 
            r.consequent.toLowerCase().includes(keyword) ||
            r.id.toLowerCase().includes(keyword)
        );
        renderRulesTable(filteredRules);
    };
}

function renderMatrixTable(rules, itemCounts) {
    const matrixTable = document.getElementById('matrixTable');
    if (!matrixTable) return;

    let topItems = Object.keys(itemCounts).sort((a,b) => itemCounts[b] - itemCounts[a]).slice(0, 6);
    if (topItems.length === 0) {
        matrixTable.innerHTML = `<tr><td class="p-4 text-slate-400">Tidak ada data untuk ditampilkan.</td></tr>`;
        return;
    }

    let liftMap = {};
    rules.forEach(r => { liftMap[`${r.antecedent}_${r.consequent}`] = parseFloat(r.lift); });

    let html = `<thead><tr class="bg-slate-100 text-slate-700 font-semibold"><th class="p-2 border border-slate-200">Produk (A \ B)</th>`;
    topItems.forEach(item => {
        let shortName = item.length > 15 ? item.substring(0, 15) + '...' : item;
        html += `<th class="p-2 border border-slate-200" title="${item}">${shortName}</th>`;
    });
    html += `</tr></thead><tbody>`;

    topItems.forEach(rowItem => {
        let shortRow = rowItem.length > 15 ? rowItem.substring(0, 15) + '...' : rowItem;
        html += `<tr><td class="p-2 font-semibold bg-slate-50 border border-slate-200 text-left" title="${rowItem}">${shortRow}</td>`;
        
        topItems.forEach(colItem => {
            if (rowItem === colItem) {
                html += `<td class="p-2 border border-slate-200 bg-slate-100 text-slate-400">-</td>`;
            } else {
                let liftVal = liftMap[`${rowItem}_${colItem}`] || liftMap[`${colItem}_${rowItem}`];
                if (liftVal) {
                    let bgColor = liftVal > 1.5 ? 'bg-emerald-100 text-emerald-800 font-bold' : 'bg-blue-50 text-blue-700';
                    html += `<td class="p-2 border border-slate-200 ${bgColor}">${liftVal.toFixed(2)}</td>`;
                } else {
                    html += `<td class="p-2 border border-slate-200 text-slate-300">0.00</td>`;
                }
            }
        });
        html += `</tr>`;
    });

    html += `</tbody>`;
    matrixTable.innerHTML = html;
}

function renderSmartPricingCards(rules) {
    const grid = document.getElementById('optimizationGrid');
    if (!grid) return;
    if (rules.length === 0) {
        grid.innerHTML = `<p class="text-xs text-slate-400 col-span-3 text-center">Belum ada rekomendasi.</p>`;
        return;
    }

    let topRules = [...rules].sort((a,b) => parseFloat(b.lift) - parseFloat(a.lift)).slice(0, 3);

    const bundleStrategies = [
        {
            badgeColor: "bg-indigo-600",
            discount: "15%",
            strategyTitle: "Strategi Display Etalase Utama",
            advice: "Letakkan kedua produk secara berdampingan di rak depan toko atau halaman utama website untuk menarik minat beli paket."
        },
        {
            badgeColor: "bg-purple-600",
            discount: "12%",
            strategyTitle: "Strategi Cross-Selling Kasir",
            advice: "Tawarkan produk kedua sebagai pelengkap (*add-on*) saat pelanggan hendak melakukan pembayaran untuk menaikkan total nilai belanja."
        },
        {
            badgeColor: "bg-blue-600",
            discount: "10%",
            strategyTitle: "Strategi Flash Sale / Bundling Terbatas",
            advice: "Jadikan paket ini sebagai penawaran khusus di akhir pekan atau hari libur guna mempercepat putaran stok barang."
        }
    ];

    grid.innerHTML = topRules.map((r, idx) => {
        const currentStrat = bundleStrategies[idx] || bundleStrategies[0];

        return `
            <div class="bg-slate-50 p-5 rounded-xl border border-slate-200 flex flex-col justify-between space-y-4 shadow-sm hover:shadow-md transition">
                <div>
                    <div class="flex justify-between items-center mb-2">
                        <span class="${currentStrat.badgeColor} text-white text-[10px] px-2.5 py-1 rounded-full font-bold">Smart Bundle #${idx+1}</span>
                        <span class="text-xs text-emerald-700 bg-emerald-50 px-2 py-0.5 rounded font-bold border border-emerald-200">Diskon Paket: ${currentStrat.discount}</span>
                    </div>
                    <h4 class="font-bold text-slate-900 text-sm mb-1">${r.antecedent} + ${r.consequent}</h4>
                    <p class="text-xs text-slate-600 mb-3">Keyakinan <strong class="text-indigo-600">${r.confidence}</strong>, Lift Ratio <strong class="text-slate-800">${r.lift}</strong>.</p>
                </div>
                <div class="bg-white p-3 rounded-xl border border-slate-200 text-xs space-y-2">
                    <p class="font-bold text-slate-800 flex items-center"><i class="fa-solid fa-bullseye text-indigo-600 mr-1.5"></i> ${currentStrat.strategyTitle}:</p>
                    <p class="text-slate-600 leading-relaxed">${currentStrat.advice}</p>
                </div>
            </div>
        `;
    }).join('');
}

// Cetak PDF via Browser (window.print)
if(downloadPdfBtn) {
    downloadPdfBtn.onclick = () => {
        if (resultsSection.classList.contains('hidden')) {
            showToast("Peringatan", "Jalankan analisis terlebih dahulu sebelum mencetak laporan.", true);
            return;
        }
        window.print();
    };
}

if(downloadExcelBtn) {
    downloadExcelBtn.onclick = () => {
        if (globalValidRules.length === 0) {
            showToast("Peringatan", "Tidak ada data aturan untuk diekspor.", true);
            return;
        }
        let exportData = globalValidRules.map(r => ({
            "ID Aturan": r.id,
            "Jika Membeli (Premis)": r.antecedent,
            "Maka Membeli (Konsekuensi)": r.consequent,
            "Support": r.support,
            "Confidence": r.confidence,
            "Lift Ratio": r.lift,
            "Status Kekuatan": r.strength
        }));
        let worksheet = XLSX.utils.json_to_sheet(exportData);
        let workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, "Aturan Asosiasi");
        XLSX.writeFile(workbook, "Laporan_Aturan_Apriori.xlsx");
        showToast("Sukses", "Laporan berhasil diekspor ke Excel!");
    };
}

function renderCharts(itemCounts, rules, validCount, droppedCount) {
    let sortedItems = Object.keys(itemCounts).map(item => ({
        name: item.length > 20 ? item.substring(0, 20) + '...' : item,
        count: itemCounts[item]
    })).sort((a, b) => b.count - a.count).slice(0, 5);

    const ctxTop = document.getElementById('chartTopProducts').getContext('2d');
    if (chartTopProdInstance) chartTopProdInstance.destroy();
    chartTopProdInstance = new Chart(ctxTop, {
        type: 'bar',
        data: {
            labels: sortedItems.map(i => i.name),
            datasets: [{ data: sortedItems.map(i => i.count), backgroundColor: '#4f46e5', borderRadius: 6 }]
        },
        options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } } }
    });

    let sangatKuatCount = rules.filter(r => r.strength === 'Sangat Kuat').length;
    let kuatCount = rules.filter(r => r.strength === 'Kuat').length;

    const ctxPie = document.getElementById('chartRuleStrength').getContext('2d');
    if (chartStrengthInstance) chartStrengthInstance.destroy();
    chartStrengthInstance = new Chart(ctxPie, {
        type: 'doughnut',
        data: {
            labels: ['Sangat Kuat', 'Kuat'],
            datasets: [{ data: [sangatKuatCount, kuatCount], backgroundColor: ['#10b981', '#3b82f6'] }]
        },
        options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'bottom' } } }
    });

    const ctxFilter = document.getElementById('chartFilterStatus').getContext('2d');
    if (chartFilterInstance) chartFilterInstance.destroy();
    chartFilterInstance = new Chart(ctxFilter, {
        type: 'bar',
        data: {
            labels: ['Sesuai', 'Tidak Sesuai'],
            datasets: [{ data: [validCount, droppedCount], backgroundColor: ['#10b981', '#ef4444'], borderRadius: 6 }]
        },
        options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } } }
    });
}

function renderFrequentItemsTable(itemCounts, totalTrans) {
    const tbody = document.getElementById('freqTableBody');
    if (!tbody) return;
    let sortedItems = Object.keys(itemCounts).map(item => ({
        name: item, count: itemCounts[item],
        support: ((itemCounts[item] / totalTrans) * 100).toFixed(1) + '%'
    })).sort((a, b) => b.count - a.count).slice(0, 10);

    tbody.innerHTML = sortedItems.map((item, idx) => `
        <tr class="hover:bg-slate-50 transition">
            <td class="p-3 font-semibold text-slate-700">#${idx + 1}</td>
            <td class="p-3 font-medium text-slate-900">${item.name}</td>
            <td class="p-3 text-slate-600">${item.count} Transaksi</td>
            <td class="p-3 font-semibold text-indigo-600">${item.support}</td>
        </tr>
    `).join('');
}

function renderRulesTable(rules) {
    const tbody = document.getElementById('rulesTableBody');
    if (!tbody) return;
    if (rules.length === 0) {
        tbody.innerHTML = `<tr><td colspan="7" class="text-center py-6 text-slate-400">Tidak ada aturan yang ditemukan.</td></tr>`;
        return;
    }
    tbody.innerHTML = rules.map(r => `
        <tr class="hover:bg-slate-50 transition">
            <td class="p-3 font-semibold text-indigo-600">${r.id}</td>
            <td class="p-3 font-medium text-slate-800">${r.antecedent}</td>
            <td class="p-3 font-medium text-slate-800">${r.consequent}</td>
            <td class="p-3 text-slate-600">${r.support}</td>
            <td class="p-3 text-slate-600 font-semibold text-indigo-600">${r.confidence}</td>
            <td class="p-3 text-slate-600">${r.lift}</td>
            <td class="p-3"><span class="px-2.5 py-1 rounded-full text-[10px] font-bold ${r.strength === 'Sangat Kuat' ? 'bg-emerald-100 text-emerald-700' : 'bg-blue-100 text-blue-700'}">${r.strength}</span></td>
        </tr>
    `).join('');
}

document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.onclick = (e) => {
        document.querySelectorAll('.tab-btn').forEach(b => {
            b.classList.remove('text-indigo-600', 'border-b-2', 'border-indigo-600', 'active');
            b.classList.add('text-slate-500');
        });
        document.querySelectorAll('.tab-content').forEach(c => c.classList.add('hidden'));

        e.target.classList.add('text-indigo-600', 'border-b-2', 'border-indigo-600', 'active');
        e.target.classList.remove('text-slate-500');
        const targetEl = document.getElementById(e.target.dataset.target);
        if(targetEl) targetEl.classList.remove('hidden');
    };
});