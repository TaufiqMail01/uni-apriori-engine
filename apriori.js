// --- SMART OFFICE ATK & TESTING ENGINE ---
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
const filterDivision = document.getElementById('filterDivision');
const searchRuleInput = document.getElementById('searchRuleInput');
const resetDataBtn = document.getElementById('resetDataBtn');
const executiveSummaryText = document.getElementById('executiveSummaryText');

const btnIcon = document.getElementById('btnIcon');
const btnText = document.getElementById('btnText');
const toastNotification = document.getElementById('toastNotification');
const toastTitle = document.getElementById('toastTitle');
const toastMsg = document.getElementById('toastMsg');

const inputSupport = document.getElementById('inputSupport');
const valSupport = document.getElementById('valSupport');
const inputConfidence = document.getElementById('inputConfidence');
const valConfidence = document.getElementById('valConfidence');

if(inputSupport) inputSupport.oninput = () => valSupport.innerText = inputSupport.value + '%';
if(inputConfidence) inputConfidence.oninput = () => valConfidence.innerText = inputConfidence.value + '%';

let rawDataStore = [];
let processedTransactions = [];
let globalValidRules = [];
let globalItemCounts = {};
let globalDivisionStats = {};
let chartTopDivInstance = null;
let chartTopProdInstance = null;
let chartStrengthInstance = null;

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
                
                if (rows.length === 0) { showToast("Error", "File Excel kosong.", true); return; }

                let headerRowIndex = 0;
                let idColIdx = -1;
                let itemColIdx = -1;
                let divColIdx = -1;
                let qtyColIdx = -1;
                let rawHeaders = rows[0].map(c => String(c).trim());

                for (let i = 0; i < Math.min(rows.length, 15); i++) {
                    let rowStr = rows[i].map(cell => String(cell).toLowerCase().trim());
                    if (idColIdx === -1) idColIdx = rowStr.findIndex(cell => /invoiceno|invoice|transaksi|id|nota|trx|nomor|no|faktur|kode_transaksi/i.test(cell));
                    if (itemColIdx === -1) itemColIdx = rowStr.findIndex(cell => /nama_barang|deskripsi|description|product|item|menu|nama\s+produk|nama\s+barang/i.test(cell));
                    if (divColIdx === -1) divColIdx = rowStr.findIndex(cell => /cabang|unit|divisi|department|dept|branch/i.test(cell));
                    if (qtyColIdx === -1) qtyColIdx = rowStr.findIndex(cell => /jumlah|qty|quantity|total/i.test(cell));
                    if (idColIdx !== -1 && itemColIdx !== -1) { headerRowIndex = i; break; }
                }

                if (idColIdx === -1) idColIdx = 0;
                if (itemColIdx === -1) itemColIdx = 1;
                if (divColIdx === -1) divColIdx = 4;

                let idColName = rawHeaders[idColIdx] || `Col ${idColIdx + 1}`;
                let itemColName = rawHeaders[itemColIdx] || `Col ${itemColIdx + 1}`;
                let divColName = rawHeaders[divColIdx] || `Col ${divColIdx + 1}`;

                detectedColumnsInfo.innerHTML = `ID: <strong>${idColName}</strong> | Unit: <strong>${divColName}</strong> | Item: <strong>${itemColName}</strong>`;

                rawDataStore = [];
                let uniqueDivisions = new Set();

                for (let i = headerRowIndex + 1; i < rows.length; i++) {
                    let r = rows[i];
                    if (!r || r.length === 0) continue;

                    let trxId = String(r[idColIdx] || '').trim();
                    let itemName = String(r[itemColIdx] || '').trim();
                    let divName = divColIdx !== -1 ? String(r[divColIdx] || 'Umum').trim() : 'Umum';
                    let qtyVal = qtyColIdx !== -1 ? parseFloat(r[qtyColIdx]) || 1 : 1;

                    if (trxId !== '' && itemName !== '' && itemName !== 'undefined') {
                        rawDataStore.push({ id: trxId, item: itemName, division: divName, qty: qtyVal });
                        uniqueDivisions.add(divName);
                    }
                }

                if (rawDataStore.length === 0) {
                    showToast("Error", "Tidak ada data valid.", true);
                    return;
                }

                if(filterDivision) {
                    filterDivision.innerHTML = `<option value="all">Semua Divisi (Gabungan)</option>`;
                    Array.from(uniqueDivisions).sort().forEach(div => {
                        let opt = document.createElement('option');
                        opt.value = div;
                        opt.innerText = `Divisi / Unit: ${div}`;
                        filterDivision.appendChild(opt);
                    });
                }

                let totalUniqueTrx = new Set(rawDataStore.map(d => d.id)).size;

                inputSupport.value = 5;
                valSupport.innerText = "5%";
                inputConfidence.value = 40;
                valConfidence.innerText = "40%";

                renderPreview(rawDataStore.slice(0, 5));
                rowCountStatus.innerText = `Total Baris: ${rawDataStore.length} (${totalUniqueTrx} Nota)`;
                document.getElementById('badgeDataset').innerHTML = `<i class="fa-solid fa-database mr-1"></i> ${file.name}`;
                runAprioriBtn.disabled = false;
                resetDataBtn.classList.remove('hidden');

                showToast("Success", `Dataset berhasil dimuat!`);
            } catch (err) {
                console.error(err);
                showToast("Error", "Gagal membaca file Excel.", true);
            }
        };
        reader.readAsArrayBuffer(file);
    };
}

if(resetDataBtn) {
    resetDataBtn.onclick = () => {
        fileInput.value = "";
        fileNameLabel.innerText = "Pilih file .xlsx / .csv";
        rowCountStatus.innerText = "Total Baris: 0";
        detectedColumnsInfo.innerHTML = "Kolom: -";
        previewSection.classList.add('hidden');
        resultsSection.classList.add('hidden');
        runAprioriBtn.disabled = true;
        resetDataBtn.classList.add('hidden');
        rawDataStore = [];
        document.getElementById('badgeDataset').innerHTML = `<i class="fa-solid fa-database mr-1"></i> Belum ada dataset`;
        showToast("Reset", "Data dibersihkan.");
    };
}

function renderPreview(previewData) {
    if(!previewTableBody) return;
    previewSection.classList.remove('hidden');
    previewTableBody.innerHTML = previewData.map((d, idx) => `
        <tr class="hover:bg-slate-50">
            <td class="p-2 border border-slate-200 text-slate-500">${idx + 1}</td>
            <td class="p-2 border border-slate-200 font-medium text-slate-800">${d.id}</td>
            <td class="p-2 border border-slate-200 font-bold text-indigo-600">${d.division}</td>
            <td class="p-2 border border-slate-200 font-medium text-slate-800">${d.item}</td>
            <td class="p-2 border border-slate-200 text-slate-600">${d.qty}</td>
        </tr>
    `).join('');
}

if(runAprioriBtn) {
    runAprioriBtn.onclick = () => {
        runAprioriBtn.disabled = true;
        btnIcon.innerHTML = `<i class="fa-solid fa-spinner fa-spin"></i>`;
        btnText.innerText = "Menjalankan Pengujian...";

        setTimeout(() => {
            try {
                const selectedDiv = filterDivision.value;
                let filteredData = rawDataStore.filter(row => selectedDiv === 'all' || row.division === selectedDiv);

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
                const totalTrans = processedTransactions.length;

                if (totalTrans === 0) {
                    showToast("Warning", "Tidak ada transaksi pada unit ini.", true);
                    runAprioriBtn.disabled = false;
                    btnIcon.innerHTML = `<i class="fa-solid fa-play"></i>`;
                    btnText.innerText = "3. Jalankan Pengujian Model";
                    return;
                }

                let itemCounts = {};
                let divisionCounts = {};
                let divisionItemQty = {};

                filteredData.forEach(row => {
                    itemCounts[row.item] = (itemCounts[row.item] || 0) + 1;
                    divisionCounts[row.division] = (divisionCounts[row.division] || 0) + 1;

                    if (!divisionItemQty[row.division]) divisionItemQty[row.division] = { totalQty: 0, items: {} };
                    divisionItemQty[row.division].totalQty += row.qty;
                    divisionItemQty[row.division].items[row.item] = (divisionItemQty[row.division].items[row.item] || 0) + row.qty;
                });

                globalItemCounts = itemCounts;
                globalDivisionStats = divisionItemQty;

                let frequent1 = {};
                Object.keys(itemCounts).forEach(item => {
                    let sup = itemCounts[item] / totalTrans;
                    if (sup >= minSupPct) frequent1[item] = itemCounts[item];
                });

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

                    if (supportAB >= minSupPct && confAtoB >= minConfPct && liftAtoB >= 1.0) {
                        rules.push({
                            id: `R${String(ruleIdCounter++).padStart(2, '0')}`,
                            antecedent: itemA, consequent: itemB,
                            support: (supportAB * 100).toFixed(1) + '%',
                            confidence: (confAtoB * 100).toFixed(1) + '%',
                            lift: liftAtoB.toFixed(2),
                            strength: liftAtoB > 1.5 ? 'Sangat Kuat' : 'Kuat'
                        });
                    }

                    let confBtoA = supportAB / supportB;
                    let liftBtoA = confBtoA / supportA;

                    if (supportAB >= minSupPct && confBtoA >= minConfPct && liftBtoA >= 1.0) {
                        rules.push({
                            id: `R${String(ruleIdCounter++).padStart(2, '0')}`,
                            antecedent: itemB, consequent: itemA,
                            support: (supportAB * 100).toFixed(1) + '%',
                            confidence: (confBtoA * 100).toFixed(1) + '%',
                            lift: liftBtoA.toFixed(2),
                            strength: liftBtoA > 1.5 ? 'Sangat Kuat' : 'Kuat'
                        });
                    }
                });

                globalValidRules = rules;

                // Update Stat Cards
                document.getElementById('statTransactions').innerText = totalTrans;
                document.getElementById('statDivisions').innerText = Object.keys(divisionCounts).length;
                document.getElementById('statProducts').innerText = Object.keys(itemCounts).length;
                document.getElementById('statRules').innerText = rules.length;

                // Update 6 Testing Modules Data
                document.getElementById('testFileStatus').innerText = "Valid & Clean";
                document.getElementById('testRowCount').innerText = rawDataStore.length + " Baris";
                document.getElementById('testUniqueTrx').innerText = totalTrans + " Nota";
                document.getElementById('testDivisionCount').innerText = Object.keys(divisionCounts).length + " Unit";
                document.getElementById('testUniqueItems').innerText = Object.keys(itemCounts).length + " Item";
                document.getElementById('testParamSupport').innerText = inputSupport.value + "%";
                document.getElementById('testParamConf').innerText = inputConfidence.value + "%";
                document.getElementById('testTotalRules').innerText = rules.length + " Aturan";

                generateExecutiveSummary(divisionCounts, itemCounts, rules, selectedDiv);

                renderRulesTable(rules);
                renderDivisionTable(divisionItemQty, divisionCounts);
                renderSmartOptimizationCards(rules);
                renderCharts(itemCounts, divisionCounts, rules);
                renderMatrixTable(rules, itemCounts);

                resultsSection.classList.remove('hidden');
                resultsSection.scrollIntoView({ behavior: 'smooth' });
                showToast("Success", "Pengujian sistem dan model selesai!");
            } catch (err) {
                console.error(err);
                showToast("Error", "Terjadi kesalahan saat memproses.", true);
            } finally {
                runAprioriBtn.disabled = false;
                btnIcon.innerHTML = `<i class="fa-solid fa-play"></i>`;
                btnText.innerText = "3. Jalankan Pengujian Model";
            }
        }, 300);
    };
}

function generateExecutiveSummary(divisionCounts, itemCounts, rules, selectedDiv) {
    if (!executiveSummaryText) return;

    let topDiv = Object.keys(divisionCounts).sort((a,b) => divisionCounts[b] - divisionCounts[a])[0] || "Umum";
    let topItem = Object.keys(itemCounts).sort((a,b) => itemCounts[b] - itemCounts[a])[0] || "N/A";
    
    let topRuleText = "Belum ada pola asosiasi yang memenuhi ambang batas parameter saat ini.";
    if (rules.length > 0) {
        let bestRule = [...rules].sort((a,b) => parseFloat(b.lift) - parseFloat(a.lift))[0];
        topRuleText = `Pola asosiasi terkuat ditemukan pada item <strong>"${bestRule.antecedent}"</strong> yang berkorelasi dengan <strong>"${bestRule.consequent}"</strong> (Confidence <strong>${bestRule.confidence}</strong>, Lift Ratio <strong>${bestRule.lift}</strong>).`;
    }

    let scopeDesc = selectedDiv === 'all' ? `seluruh unit kerja instansi` : `unit kerja <strong>${selectedDiv}</strong>`;

    executiveSummaryText.innerHTML = `
        Hasil pengujian sistem untuk ${scopeDesc} mencatat total <strong>${Object.keys(divisionCounts).length} Unit</strong> aktif dengan volume transaksi tertinggi dipegang oleh <strong>Divisi ${topDiv}</strong>. Item dominan yang sering diajukan adalah <strong>${topItem}</strong>. ${topRuleText}
    `;
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
        matrixTable.innerHTML = `<tr><td class="p-4 text-slate-400">Tidak ada data matrik.</td></tr>`;
        return;
    }

    let liftMap = {};
    rules.forEach(r => { liftMap[`${r.antecedent}_${r.consequent}`] = parseFloat(r.lift); });

    let html = `<thead><tr class="bg-slate-100 text-slate-700 font-semibold"><th class="p-2 border border-slate-200">Item (A \ B)</th>`;
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
                    let bgColor = liftVal > 1.5 ? 'bg-emerald-100 text-emerald-800 font-bold' : 'bg-indigo-100 text-indigo-800';
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

function renderSmartOptimizationCards(rules) {
    const grid = document.getElementById('optimizationGrid');
    if (!grid) return;
    if (rules.length === 0) {
        grid.innerHTML = `<p class="text-xs text-slate-400 col-span-3 text-center">Belum ada aturan yang memenuhi parameter saat ini.</p>`;
        return;
    }

    let topRules = [...rules].sort((a,b) => parseFloat(b.lift) - parseFloat(a.lift)).slice(0, 3);

    const strategies = [
        { badge: "bg-indigo-600", title: "Rak Penyimpanan Utama", advice: "Letakkan kedua item ATK ini berdampingan di rak gudang utama agar staf logistik dapat mengambilnya bersamaan." },
        { badge: "bg-purple-600", title: "Paket Distribusi Divisi", advice: "Sediakan kedua barang ini dalam satu paket pengiriman rutin ke setiap unit kerja." },
        { badge: "bg-blue-600", title: "Manajemen Safety Stock", advice: "Tetapkan reorder point bersama agar pengadaan grosir bulanan selalu sinkron." }
    ];

    grid.innerHTML = topRules.map((r, idx) => {
        let strat = strategies[idx] || strategies[0];
        return `
            <div class="bg-slate-50 p-4 rounded-xl border border-slate-200 flex flex-col justify-between space-y-3 shadow-sm">
                <div>
                    <div class="flex justify-between items-center mb-2">
                        <span class="${strat.badge} text-white text-[10px] px-2.5 py-0.5 rounded font-bold">Bundling #${idx+1}</span>
                        <span class="text-[11px] text-emerald-700 bg-emerald-50 px-2 py-0.5 rounded font-bold border border-emerald-200">Lift: ${r.lift}</span>
                    </div>
                    <h4 class="font-bold text-slate-900 text-xs mb-1">${r.antecedent} + ${r.consequent}</h4>
                    <p class="text-[11px] text-slate-600 mb-2">Confidence: <strong class="text-indigo-600">${r.confidence}</strong></p>
                </div>
                <div class="bg-white p-2.5 rounded-lg border border-slate-200 text-[11px] space-y-1">
                    <p class="font-bold text-slate-800"><i class="fa-solid fa-bullseye text-indigo-600 mr-1"></i> ${strat.title}:</p>
                    <p class="text-slate-600 leading-relaxed">${strat.advice}</p>
                </div>
            </div>
        `;
    }).join('');
}

if(downloadPdfBtn) { downloadPdfBtn.onclick = () => { window.print(); }; }

if(downloadExcelBtn) {
    downloadExcelBtn.onclick = () => {
        if (globalValidRules.length === 0) {
            showToast("Warning", "Tidak ada data aturan untuk diekspor.", true);
            return;
        }
        let exportData = globalValidRules.map(r => ({
            "ID": r.id, "Premis (A)": r.antecedent, "Konsekuen (B)": r.consequent,
            "Support": r.support, "Confidence": r.confidence, "Lift": r.lift, "Status": r.strength
        }));
        let worksheet = XLSX.utils.json_to_sheet(exportData);
        let workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, "Apriori Rules");
        XLSX.writeFile(workbook, "Laporan_Pengujian_Apriori.xlsx");
        showToast("Success", "Laporan diekspor ke Excel!");
    };
}

function renderCharts(itemCounts, divisionCounts, rules) {
    let sortedDivs = Object.keys(divisionCounts).map(d => ({ name: d, count: divisionCounts[d] })).sort((a, b) => b.count - a.count).slice(0, 5);
    const ctxDiv = document.getElementById('chartTopDivisions').getContext('2d');
    if (chartTopDivInstance) chartTopDivInstance.destroy();
    chartTopDivInstance = new Chart(ctxDiv, {
        type: 'bar',
        data: { labels: sortedDivs.map(d => d.name), datasets: [{ data: sortedDivs.map(d => d.count), backgroundColor: '#3b82f6', borderRadius: 6 }] },
        options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } } }
    });

    let sortedItems = Object.keys(itemCounts).map(item => ({
        name: item.length > 18 ? item.substring(0, 18) + '...' : item, fullName: item, count: itemCounts[item]
    })).sort((a, b) => b.count - a.count).slice(0, 5);

    const ctxTop = document.getElementById('chartTopProducts').getContext('2d');
    if (chartTopProdInstance) chartTopProdInstance.destroy();
    chartTopProdInstance = new Chart(ctxTop, {
        type: 'bar',
        data: { labels: sortedItems.map(i => i.name), datasets: [{ data: sortedItems.map(i => i.count), backgroundColor: '#4f46e5', borderRadius: 6 }] },
        options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false }, tooltip: { callbacks: { title: ctx => sortedItems[ctx[0].dataIndex].fullName } } } }
    });

    let sangatKuat = rules.filter(r => r.strength === 'Sangat Kuat').length;
    let kuat = rules.filter(r => r.strength === 'Kuat').length;
    const ctxPie = document.getElementById('chartRuleStrength').getContext('2d');
    if (chartStrengthInstance) chartStrengthInstance.destroy();
    chartStrengthInstance = new Chart(ctxPie, {
        type: 'doughnut',
        data: { labels: ['Sangat Kuat', 'Kuat'], datasets: [{ data: [sangatKuat, kuat], backgroundColor: ['#10b981', '#3b82f6'] }] },
        options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'bottom' } } }
    });
}

function renderDivisionTable(divisionItemQty, divisionCounts) {
    const tbody = document.getElementById('divisionTableBody');
    if (!tbody) return;
    let sortedDivs = Object.keys(divisionCounts).sort((a,b) => divisionCounts[b] - divisionCounts[a]);
    tbody.innerHTML = sortedDivs.map(div => {
        let stats = divisionItemQty[div];
        let topItem = "N/A", maxQty = -1;
        for (let item in stats.items) {
            if (stats.items[item] > maxQty) { maxQty = stats.items[item]; topItem = item; }
        }
        return `
            <tr class="hover:bg-slate-50">
                <td class="p-2.5 border border-slate-200 font-bold text-indigo-600">${div}</td>
                <td class="p-2.5 border border-slate-200 font-semibold text-slate-700">${divisionCounts[div]} Nota</td>
                <td class="p-2.5 border border-slate-200 text-slate-700">${stats.totalQty} Unit</td>
                <td class="p-2.5 border border-slate-200 font-medium text-slate-900">${topItem} <span class="text-[11px] text-indigo-500 font-bold">(${maxQty} pcs)</span></td>
            </tr>
        `;
    }).join('');
}

function renderRulesTable(rules) {
    const tbody = document.getElementById('rulesTableBody');
    if (!tbody) return;
    if (rules.length === 0) {
        tbody.innerHTML = `<tr><td colspan="7" class="text-center py-6 text-slate-400">Tidak ada aturan ditemukan pada parameter ini.</td></tr>`;
        return;
    }
    tbody.innerHTML = rules.map(r => `
        <tr class="hover:bg-slate-50">
            <td class="p-2.5 border border-slate-200 font-semibold text-indigo-600">${r.id}</td>
            <td class="p-2.5 border border-slate-200 font-medium text-slate-800">${r.antecedent}</td>
            <td class="p-2.5 border border-slate-200 font-medium text-slate-800">${r.consequent}</td>
            <td class="p-2.5 border border-slate-200 text-slate-600">${r.support}</td>
            <td class="p-2.5 border border-slate-200 text-slate-600 font-semibold text-indigo-600">${r.confidence}</td>
            <td class="p-2.5 border border-slate-200 text-slate-600">${r.lift}</td>
            <td class="p-2.5 border border-slate-200"><span class="px-2 py-0.5 rounded text-[10px] font-bold ${r.strength === 'Sangat Kuat' ? 'bg-emerald-100 text-emerald-700' : 'bg-blue-100 text-blue-700'}">${r.strength}</span></td>
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