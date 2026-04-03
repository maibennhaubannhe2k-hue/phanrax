const SANS = ['tiktok', 'shopee', 'ngoai'];
let floorResults = { tiktok: new Map(), shopee: new Map(), ngoai: new Map() };
let historyStack = [];

function init() {
    SANS.forEach(san => {
        const dId = `dict-${san}`, oId = `order-${san}`;
        if (localStorage.getItem(dId)) document.getElementById(dId).innerHTML = localStorage.getItem(dId);
        else renderEmpty(dId, 4, false);
        if (localStorage.getItem(oId)) document.getElementById(oId).innerHTML = localStorage.getItem(oId);
        else renderEmpty(oId, 4, true);

        const dWrap = document.getElementById(dId).closest('.excel-wrapper');
        const oWrap = document.getElementById(oId).closest('.excel-wrapper');
        dWrap.addEventListener('paste', e => handlePaste(e, dId, 4, false));
        oWrap.addEventListener('paste', e => handlePaste(e, oId, 4, true));
    });
    initCellSelection();
    initUndo();
}

// LƯU LỊCH SỬ ĐỂ CTRL + Z
function saveHistory() {
    const state = {};
    SANS.forEach(s => {
        state[`dict-${s}`] = document.getElementById(`dict-${s}`).innerHTML;
        state[`order-${s}`] = document.getElementById(`order-${s}`).innerHTML;
    });
    historyStack.push(JSON.stringify(state));
    if (historyStack.length > 30) historyStack.shift();
}

function initUndo() {
    document.addEventListener('keydown', e => {
        if ((e.ctrlKey || e.metaKey) && e.key === 'z') {
            if (historyStack.length > 0) {
                e.preventDefault();
                const last = JSON.parse(historyStack.pop());
                SANS.forEach(s => {
                    document.getElementById(`dict-${s}`).innerHTML = last[`dict-${s}`];
                    document.getElementById(`order-${s}`).innerHTML = last[`order-${s}`];
                });
                saveData();
                alert("Đã quay lại bước trước!");
            }
        }
    });
}

function renderEmpty(id, cols, hasRes) {
    let html = '';
    for(let i=1; i<=20; i++) {
        html += `<tr><td>${i}</td>` + `<td contenteditable="true" onfocus="saveHistory()"></td>`.repeat(cols) + (hasRes ? `<td></td>` : '') + `</tr>`;
    }
    document.getElementById(id).innerHTML = html;
}

function saveData() {
    SANS.forEach(s => {
        localStorage.setItem(`dict-${s}`, document.getElementById(`dict-${s}`).innerHTML);
        localStorage.setItem(`order-${s}`, document.getElementById(`order-${s}`).innerHTML);
    });
}

function handlePaste(e, tbodyId, colCount, hasRes) {
    saveHistory();
    e.preventDefault();
    const text = (e.clipboardData || window.clipboardData).getData('text/plain');
    if (!text) return;
    const lines = text.trim().split('\n');
    if (lines[0].toLowerCase().includes('sku')) lines.shift();

    const tbody = document.getElementById(tbodyId);
    let sRow = 0, sCol = 1;
    const active = document.activeElement;
    if (active?.closest('tbody')?.id === tbodyId) {
        sCol = active.cellIndex; sRow = active.parentElement.rowIndex;
    }

    lines.forEach((line, i) => {
        const cols = line.split('\t');
        let rIdx = sRow + i;
        if (!tbody.rows[rIdx]) {
            const tr = document.createElement('tr');
            tr.innerHTML = `<td></td>` + `<td contenteditable="true" onfocus="saveHistory()"></td>`.repeat(colCount) + (hasRes ? `<td></td>` : '');
            tbody.appendChild(tr);
        }
        cols.forEach((val, j) => {
            const cIdx = sCol + j;
            if (cIdx > 0 && cIdx <= colCount) tbody.rows[rIdx].cells[cIdx].innerText = val.trim();
        });
    });
    Array.from(tbody.rows).forEach((r, idx) => r.cells[0].innerText = idx + 1);
    saveData();
}

function parseCombo(str) {
    return str.split(',').map(s => {
        const m = s.trim().match(/^(\d+)\s*(.*)$/);
        return m ? { qty: parseInt(m[1]), name: m[2].trim() } : { qty: 1, name: s.trim() };
    });
}

// CHỨC NĂNG PHÂN RÃ SÀN
function processFloor(san) {
    const dict = new Map();
    document.querySelectorAll(`#dict-${san} tr`).forEach(r => {
        const c = r.cells;
        const key = `${c[1].innerText.trim()}|${c[2].innerText.trim()}|${c[3].innerText.trim()}`;
        if(c[4].innerText.trim() && key !== "||") dict.set(key, parseCombo(c[4].innerText.trim()));
    });

    const res = new Map(), errs = [];
    document.querySelectorAll(`#order-${san} tr`).forEach(r => {
        const c = r.cells;
        const key = `${c[1].innerText.trim()}|${c[2].innerText.trim()}|${c[3].innerText.trim()}`;
        let q = parseInt(c[4].innerText.replace(/[^\d]/g, '')) || 0;
        if(key === "||" || (c[1].innerText === "" && c[2].innerText === "")) return;

        const items = dict.get(key);
        if(items && q > 0) {
            r.className = "row-success";
            let txt = [];
            items.forEach(it => {
                const fQ = it.qty * q;
                res.set(it.name, (res.get(it.name) || 0) + fQ);
                txt.push(`${fQ}x[${it.name}]`);
            });
            c[5].innerText = "✅ " + txt.join(' + ');
        } else if (q > 0) {
            r.className = "row-error";
            c[5].innerText = "❌ Lỗi";
            errs.push({ line: c[0].innerText, a: c[1].innerText, b: c[2].innerText, c: c[3].innerText });
        }
    });

    floorResults[san] = res;
    const eBox = document.getElementById(`error-${san}`), eBody = document.getElementById(`error-body-${san}`);
    if(errs.length > 0) {
        eBox.classList.remove('hidden');
        eBody.innerHTML = errs.map(e => `<tr><td>${e.line}</td><td>${e.a}</td><td>${e.b}</td><td>${e.c}</td></tr>`).join('');
    } else eBox.classList.add('hidden');

    const rBox = document.getElementById(`result-${san}`), rBody = document.getElementById(`res-body-${san}`);
    if(res.size > 0) {
        rBox.classList.remove('hidden');
        rBody.innerHTML = Array.from(res, ([n, q]) => `<tr><td>${n}</td><td>${q}</td></tr>`).join('');
    } else rBox.classList.add('hidden');
    saveData();
}

// CHỨC NĂNG XÓA TỪNG BẢNG
function clearTable(tbodyId) {
    if(confirm("Bạn muốn xóa bảng này?")) {
        saveHistory();
        const isOrder = tbodyId.includes('order');
        renderEmpty(tbodyId, 4, isOrder);
        const san = tbodyId.split('-')[1];
        if(isOrder) {
            document.getElementById(`error-${san}`).classList.add('hidden');
            document.getElementById(`result-${san}`).classList.add('hidden');
        }
        saveData();
    }
}

// XÓA VÙNG CHỌN
function clearSelection() {
    const selected = document.querySelectorAll('.selected-cell');
    if(selected.length === 0) return alert("Bôi đen vùng cần xóa!");
    saveHistory();
    selected.forEach(c => { if(c.contentEditable === "true") c.innerText = ""; });
    saveData();
}

// XÓA TẤT CẢ
function clearAllData() {
    if(confirm("Xóa sạch TẤT CẢ dữ liệu hệ thống?")) {
        localStorage.clear();
        location.reload();
    }
}

function combineAll() {
    const final = new Map();
    SANS.forEach(s => { if(floorResults[s]) floorResults[s].forEach((v, k) => final.set(k, (final.get(k) || 0) + v)); });
    const sorted = Array.from(final, ([n, q]) => ({n, q})).sort((a,b) => b.q - a.q);
    document.getElementById('final-result-body').innerHTML = sorted.map((it, i) => `<tr><td>${i+1}</td><td>${it.n}</td><td style="color:red; font-weight:bold;">${it.q}</td></tr>`).join('');
    openTab(null, 'tab-tonghop');
}

function openTab(evt, name) {
    document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
    document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
    document.getElementById(name).classList.add('active');
    if(evt) evt.currentTarget.classList.add('active');
    else document.querySelector(`button[onclick*='${name}']`).classList.add('active');
}

function initCellSelection() {
    let isSel = false, startC = null;
    document.addEventListener('mousedown', e => {
        const c = e.target.closest('td');
        if (!c || c.cellIndex === 0 || !c.closest('tbody')) return;
        isSel = true; startC = c;
        document.querySelectorAll('.selected-cell').forEach(x => x.classList.remove('selected-cell'));
        c.classList.add('selected-cell');
    });
    document.addEventListener('mouseover', e => {
        if (!isSel) return;
        const cur = e.target.closest('td');
        if (!cur || cur.cellIndex === 0 || cur.closest('tbody') !== startC.closest('tbody')) return;
        document.querySelectorAll('.selected-cell').forEach(x => x.classList.remove('selected-cell'));
        const r1 = startC.parentElement.rowIndex, c1 = startC.cellIndex;
        const r2 = cur.parentElement.rowIndex, c2 = cur.cellIndex;
        const rs = startC.closest('table').rows;
        for(let i=Math.min(r1,r2); i<=Math.max(r1,r2); i++)
            for(let j=Math.min(c1,c2); j<=Math.max(c1,c2); j++)
                if(rs[i]?.cells[j] && rs[i].cells[j].cellIndex !== 0) rs[i].cells[j].classList.add('selected-cell');
    });
    document.addEventListener('mouseup', () => isSel = false);
}

function exportExcel() {
    const table = document.getElementById('final-result-body');
    const data = Array.from(table.rows).map(r => ({"STT": r.cells[0].innerText, "Mặt Hàng": r.cells[1].innerText, "Số Lượng": r.cells[2].innerText}));
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, "TongHop");
    XLSX.writeFile(wb, "Phieu_Tong_Xuat_Kho.xlsx");
}

init();