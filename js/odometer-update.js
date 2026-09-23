/* =========================================================
   UPDATE DATA — writes to the Google Sheets master data via
   a Google Apps Script Web App (see google-apps-script/Code.gs)
   Two ways in: Upload Excel (bulk) and Update Manual (single).
   ========================================================= */

/* TODO: fill these in after deploying the Apps Script (see
   google-apps-script/Code.gs for step-by-step instructions). */
const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbz1f_qclndhzs8dd7XV4yCCAD-L5smpNHWRd86Sr7HajtKpMALdutaSqXThNeZGb1DQ/exec';
const APP_SECRET = 'transkonjaya';

const NOT_CONNECTED_MESSAGE =
    'Fitur ini belum terhubung ke Google Sheets. Deploy dulu google-apps-script/Code.gs ' +
    'sebagai Web App, lalu isi APPS_SCRIPT_URL di js/odometer-update.js.';


document.addEventListener('DOMContentLoaded', () => {
    setupNavDropdown();
    setupUpdateTabs();
    setupManualForm();
    setupExcelUpload();
});


/* ---------------------------------------------------------
   Navbar dropdown (Update Data)
--------------------------------------------------------- */

function setupNavDropdown() {
    const dropdown = document.getElementById('navUpdateData');
    if (!dropdown) return;

    const trigger = dropdown.querySelector('.nav-dropdown-trigger');
    const items = dropdown.querySelectorAll('.nav-dropdown-item');

    trigger?.addEventListener('click', (e) => {
        e.stopPropagation();
        dropdown.classList.toggle('open');
    });

    document.addEventListener('click', () => dropdown.classList.remove('open'));

    items.forEach(item => {
        item.addEventListener('click', () => {
            dropdown.classList.remove('open');
            switchUpdateTab(item.dataset.tab);
        });
    });
}


/* ---------------------------------------------------------
   Tab switcher (Upload Excel / Update Manual)
--------------------------------------------------------- */

function setupUpdateTabs() {
    const tabs = document.querySelectorAll('.report-tab[data-updatetab]');
    tabs.forEach(tab => {
        tab.addEventListener('click', () => switchUpdateTab(tab.dataset.updatetab));
    });
}

function switchUpdateTab(key) {
    document.querySelectorAll('.report-tab[data-updatetab]').forEach(t => {
        t.classList.toggle('active', t.dataset.updatetab === key);
    });
    document.querySelectorAll('.report-panel[data-updatetab]').forEach(p => {
        p.hidden = p.dataset.updatetab !== key;
    });
}


/* ---------------------------------------------------------
   Manual update (single TK Number)
--------------------------------------------------------- */

function setupManualForm() {
    const form = document.getElementById('odometerForm');
    const messageBox = document.getElementById('odometerFormMessage');
    const submitBtn = document.getElementById('odometerSubmitBtn');

    if (!form) return;

    form.addEventListener('submit', async (e) => {
        e.preventDefault();

        const tkNumber = document.getElementById('inputTkNumber').value.trim();
        const odometer = document.getElementById('inputOdometer').value;

        if (isNotConnected()) {
            showFormMessage(messageBox, 'error', NOT_CONNECTED_MESSAGE);
            return;
        }

        setButtonLoading(submitBtn, true, 'Simpan ke Spreadsheet');

        try {
            // No updateDate sent — the Apps Script fills it in with
            // today() automatically when the field is left out.
            const result = await callAppsScript({ tkNumber, odometer });

            if (result.success) {
                showFormMessage(
                    messageBox, 'success',
                    `${result.tkNumber} berhasil di-update: ${fmtKm(result.previousOdometer)} KM → ${fmtKm(result.newOdometer)} KM (${result.updateDate}).`
                );
                form.reset();
            } else {
                showFormMessage(messageBox, 'error', result.error || 'Gagal menyimpan data.');
            }
        } catch (err) {
            showFormMessage(messageBox, 'error', 'Tidak bisa menghubungi Google Sheets. Cek koneksi internet atau URL Apps Script.');
        } finally {
            setButtonLoading(submitBtn, false, 'Simpan ke Spreadsheet');
        }
    });
}


/* ---------------------------------------------------------
   Excel bulk upload
--------------------------------------------------------- */

function setupExcelUpload() {
    const dropZone = document.getElementById('excelDropZone');
    const fileInput = document.getElementById('excelFileInput');
    const fileNameLabel = document.getElementById('excelFileName');
    const previewWrap = document.getElementById('excelPreviewWrap');
    const previewTable = document.getElementById('excelPreviewTable');
    const previewTitle = document.getElementById('excelPreviewTitle');
    const uploadBtn = document.getElementById('excelUploadBtn');
    const messageBox = document.getElementById('excelFormMessage');

    if (!dropZone || !fileInput) return;

    let parsedRows = [];

    dropZone.addEventListener('click', () => fileInput.click());

    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.classList.add('dragover');
    });
    dropZone.addEventListener('dragleave', () => dropZone.classList.remove('dragover'));
    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.classList.remove('dragover');
        if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);
    });

    fileInput.addEventListener('change', () => {
        if (fileInput.files.length) handleFile(fileInput.files[0]);
    });

    function handleFile(file) {
        messageBox.textContent = '';
        messageBox.className = 'form-message';
        fileNameLabel.textContent = file.name;

        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const wb = XLSX.read(e.target.result, { type: 'array', cellDates: true });
                const sheet = wb.Sheets[wb.SheetNames[0]];
                const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
                parsedRows = mapRows(rows);
                renderPreview(parsedRows);
            } catch (err) {
                showFormMessage(messageBox, 'error', 'Gagal membaca file. Pastikan formatnya .xlsx, .xls, atau .csv.');
            }
        };
        reader.readAsArrayBuffer(file);
    }

    function mapRows(rows) {
        return rows.map(row => {
            const keys = Object.keys(row);
            const tkKey = keys.find(k => /^tk[\s_]?number$/i.test(k.trim()));
            const odoKey = keys.find(k => /odometer/i.test(k) || /^km$/i.test(k.trim()));

            const tkNumber = tkKey ? String(row[tkKey]).trim().toUpperCase() : '';
            const odometerRaw = odoKey ? row[odoKey] : '';
            const odometer = Number(String(odometerRaw).replace(/[^\d.]/g, ''));

            // updateDate is intentionally not read from the file —
            // the Apps Script stamps it with today() on upload.
            return { tkNumber, odometer, status: 'pending' };
        }).filter(r => r.tkNumber);
    }

    function renderPreview(rows) {
        if (!rows.length) {
            previewWrap.hidden = true;
            showFormMessage(messageBox, 'error', 'Tidak ada baris valid terbaca. Cek nama kolom di file kamu (TK Number, Odometer).');
            return;
        }

        previewWrap.hidden = false;
        previewTitle.textContent = `${rows.length} baris siap di-upload`;

        previewTable.innerHTML = `
            <thead>
                <tr>
                    <th>TK NUMBER</th>
                    <th>ODOMETER (KM)</th>
                    <th>STATUS</th>
                </tr>
            </thead>
            <tbody>
                ${rows.map((r, i) => `
                    <tr id="excelRow-${i}">
                        <td><strong>${r.tkNumber}</strong></td>
                        <td>${r.odometer ? fmtKm(r.odometer) : '<span class="cell-placeholder">–</span>'}</td>
                        <td><span class="upload-status pending">Menunggu</span></td>
                    </tr>
                `).join('')}
            </tbody>
        `;
    }

    uploadBtn?.addEventListener('click', async () => {
        if (isNotConnected()) {
            showFormMessage(messageBox, 'error', NOT_CONNECTED_MESSAGE);
            return;
        }
        if (!parsedRows.length) return;

        setButtonLoading(uploadBtn, true, 'Upload ke Spreadsheet');
        showFormMessage(messageBox, '', '');

        try {
            const result = await callAppsScript({
                updates: parsedRows.map(({ tkNumber, odometer }) => ({ tkNumber, odometer })),
            });

            if (result.success && result.results) {
                result.results.forEach((r, i) => {
                    const row = document.getElementById(`excelRow-${i}`);
                    if (!row) return;
                    const statusCell = row.querySelector('.upload-status');
                    if (r.success) {
                        statusCell.textContent = 'Berhasil';
                        statusCell.className = 'upload-status success';
                    } else {
                        statusCell.textContent = r.error || 'Gagal';
                        statusCell.className = 'upload-status error';
                        statusCell.title = r.error || '';
                    }
                });

                showFormMessage(
                    messageBox,
                    result.failCount ? 'error' : 'success',
                    `Selesai: ${result.successCount} berhasil, ${result.failCount} gagal dari ${result.total} baris.`
                );
            } else {
                showFormMessage(messageBox, 'error', result.error || 'Gagal mengupload data.');
            }
        } catch (err) {
            showFormMessage(messageBox, 'error', 'Tidak bisa menghubungi Google Sheets. Cek koneksi internet atau URL Apps Script.');
        } finally {
            setButtonLoading(uploadBtn, false, 'Upload ke Spreadsheet');
        }
    });
}


/* ---------------------------------------------------------
   Shared helpers
--------------------------------------------------------- */

function isNotConnected() {
    return APPS_SCRIPT_URL.includes('PASTE_YOUR_DEPLOYED_WEB_APP_URL_HERE');
}

async function callAppsScript(payload) {
    const res = await fetch(APPS_SCRIPT_URL, {
        method: 'POST',
        // Plain string body (no custom Content-Type header) keeps this a
        // "simple request" so the browser skips CORS preflight, which
        // Apps Script Web Apps cannot answer.
        body: JSON.stringify({ ...payload, secret: APP_SECRET }),
    });
    return res.json();
}

function setButtonLoading(btn, isLoading, label) {
    if (!btn) return;
    btn.disabled = isLoading;
    btn.innerHTML = isLoading
        ? '<i class="fa-solid fa-spinner fa-spin"></i> Memproses...'
        : `<i class="fa-solid fa-cloud-arrow-up"></i> ${label}`;
}

function showFormMessage(box, type, text) {
    if (!box) return;
    box.textContent = text;
    box.className = 'form-message' + (type ? ' ' + type : '');
}

function fmtKm(n) {
    return Number(n).toLocaleString('en-US');
}
