const STORAGE_KEY = 'test_case_manager_v1';
let items = [];

const $ = (sel) => document.querySelector(sel);
function uid() { return 'id-' + Math.random().toString(36).slice(2, 10); }
function load() { try { items = JSON.parse(localStorage.getItem(STORAGE_KEY)) || []; } catch { items = []; } }
function save() { localStorage.setItem(STORAGE_KEY, JSON.stringify(items)); }

function groupBy(arr, key) {
  return arr.reduce((acc, cur) => { (acc[cur[key]] ||= []).push(cur); return acc; }, {});
}
function escapeHtml(str) {
  return String(str).replace(/[&<>"']+/g, (m) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[m]));
}

// Rendering
function render() {
  const accordion = $('#categoryAccordion');
  accordion.innerHTML = '';
  const grouped = groupBy(items, 'category');
  const hasItems = Object.keys(grouped).length > 0;
  $('#emptyState').style.display = hasItems ? 'none' : 'block';

  let idx = 0;
  for (const [category, rows] of Object.entries(grouped)) {
    const accId = 'acc-' + (idx++);
    const total = rows.length;
    const passed = rows.filter(r => r.status === 'Passed').length;
    const failed = total - passed;

    const header = `
      <button class="accordion-button collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#${accId}">
        <div class="d-flex align-items-center w-100">
          <div class="me-auto fw-semibold">${category}</div>
          <div class="d-flex gap-2 small">
            <span class="status-pill status-passed">Passed: ${passed}</span>
            <span class="status-pill status-failed">Failed: ${failed}</span>
            <span class="badge text-bg-secondary">Total: ${total}</span>
          </div>
        </div>
      </button>`;

    const tableRows = rows.map(r => `
      <tr>
        <td class="text-wrap">${escapeHtml(r.testCase)}</td>
        <td><span class="status-pill ${r.status === 'Passed' ? 'status-passed' : 'status-failed'}">${r.status}</span></td>
        <td class="text-wrap">${escapeHtml(r.remarks || '')}</td>
        <td class="text-nowrap text-end">
          <button class="btn btn-sm btn-outline-light me-1" onclick="onEdit('${r.id}')"><i class="bi bi-pencil"></i></button>
          <button class="btn btn-sm btn-outline-danger" onclick="onDelete('${r.id}')"><i class="bi bi-trash"></i></button>
        </td>
      </tr>`).join('');

    const table = `
      <div id="${accId}" class="accordion-collapse collapse" data-bs-parent="#categoryAccordion">
        <div class="accordion-body p-0">
          <div class="table-responsive" style="max-height: 420px;">
            <table class="table table-dark table-striped align-middle mb-0">
              <thead><tr><th style="width:45%">Test case</th><th style="width:15%">Status</th><th style="width:30%">Remarks</th><th style="width:10%" class="text-end">Actions</th></tr></thead>
              <tbody>${tableRows || `<tr><td colspan="4" class="text-center text-secondary py-4">No items</td></tr>`}</tbody>
            </table>
          </div>
        </div>
      </div>`;

    const card = document.createElement('div');
    card.className = 'accordion-item';
    card.innerHTML = `<h2 class="accordion-header">${header}</h2>${table}`;
    accordion.appendChild(card);
  }
}

// CRUD
$('#testForm').addEventListener('submit', (e) => {
  e.preventDefault();
  const category = $('#category').value;
  const testCase = $('#testCase').value.trim();
  const status = $('#status').value;
  const remarks = $('#remarks').value.trim();
  if (!category || !testCase) return;
  items.push({ id: uid(), category, testCase, status, remarks });
  save(); render(); e.target.reset();
});

window.onDelete = (id) => { items = items.filter(i => i.id !== id); save(); render(); }

window.onEdit = (id) => {
  const it = items.find(i => i.id === id);
  if (!it) return;
  $('#edit-id').value = it.id;
  $('#edit-category').value = it.category;
  $('#edit-testCase').value = it.testCase;
  $('#edit-status').value = it.status;
  $('#edit-remarks').value = it.remarks || '';
  bootstrap.Modal.getOrCreateInstance($('#editModal')).show();
}

$('#saveEditBtn').addEventListener('click', () => {
  const id = $('#edit-id').value;
  const idx = items.findIndex(i => i.id === id);
  if (idx === -1) return;
  items[idx] = { ...items[idx], category: $('#edit-category').value, testCase: $('#edit-testCase').value.trim(), status: $('#edit-status').value, remarks: $('#edit-remarks').value.trim() };
  save(); render(); bootstrap.Modal.getInstance($('#editModal')).hide();
});

$('#clearAllBtn').addEventListener('click', () => { if (confirm('Clear all test cases?')) { items = []; save(); render(); } });
// Excel export
$('#exportExcelBtn').addEventListener('click', async () => {
  if (!items.length) { 
    alert('No data to export.'); 
    return; 
  }

  const workbook = new ExcelJS.Workbook();
  workbook.created = workbook.modified = new Date();
  const grouped = groupBy(items, 'category');

  for (const [category, rows] of Object.entries(grouped)) {
    const sheet = workbook.addWorksheet(category.substring(0, 31));

    // Define columns
    sheet.columns = [
      { header: 'Test Case', key: 'testCase', width: 60 },
      { header: 'Status', key: 'status', width: 14 },
      { header: 'Remarks', key: 'remarks', width: 50 }
    ];

    // Title Row (merged cells)
    sheet.mergeCells('A1:C1');
    const title = sheet.getCell('A1');
    title.value = category + ' – Test Cases';
    title.alignment = { horizontal: 'center', vertical: 'middle' };
    title.font = { size: 14, bold: true, color: { argb: 'FFFFFFFF' } };
    title.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0B1A35' } };
    sheet.getRow(1).height = 24;

    // Header row
    const headerRow = sheet.getRow(2);
    headerRow.values = ['Test Case', 'Status', 'Remarks'];
    headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    headerRow.alignment = { horizontal: 'center', vertical: 'middle' };
    headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E90FF' } };

    // Add data rows
    rows.forEach(r => {
      sheet.addRow([r.testCase, r.status, r.remarks || '']);
    });
  }

  // Export file
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'TestCases.xlsx';
  a.click();
  URL.revokeObjectURL(url);
});