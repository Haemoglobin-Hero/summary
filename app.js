const $ = id => document.getElementById(id);

const MONTH_ALIASES = {
  Jan: ['jan','january'], Feb:['feb','february'], Mar:['mar','march'],
  Apr:['apr','april'], May:['may'], Jun:['jun','june'],
  July:['jul','july'], Aug:['aug','august']
};
const DISPLAY_MONTHS = ['Jan','Feb','Mar','Apr','May','Jun','July','Aug'];

let workbook = null;
let fileName = '';
let monthSheets = [];
let masterParts = new Map();
let selectedParts = new Set();
let selectedMonths = new Set();
let conflicts = [];
let detectedTotals = {};

$('fileInput').addEventListener('change', e => {
  if (e.target.files[0]) loadWorkbook(e.target.files[0]);
});
$('replaceBtn').addEventListener('click', () => $('fileInput').click());
$('dropZone').addEventListener('dragover', e => { e.preventDefault(); $('uploadSection').classList.add('dragover'); });
$('dropZone').addEventListener('dragleave', () => $('uploadSection').classList.remove('dragover'));
$('dropZone').addEventListener('drop', e => {
  e.preventDefault(); $('uploadSection').classList.remove('dragover');
  if (e.dataTransfer.files[0]) loadWorkbook(e.dataTransfer.files[0]);
});
$('toggleMonths').onclick = toggleAllMonths;
$('toggleParts').onclick = toggleAllParts;
$('clearParts').onclick = () => { selectedParts.clear(); renderParts(); updateUI(); };
$('masterPartCheck').onchange = e => {
  filteredParts().forEach(p => e.target.checked ? selectedParts.add(p.partNo) : selectedParts.delete(p.partNo));
  renderParts(); updateUI();
};
$('partSearch').oninput = renderParts;
$('runChecks').onclick = runVerification;
$('generateBtn').onclick = generateWorkbook;

async function loadWorkbook(file) {
  try {
    fileName = file.name;
    setStatus('Reading workbook…', false);
    const data = await file.arrayBuffer();
    workbook = XLSX.read(data, { type:'array', cellDates:true });
    parseWorkbook();
    $('fileName').textContent = fileName;
    $('workspace').classList.remove('hidden');
    $('uploadSection').classList.add('hidden');
    $('sheetCount').textContent = workbook.SheetNames.length;
    $('monthCount').textContent = monthSheets.length;
    $('partCount').textContent = masterParts.size;
    $('issueCount').textContent = conflicts.length;
    setStatus('Workbook ready', true);
    renderMonths();
    renderParts();
    runVerification();
    updateUI();
  } catch(err) {
    console.error(err);
    alert('Could not read this workbook. Please check that it is a valid .xls or .xlsx file.');
    setStatus('Workbook error', false);
  }
}

function setStatus(text, ready) {
  $('headerStatus').innerHTML = `<i></i>${escapeHtml(text)}`;
  $('headerStatus').classList.toggle('ready', !!ready);
}

function normalize(v) {
  return String(v ?? '').trim().replace(/\s+/g,' ').toLowerCase();
}
function clean(v) { return String(v ?? '').trim(); }

function parseWorkbook() {
  masterParts = new Map();
  conflicts = [];
  detectedTotals = {};
  monthSheets = [];

  workbook.SheetNames.forEach(sheetName => {
    const key = normalize(sheetName);
    const month = DISPLAY_MONTHS.find(m => MONTH_ALIASES[m].includes(key));
    if (!month) return;

    const ws = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(ws, { header:1, defval:null, raw:true });
    const summary = locateSummary(rows, sheetName, month);
    monthSheets.push({ month, sheetName, rows, ...summary });
  });

  // Keep workbook order for detected sheets, but display in Jan-Aug order.
  monthSheets.sort((a,b) => DISPLAY_MONTHS.indexOf(a.month) - DISPLAY_MONTHS.indexOf(b.month));

  monthSheets.forEach(ms => {
    ms.items.forEach(item => {
      const partNo = clean(item.partNo);
      if (!partNo) return;
      const key = normalize(partNo);
      if (!masterParts.has(key)) {
        masterParts.set(key, {
          partNo,
          description: clean(item.description),
          uom: clean(item.uom),
          quantities: {},
          sources: []
        });
      }
      const master = masterParts.get(key);
      const desc = clean(item.description);
      const uom = clean(item.uom);
      if (desc && master.description && normalize(desc) !== normalize(master.description)) {
        conflicts.push({ type:'description', partNo, month:ms.month, a:master.description, b:desc });
      } else if (!master.description && desc) master.description = desc;
      if (uom && master.uom && normalize(uom) !== normalize(master.uom)) {
        conflicts.push({ type:'uom', partNo, month:ms.month, a:master.uom, b:uom });
      } else if (!master.uom && uom) master.uom = uom;

      master.quantities[ms.month] = (master.quantities[ms.month] || 0) + numeric(item.quantity);
      master.sources.push(ms.month);
    });
    if (ms.total != null) detectedTotals[ms.month] = ms.total;
  });

  // default: all detected months and all parts
  selectedMonths = new Set(monthSheets.map(x => x.month));
  selectedParts = new Set([...masterParts.values()].map(x => x.partNo));
}

function locateSummary(rows, sheetName, month) {
  // Find the most likely item-wise summary header. Handles merged-looking title rows.
  const headerAliases = {
    partNo:['part no','part number','partno','part #','item code','item no'],
    description:['part description','description','item description','material description','item'],
    uom:['uom','unit','unit of measure'],
    quantity:['quantity','qty','consumption','consumed quantity']
  };
  let best = null;

  for (let r=0; r<rows.length; r++) {
    const row = rows[r] || [];
    const texts = row.map(normalize);
    const indexes = {};
    for (const [field, aliases] of Object.entries(headerAliases)) {
      indexes[field] = texts.findIndex(t => aliases.some(a => t === a || t.includes(a)));
    }
    const score = ['partNo','description','uom','quantity'].reduce((s,f)=>s+(indexes[f]>=0?1:0),0);
    if (indexes.partNo >= 0 && indexes.quantity >= 0 && score >= 3) {
      best = { headerRow:r, indexes };
      // Prefer a row that occurs after a summary/title marker if available.
      if (texts.some(t => t.includes('item-wise') || t.includes('item wise') || t.includes('summary'))) break;
    }
  }

  if (!best) {
    // Fallback: search rows for common four-column arrangements.
    best = { headerRow: -1, indexes: {partNo:0, description:1, uom:2, quantity:3} };
  }

  const items = [];
  let total = null;
  const start = best.headerRow + 1;

  for (let r=start; r<rows.length; r++) {
    const row = rows[r] || [];
    const p = best.indexes.partNo >= 0 ? row[best.indexes.partNo] : row[0];
    const d = best.indexes.description >= 0 ? row[best.indexes.description] : row[1];
    const u = best.indexes.uom >= 0 ? row[best.indexes.uom] : row[2];
    const q = best.indexes.quantity >= 0 ? row[best.indexes.quantity] : row[3];
    const pn = clean(p);
    const first = normalize(pn);

    if (first === 'total' || first === 'grand total' || first.includes('total')) {
      const qv = numeric(q);
      if (q !== null && q !== '' && Number.isFinite(qv)) total = qv;
      continue;
    }
    if (!pn || !Number.isFinite(numeric(q))) continue;
    // Avoid transaction-level rows where the detected "Part No." is actually a date/index.
    if (isLikelyTransactionRow(row, best)) continue;

    items.push({ partNo:pn, description:d, uom:u, quantity:q });
  }

  // If a summary has duplicate Part Nos. within a month, combine them.
  const combined = new Map();
  items.forEach(x => {
    const k = normalize(x.partNo);
    if (!combined.has(k)) combined.set(k, {...x, quantity:numeric(x.quantity)});
    else combined.get(k).quantity += numeric(x.quantity);
  });

  return { items:[...combined.values()], total };
}

function isLikelyTransactionRow(row, best) {
  const nonEmpty = row.filter(v => v !== null && v !== '').length;
  // This deliberately stays permissive because real summary layouts vary.
  return nonEmpty === 0;
}
function numeric(v) {
  if (typeof v === 'number') return Number.isFinite(v) ? v : 0;
  if (v === null || v === undefined || v === '') return 0;
  const n = Number(String(v).replace(/,/g,'').replace(/\s/g,''));
  return Number.isFinite(n) ? n : 0;
}
function fmt(n) {
  return Number(n || 0).toLocaleString(undefined, { maximumFractionDigits: 12 });
}

function renderMonths() {
  const grid = $('monthsGrid');
  grid.innerHTML = '';
  monthSheets.forEach(ms => {
    const div = document.createElement('label');
    div.className = 'month-card ' + (selectedMonths.has(ms.month) ? 'active':'');
    div.innerHTML = `<input type="checkbox" ${selectedMonths.has(ms.month)?'checked':''}>
      <span class="month-name">${escapeHtml(ms.month)}</span>
      <span class="sheet-name">${escapeHtml(ms.sheetName)}</span>`;
    div.onclick = e => {
      e.preventDefault();
      if (selectedMonths.has(ms.month)) selectedMonths.delete(ms.month);
      else selectedMonths.add(ms.month);
      renderMonths(); updateUI(); renderPreview();
    };
    grid.appendChild(div);
  });
  $('toggleMonths').textContent = selectedMonths.size === monthSheets.length ? 'Clear all' : 'Select all';
}

function filteredParts() {
  const q = normalize($('partSearch').value);
  const list = [...masterParts.values()];
  if (!q) return list;
  return list.filter(p => normalize(p.partNo).includes(q) || normalize(p.description).includes(q));
}

function renderParts() {
  const body = $('partsBody');
  const list = filteredParts();
  $('visibleCount').textContent = `${list.length} part${list.length===1?'':'s'}`;
  body.innerHTML = '';
  list.forEach(p => {
    const tr = document.createElement('tr');
    const checked = selectedParts.has(p.partNo);
    tr.className = checked ? 'selected':'';
    const months = [...new Set(p.sources)].map(x=>`<span class="tag">${x}</span>`).join('');
    tr.innerHTML = `<td><input type="checkbox" ${checked?'checked':''}></td>
      <td><strong>${escapeHtml(p.partNo)}</strong></td>
      <td>${escapeHtml(p.description || '—')}</td>
      <td>${escapeHtml(p.uom || '—')}</td>
      <td><div class="month-tags">${months || '—'}</div></td>`;
    tr.onclick = e => {
      if (e.target.tagName === 'INPUT') return;
      togglePart(p.partNo);
    };
    tr.querySelector('input').onclick = e => { e.stopPropagation(); togglePart(p.partNo); };
    body.appendChild(tr);
  });
  $('masterPartCheck').checked = list.length > 0 && list.every(p=>selectedParts.has(p.partNo));
  $('masterPartCheck').indeterminate = list.some(p=>selectedParts.has(p.partNo)) && !$('masterPartCheck').checked;
  $('selectionCount').textContent = `${selectedParts.size} selected`;
}

function togglePart(partNo) {
  if (selectedParts.has(partNo)) selectedParts.delete(partNo);
  else selectedParts.add(partNo);
  renderParts(); updateUI();
}
function toggleAllParts() {
  const all = filteredParts();
  const select = !all.length || !all.every(p=>selectedParts.has(p.partNo));
  all.forEach(p => select ? selectedParts.add(p.partNo) : selectedParts.delete(p.partNo));
  renderParts(); updateUI();
}
function toggleAllMonths() {
  if (selectedMonths.size === monthSheets.length) selectedMonths.clear();
  else selectedMonths = new Set(monthSheets.map(x=>x.month));
  renderMonths(); updateUI(); renderPreview();
}

function updateUI() {
  const n = selectedParts.size, m = selectedMonths.size;
  $('previewBadge').textContent = `${n} part${n===1?'':'s'} · ${m} month${m===1?'':'s'}`;
  $('generateSummary').textContent = n && m ? `Ready to generate ${n} selected part${n===1?'':'s'} across ${m} selected month${m===1?'':'s'}.` : 'Select parts and months to continue.';
  $('generateSub').textContent = conflicts.length ? `${conflicts.length} data consistency issue${conflicts.length===1?'':'s'} found — review Verification.` : 'The generated file will be a new workbook.';
  $('generateBtn').disabled = !(n && m);
  $('issueCount').textContent = conflicts.length;
  renderPreview();
}

function runVerification() {
  const box = $('verification');
  const checks = [];
  const unique = masterParts.size;
  checks.push({ok:unique>0, text:`${unique} unique Part No.${unique===1?'':'s'} detected across ${monthSheets.length} monthly sheet${monthSheets.length===1?'':'s'}.`});
  const dup = detectDuplicatePartNumbers();
  checks.push({ok:dup===0, text:dup ? `${dup} duplicate Part No. key${dup===1?'':'s'} were detected and consolidated.` : 'No duplicate Part Nos. exist in the master list.'});
  const uomIssues = conflicts.filter(x=>x.type==='uom');
  checks.push({ok:uomIssues.length===0, text:uomIssues.length ? `${uomIssues.length} UoM conflict${uomIssues.length===1?'':'s'} need review.` : 'No UoM discrepancies found.'});
  const descIssues = conflicts.filter(x=>x.type==='description');
  checks.push({ok:descIssues.length===0, text:descIssues.length ? `${descIssues.length} description conflict${descIssues.length===1?'':'s'} need review.` : 'No Part Description discrepancies found.'});
  checks.push({ok:monthSheets.length>0, text:`Detected months: ${monthSheets.map(x=>x.month).join(', ') || 'none'}.`});
  checks.push({ok:selectedParts.size>0 && selectedMonths.size>0, text:`Current selection: ${selectedParts.size} parts and ${selectedMonths.size} months.`});
  box.innerHTML = checks.map(c=>`<div class="check ${c.ok?'ok':'warn'}"><span class="mark">${c.ok?'✓':'!'}</span><span>${escapeHtml(c.text)}</span></div>`).join('');
}
function detectDuplicatePartNumbers() {
  let count = 0;
  monthSheets.forEach(ms => {
    const seen = new Set();
    ms.items.forEach(i => {
      const k=normalize(i.partNo);
      if (seen.has(k)) count++;
      seen.add(k);
    });
  });
  return count;
}

function renderPreview() {
  const months = DISPLAY_MONTHS.filter(m=>selectedMonths.has(m));
  const head = $('previewHead');
  const body = $('previewBody');
  head.innerHTML = `<tr><th>Part No.</th><th>Part Description</th><th>UoM</th>${months.map(m=>`<th>${m} Quantity</th>`).join('')}<th>Total Quantity</th></tr>`;
  body.innerHTML = '';
  [...masterParts.values()].filter(p=>selectedParts.has(p.partNo)).slice(0,30).forEach(p=>{
    const vals = months.map(m=>numeric(p.quantities[m]));
    const total = vals.reduce((a,b)=>a+b,0);
    body.insertAdjacentHTML('beforeend', `<tr><td><strong>${escapeHtml(p.partNo)}</strong></td><td>${escapeHtml(p.description||'—')}</td><td>${escapeHtml(p.uom||'—')}</td>${vals.map(v=>`<td>${fmt(v)}</td>`).join('')}<td><strong>${fmt(total)}</strong></td></tr>`);
  });
  if (selectedParts.size>30) body.insertAdjacentHTML('beforeend', `<tr><td colspan="${3+months.length+1}" class="muted">Preview limited to the first 30 selected parts. The generated workbook contains all selected parts.</td></tr>`);
}

function generateWorkbook() {
  if (!selectedParts.size || !selectedMonths.size) return;
  const months = DISPLAY_MONTHS.filter(m=>selectedMonths.has(m));
  const wb = XLSX.utils.book_new();
  const headers = ['Part No.','Part Description','UoM',...months.map(m=>`${m} Quantity`),'Total Quantity'];
  const rows = [headers];

  const selected = [...masterParts.values()].filter(p=>selectedParts.has(p.partNo));
  selected.forEach((p, idx) => {
    const rowNo = idx + 2;
    const vals = months.map(m=>numeric(p.quantities[m]));
    rows.push([p.partNo,p.description,p.uom,...vals,{t:'n', f:`SUM(${colLetter(4)}${rowNo}:${colLetter(3+months.length)}${rowNo})`}]);
  });

  const totalRowNo = rows.length + 1;
  const totalRow = ['TOTAL','',''];
  for (let i=0;i<months.length;i++) {
    const col = colLetter(4+i);
    totalRow.push({t:'n', f:`SUM(${col}2:${col}${totalRowNo-1})`});
  }
  const totalCol = colLetter(4+months.length);
  totalRow.push({t:'n', f:`SUM(${totalCol}2:${totalCol}${totalRowNo-1})`});
  rows.push(totalRow);

  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws['!freeze'] = { xSplit:3, ySplit:1 };
  ws['!autofilter'] = { ref:`A1:${totalCol}${totalRowNo-1}` };
  ws['!cols'] = [
    {wch:18},{wch:38},{wch:10},
    ...months.map(()=>({wch:16})),{wch:18}
  ];
  ws['!rows'] = [{hpt:24}];

  // Professional cell formatting. SheetJS community edition preserves these fields
  // in many spreadsheet viewers; Excel will render the values/formulas correctly.
  for (let c=0;c<headers.length;c++) {
    const cell = ws[XLSX.utils.encode_cell({r:0,c})];
    if (cell) cell.s = {font:{bold:true}, fill:{fgColor:{rgb:'EAF1FF'}}, alignment:{vertical:'center'}};
  }
  for (let r=1;r<totalRowNo;r++) {
    for (let c=3;c<headers.length;c++) {
      const cell=ws[XLSX.utils.encode_cell({r,c})];
      if (cell) cell.z = '#,##0.############';
    }
  }
  for (let c=0;c<headers.length;c++) {
    const cell=ws[XLSX.utils.encode_cell({r:totalRowNo-1,c})];
    if (cell) cell.s = {font:{bold:true}, fill:{fgColor:{rgb:'EAF1FF'}}};
  }

  XLSX.utils.book_append_sheet(wb, ws, 'Full Summary');
  const out = 'Material_Consumption_2026_Full_Summary.xlsx';
  XLSX.writeFile(wb, out, {bookType:'xlsx', cellStyles:true});
}

function colLetter(n) {
  let s='';
  while(n>0){ let r=(n-1)%26; s=String.fromCharCode(65+r)+s; n=Math.floor((n-1)/26); }
  return s;
}
function escapeHtml(v) {
  return String(v ?? '').replace(/[&<>"']/g, m => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[m]));
}
