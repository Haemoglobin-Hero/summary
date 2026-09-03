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
let expandedCategories = new Set();

$('fileInput').addEventListener('change', e => { if (e.target.files[0]) loadWorkbook(e.target.files[0]); });
$('replaceBtn').addEventListener('click', () => $('fileInput').click());
$('dropZone').addEventListener('dragover', e => { e.preventDefault(); $('uploadSection').classList.add('dragover'); });
$('dropZone').addEventListener('dragleave', () => $('uploadSection').classList.remove('dragover'));
$('dropZone').addEventListener('drop', e => { e.preventDefault(); $('uploadSection').classList.remove('dragover'); if (e.dataTransfer.files[0]) loadWorkbook(e.dataTransfer.files[0]); });
$('toggleMonths').onclick = toggleAllMonths;
$('toggleParts').onclick = toggleAllParts;
$('clearParts').onclick = () => { selectedParts.clear(); renderParts(); updateUI(); };
$('masterPartCheck').onchange = e => { filteredParts().forEach(p => e.target.checked ? selectedParts.add(p.partNo) : selectedParts.delete(p.partNo)); renderParts(); updateUI(); };
$('partSearch').oninput = renderParts;
$('runChecks').onclick = runVerification;
$('generateBtn').onclick = generateWorkbook;

async function loadWorkbook(file) {
  try {
    fileName = file?.name || 'workbook';
    const ext = fileName.includes('.') ? fileName.split('.').pop().toLowerCase() : '';
    if (!['xls','xlsx'].includes(ext)) throw new Error(`Unsupported file type: .${ext || 'unknown'}`);
    if (typeof XLSX === 'undefined') throw new Error('Excel parser library did not load. Refresh the page and try again.');

    setStatus(`Reading ${ext.toUpperCase()} workbook…`, false);

    // Use ArrayBuffer when available, with a FileReader fallback for older browsers.
    let data;
    if (typeof file.arrayBuffer === 'function') {
      data = await file.arrayBuffer();
    } else {
      data = await new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result);
        reader.onerror = () => reject(new Error('The browser could not read the selected file.'));
        reader.readAsArrayBuffer(file);
      });
    }
    if (!data || data.byteLength === 0) throw new Error('The selected file is empty.');

    // SheetJS supports both legacy BIFF .xls and OOXML .xlsx with type:'array'.
    workbook = XLSX.read(data, { type:'array', cellDates:true, cellNF:false, cellText:true });
    if (!workbook || !workbook.SheetNames || !workbook.SheetNames.length) {
      throw new Error('No worksheets were found in the workbook.');
    }

    parseWorkbook();
    $('fileName').textContent = fileName;
    $('workspace').classList.remove('hidden');
    $('uploadSection').classList.add('hidden');
    $('sheetCount').textContent = workbook.SheetNames.length;
    $('monthCount').textContent = monthSheets.length;
    $('partCount').textContent = masterParts.size;
    $('issueCount').textContent = conflicts.length;
    setStatus('Workbook ready', true);
    renderMonths(); renderParts(); runVerification(); updateUI();
  } catch(err) {
    console.error('Workbook load failed:', err);
    const reason = err?.message || String(err);
    alert(`Could not read “${fileName}”.\n\n${reason}\n\nPlease make sure it is a valid .xls or .xlsx workbook.`);
    setStatus('Workbook error', false);
  }
}

function setStatus(text, ready) { $('headerStatus').innerHTML = `<i></i>${escapeHtml(text)}`; $('headerStatus').classList.toggle('ready', !!ready); }
function normalize(v) { return String(v ?? '').trim().replace(/\s+/g,' ').toLowerCase(); }
function clean(v) { return String(v ?? '').trim(); }

function parseWorkbook() {
  masterParts = new Map(); conflicts = []; detectedTotals = {}; monthSheets = []; expandedCategories = new Set();
  workbook.SheetNames.forEach(sheetName => {
    const key = normalize(sheetName);
    const month = DISPLAY_MONTHS.find(m => MONTH_ALIASES[m].includes(key));
    if (!month) return;
    const ws = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(ws, { header:1, defval:null, raw:true });
    const summary = locateSummary(rows, sheetName, month);
    monthSheets.push({ month, sheetName, rows, ...summary });
  });
  monthSheets.sort((a,b) => DISPLAY_MONTHS.indexOf(a.month) - DISPLAY_MONTHS.indexOf(b.month));
  monthSheets.forEach(ms => {
    ms.items.forEach(item => {
      const partNo = clean(item.partNo); if (!partNo) return;
      const key = normalize(partNo);
      if (!masterParts.has(key)) masterParts.set(key, { partNo, description:clean(item.description), transactionCode:clean(item.transactionCode), uom:clean(item.uom), quantities:{}, sources:[] });
      const master = masterParts.get(key);
      const desc = clean(item.description), code = clean(item.transactionCode), uom = clean(item.uom);
      if (desc && master.description && normalize(desc) !== normalize(master.description)) conflicts.push({type:'description',partNo,month:ms.month,a:master.description,b:desc});
      else if (!master.description && desc) master.description = desc;
      if (code && master.transactionCode && normalize(code) !== normalize(master.transactionCode)) conflicts.push({type:'transactionCode',partNo,month:ms.month,a:master.transactionCode,b:code});
      else if (!master.transactionCode && code) master.transactionCode = code;
      if (uom && master.uom && normalize(uom) !== normalize(master.uom)) conflicts.push({type:'uom',partNo,month:ms.month,a:master.uom,b:uom});
      else if (!master.uom && uom) master.uom = uom;
      master.quantities[ms.month] = (master.quantities[ms.month] || 0) + numeric(item.quantity);
      master.sources.push(ms.month);
    });
    if (ms.total != null) detectedTotals[ms.month] = ms.total;
  });
  selectedMonths = new Set(monthSheets.map(x=>x.month));
  selectedParts = new Set([...masterParts.values()].map(x=>x.partNo));
}

function locateSummary(rows, sheetName, month) {
  const headerAliases = {
    partNo:['part no','part number','partno','part #','item code','item no'],
    description:['part description','description','item description','material description'],
    transactionCode:['transaction code','transactioncode','transaction cd','trans code','transaction'],
    uom:['uom','unit','unit of measure'],
    quantity:['quantity','qty','consumption','consumed quantity']
  };
  let best = null;
  for (let r=0;r<rows.length;r++) {
    const row=rows[r]||[], texts=row.map(normalize), indexes={};
    for (const [field,aliases] of Object.entries(headerAliases)) indexes[field]=texts.findIndex(t=>aliases.some(a=>t===a || t.includes(a)));
    const score=['partNo','description','transactionCode','uom','quantity'].reduce((s,f)=>s+(indexes[f]>=0?1:0),0);
    if (indexes.partNo>=0 && indexes.quantity>=0 && score>=3) {
      best={headerRow:r,indexes};
      if (texts.some(t=>t.includes('item-wise')||t.includes('item wise')||t.includes('summary'))) break;
    }
  }
  if (!best) best={headerRow:-1,indexes:{partNo:0,description:1,transactionCode:2,uom:3,quantity:4}};
  const items=[], start=best.headerRow+1; let total=null;
  for (let r=start;r<rows.length;r++) {
    const row=rows[r]||[];
    const p=best.indexes.partNo>=0?row[best.indexes.partNo]:row[0];
    const d=best.indexes.description>=0?row[best.indexes.description]:row[1];
    const tc=best.indexes.transactionCode>=0?row[best.indexes.transactionCode]:row[2];
    const u=best.indexes.uom>=0?row[best.indexes.uom]:row[3];
    const q=best.indexes.quantity>=0?row[best.indexes.quantity]:row[4];
    const pn=clean(p), first=normalize(pn);
    if (first==='total'||first==='grand total'||first.includes('total')) { const qv=numeric(q); if(q!==null&&q!==''&&Number.isFinite(qv)) total=qv; continue; }
    if(!pn || !Number.isFinite(numeric(q)) || isLikelyTransactionRow(row,best)) continue;
    items.push({partNo:pn,description:d,transactionCode:tc,uom:u,quantity:q});
  }
  const combined=new Map();
  items.forEach(x=>{const k=normalize(x.partNo); if(!combined.has(k)) combined.set(k,{...x,quantity:numeric(x.quantity)}); else {const existing=combined.get(k); existing.quantity+=numeric(x.quantity); if(!existing.description&&x.description) existing.description=x.description; if(!existing.transactionCode&&x.transactionCode) existing.transactionCode=x.transactionCode; if(!existing.uom&&x.uom) existing.uom=x.uom;}});
  return {items:[...combined.values()],total};
}
function isLikelyTransactionRow(row,best){ return row.filter(v=>v!==null&&v!=='').length===0; }
function numeric(v){ if(typeof v==='number') return Number.isFinite(v)?v:0; if(v===null||v===undefined||v==='') return 0; const n=Number(String(v).replace(/,/g,'').replace(/\s/g,'')); return Number.isFinite(n)?n:0; }
function fmt(n){ return Number(n||0).toLocaleString(undefined,{maximumFractionDigits:12}); }

function renderMonths(){
  const grid=$('monthsGrid'); grid.innerHTML='';
  monthSheets.forEach(ms=>{const div=document.createElement('label'); div.className='month-card '+(selectedMonths.has(ms.month)?'active':''); div.innerHTML=`<input type="checkbox" ${selectedMonths.has(ms.month)?'checked':''}><span class="month-name">${escapeHtml(ms.month)}</span><span class="sheet-name">${escapeHtml(ms.sheetName)}</span>`; div.onclick=e=>{e.preventDefault(); if(selectedMonths.has(ms.month))selectedMonths.delete(ms.month);else selectedMonths.add(ms.month); renderMonths();updateUI();renderPreview();}; grid.appendChild(div);});
  $('toggleMonths').textContent=selectedMonths.size===monthSheets.length?'Clear all':'Select all';
}
function filteredParts(){const q=normalize($('partSearch').value),list=[...masterParts.values()]; if(!q)return list; return list.filter(p=>normalize(p.partNo).includes(q)||normalize(p.description).includes(q)||normalize(p.transactionCode).includes(q));}
function getPartCategory(partNo){const raw=clean(partNo),normalized=raw.replace(/\s+/g,'').toUpperCase(); const match=normalized.match(/^([A-Z][A-Z0-9]*?)(?:[-_]?\d+|[-_])/); if(match&&match[1])return `${match[1]} Set`; const fallback=normalized.match(/^([A-Z]{2,})\d+/); if(fallback&&fallback[1])return `${fallback[1]} Set`; const firstToken=raw.split(/[-_\s]+/)[0]; if(/^[A-Za-z]{2,}$/.test(firstToken)&&firstToken.length<=12)return `${firstToken.toUpperCase()} Set`; return 'Other Set';}
function groupParts(parts){const groups=new Map();parts.forEach(p=>{const category=getPartCategory(p.partNo);if(!groups.has(category))groups.set(category,[]);groups.get(category).push(p);});return [...groups.entries()].sort((a,b)=>a[0].localeCompare(b[0],undefined,{numeric:true,sensitivity:'base'})).map(([category,items])=>[category,items.sort((a,b)=>a.partNo.localeCompare(b.partNo,undefined,{numeric:true,sensitivity:'base'}))]);}
function renderParts(){
  const body=$('partsBody'),list=filteredParts(),groups=groupParts(list); $('visibleCount').textContent=`${list.length} part${list.length===1?'':'s'} · ${groups.length} set${groups.length===1?'':'s'}`; body.innerHTML='';
  groups.forEach(([category,items])=>{
    const categoryId=`category-${normalize(category).replace(/[^a-z0-9]+/g,'-')}`,expanded=expandedCategories.has(category),selectedCount=items.filter(p=>selectedParts.has(p.partNo)).length,allSelected=items.length>0&&selectedCount===items.length,someSelected=selectedCount>0&&!allSelected;
    const groupRow=document.createElement('tr'); groupRow.className='part-category-row'; groupRow.innerHTML=`<td class="category-check-cell"><input type="checkbox" aria-label="Select ${escapeHtml(category)}" ${allSelected?'checked':''}></td><td colspan="5"><div class="category-title"><button type="button" class="category-toggle" aria-expanded="${expanded}" aria-controls="${categoryId}"><span class="category-chevron">${expanded?'▾':'▸'}</span><strong>${escapeHtml(category)}</strong><span class="category-count">${selectedCount}/${items.length}</span></button></div></td>`;
    const categoryCheck=groupRow.querySelector('input'); categoryCheck.indeterminate=someSelected; categoryCheck.onclick=e=>{e.stopPropagation();items.forEach(p=>allSelected?selectedParts.delete(p.partNo):selectedParts.add(p.partNo));renderParts();updateUI();};
    groupRow.querySelector('.category-toggle').onclick=e=>{e.stopPropagation();if(expanded)expandedCategories.delete(category);else expandedCategories.add(category);renderParts();}; body.appendChild(groupRow);
    if(!expanded)return;
    items.forEach(p=>{const tr=document.createElement('tr'),checked=selectedParts.has(p.partNo);tr.className=checked?'selected part-child-row':''; const months=[...new Set(p.sources)].map(x=>`<span class="tag">${x}</span>`).join(''); tr.innerHTML=`<td></td><td><strong>${escapeHtml(p.partNo)}</strong></td><td>${escapeHtml(p.description||'—')}</td><td>${escapeHtml(p.transactionCode||'—')}</td><td>${escapeHtml(p.uom||'—')}</td><td><div class="month-tags">${months||'—'}</div></td>`; tr.onclick=e=>{if(e.target.tagName==='INPUT')return;togglePart(p.partNo);}; body.appendChild(tr);});
  });
  const visibleSelected=list.filter(p=>selectedParts.has(p.partNo)).length; $('masterPartCheck').checked=list.length>0&&visibleSelected===list.length; $('masterPartCheck').indeterminate=visibleSelected>0&&visibleSelected<list.length; $('selectionCount').textContent=`${selectedParts.size} selected`;
}
function togglePart(partNo){if(selectedParts.has(partNo))selectedParts.delete(partNo);else selectedParts.add(partNo);renderParts();updateUI();}
function toggleAllParts(){const all=filteredParts(),select=!all.length||!all.every(p=>selectedParts.has(p.partNo));all.forEach(p=>select?selectedParts.add(p.partNo):selectedParts.delete(p.partNo));renderParts();updateUI();}
function toggleAllMonths(){if(selectedMonths.size===monthSheets.length)selectedMonths.clear();else selectedMonths=new Set(monthSheets.map(x=>x.month));renderMonths();updateUI();renderPreview();}
function updateUI(){const n=selectedParts.size,m=selectedMonths.size;$('previewBadge').textContent=`${n} part${n===1?'':'s'} · ${m} month${m===1?'':'s'}`;$('generateSummary').textContent=n&&m?`Ready to generate ${n} selected part${n===1?'':'s'} across ${m} selected month${m===1?'':'s'}.`:'Select parts and months to continue.';$('generateSub').textContent=conflicts.length?`${conflicts.length} data consistency issue${conflicts.length===1?'':'s'} found — review Verification.`:'The generated file will be a new workbook.';$('generateBtn').disabled=!(n&&m);$('issueCount').textContent=conflicts.length;renderPreview();}
function runVerification(){const box=$('verification'),checks=[],unique=masterParts.size;checks.push({ok:unique>0,text:`${unique} unique Part No.${unique===1?'':'s'} detected across ${monthSheets.length} monthly sheet${monthSheets.length===1?'':'s'}.`});const dup=detectDuplicatePartNumbers();checks.push({ok:dup===0,text:dup?`${dup} duplicate Part No. key${dup===1?'':'s'} were detected and consolidated.`:'No duplicate Part Nos. exist in the master list.'});const uomIssues=conflicts.filter(x=>x.type==='uom');checks.push({ok:uomIssues.length===0,text:uomIssues.length?`${uomIssues.length} UoM conflict${uomIssues.length===1?'':'s'} need review.`:'No UoM discrepancies found.'});const descIssues=conflicts.filter(x=>x.type==='description');checks.push({ok:descIssues.length===0,text:descIssues.length?`${descIssues.length} description conflict${descIssues.length===1?'':'s'} need review.`:'No Part Description discrepancies found.'});const tcIssues=conflicts.filter(x=>x.type==='transactionCode');checks.push({ok:tcIssues.length===0,text:tcIssues.length?`${tcIssues.length} Transaction Code conflict${tcIssues.length===1?'':'s'} need review.`:'No Transaction Code discrepancies found.'});checks.push({ok:monthSheets.length>0,text:`Detected months: ${monthSheets.map(x=>x.month).join(', ')||'none'}.`});checks.push({ok:selectedParts.size>0&&selectedMonths.size>0,text:`Current selection: ${selectedParts.size} parts and ${selectedMonths.size} months.`});box.innerHTML=checks.map(c=>`<div class="check ${c.ok?'ok':'warn'}"><span class="mark">${c.ok?'✓':'!'}</span><span>${escapeHtml(c.text)}</span></div>`).join('');}
function detectDuplicatePartNumbers(){let count=0;monthSheets.forEach(ms=>{const seen=new Set();ms.items.forEach(i=>{const k=normalize(i.partNo);if(seen.has(k))count++;seen.add(k);});});return count;}
function renderPreview(){const months=DISPLAY_MONTHS.filter(m=>selectedMonths.has(m)),head=$('previewHead'),body=$('previewBody');head.innerHTML=`<tr><th>Part No.</th><th>Part Description</th><th>Transaction Code</th><th>UoM</th>${months.map(m=>`<th>${m} Quantity</th>`).join('')}<th>Total Quantity</th></tr>`;body.innerHTML='';[...masterParts.values()].filter(p=>selectedParts.has(p.partNo)).slice(0,30).forEach(p=>{const vals=months.map(m=>numeric(p.quantities[m])),total=vals.reduce((a,b)=>a+b,0);body.insertAdjacentHTML('beforeend',`<tr><td><strong>${escapeHtml(p.partNo)}</strong></td><td>${escapeHtml(p.description||'—')}</td><td>${escapeHtml(p.transactionCode||'—')}</td><td>${escapeHtml(p.uom||'—')}</td>${vals.map(v=>`<td>${fmt(v)}</td>`).join('')}<td><strong>${fmt(total)}</strong></td></tr>`);});if(selectedParts.size>30)body.insertAdjacentHTML('beforeend',`<tr><td colspan="${4+months.length+1}" class="muted">Preview limited to the first 30 selected parts. The generated workbook contains all selected parts.</td></tr>`);}
function generateWorkbook(){if(!selectedParts.size||!selectedMonths.size)return;const months=DISPLAY_MONTHS.filter(m=>selectedMonths.has(m)),wb=XLSX.utils.book_new();const headers=['Part No.','Part Description','Transaction Code','UoM',...months.map(m=>`${m} Quantity`),'Total Quantity'];const rows=[headers];const selected=[...masterParts.values()].filter(p=>selectedParts.has(p.partNo));selected.forEach((p,idx)=>{const rowNo=idx+2,vals=months.map(m=>numeric(p.quantities[m]));rows.push([p.partNo,p.description,p.transactionCode,p.uom,...vals,{t:'n',f:`SUM(${colLetter(5)}${rowNo}:${colLetter(4+months.length)}${rowNo})`}]);});const totalRowNo=rows.length+1,totalRow=['TOTAL','','',''];for(let i=0;i<months.length;i++){const col=colLetter(5+i);totalRow.push({t:'n',f:`SUM(${col}2:${col}${totalRowNo-1})`});}const totalCol=colLetter(5+months.length);totalRow.push({t:'n',f:`SUM(${totalCol}2:${totalCol}${totalRowNo-1})`});rows.push(totalRow);const ws=XLSX.utils.aoa_to_sheet(rows);ws['!freeze']={xSplit:4,ySplit:1};ws['!autofilter']={ref:`A1:${totalCol}${totalRowNo-1}`};ws['!cols']=[{wch:18},{wch:38},{wch:20},{wch:10},...months.map(()=>({wch:16})),{wch:18}];ws['!rows']=[{hpt:24}];for(let c=0;c<headers.length;c++){const cell=ws[XLSX.utils.encode_cell({r:0,c})];if(cell)cell.s={font:{bold:true},fill:{fgColor:{rgb:'EAF1FF'}},alignment:{vertical:'center'}};}for(let r=1;r<totalRowNo;r++)for(let c=4;c<headers.length;c++){const cell=ws[XLSX.utils.encode_cell({r,c})];if(cell)cell.z='#,##0.############';}for(let c=0;c<headers.length;c++){const cell=ws[XLSX.utils.encode_cell({r:totalRowNo-1,c})];if(cell)cell.s={font:{bold:true},fill:{fgColor:{rgb:'EAF1FF'}}};}XLSX.utils.book_append_sheet(wb,ws,'Full Summary');XLSX.writeFile(wb,'Material_Consumption_2026_Full_Summary.xlsx',{bookType:'xlsx',cellStyles:true});}
function colLetter(n){let s='';while(n>0){let r=(n-1)%26;s=String.fromCharCode(65+r)+s;n=Math.floor((n-1)/26);}return s;}
function escapeHtml(v){return String(v??'').replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[m]));}
