import { calculate, KEYS, datesInMonth, previousMonth, DEFAULT_HOLIDAYS, validDate } from './rules.js';
import { readWorkbook, parseSheet, combineEmployees, reportWorkbook, rosterBlocks } from './excel.js';
import { sampleRows, templateRows } from './sample.js';
const $=id=>document.getElementById(id);
const XLSX=globalThis.XLSX;
const state={current:null,previous:null,report:null,demo:false,busy:false};
const fmt=n=>Number(n.toFixed(1)).toString();
function node(tag,value,className) {const e=document.createElement(tag);e.textContent=value;if(className)e.className=className;return e;}
function message(value,error=false) {$('message').replaceChildren(...(value?[node('p',value,`feedback${error?' error':''}`)]:[]));}
function invalidate() {state.report=null;$('export').disabled=true;$('report').hidden=true;$('empty').hidden=false;$('result-caption').textContent='設定已更新，請重新辨識班表並計算。';message('');}
function refreshBlocks(which) {
  const prefix=which==='current'?'':'prev-',book=state[which],select=$(prefix+'block');
  select.replaceChildren();$(prefix+'block-field').hidden=true;
  if(!book)return;
  try {
    const blocks=rosterBlocks(XLSX,book.wb,$(prefix+'sheet').value);
    if(!blocks.length)return;
    const placeholder=node('option','請選擇適用三班制的區段');placeholder.value='';select.append(placeholder);
    blocks.forEach((b,i)=>{const option=node('option',b.label+(i?'（本版未納入）':'（上方三班制）'));option.value=b.id;option.disabled=i>0;select.append(option);});
    $(prefix+'block-field').hidden=false;
  }catch(error){message(error.message,true);}
}
function setBook(which,wb,name) {
  state[which]={wb,name};
  const select=$(which==='current'?'sheet':'prev-sheet');select.replaceChildren(...wb.SheetNames.map(s=>{const o=node('option',s);o.value=s;return o;}));
  select.disabled=false;
  refreshBlocks(which);
  if(which==='current') {$('calculate').disabled=false;$('file-state').textContent=`已讀取 ${name}，共 ${wb.SheetNames.length} 張工作表。請選擇要計算的班表。`;}
}
function busy(value) {
  state.busy=value;
  for (const id of ['file','prev-file','sample','calculate','month','template','sheet','prev-sheet','block','prev-block','clear-prev','holidays','prev-holidays','confirmed','prev-confirmed']) {
    $(id).disabled=value || (id==='calculate' && !state.current) || (id==='sheet' && !state.current) || (id==='prev-sheet' && !state.previous);
  }
}
async function importFile(event,which) {
  const file=event.target.files[0];if(!file)return;
  invalidate();state[which]=null;
  refreshBlocks(which);
  if(which==='current') {
    if(state.demo)$('confirmed').checked=false;
    state.demo=false;
    $('holiday-source').textContent='請依本月原班表與國定本日清單核對；預填值不代表正式核定。';
  }
  busy(true);
  try {
    if(!/\.xlsx$/i.test(file.name))throw new Error('請匯入 .xlsx 格式。');
    if(file.size>20*1024*1024)throw new Error('檔案超過 20MB，請另存所需工作表。');
    setBook(which,readWorkbook(XLSX,await file.arrayBuffer()),file.name);
    message('檔案已在本機讀取。請核對工作表、結算年月及國定日期，再按「辨識班表並計算」。');
  } catch(error) {
    const select=$(which==='current'?'sheet':'prev-sheet');select.replaceChildren(node('option','讀取失敗，請重新選檔'));
    if(which==='current')$('file-state').textContent='本月檔案讀取失敗。';
    message(`無法匯入：${error.message}`,true);
  } finally {busy(false);event.target.value='';}
}
function holidaysFrom(id,month) {
  const values=$(id).value.split(/[\s,，、;；]+/).filter(Boolean);
  if(values.some(d=>!validDate(d)||!d.startsWith(month)))throw new Error(`${month} 國定日期請使用有效的 YYYY-MM-DD，且須在該月內。`);
  return values;
}
function run() {
  invalidate();if(!state.current)return;
  try {
    const month=$('month').value;datesInMonth(month);
    const current=parseSheet(XLSX,state.current.wb,$('sheet').value,month,$('block').value);
    const prev=previousMonth(month), previous=state.previous?parseSheet(XLSX,state.previous.wb,$('prev-sheet').value,prev,$('prev-block').value):null;
    const holidays=[...holidaysFrom('holidays',month),...(previous?holidaysFrom('prev-holidays',prev):[])];
    const confirmedMonths=[...($('confirmed').checked?[month]:[]),...(previous && $('prev-confirmed').checked?[prev]:[])];
    const inputIssues=[...current.issues,...(previous?.issues || []).map(x=>`前月：${x}`)];
    const employees=combineEmployees(current.employees,previous?.employees || []);
    state.report=calculate({month,employees,holidays,confirmedMonths,inputIssues});
    state.report.metadata={'本月檔案':state.current.name,'本月工作表':$('sheet').value,'前月檔案':state.previous?.name || '未提供','前月工作表':previous?.sheetName || '未提供','資料性質':state.demo?'合成範例，非正式班表':'使用者匯入，仍需核對原檔','符號圖例':'日、夜、$日、$夜、*日、*夜、例、休、公休、特休；家照、公假、身調假、補休（可附月/日）、10/10 國慶日未出勤皆 0H','國定日期來源':'頁面輸入且勾選確認；2026-10 預填來自規則草稿'};
    Object.assign(state.report.metadata,{'本月計算區段':current.selection,'前月計算區段':previous?.selection || '未提供','排除範圍':current.blockCount?'僅計所選區段，其他區段未納入':'無區段篩選'});
    renderReport();
    message(`${state.demo?'合成範例演練。':''}已辨識「${current.sheetName}」第 ${current.headerRow} 列表頭：${current.layout==='wide'?'日期橫排':'逐筆直列'}，${employees.length} 位員工、${current.dateCount} 個日期。${current.blockCount?`僅計算 ${current.selection}，其他區段未納入。`:''}請對照原班表核對姓名、日期與班別。`);
  } catch(error) {message(`未完成計算：${error.message}`,true);}
}
function renderReport() {
  const report=state.report;$('empty').hidden=true;$('report').hidden=false;$('export').disabled=false;
  $('result-caption').textContent=`${report.month} · ${report.employees.length} 位員工 · 單位 H${state.demo?' · 合成範例':''}`;
  $('report-state').className=`report-state${report.confirmed?'':' pending'}`;
  $('report-state').textContent=report.confirmed?'已依輸入規則完成計算，仍需總務覆核；此狀態不代表正式核定。':'暫計：有未解決資料／規則疑義，下列總數僅為可計算部分。';
  const body=document.querySelector('#summary tbody');body.replaceChildren();
  report.employees.forEach((p,i)=>{
    const tr=document.createElement('tr'), first=document.createElement('td'), link=node('button',p.name+(p.id?` · ${p.id}`:''),'person-link');
    link.type='button';link.addEventListener('click',()=>{$('employee').value=String(i);renderDaily();$('employee').focus();});first.append(link);tr.append(first);
    KEYS.forEach(k=>tr.append(node('td',fmt(p.totals[k]))));tr.append(node('td',fmt(p.totals.total)),node('td',p.confirmed?'待總務覆核':'暫計','status'+(p.confirmed?'':' pending')));body.append(tr);
  });
  const total=document.createElement('tr');total.append(node('td','全體合計'));KEYS.forEach(k=>total.append(node('td',fmt(report.totals[k]))));total.append(node('td',fmt(report.totals.total)),node('td',report.confirmed?'待總務覆核':'暫計'));document.querySelector('#summary tfoot').replaceChildren(total);
  const issues=[...report.issues.map(m=>({name:'全體',date:'',message:m})),...report.employees.flatMap(p=>p.issues.map(i=>({...i,name:p.name})))];
  $('issue-count').textContent=`（${issues.length}）`;
  $('issues').replaceChildren(...(issues.length?issues.map(i=>{const li=document.createElement('li');li.append(node('strong',`${i.name} ${i.date}　`),document.createTextNode(i.message));return li;}):[node('li','沒有偵測到資料異常；仍請對照原始班表及規則覆核。')]));
  $('employee').replaceChildren(...report.employees.map((p,i)=>{const o=node('option',p.name+(p.id?` · ${p.id}`:''));o.value=String(i);return o;}));renderDaily();
}
function renderDaily() {
  const p=state.report?.employees[Number($('employee').value)];if(!p)return;
  const body=document.querySelector('#daily tbody');body.replaceChildren();
  for(const r of p.detail) {
    const tr=document.createElement('tr');if(r.isHoliday)tr.className='holiday';const h=r.hours;
    for(const value of [r.date+(r.isHoliday?' 國定':''),r.symbol || '空白',fmt(h.ordinary),[h.rest02,h.rest38,h.rest912].map(fmt).join(' / '),[h.holiday08,h.holiday910,h.holiday1112].map(fmt).join(' / '),fmt(h.holidayRest),fmt(KEYS.reduce((n,k)=>n+h[k],0))])tr.append(node('td',value));
    const note=document.createElement('td');note.append(node('span',r.note || '待核對'));
    if(r.transfer)note.append(node('span',r.transfer,'source'));
    if(r.source)note.append(node('span',r.source,'source'));
    if(r.issues.length)note.append(node('p',r.issues.join('；'),'error-text'));
    tr.append(note);body.append(tr);
  }
}
function saveBook(wb,name) {try {XLSX.writeFile(wb,name,{bookType:'xlsx'});}catch(e){message(`匯出失敗：${e.message}`,true);}}
$('file').addEventListener('change',e=>importFile(e,'current'));
$('prev-file').addEventListener('change',e=>importFile(e,'previous'));
for(const id of ['block','prev-block','holidays','prev-holidays','confirmed','prev-confirmed'])$(id).addEventListener('input',invalidate);
for(const [id,which] of [['sheet','current'],['prev-sheet','previous']])$(id).addEventListener('change',()=>{invalidate();refreshBlocks(which);});
$('month').addEventListener('input',()=>{invalidate();state.demo=false;$('confirmed').checked=false;$('prev-confirmed').checked=false;$('holidays').value=(DEFAULT_HOLIDAYS[$('month').value] || []).join('\n');$('prev-holidays').value='';$('holiday-source').textContent=DEFAULT_HOLIDAYS[$('month').value]?'規則草稿指定本日，請核對後勾選。':'此月份無預設國定清單，請向總務取得後輸入並確認。';});
$('clear-prev').addEventListener('click',()=>{state.previous=null;$('prev-sheet').replaceChildren(node('option','尚未提供'));$('prev-sheet').disabled=true;$('prev-confirmed').checked=false;refreshBlocks('previous');invalidate();});
$('calculate').addEventListener('click',run);$('employee').addEventListener('change',renderDaily);
$('export').addEventListener('click',()=>{if(state.report)saveBook(reportWorkbook(XLSX,state.report,state.report.metadata),`三班制時數_${state.report.month}_${state.report.confirmed?'待覆核':'暫計'}${state.demo?'_合成範例':''}.xlsx`);});
$('template').addEventListener('click',()=>{try {const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(templateRows($('month').value)),'班表');saveBook(wb,`空白班表範本_${$('month').value}.xlsx`);}catch(e){message(e.message,true);}});
$('sample').addEventListener('click',()=>{
  invalidate();$('month').value='2026-10';$('holidays').value=DEFAULT_HOLIDAYS['2026-10'].join('\n');$('confirmed').checked=true;state.previous=null;$('prev-sheet').replaceChildren(node('option','尚未提供'));$('prev-sheet').disabled=true;$('prev-confirmed').checked=false;
  refreshBlocks('previous');
  $('holiday-source').textContent='合成範例固定使用草稿 10/10、10/25；正式資料請重新核對。';
  const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(sampleRows()),'合成範例');setBook('current',wb,'內建合成範例（非正式班表）');state.demo=true;run();
});
if(!XLSX){message('本機 SheetJS 程式庫載入失敗，請確認 dist/vendor/xlsx.full.min.js 完整存在。',true);busy(true);}
