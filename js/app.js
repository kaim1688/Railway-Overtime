import { KEYS, REGIMES, datesInMonth, previousMonth } from './rules.js';
import { readWorkbook, reportWorkbook, rosterBlocks, rosterMonth } from './excel.js';
import { defaultCalendars, parseHolidayInput, calculateImportedRoster } from './workflow.js';
import { sampleRows, templateRows } from './sample.js';
const $=id=>document.getElementById(id);
const XLSX=globalThis.XLSX;
const state={current:null,previous:null,report:null,demo:false,busy:false,calendars:defaultCalendars(),dateDrafts:new Set()};
const fmt=n=>Number(n.toFixed(1)).toString();
function node(tag,value,className) {const e=document.createElement(tag);e.textContent=value;if(className)e.className=className;return e;}
function message(value,error=false) {$('message').replaceChildren(...(value?[node('p',value,`feedback${error?' error':''}`)]:[]));}
function invalidate() {state.report=null;$('export').disabled=true;$('report').hidden=true;$('empty').hidden=false;$('result-caption').textContent='等待完整的班表與日期設定。';message('');}
function refreshDates() {
  const month=$('month').value;
  for(const prefix of ['', 'prev-']) {
    let target=month;
    if(prefix && month)target=previousMonth(month);
    const entry=state.calendars[target];
    $(prefix+'holidays').value=entry?.dates.join('\n') || '';
    $(prefix+'holiday-source').textContent=entry?entry.source:'此月份尚未設定國定本日；空白不代表沒有國定日。';
    $(prefix+'holiday-summary').textContent=target?(entry?`${target}：${entry.dates.length?entry.dates.map(d=>d.slice(5).replace('-','/')).join('、'):'無國定日'}`:`${target}：日期尚未設定`):'請先指定結算年月。';
    state.dateDrafts.delete(prefix);
  }
}
function inspectCurrent() {
  $('regime-field').hidden=true;
  const book=state.current,sheet=$('sheet').value;
  if(!book || !sheet)return;
  const blocks=rosterBlocks(XLSX,book.wb,sheet);
  $('regime-field').hidden=blocks.length>0;
  $('month').value=rosterMonth(XLSX,book.wb,sheet);
  refreshDates();
}
function setBook(which,wb,name) {
  state[which]={wb,name};
  const prefix=which==='current'?'':'prev-',select=$(prefix+'sheet');
  const options=wb.SheetNames.map(s=>{const o=node('option',s);o.value=s;return o;});
  if(wb.SheetNames.length>1){const o=node('option','請選擇班表工作表');o.value='';options.unshift(o);}
  select.replaceChildren(...options);select.disabled=false;
  if(which==='current') {
    $('sheet-field').hidden=wb.SheetNames.length===1;
    $('file-state').textContent=`已讀取 ${name}（${wb.SheetNames.length} 張工作表）。`;
    inspectCurrent();
  }
}
function busy(value) {
  state.busy=value;
  for(const id of ['file','prev-file','sample','month','template','sheet','prev-sheet','regime','clear-prev','holidays','prev-holidays','apply-holidays','prev-apply-holidays','no-holidays','prev-no-holidays']) {
    $(id).disabled=value || (id==='sheet' && !state.current) || (id==='prev-sheet' && !state.previous);
  }
}
async function importFile(event,which) {
  const file=event.target.files[0];if(!file)return;
  invalidate();state[which]=null;
  if(which==='current'){state.demo=false;$('regime').value='';$('month').value='';refreshDates();}
  busy(true);
  try {
    if(!/\.xlsx$/i.test(file.name))throw new Error('請匯入 .xlsx 格式。');
    if(file.size>20*1024*1024)throw new Error('檔案超過 20MB，請另存所需工作表。');
    setBook(which,readWorkbook(XLSX,await file.arrayBuffer()),file.name);
    run();
  } catch(error) {
    state[which]=null;
    const select=$(which==='current'?'sheet':'prev-sheet');select.replaceChildren(node('option','讀取失敗，請重新選檔'));
    if(which==='current')$('file-state').textContent='本月檔案讀取失敗。';
    message(`無法匯入：${error.message}`,true);
  } finally {busy(false);event.target.value='';}
}
function setDates(prefix,empty=false) {
  invalidate();
  try {
    const month=prefix?previousMonth($('month').value):$('month').value;
    datesInMonth(month);
    const dates=empty?[]:parseHolidayInput($(prefix+'holidays').value,month);
    state.calendars[month]={dates,source:empty?'本次操作明確指定無國定日':'本次操作輸入的國定本日；僅保留於目前分頁'};
    state.dateDrafts.delete(prefix);
    $(prefix+'holidays').value=dates.join('\n');
    $(prefix+'holiday-source').textContent=state.calendars[month].source;
    $(prefix+'holiday-summary').textContent=`${month}：${dates.length?dates.map(d=>d.slice(5).replace('-','/')).join('、'):'無國定日'}`;
    run();
  }catch(error){message(error.message,true);}
}
function run() {
  invalidate();if(!state.current)return;
  try {
    const month=$('month').value;
    if(!$('sheet').value)throw new Error('檔案有多張工作表，請選擇本月班表。');
    if(!month)throw new Error('無法辨識班表年月，請指定結算年月。');
    if(state.dateDrafts.size)throw new Error('國定日期仍在編輯或有錯誤，請完成日期設定。');
    const {report,parsed:current,prior:previous}=calculateImportedRoster(XLSX,{
      current:{...state.current,sheetName:$('sheet').value},previous:state.previous?{...state.previous,sheetName:$('prev-sheet').value}:null,
      month,regime:$('regime').value,calendars:state.calendars});
    state.report=report;
    report.metadata={'本月檔案':state.current.name,'本月工作表':current.sheetName,'前月檔案':state.previous?.name || '未提供','前月工作表':previous?.sheetName || '未提供',
      '資料性質':state.demo?'合成範例，非正式班表':'使用者匯入，依設定規則試算','國定日期來源':state.calendars[month].source,
      '前月國定日期來源':previous?(state.calendars[previousMonth(month)]?.source || '尚未設定'):'未使用',
      '本月計算範圍':current.selection,'前月計算範圍':previous?.selection || '未使用',
      '適用班制':[...new Set(report.employees.map(p=>REGIMES[p.regime]))].join('、'),
      '排除人員':current.excluded.map(p=>`${p.name}（${p.source}，${p.reason}）`).join('；') || '無',
      '前三位排除依據':current.excluded.some(p=>p.reason.startsWith('下方'))?'原版下方第1位站長、第2–3位總務；依既定位置排除':'未套用位置排除'};
    $('import-info').textContent=`${state.current.name}｜${current.sheetName}｜表頭第 ${current.headerRows?.join('、') || current.headerRow} 列｜${current.dateCount} 個日期｜${current.selection}`;
    renderReport();
  }catch(error){message(`未完成計算：${error.message}`,true);}
}
function renderReport() {
  const report=state.report;$('empty').hidden=true;$('report').hidden=false;$('export').disabled=false;
  $('result-caption').textContent=`${report.month} · ${report.employees.length} 位員工（三班制 ${report.employees.filter(p=>p.regime==='threeShift').length}／常日班 ${report.employees.filter(p=>p.regime==='regular').length}）· 單位 H${state.demo?' · 合成範例':''}`;
  $('excluded').textContent=report.excluded.length?`未計算 ${report.excluded.length} 位：`+report.excluded.map(p=>`${p.name}（${p.source}，${p.reason}）`).join('；'):'此範圍沒有排除人員。';
  $('report-state').className=`report-state${report.confirmed?'':' pending'}`;
  $('report-state').textContent=report.confirmed?`計算完成：${report.employees.length} 人，無資料異常。`:`需核對：${report.employees.filter(p=>!p.confirmed).length} 人有資料疑義，總數為可計算部分的暫計。`;
  const body=document.querySelector('#summary tbody');body.replaceChildren();
  report.employees.forEach((p,i)=>{
    const tr=document.createElement('tr'), first=document.createElement('td'), link=node('button',p.name+(p.id?` · ${p.id}`:''),'person-link');
    link.type='button';link.addEventListener('click',()=>{$('employee').value=String(i);renderDaily();$('employee').focus();});first.append(link,node('span',REGIMES[p.regime],'source'));tr.append(first);
    KEYS.forEach(k=>tr.append(node('td',fmt(p.totals[k]))));tr.append(node('td',fmt(p.totals.total),'total-hours'));if(!report.confirmed)tr.append(node('td',p.confirmed?'':'需核對','status'+(p.confirmed?'':' pending')));body.append(tr);
  });
  const total=document.createElement('tr');total.append(node('td','全體合計'));KEYS.forEach(k=>total.append(node('td',fmt(report.totals[k]))));total.append(node('td',fmt(report.totals.total),'total-hours'));if(!report.confirmed)total.append(node('td','暫計'));$('status-heading').hidden=report.confirmed;document.querySelector('#summary tfoot').replaceChildren(total);
  const issues=[...report.issues.map(m=>({name:'全體',date:'',message:m})),...report.employees.flatMap(p=>p.issues.map(i=>({...i,name:p.name})))];
  $('metric-people').textContent=String(report.employees.length);
  $('metric-hours').textContent=fmt(report.totals.total);
  $('metric-hours-label').textContent=report.confirmed?'本月總時數':'本月暫計時數';
  $('metric-issues').textContent=String(issues.length);
  $('metric-check').classList.toggle('pending',issues.length>0);
  $('metric-check-note').textContent=issues.length?'請查看下方待核對事項':'未偵測到資料異常';
  $('issues-section').hidden=!issues.length;
  $('issue-count').textContent=`（${issues.length}）`;
  $('issues').replaceChildren(...(issues.length?issues.map(i=>{const li=document.createElement('li');li.append(node('strong',`${i.name} ${i.date}　`),document.createTextNode(i.message));return li;}):[]));
  $('employee').replaceChildren(...report.employees.map((p,i)=>{const o=node('option',p.name+' · '+REGIMES[p.regime]+(p.id?` · ${p.id}`:''));o.value=String(i);return o;}));renderDaily();
}
function renderDaily() {
  const p=state.report?.employees[Number($('employee').value)];if(!p)return;
  document.querySelectorAll('#summary tbody tr').forEach((row,i)=>row.classList.toggle('selected',i===Number($('employee').value)));
  const body=document.querySelector('#daily tbody');body.replaceChildren();
  for(const r of p.detail) {
    const tr=document.createElement('tr');if(r.isHoliday)tr.className='holiday';const h=r.hours;
    for(const value of [r.date+(r.isHoliday?' 國定':''),r.symbol || '空白',fmt(h.ordinary),[h.rest02,h.rest38,h.rest912].map(fmt).join(' / '),[h.holiday08,h.holiday910,h.holiday1112].map(fmt).join(' / '),fmt(h.holidayRest),fmt(KEYS.reduce((n,k)=>n+h[k],0))])tr.append(node('td',value));
    const note=document.createElement('td');note.append(node('span',r.note || '待核對'));
    if(r.source)note.append(node('span',r.source,'source'));
    if(r.transfer)note.append(node('span',r.transfer,'source'));
    if(r.issues.length)note.append(node('p',r.issues.join('；'),'error-text'));
    tr.append(note);body.append(tr);
  }
  const h=p.totals, total=document.createElement('tr'), label=node('th',`${p.name} 個人總計`);
  label.scope='row';label.colSpan=2;total.append(label);
  for(const value of [fmt(h.ordinary),[h.rest02,h.rest38,h.rest912].map(fmt).join(' / '),[h.holiday08,h.holiday910,h.holiday1112].map(fmt).join(' / '),fmt(h.holidayRest),fmt(h.total),p.confirmed&&!state.report.issues.length?'單位：小時（H）':'暫計：有待核對事項，僅含可計算時數'])total.append(node('td',value));
  document.querySelector('#daily tfoot').replaceChildren(total);
}
function saveBook(wb,name) {try {XLSX.writeFile(wb,name,{bookType:'xlsx'});}catch(e){message(`匯出失敗：${e.message}`,true);}}
$('file').addEventListener('change',e=>importFile(e,'current'));
$('prev-file').addEventListener('change',e=>importFile(e,'previous'));
$('sheet').addEventListener('change',()=>{invalidate();$('regime').value='';try{inspectCurrent();run();}catch(e){message(e.message,true);}});
$('prev-sheet').addEventListener('change',run);
$('regime').addEventListener('change',run);
$('month').addEventListener('change',()=>{invalidate();state.demo=false;try{datesInMonth($('month').value);refreshDates();run();}catch(e){message(e.message,true);}});
for(const prefix of ['', 'prev-']) {
  $(prefix+'holidays').addEventListener('input',()=>{state.dateDrafts.add(prefix);invalidate();message('日期編輯中，完成後按「套用日期」重新計算。');});
  $(prefix+'apply-holidays').addEventListener('click',()=>setDates(prefix));
  $(prefix+'no-holidays').addEventListener('click',()=>setDates(prefix,true));
}
$('clear-prev').addEventListener('click',()=>{state.previous=null;$('prev-sheet').replaceChildren(node('option','尚未提供'));$('prev-sheet').disabled=true;state.dateDrafts.delete('prev-');run();});
$('employee').addEventListener('change',renderDaily);
$('export').addEventListener('click',()=>{if(state.report)saveBook(reportWorkbook(XLSX,state.report,state.report.metadata),`班表時數_${state.report.month}${state.report.confirmed?'':'_暫計'}${state.demo?'_合成範例':''}.xlsx`);});
$('template').addEventListener('click',()=>{try{const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(templateRows($('month').value)),'班表');saveBook(wb,`空白班表範本_${$('month').value}.xlsx`);}catch(e){message(e.message,true);}});
$('sample').addEventListener('click',()=>{
  invalidate();state.previous=null;$('prev-sheet').replaceChildren(node('option','尚未提供'));$('prev-sheet').disabled=true;
  const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(sampleRows()),'合成範例');
  setBook('current',wb,'內建合成範例（非正式班表）');$('regime').value='threeShift';state.demo=true;run();
});
refreshDates();
if(!XLSX){message('本機 SheetJS 程式庫載入失敗，請確認 dist/vendor/xlsx.full.min.js 完整存在。',true);busy(true);}
