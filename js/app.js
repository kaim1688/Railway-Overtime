import { KEYS, LABELS, REGIMES, datesInMonth, previousMonth, STARRED_ATTENDANCE_DATE, TYPHOON_KEYS } from './rules.js?v=8ec2bf395a99';
import { readWorkbook, reportWorkbook, rosterBlocks, rosterMonth, hasRosterLayout } from './excel.js?v=8ec2bf395a99';
import { defaultCalendars, parseHolidayInput, calculateImportedRoster } from './workflow.js?v=8ec2bf395a99';
import { sampleRows, templateRows } from './sample.js?v=8ec2bf395a99';
import {initPrinting,clearPrint} from './print.js?v=8ec2bf395a99';
import {initCatEffect} from './cat.js?v=8ec2bf395a99';
import {initLanding} from './landing.js?v=8ec2bf395a99';
const $=id=>document.getElementById(id);
const XLSX=globalThis.XLSX;
const mobile=window.matchMedia('(max-width:720px)');
mobile.addEventListener('change',()=>{if(!mobile.matches)$('setup-fold').open=true;});
const state={current:null,previous:null,report:null,demo:false,busy:false,calendars:defaultCalendars(),dateDrafts:new Set()};
let receiptTimer=null;
function clearExportReceipt() {
  clearTimeout(receiptTimer);receiptTimer=null;
  $('export-receipt').hidden=true;$('export-receipt').classList.remove('receipt-motion');
}
function showExportReceipt() {
  if(receiptTimer!==null)return;
  const receipt=$('export-receipt');receipt.hidden=false;
  receipt.classList.toggle('receipt-motion',document.documentElement.dataset.motion==='on');
  receiptTimer=setTimeout(clearExportReceipt,1800);
}
new MutationObserver(()=>{
  if(document.documentElement.dataset.screen!=='workbench')clearExportReceipt();
  if(document.documentElement.dataset.motion!=='on')$('export-receipt').classList.remove('receipt-motion');
}).observe(document.documentElement,{attributes:true,attributeFilter:['data-screen','data-motion']});
document.addEventListener('visibilitychange',()=>{if(document.hidden)clearExportReceipt();});
document.addEventListener('keydown',event=>{if(event.key==='Escape')clearExportReceipt();});
const fmt=n=>Number(n.toFixed(1)).toString();
const specialDateNote=month=>STARRED_ATTENDANCE_DATE.startsWith(month || 'none')?'；11/28 出勤與三班例休另依特例':'';
function node(tag,value,className) {const e=document.createElement(tag);e.textContent=value;if(className)e.className=className;return e;}
function message(value,error=false) {$('message').replaceChildren(...(value?[node('p',value,`feedback${error?' error':''}`)]:[]));}
function invalidate() {clearExportReceipt();$('setup-fold').open=true;$('setup-caption').textContent='選擇或調整班表設定';state.report=null;$('typhoon-section').hidden=true;clearPrint();$('print-summary').disabled=true;$('print-employee').disabled=true;$('export').disabled=true;$('report').hidden=true;$('empty').hidden=false;$('result-caption').textContent='等待完整的班表與日期設定。';message('');}
function refreshDates() {
  const month=$('month').value;
  for(const prefix of ['', 'prev-']) {
    let target=month;
    if(prefix && month)target=previousMonth(month);
    const entry=state.calendars[target];
    $(prefix+'holidays').value=entry?.dates.join('\n') || '';
    $(prefix+'holiday-source').textContent=entry?entry.source:'此月份尚未設定國定本日；空白不代表沒有國定日。';
    $(prefix+'holiday-summary').textContent=target?(entry?`${target}：${entry.dates.length?entry.dates.map(d=>d.slice(5).replace('-','/')).join('、'):'無一般國定日'}${specialDateNote(target)}`:`${target}：日期尚未設定`):'請先指定結算年月。';
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
    const workbook=readWorkbook(XLSX,await file.arrayBuffer());
    if(which==='current') {
      let layoutError;
      const found=workbook.SheetNames.some(name=>{
        try{return hasRosterLayout(XLSX,workbook,name);}
        catch(error){layoutError ??= error;return false;}
      });
      if(!found)throw layoutError || new Error('找不到可辨識的班表。請改選含姓名、日期與班別的班表，或使用空白範本。');
    }
    setBook(which,workbook,file.name);
    run();
  } catch(error) {
    state[which]=null;
    const select=$(which==='current'?'sheet':'prev-sheet');select.replaceChildren(node('option','讀取失敗，請重新選檔'));
    if(which==='current')$('file-state').textContent='本月檔案讀取失敗。';
    message(`無法匯入：${error.message}`,true);
    return error.message;
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
    $(prefix+'holiday-summary').textContent=`${month}：${dates.length?dates.map(d=>d.slice(5).replace('-','/')).join('、'):'無一般國定日'}${specialDateNote(month)}`;
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
    $('setup-caption').textContent=`${report.month} · ${report.employees.length} 人 · 點開更換班表`;
    if(mobile.matches){$('setup-fold').open=false;$('results-title').scrollIntoView({block:'start'});}
  }catch(error){message(`未完成計算：${error.message}`,true);}
}
function selectEmployee(index) {
  $('employee').value=String(index);renderDaily();
  $('daily-section').scrollIntoView({block:'start'});$('employee').focus({preventScroll:true});
}
function renderReport() {
  const report=state.report;$('empty').hidden=true;$('report').hidden=false;$('export').disabled=false;
  $('print-summary').disabled=!report.employees.length;
  $('result-caption').textContent=`${report.month} · ${report.employees.length} 位員工（三班制 ${report.employees.filter(p=>p.regime==='threeShift').length}／常日班 ${report.employees.filter(p=>p.regime==='regular').length}）· 單位 H${state.demo?' · 合成範例':''}`;
  $('excluded').textContent=report.excluded.length?`未計算 ${report.excluded.length} 位：`+report.excluded.map(p=>`${p.name}（${p.source}，${p.reason}）`).join('；'):'此範圍沒有排除人員。';
  $('report-state').className=`report-state${report.confirmed?'':' pending'}`;
  $('report-state').textContent=report.confirmed?`計算完成：${report.employees.length} 人，無資料異常。`:`需核對：${report.employees.filter(p=>!p.confirmed).length} 人有資料疑義，總數為可計算部分的暫計。`;
  const body=document.querySelector('#summary tbody');body.replaceChildren();
  report.employees.forEach((p,i)=>{
    const tr=document.createElement('tr'), first=document.createElement('td'), link=node('button',p.name+(p.id?` · ${p.id}`:''),'person-link');
    link.type='button';link.addEventListener('click',()=>selectEmployee(i));first.append(link,node('span',REGIMES[p.regime],'source'));tr.append(first);
    KEYS.forEach(k=>tr.append(node('td',fmt(p.totals[k]))));tr.append(node('td',fmt(p.totals.total),'total-hours'));if(!report.confirmed)tr.append(node('td',p.confirmed?'':'需核對','status'+(p.confirmed?'':' pending')));body.append(tr);
  });
  $('mobile-summary').replaceChildren(...report.employees.map((p,i)=>{
    const detail=document.createElement('details'),summary=document.createElement('summary');
    summary.append(node('span',p.name+(p.id?` · ${p.id}`:''),'mobile-person'),node('strong',`${fmt(p.totals.total)} H`));
    if(!p.confirmed || report.issues.length)summary.append(node('span','需核對／暫計','mobile-warning'));
    const content=node('div','','mobile-person-detail'),list=document.createElement('dl');
    content.append(node('p',REGIMES[p.regime],'help'));
    KEYS.forEach((key,index)=>{const group=document.createElement('div');group.append(node('dt',LABELS[index]),node('dd',`${fmt(p.totals[key])} H`));list.append(group);});
    const button=node('button','查看逐日明細');button.type='button';button.addEventListener('click',()=>selectEmployee(i));
    content.append(list,button);detail.append(summary,content);return detail;
  }));
  const total=document.createElement('tr');total.append(node('td','全體合計'));KEYS.forEach(k=>total.append(node('td',fmt(report.totals[k]))));total.append(node('td',fmt(report.totals.total),'total-hours'));if(!report.confirmed)total.append(node('td','暫計'));$('status-heading').hidden=report.confirmed;document.querySelector('#summary tfoot').replaceChildren(total);
  const issues=[...report.issues.map(m=>({name:'全體',date:'',message:m})),...report.employees.flatMap(p=>p.issues.map(i=>({...i,name:p.name})))];
  $('metric-people').textContent=String(report.employees.length);
  $('metric-hours').textContent=fmt(report.totals.total);
  $('metric-hours-label').textContent=report.typhoonDates.length?(report.confirmed?'一般加班時數':'一般暫計時數'):(report.confirmed?'本月總時數':'本月暫計時數');
  $('typhoon-section').hidden=!report.typhoonDates.length;
  $('summary-scope').textContent=report.typhoonDates.length?'颱風日的休息日出勤優先列於本表休息日級距，颱風不另計；其餘颱風時數請見頁面最後的獨立表。':'';
  if(report.typhoonDates.length)renderTyphoon(report);
  $('metric-issues').textContent=String(issues.length);
  $('metric-check').classList.toggle('pending',issues.length>0);
  $('metric-check-note').textContent=issues.length?'請查看下方待核對事項':'未偵測到資料異常';
  $('issues-section').hidden=!issues.length;
  $('issue-count').textContent=`（${issues.length}）`;
  $('issues').replaceChildren(...(issues.length?issues.map(i=>{const li=document.createElement('li');li.append(node('strong',`${i.name} ${i.date}　`),document.createTextNode(i.message));return li;}):[]));
  $('employee').replaceChildren(...report.employees.map((p,i)=>{const o=node('option',p.name+' · '+REGIMES[p.regime]+(p.id?` · ${p.id}`:''));o.value=String(i);return o;}));renderDaily();
}
function renderTyphoon(report) {
  $('typhoon-status-heading').hidden=report.confirmed;
  $('typhoon-caption').textContent=`${report.typhoonDates.join('、')} · ${report.confirmed?'颱風出勤合計':'颱風出勤暫計'} ${fmt(report.typhoonTotals.total)} H · 不併入上方一般／國定時數`;
  $('typhoon-summary-body').replaceChildren(...report.employees.map(p=>{
    const tr=document.createElement('tr');
    for(const value of [p.name,...TYPHOON_KEYS.map(k=>fmt(p.typhoonTotals[k])),fmt(p.typhoonTotals.total),...(report.confirmed?[]:[p.confirmed?'':'需核對／暫計'])])tr.append(node('td',value));
    return tr;
  }));
  const total=document.createElement('tr');total.append(node('th','全體合計'));
  for(const value of [...TYPHOON_KEYS.map(k=>fmt(report.typhoonTotals[k])),fmt(report.typhoonTotals.total),...(report.confirmed?[]:['暫計'])])total.append(node('td',value));
  $('typhoon-summary-total').replaceChildren(total);
  $('typhoon-daily-body').replaceChildren(...report.employees.flatMap(p=>p.detail.filter(r=>r.isTyphoonDate).map(r=>{
    const tr=document.createElement('tr');
    for(const value of [p.name,r.date,r.symbol || '空白',...TYPHOON_KEYS.map(k=>fmt(r.typhoonHours[k])),fmt(TYPHOON_KEYS.reduce((n,k)=>n+r.typhoonHours[k],0))])tr.append(node('td',value));
    const note=node('td',r.note || '待核對');if(r.source)note.append(node('span',r.source,'source'));if(r.issues.length)note.append(node('p',r.issues.join('；'),'error-text'));tr.append(note);return tr;
  })));
}
function renderDaily() {
  const p=state.report?.employees[Number($('employee').value)];if(!p)return;
  clearPrint();$('print-employee').disabled=false;
  $('daily-context').textContent=`${state.report.month} · ${p.name} · ${p.confirmed&&!state.report.issues.length?(state.report.typhoonDates.length?'一般總計':'總計'):'暫計'} ${fmt(p.totals.total)} H`;
  document.querySelectorAll('#summary tbody tr').forEach((row,i)=>row.classList.toggle('selected',i===Number($('employee').value)));
  const body=document.querySelector('#daily tbody');body.replaceChildren();
  for(const r of p.detail) {
    const tr=document.createElement('tr');if(r.isHoliday || r.isSpecialAttendance)tr.className='holiday';const h=r.hours;
    for(const value of [r.date+(r.isTyphoonDate?' 颱風':r.isSpecialAttendance?' 出勤特例':r.isHoliday?' 國定':''),r.symbol || '空白',fmt(h.ordinary),fmt(h.ordinary34),[h.rest02,h.rest38,h.rest912].map(fmt).join(' / '),[h.holiday08,h.holiday910,h.holiday1112].map(fmt).join(' / '),fmt(h.holidayRest),fmt(KEYS.reduce((n,k)=>n+h[k],0))])tr.append(node('td',value));
    const note=document.createElement('td');note.append(node('span',r.note || '待核對'));
    if(r.source)note.append(node('span',r.source,'source'));
    if(r.transfer)note.append(node('span',r.transfer,'source'));
    if(r.issues.length)note.append(node('p',r.issues.join('；'),'error-text'));
    tr.append(note);
    [...tr.children].forEach((cell,index)=>{cell.dataset.label=['日期','原班別',...LABELS.slice(0,2),'休息日 2／6／3','國定出勤 8／2／1','國定例／休','小計','依據與核對事項'][index];});
    body.append(tr);
  }
  const h=p.totals, total=document.createElement('tr'), label=node('th',`${p.name} 個人總計`);
  label.scope='row';label.colSpan=2;total.append(label);
  for(const value of [fmt(h.ordinary),fmt(h.ordinary34),[h.rest02,h.rest38,h.rest912].map(fmt).join(' / '),[h.holiday08,h.holiday910,h.holiday1112].map(fmt).join(' / '),fmt(h.holidayRest),fmt(h.total),p.confirmed&&!state.report.issues.length?'單位：小時（H）':'暫計：有待核對事項，僅含可計算時數'])total.append(node('td',value));
  [...total.children].forEach((cell,index)=>{cell.dataset.label=['個人總計',...LABELS.slice(0,2),'休息日 2／6／3','國定出勤 8／2／1','國定例／休','總時數','核對狀態'][index];});
  document.querySelector('#daily tfoot').replaceChildren(total);
}
function saveBook(wb,name) {try {XLSX.writeFile(wb,name,{bookType:'xlsx'});return true;}catch(e){clearExportReceipt();message(`匯出失敗：${e.message}`,true);return false;}}
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
$('export').addEventListener('click',()=>{
  if(!state.report)return;
  try {
    if(saveBook(reportWorkbook(XLSX,state.report,state.report.metadata),`班表時數_${state.report.month}${state.report.confirmed?'':'_暫計'}${state.demo?'_合成範例':''}.xlsx`))showExportReceipt();
  }catch(e){clearExportReceipt();message(`匯出失敗：${e.message}`,true);}
});
$('template').addEventListener('click',()=>{try{const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(templateRows($('month').value)),'班表');saveBook(wb,`空白班表範本_${$('month').value}.xlsx`);}catch(e){message(e.message,true);}});
$('sample').addEventListener('click',()=>{
  invalidate();state.previous=null;$('prev-sheet').replaceChildren(node('option','尚未提供'));$('prev-sheet').disabled=true;
  const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(sampleRows()),'合成範例');
  setBook('current',wb,'內建合成範例（非正式班表）');$('regime').value='threeShift';state.demo=true;run();
});
$('summary-view-toggle').addEventListener('click',()=>{
  const full=$('report').classList.toggle('full-table');$('summary-view-toggle').setAttribute('aria-pressed',String(full));$('summary-view-toggle').textContent=full?'員工摘要':'完整表格';
});
initCatEffect($('cat-greeting'),document.querySelector('.rail-panel'));
initPrinting(()=>state.report,()=>Number($('employee').value));
refreshDates();
if(!XLSX){message('本機 SheetJS 程式庫載入失敗，請確認 dist/vendor/xlsx.full.min.js 完整存在。',true);busy(true);}
initLanding({
  importFile:file=>importFile({target:{files:[file],value:''}},'current'),
  loadSample:()=>{$('sample').click();return state.report ? null : $('message').textContent;},
  hasReport:()=>Boolean(state.report)
});
