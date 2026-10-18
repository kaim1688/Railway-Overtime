import { datesInMonth, validDate, KEYS, LABELS } from './rules.js';

const text = v => String(v ?? '').trim();
const NAME = ['姓名','員工姓名','員工'];
const ID = ['員工編號','工號','員編'];
const DATE = ['日期','出勤日期'];
const SHIFT = ['班別','班次'];
export function readWorkbook(XLSX, data) {
  if (data.byteLength > 20 * 1024 * 1024) throw new Error('檔案超過 20MB，請先另存所需班表工作表。');
  const bytes = new Uint8Array(data);
  if (bytes[0]!==0x50 || bytes[1]!==0x4b) throw new Error('不是有效的 .xlsx 檔案，請以 Excel 另存為 .xlsx。');
  return XLSX.read(data, {type:'array', cellDates:true, cellFormula:true});
}
export function dateValue(value, month, XLSX, date1904=false) {
  if (value instanceof Date) return value.toISOString().slice(0,10);
  if (typeof value==='number' && value>10000) {
    const p=XLSX.SSF.parse_date_code(value,{date1904});
    if (!p) return null;
    return `${p.y}-${String(p.m).padStart(2,'0')}-${String(p.d).padStart(2,'0')}`;
  }
  const s=text(value);
  let m=s.match(/^(\d{3,4})[年/.-](\d{1,2})[月/.-](\d{1,2})日?$/);
  if (m) {
    const y=Number(m[1])+(m[1].length===3 ? 1911:0);
    return `${y}-${m[2].padStart(2,'0')}-${m[3].padStart(2,'0')}`;
  }
  m=s.match(/^(\d{1,2})[月/.-](\d{1,2})日?$/);
  if (m) return `${month.slice(0,4)}-${m[1].padStart(2,'0')}-${m[2].padStart(2,'0')}`;
  m=s.match(/^(\d{1,2})日?$/);
  return m ? `${month}-${m[1].padStart(2,'0')}` : null;
}
function sheetRows(XLSX, workbook, sheetName) {
  const ws=workbook.Sheets[sheetName];
  if (!ws || !ws['!ref']) throw new Error('工作表不存在或為空白。');
  const range=XLSX.utils.decode_range(ws['!ref']);
  if (range.e.r>10000 || range.e.c>366) throw new Error('工作表超過 10,001 列或 367 欄，請另存班表範圍後再匯入。');
  return {ws,rows:XLSX.utils.sheet_to_json(ws,{header:1,raw:true,defval:null,range:0,blankrows:true})};
}
function findRosterBlocks(rows,ws) {
  const blocks=[];
  for (let r=2;r<rows.length;r++) {
    const nameCol=rows[r-1].findIndex(v=>text(v)==='星期');
    const title=rows[r-2].map(text).join(' ').match(/(\d{3,4})年\s*(\d{1,2})月/);
    if (nameCol<0 || !title || text(rows[r][nameCol])) continue;
    const month=`${Number(title[1])+(title[1].length===3?1911:0)}-${title[2].padStart(2,'0')}`;
    if (!validDate(month+'-01')) continue;
    const columns=rows[r].map((v,c)=>({c,day:text(v)})).filter(x=>/^\d{1,2}$/.test(x.day) && x.c>nameCol);
    if (columns.length<28) continue;
    blocks.push({id:String(r),layout:'wide',r,titleRow:r-2,startRow:r+1,nameCol,idCol:-1,month,
      columns:columns.map(x=>({c:x.c,date:`${month}-${x.day.padStart(2,'0')}`}))});
  }
  return blocks.map((block,i)=>{
    let endRow=(blocks[i+1]?.titleRow ?? rows.length)-1;
    const footer=(ws['!merges'] || []).filter(m=>m.s.r>block.r && m.s.r<=endRow && m.s.c<=block.nameCol && m.e.c>=block.columns[1].c && /\d{1,2}:\d{2}\s*[~～–-]\s*\d{1,2}:\d{2}|班別.*圖例/.test(text(rows[m.s.r][m.s.c]))).sort((a,b)=>a.s.r-b.s.r)[0];
    if(footer)endRow=footer.s.r-1;
    while(endRow>block.r && !rows[endRow].some(v=>text(v)))endRow--;
    return {...block,endRow,label:`第 ${i+1} 區｜${block.month}｜第 ${block.startRow+1}–${endRow+1} 列`,blockCount:blocks.length};
  });
}
export function rosterBlocks(XLSX,workbook,sheetName) {
  const {ws,rows}=sheetRows(XLSX,workbook,sheetName);
  return findRosterBlocks(rows,ws);
}
export function parseSheet(XLSX, workbook, sheetName, month, blockId='') {
  datesInMonth(month);
  const {ws,rows}=sheetRows(XLSX,workbook,sheetName);
  const issues=[], candidates=[];
  const date1904=!!workbook.Workbook?.WBProps?.date1904;
  for (let r=0;r<Math.min(rows.length,30);r++) {
    const labels=rows[r].map(text), nameCol=labels.findIndex(v=>NAME.includes(v));
    if (nameCol<0) continue;
    const idCol=labels.findIndex(v=>ID.includes(v)), dateCol=labels.findIndex(v=>DATE.includes(v)), shiftCol=labels.findIndex(v=>SHIFT.includes(v));
    if (dateCol>=0 && shiftCol>=0) { candidates.push({layout:'long',r,nameCol,idCol,dateCol,shiftCol}); continue; }
    for (let dr=r;dr<=Math.min(r+2,rows.length-1);dr++) {
      const columns=rows[dr].map((v,c)=>({c,date:dateValue(v,month,XLSX,date1904)})).filter(x=>x.c!==nameCol && x.c!==idCol && x.date);
      if (columns.length>=2) { candidates.push({layout:'wide',r:dr,nameCol,idCol,columns}); break; }
    }
  }
  const blocks=candidates.length ? [] : findRosterBlocks(rows,ws);
  if(blocks.length && !blockId)throw new Error('請選擇班表區段，並確認該區段人員適用三班制；不自動合併不同班制。');
  const schema=blocks.length ? blocks.find(b=>b.id===blockId) : candidates[0];
  if (blocks.length && !schema)throw new Error('所選班表區段無效，請重新選擇。');
  if (!blocks.length && candidates.length!==1) throw new Error(candidates.length ? '找到多組姓名表頭，請將每組另存工作表，避免混算。' : '找不到明確表頭。支援姓名表頭，或「年月／星期／完整日期」的分區班表；可先下載範本。');
  if(blocks.length) {
    if(schema.month!==month)throw new Error(`區段標示 ${schema.month}，與選定 ${month} 不符。`);
    const expected=datesInMonth(month);
    if(schema.columns.length!==expected.length || schema.columns.some((col,i)=>col.date!==expected[i] || col.c!==schema.nameCol+i+1))throw new Error('區段日期欄必須完整、依序且連續，請核對缺日／重複欄。');
    if([schema.titleRow,schema.r-1,schema.r].some(r=>rows[r].some((_,c)=>ws[XLSX.utils.encode_cell({r,c})]?.f)))throw new Error('區段表頭包含公式，請另存為值後核對。');
    if(schema.columns.some(col=>text(rows[schema.r-1][col.c])!=='日一二三四五六'[new Date(col.date+'T00:00:00Z').getUTCDay()]))throw new Error('星期列與結算年月日期不符，請核對班表版本。');
  }
  const titleText=[sheetName,...rows.slice(schema.titleRow ?? 0,schema.r).flat().filter(v=>typeof v==='string')].join(' ');
  for (const m of titleText.matchAll(/(\d{3,4})年\s*(\d{1,2})月/g)) {
    const detected=`${Number(m[1])+(m[1].length===3?1911:0)}-${m[2].padStart(2,'0')}`;
    if (detected!==month) throw new Error(`表頭／工作表標示 ${detected}，與選定 ${month} 不符。`);
  }
  const cell=(r,c)=>ws[XLSX.utils.encode_cell({r,c})];
  const ref=(r,c)=>`${sheetName}!${XLSX.utils.encode_cell({r,c})}`;
  function safe(r,c) {
    const v=cell(r,c);
    if (v?.f || v?.t==='e') { issues.push(`${ref(r,c)} 含公式／錯誤值，未採用快取結果，請另存為值後核對。`); return null; }
    const merged=(ws['!merges'] || []).find(m=>r>=m.s.r && r<=m.e.r && c>=m.s.c && c<=m.e.c && (m.s.r!==m.e.r || m.s.c!==m.e.c));
    if (merged && r>schema.r) { issues.push(`${ref(r,c)} 位於合併資料儲存格，請拆分並填妥，不自動填值。`); return null; }
    return rows[r]?.[c];
  }
  if (schema.layout==='wide') {
    const seen=new Set();
    for (const col of schema.columns) {
      if (!validDate(col.date) || !col.date.startsWith(month)) throw new Error(`日期欄 ${col.date} 無效或不屬於 ${month}。請核對月份。`);
      if (seen.has(col.date)) throw new Error(`日期欄 ${col.date} 重複，請先修正。`);
      if (cell(schema.r,col.c)?.f) throw new Error('日期表頭包含公式，請另存為值後核對。');
      seen.add(col.date);
    }
  }
  const people=new Map();
  for (let r=schema.r+1;r<=(schema.endRow ?? rows.length-1);r++) {
    if (!rows[r].some(v=>v!==null && text(v)!=='')) continue;
    const rawName=safe(r,schema.nameCol), name=text(rawName);
    if (!name || typeof rawName!=='string' || NAME.includes(name)) { issues.push(`第 ${r+1} 列姓名缺漏／不明，該列未納入，請核對是否為員工。`); continue; }
    if (/^(合計|總計|備註|說明)/.test(name)) { issues.push(`第 ${r+1} 列「${name}」可能是備註或合計，未納入，請人工核對。`); continue; }
    const id=schema.idCol>=0 ? text(safe(r,schema.idCol)) : '';
    const key=id ? `id:${id}` : `name:${name}`;
    if (!people.has(key)) people.set(key,{name,id,records:[]});
    const p=people.get(key);
    if (p.name!==name) issues.push(`員工編號 ${id} 對應不同姓名，請核對。`);
    if (schema.layout==='wide' && p.records.length) issues.push(`員工「${name}」有重複橫排資料；同名者請補上不同員工編號。`);
    if (schema.layout==='wide') for (const col of schema.columns) p.records.push({date:col.date,symbol:text(safe(r,col.c)),source:ref(r,col.c)});
    else {
      const date=dateValue(safe(r,schema.dateCol),month,XLSX,date1904);
      if (!date || !validDate(date) || !date.startsWith(month)) { issues.push(`${ref(r,schema.dateCol)} 日期不明或非 ${month}，該列未納入。`); continue; }
      p.records.push({date,symbol:text(safe(r,schema.shiftCol)),source:ref(r,schema.shiftCol)});
    }
  }
  if (!people.size) throw new Error('找不到員工資料，請核對工作表及姓名欄。');
  return {employees:[...people.values()],issues,sheetName,month,layout:schema.layout,headerRow:schema.r+1,
    selection:schema.label || '全工作表',blockCount:blocks.length,
    dateCount:schema.columns?.length ?? new Set([...people.values()].flatMap(p=>p.records.map(r=>r.date))).size};
}
export function combineEmployees(current, previous) {
  const key=p=>p.id ? `id:${p.id}` : `name:${p.name}`;
  return current.map(p=> {
    const old=previous.find(q=>key(q)===key(p) && q.name===p.name);
    return {...p,records:[...(old?.records || []),...p.records]};
  });
}
export function reportWorkbook(XLSX, report, metadata={}) {
  const wb=XLSX.utils.book_new();
  const append=(name,rows)=> {
    const ws=XLSX.utils.aoa_to_sheet(rows);
    ws['!cols']=rows[0].map((_,i)=>({wch:i===0 ? 18 : 20}));
    ws['!autofilter']={ref:ws['!ref']};
    XLSX.utils.book_append_sheet(wb,ws,name);
  };
  append('統計總表', [['結算年月','姓名','員工編號','核對狀態',...LABELS,'國定 0–8H 合計','總計費時數'],
    ...report.employees.map(p=>[report.month,p.name,p.id,p.confirmed?'規則內核對通過（非核定）':'暫計／待核對',...KEYS.map(k=>p.totals[k]),p.totals.holiday08Combined,p.totals.total]),
    [report.month,'全體合計','',report.confirmed?'規則內核對通過（非核定）':'暫計／待核對',...KEYS.map(k=>report.totals[k]),report.totals.holiday08Combined,report.totals.total]]);
  append('每日明細', [['姓名','員工編號','日期','原班別','國定本日',...LABELS,'當日合計','計算依據','轉移來源 → 目標','資料位置','異常'],
    ...report.employees.flatMap(p=>p.detail.map(r=>[p.name,p.id,r.date,r.symbol,r.isHoliday?'是':'否',...KEYS.map(k=>r.hours[k]),KEYS.reduce((n,k)=>n+r.hours[k],0),r.note,r.transfer,r.source || '',r.issues.join('；')]))]);
  append('待核對事項',[['姓名','日期','問題'],...report.issues.map(m=>['全體','',m]),...report.employees.flatMap(p=>p.issues.map(i=>[p.name,i.date,i.message]))]);
  append('計算設定',[['項目','內容'],['規則版本',report.ruleVersion],['用途','依提供規則內部試算，非核定薪資；只計時數'],['國定本日',report.holidays.join('、')],['已人工確認月份',report.confirmedMonths.join('、')],...Object.entries(metadata).map(([k,v])=>[k,String(v)])]);
  return wb;
}
