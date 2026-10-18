export const RULE_VERSION = '使用者規則草稿 v1.1（2026-10-17）/ MVP 0.1.0';
export const KEYS = ['ordinary', 'rest02', 'rest38', 'rest912', 'holiday08', 'holiday910', 'holiday1112', 'holidayRest'];
export const LABELS = ['普通延長', '休息日 0–2H', '休息日 3–8H', '休息日 9–12H', '國定出勤 0–8H', '國定 9–10H', '國定 11–12H', '國定例／休 0–8H'];
export const SYMBOLS = Object.freeze({ '日':'day', '夜':'night', '$日':'big', '$夜':'big', '*日':'holiday', '*夜':'holiday', '例':'rest', '休':'rest', '公休':'off', '特休':'leave', '家照':'noWork', '公假':'noWork', '身調假':'noWork', '補休':'noWork', '國慶日':'holidayOff' });
export const DEFAULT_HOLIDAYS = { '2026-10': ['2026-10-10', '2026-10-25'] };
const zero = () => Object.fromEntries(KEYS.map(k => [k, 0]));
const round = n => Math.round(n * 10) / 10;
export function hoursFor(type) {
  const h = zero();
  if (type === 'ordinary') h.ordinary = 1.2;
  if (type === 'big') Object.assign(h, { rest02:2, rest38:6, rest912:3 });
  if (type === 'holiday') Object.assign(h, { holiday08:8, holiday910:2, holiday1112:1 });
  if (type === 'holidayRest') h.holidayRest = 8;
  return h;
}
export function sumHours(items) {
  const h = zero();
  for (const item of items) for (const k of KEYS) h[k] = round(h[k] + item[k]);
  return { ...h, holiday08Combined: round(h.holiday08 + h.holidayRest), total: round(KEYS.reduce((n,k)=>n+h[k],0)) };
}
export function validDate(s) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) return false;
  const d = new Date(s + 'T00:00:00Z');
  return Number.isFinite(d.getTime()) && d.toISOString().slice(0,10) === s;
}
export function datesInMonth(month) {
  if (!/^\d{4}-\d{2}$/.test(month) || !validDate(month+'-01')) throw new Error('請選擇有效結算年月。');
  const [y,m] = month.split('-').map(Number);
  return Array.from({length:new Date(Date.UTC(y,m,0)).getUTCDate()},(_,i)=>`${month}-${String(i+1).padStart(2,'0')}`);
}
export function previousMonth(month) {
  datesInMonth(month);
  const d = new Date(month+'-01T00:00:00Z'); d.setUTCMonth(d.getUTCMonth()-1);
  return d.toISOString().slice(0,7);
}
const before = date => new Date(Date.parse(date+'T00:00:00Z')-86400000).toISOString().slice(0,10);
export function calculate({ month, employees, holidays, confirmedMonths = [], inputIssues = [] }) {
  const days = datesInMonth(month), prev = previousMonth(month);
  if (!Array.isArray(holidays) || holidays.some(d=>!validDate(d))) throw new Error('國定日期必須是有效的 YYYY-MM-DD。');
  if (new Set(holidays).size !== holidays.length) throw new Error('國定日期重複，請核對。');
  const holidaySet = new Set(holidays), reviewed = new Set(confirmedMonths);
  const globalIssues = [...inputIssues];
  if (!reviewed.has(month)) globalIssues.push('本月國定本日清單尚未人工確認，全部時數暫停計入。');
  const results = employees.map(employee => {
    const issues = [], byDate = new Map(), transfers = [];
    const warn = (row, message) => {
      row.issues.push(message); issues.push({ date:row.date, message });
    };
    for (const record of employee.records) {
      if (!validDate(record.date) || ![month,prev].includes(record.date.slice(0,7))) {
        issues.push({date:record.date || '',message:'日期無效或不在本月／前月範圍。'}); continue;
      }
      if (byDate.has(record.date)) {
        const existing = byDate.get(record.date); existing.duplicate = true;
        warn(existing,'同一員工同一天有重複資料，不選用其中一筆。'); continue;
      }
      byDate.set(record.date,{...record, symbol:String(record.symbol ?? '').trim(), hours:zero(), issues:[], note:'', transfer:''});
    }
    for (const date of days) if (!byDate.has(date)) byDate.set(date,{date,symbol:'',hours:zero(),issues:[],note:'',transfer:''});
    const rows = [...byDate.values()].sort((a,b)=>a.date.localeCompare(b.date));
    for (const row of rows) {
      row.kind = Object.hasOwn(SYMBOLS,row.symbol) ? SYMBOLS[row.symbol] : undefined;
      const compLeave = /^補休(?:\((\d{1,2})\/(\d{1,2})\)|（(\d{1,2})\/(\d{1,2})）)$/.exec(row.symbol);
      if (compLeave && validDate(`2000-${(compLeave[1] || compLeave[3]).padStart(2,'0')}-${(compLeave[2] || compLeave[4]).padStart(2,'0')}`)) row.kind='noWork';
      row.isHoliday = holidaySet.has(row.date);
      if (row.duplicate) continue;
      if (!row.kind) { warn(row,row.symbol ? `無法辨識班別「${row.symbol}」。` : '班別空白／缺少當日資料。'); continue; }
      if (!reviewed.has(row.date.slice(0,7))) { warn(row,'該月國定本日清單未確認。'); continue; }
      if (row.kind==='noWork') row.note='未出勤：加班 0H（使用者確認）';
      else if (row.kind==='holidayOff') {
        if (row.isHoliday && row.date.endsWith('-10-10')) row.note='國慶日未出勤：加班 0H（使用者確認）';
        else warn(row,'國慶日標記與 10/10 國定本日清單不符，請核對日期／班別。');
      } else if (row.isHoliday) {
        if (row.kind === 'holiday') { row.hours = hoursFor('holiday'); row.note='國定本日出勤 8／2／1H'; }
        else if (row.kind === 'rest') { row.hours = hoursFor('holidayRest'); row.note='國定例／休 8H'; }
        else if (row.kind === 'leave') row.note='國定特休：0H、不扣特休額度';
        else if (row.kind === 'off') row.note='國定公休，待回溯普通日班';
        else warn(row,'國定本日出現普通班／大班，規則未定義，請核對符號。');
      } else {
        if (['day','night'].includes(row.kind)) { row.hours=hoursFor('ordinary'); row.note='普通延長 1.2H'; }
        else if (row.kind==='big') { row.hours=hoursFor('big'); row.note='休息日大班 2／6／3H'; }
        else if (row.kind==='holiday') warn(row,'非國定本日出現星號出勤，請核對日期／班別。');
        else row.note='本版不計加班';
      }
    }
    // A known off-day stays searchable even when its own transfer needs review.
    for (const row of rows) row.inputUncertain = row.issues.length > 0;
    const occupied = new Set();
    for (const source of rows.filter(r=>r.isHoliday && r.kind==='off' && !r.issues.length)) {
      let cursor = before(source.date), target;
      while (true) {
        const candidate = byDate.get(cursor);
        if (!candidate) { warn(source,`轉移資料不足：缺 ${cursor} 班表，請補齊本月／前月資料。`); break; }
        if (candidate.inputUncertain) { warn(source,`回溯 ${cursor} 有資料疑義，無法確認最近普通日班。`); break; }
        if (candidate.kind==='day' && !candidate.isHoliday) {
          if (occupied.has(cursor)) warn(source,`轉移衝突：最近普通日班 ${cursor} 已被占用，不改選更早日。`);
          else target=candidate;
          break;
        }
        cursor=before(cursor);
      }
      if (!target) continue;
      occupied.add(target.date);
      const crossMonth = source.date.slice(0,7)!==target.date.slice(0,7);
      const label = `${source.date} → ${target.date}`;
      source.transfer=label;
      transfers.push({source:source.date,target:target.date,crossMonth});
      if (crossMonth) {
        warn(source,`跨月歸屬待總務確認：候選 ${label}；取消 1.2H、改計 8／2／1H 尚未入帳。`);
        source.note='跨月候選，不計入本月';
      } else {
        target.hours=hoursFor('holiday'); target.transfer=label;
        target.note='國定轉移：取消普通 1.2H，改計 8／2／1H';
        source.note='已轉移至普通日班；來源日不另計';
      }
    }
    const detail = rows.filter(r=>r.date.startsWith(month));
    const totals = sumHours(detail.map(r=>r.hours));
    const confirmed = !issues.length && !globalIssues.length;
    return {name:employee.name, id:employee.id || '', detail, issues, transfers, totals, confirmed};
  });
  if (!results.length) globalIssues.push('沒有可計算的員工。');
  return {month, holidays, confirmedMonths, ruleVersion:RULE_VERSION, employees:results, issues:globalIssues,
    confirmed:results.length>0 && !globalIssues.length && results.every(r=>r.confirmed), totals:sumHours(results.map(r=>r.totals))};
}
