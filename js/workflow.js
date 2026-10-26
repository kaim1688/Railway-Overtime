import {calculate,datesInMonth,previousMonth,validDate,DEFAULT_HOLIDAYS} from './rules.js?v=b3fbe0184271';
import {parseRoster,combineEmployees,rosterBlocks} from './excel.js?v=b3fbe0184271';
import {CALENDAR_VERSION,CALENDAR_SOURCES} from './holidays.js?v=b3fbe0184271';

export function defaultCalendars() {
  return Object.fromEntries(Object.entries(DEFAULT_HOLIDAYS).map(([month,dates])=>[month,{dates:[...dates],
    source:`官方資料核對後內建；僅一般國定本日，不含補假。${CALENDAR_VERSION}。來源：${CALENDAR_SOURCES.law}；${CALENDAR_SOURCES[month.slice(0,4)]}`} ]));
}
export function parseHolidayInput(value,month) {
  datesInMonth(month);
  const dates=value.split(/[\s,，、;；]+/).filter(Boolean);
  if(!dates.length)throw new Error(`${month} 日期尚未設定；若確實沒有國定日，請按「本月無國定日」。`);
  if(dates.some(d=>!validDate(d)||!d.startsWith(month)))throw new Error(`${month} 國定日期須為該月份有效的 YYYY-MM-DD。`);
  if(new Set(dates).size!==dates.length)throw new Error('國定日期重複，請修正。');
  return dates.sort();
}
export function calculateImportedRoster(XLSX,{current,previous,month,regime,calendars}) {
  datesInMonth(month);
  if(!current?.sheetName)throw new Error('檔案有多張工作表，請選擇本月班表。');
  if(!rosterBlocks(XLSX,current.wb,current.sheetName).length && !regime)throw new Error('此為姓名表頭版型，請指定適用班制。');
  const parsed=parseRoster(XLSX,current.wb,current.sheetName,month,'all',regime);
  if(!Object.hasOwn(calendars,month))throw new Error(`${month} 國定日期尚未設定，請展開「查看／修改日期」補上清單，或明確指定本月無國定日。`);
  const prevMonth=previousMonth(month);
  let prior=null;
  if(previous && parsed.employees.some(p=>p.regime==='threeShift')) {
    if(!previous.sheetName)throw new Error('請選擇前月工作表。');
    const blocks=rosterBlocks(XLSX,previous.wb,previous.sheetName);
    prior=parseRoster(XLSX,previous.wb,previous.sheetName,prevMonth,blocks[0]?.id || '','threeShift');
  }
  const hasPriorDates=Object.hasOwn(calendars,prevMonth);
  const report=calculate({month,employees:combineEmployees(parsed.employees,prior?.employees || []),
    holidays:[...calendars[month].dates,...(prior && hasPriorDates?calendars[prevMonth].dates:[])],
    confirmedMonths:[month,...(prior && hasPriorDates?[prevMonth]:[])],
    inputIssues:[...parsed.issues,...(prior?.issues || []).map(x=>`前月：${x}`)]});
  report.excluded=parsed.excluded;
  return {report,parsed,prior};
}
