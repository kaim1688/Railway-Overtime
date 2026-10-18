import { datesInMonth } from './rules.js';
export function templateRows(month) {
  return [['姓名','員工編號',...datesInMonth(month)],['請填員工姓名','',...datesInMonth(month).map(()=>'')]];
}
export function sampleRows() {
  const days=datesInMonth('2026-10');
  const symbols=days.map(()=>'休');
  const ordinary=[1,2,3,4,5,6,7,8,9,11,12,13,14,15];
  ordinary.forEach((d,i)=>symbols[d-1]=i<4?'日':'夜');
  symbols[17-1]='$日';symbols[20-1]='$夜';symbols[10-1]='例';symbols[25-1]='休';
  return [['合成測試資料，非任何員工正式班表'],['姓名','員工編號',...days],['範例員工','DEMO-001',...symbols]];
}
