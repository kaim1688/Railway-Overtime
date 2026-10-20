import {REGIMES} from './rules.js';

export function printModel(report,index=null) {
  if(!report?.employees.length)return null;
  const person=index===null?null:report.employees[index];
  if(index!==null && !person)return null;
  const people=person?[person]:report.employees;
  const warnings=[...report.issues,...people.flatMap(p=>p.issues.map(i=>`${p.name} ${i.date || ''}：${i.message}`))];
  const confirmed=person?person.confirmed&&!report.issues.length:report.confirmed;
  return {
    table:person?'daily':'summary',
    title:person?'台鐵班表・個人逐日計算明細':'台鐵班表・全體加班時數統計表',
    caption:`${report.month}｜${person?`${person.name}${person.id?`（${person.id}）`:''}・${REGIMES[person.regime]}`:`全體 ${people.length} 人`}｜單位：小時（H）`,
    status:confirmed?'依設定規則試算，只計時數。':'暫計／待核對：僅含可計算時數，請先核對下列事項。',
    dataType:report.metadata?.['資料性質'] || '',warnings
  };
}

export function clearPrint() {document.getElementById('print-report').replaceChildren();}

function prepare(report,index) {
  clearPrint();
  const model=printModel(report,index),area=document.getElementById('print-report');
  const text=(tag,value)=>{const e=document.createElement(tag);e.textContent=value;return e;};
  if(!model){area.append(text('p','尚未完成計算，請先匯入班表並完成日期設定。'));return false;}
  const table=document.getElementById(model.table).cloneNode(true);
  table.removeAttribute('id');table.className=`print-${model.table}`;
  table.querySelectorAll('[id]').forEach(e=>e.removeAttribute('id'));
  table.querySelectorAll('button').forEach(e=>e.replaceWith(text('span',e.textContent)));
  area.append(text('h1',model.title),text('p',model.caption),text('p',model.dataType),text('p',model.status),table);
  if(model.warnings.length) {
    const list=document.createElement('ul');
    list.append(...model.warnings.map(w=>text('li',w)));
    area.append(text('h2','待核對事項'),list);
  }
  return true;
}

export function initPrinting(getReport,getEmployeeIndex) {
  for(const [id,index] of [['print-summary',()=>null],['print-employee',getEmployeeIndex]]) {
    document.getElementById(id).addEventListener('click',()=>{
      if(prepare(getReport(),index()))window.print();
    });
  }
  // Browser Ctrl+P defaults to the full summary; never reuse a prior employee printout.
  window.addEventListener('beforeprint',()=>{
    if(!document.getElementById('print-report').hasChildNodes())prepare(getReport(),null);
  });
  window.addEventListener('afterprint',clearPrint);
}
