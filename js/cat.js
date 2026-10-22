export function initCatGreeting(button,bubble,rail) {
  const messages=['小潘，喵🐾','班表交給我，最後核對交給你。','小潘辛苦了。','今日值勤：陪小潘核對班表。','慢慢核對。','喝杯飲料～'];
  let previous=-1,timer;
  const close=()=>{clearTimeout(timer);bubble.textContent='';delete rail.dataset.greeting;};
  button.addEventListener('click',()=>{
    clearTimeout(timer);
    let index=Math.floor(Math.random()*(messages.length-(previous<0?0:1)));
    if(previous>=0 && index>=previous)index++;
    bubble.textContent=messages[index];previous=index;
    rail.dataset.greeting='true';
    timer=setTimeout(close,5000);
  });
  button.addEventListener('keydown',event=>{if(event.key==='Escape')close();});
}
