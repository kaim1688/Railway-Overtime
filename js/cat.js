export function initCatEffect(button,rail) {
  let timer;
  const close=()=>{clearTimeout(timer);delete rail.dataset.greeting;delete rail.dataset.spark;};
  button.addEventListener('click',()=>{
    close();rail.dataset.greeting='true';rail.dataset.spark='true';
    timer=setTimeout(close,850);
  });
  button.addEventListener('keydown',event=>{if(event.key==='Escape')close();});
}
