export function waitForPresentation(ms, signal) {
  if (signal.aborted) return Promise.resolve();
  return new Promise(resolve => {
    const finish = () => {
      clearTimeout(timer);
      signal.removeEventListener('abort', finish);
      resolve();
    };
    const timer = setTimeout(finish, ms);
    signal.addEventListener('abort', finish, {once:true});
  });
}

export function initLanding({importFile, loadSample, hasReport}) {
  const host = document.getElementById('landing');
  const root = host.attachShadow({mode:'open'});
  root.append(document.getElementById('landing-template').content.cloneNode(true));
  const $ = id => root.getElementById(id);
  const page = root.querySelector('.landing-page');
  const workbench = document.getElementById('workbench');
  const appearance = document.documentElement;
  const reduced = matchMedia('(prefers-reduced-motion: reduce)');
  let busy = false;
  let presentation = null;
  let petTimer = null;
  let trainTimer = null;

  function updatePetControl() {
    $('pet-cat').disabled = busy || !$('error').hidden || page.classList.contains('confused') || page.classList.contains('offering');
    $('pet-cat').setAttribute('aria-busy',String(petTimer !== null));
    $('station-sign').disabled = $('pet-cat').disabled;
    $('station-sign').setAttribute('aria-busy',String(trainTimer !== null));
  }
  function stopTrain() {
    if (trainTimer === null) return;
    clearTimeout(trainTimer);trainTimer = null;
    page.classList.remove('train-running');
    if (!busy && $('error').hidden && !page.classList.contains('offering')) {
      $('bubble').textContent = '喵，今天也辛苦了。';
      $('announcement').textContent = '班表僅在本機處理，不會上傳。';
    }
    updatePetControl();
  }
  $('station-sign').addEventListener('click',()=>{
    if ($('station-sign').disabled || trainTimer !== null) return;
    stopPet();
    const motion = appearance.dataset.motion === 'on' && !reduced.matches;
    $('bubble').textContent = motion ? '喵，小火車來了！' : '喵，陪你等車。';
    $('announcement').textContent = motion ? '小火車經過橘貓小站。' : '橘貓陪你等車。';
    page.classList.toggle('train-running',motion);
    trainTimer = setTimeout(stopTrain,2100);
    updatePetControl();
  });
  function stopPet() {
    if (petTimer === null) return;
    clearTimeout(petTimer);
    petTimer = null;
    page.classList.remove('petting');
    if (!busy && $('error').hidden && !page.classList.contains('offering')) {
      $('bubble').textContent = '喵，今天也辛苦了。';
      $('announcement').textContent = '班表僅在本機處理，不會上傳。';
    }
    updatePetControl();
  }
  $('pet-cat').addEventListener('click',()=>{
    if ($('pet-cat').disabled || petTimer !== null) return;
    stopTrain();
    const motion = appearance.dataset.motion === 'on' && !reduced.matches;
    $('bubble').textContent = '喵，休息一下。';
    $('announcement').textContent = motion ? '橘貓伸了個懶腰。喵，休息一下。' : '喵，休息一下。';
    page.classList.toggle('petting',motion);
    petTimer = setTimeout(stopPet,motion ? 1100 : 2000);
    updatePetControl();
  });
  root.addEventListener('keydown',event=>{if(event.key==='Escape'){stopPet();stopTrain();}});

  function syncAppearance() {
    const dark = appearance.dataset.theme === 'dark';
    const motion = appearance.dataset.motion === 'on';
    host.dataset.theme = dark ? 'dark' : 'light';
    $('theme-label').textContent = dark ? '淺色' : '深色';
    $('theme').setAttribute('aria-label', dark ? '切換淺色模式' : '切換深色模式');
    $('theme').setAttribute('aria-pressed', String(dark));
    $('motion').textContent = reduced.matches ? '靜態' : motion ? '動畫開' : '動畫關';
    $('motion').setAttribute('aria-pressed', String(motion));
    $('motion').setAttribute('aria-label', reduced.matches ? '已依系統設定減少動畫' : motion ? '關閉動畫' : '啟用動畫');
    $('motion').disabled = reduced.matches;
    page.classList.toggle('no-motion', !motion);
    if (!motion) { presentation?.abort(); stopPet(); stopTrain(); }
    pauseMotion();
  }
  function pauseMotion() {
    page.classList.toggle('paused', document.hidden || host.hidden);
    if (document.hidden || host.hidden) {stopPet();stopTrain();}
  }
  new MutationObserver(syncAppearance).observe(appearance,{attributes:true,attributeFilter:['data-theme','data-motion']});
  document.addEventListener('visibilitychange',pauseMotion);
  syncAppearance();
  // Enable transitions only after the shadow stylesheet's initial theme has painted.
  const sceneStyles = root.querySelector('link[rel="stylesheet"]');
  const enableTransitions = () => requestAnimationFrame(()=>requestAnimationFrame(()=>page.classList.add('scene-ready')));
  if (sceneStyles.sheet) enableTransitions();
  else sceneStyles.addEventListener('load',enableTransitions,{once:true});
  $('theme').addEventListener('click',()=>document.getElementById('theme-toggle').click());
  $('motion').addEventListener('click',()=>document.getElementById('motion-toggle').click());

  function openWorkbench() {
    host.hidden = true;
    workbench.hidden = false;
    appearance.dataset.screen = 'workbench';
    pauseMotion();
    const target = hasReport() ? document.getElementById('results-title') : document.querySelector('#setup-fold summary');
    if (!hasReport()) document.getElementById('setup-fold').open = true;
    target.setAttribute('tabindex','-1');
    target.focus({preventScroll:true});
    target.scrollIntoView({block:'start'});
  }
  document.getElementById('back-home').addEventListener('click',()=>{
    workbench.hidden = true;
    host.hidden = false;
    appearance.dataset.screen = 'landing';
    pauseMotion();
    $('choose').focus({preventScroll:true});
    window.scrollTo(0,0);
  });
  $('open-workbench').addEventListener('click',openWorkbench);
  $('choose').addEventListener('click',()=>$('file').click());

  async function process(file, sample=false) {
    if (busy || (!file && !sample)) return;
    busy = true;
    stopPet();
    stopTrain();
    clearOffering();
    page.classList.remove('confused');
    presentation = new AbortController();
    const signal = presentation.signal;
    if (appearance.dataset.motion !== 'on') presentation.abort();
    const source = $('dropzone').getBoundingClientRect();
    const destination = root.querySelector('.receiving-tray').getBoundingClientRect();
    const sheet = $('handoff-file');
    for (const [name,value] of Object.entries({
      'from-x':source.left+source.width/2-18, 'from-y':source.top+source.height/2-22,
      'to-x':destination.left+destination.width/2-18, 'to-y':destination.top-22
    })) sheet.style.setProperty('--'+name,value+'px');
    $('ready').hidden = true;
    $('working').hidden = false;
    $('error').hidden = true;
    $('demo').disabled = true;
    $('open-workbench').disabled = true;
    $('stage-title').textContent = signal.aborted ? '橘貓正在整理班表' : '班表交給橘貓了';
    $('filename').textContent = sample ? '合成範例（非正式班表）' : file.name;
    $('bubble').textContent = '收到，交給我吧！';
    $('announcement').textContent = '讀取與計算中，資料留在本機。';
    page.classList.add('busy');
    if (!signal.aborted) page.classList.add('receiving');
    const received = waitForPresentation(800,signal).then(()=>{
      page.classList.remove('receiving');
      if (!signal.aborted) {
        $('stage-title').textContent = '橘貓正在整理班表';
        $('bubble').textContent = '一格一格，仔細整理。';
      }
    });
    const minimumWork = waitForPresentation(2100,signal);
    $('working').setAttribute('tabindex','-1');
    $('working').focus({preventScroll:true});
    try {
      // Give the busy state one paint before synchronous workbook parsing.
      await new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve)));
      const error = sample ? await loadSample() : await importFile(file);
      if (error) throw new Error(error);
      if (hasReport()) {
        await Promise.all([received,minimumWork]);
        if (!signal.aborted) {
          page.classList.add('complete');
          $('stage-title').textContent = '好了，來核對一下！';
          $('bubble').textContent = '好了！';
          $('announcement').textContent = '計算完成，準備顯示結果。';
          await waitForPresentation(900,signal);
        }
      }
      openWorkbench();
    } catch(error) {
      page.classList.add('confused');
      $('bubble').textContent = '這是啥？我看不懂耶。';
      $('error').textContent = error.message || '無法讀取班表，請重新選擇 .xlsx 檔案。';
      $('error').hidden = false;
      $('announcement').textContent = '未完成讀取，請重新選檔。';
    } finally {
      presentation.abort();
      await received;
      presentation = null;
      busy = false;
      $('file').value = '';
      $('ready').hidden = false;
      $('working').hidden = true;
      $('demo').disabled = false;
      $('open-workbench').disabled = false;
      page.classList.remove('busy','receiving','complete');
      updatePetControl();
      if ($('error').hidden) $('bubble').textContent = '喵，今天也辛苦了。';
      if ($('error').hidden) $('announcement').textContent = '班表僅在本機處理，不會上傳。';
      else $('choose').focus({preventScroll:true});
    }
  }
  $('file').addEventListener('change',e=>process(e.target.files[0]));
  $('demo').addEventListener('click',()=>process(null,true));
  let depth = 0;
  const drop = $('dropzone');
  function offerFile(active) {
    if (active) {stopPet();stopTrain();}
    drop.classList.toggle('drag-over', active);
    page.classList.toggle('offering', active);
    updatePetControl();
    if (!busy && $('error').hidden) $('bubble').textContent = active ? '放這裡，我接住！' : '喵，今天也辛苦了。';
  }
  function clearOffering() { depth = 0; offerFile(false); }
  host.addEventListener('dragover',e=>e.preventDefault());
  host.addEventListener('drop',e=>{e.preventDefault();clearOffering();});
  window.addEventListener('dragend',clearOffering);
  window.addEventListener('blur',clearOffering);
  drop.addEventListener('dragenter',e=>{
    e.preventDefault();
    if (busy || !Array.from(e.dataTransfer?.types || []).includes('Files')) return;
    depth++;
    offerFile(true);
  });
  drop.addEventListener('dragleave',()=>{depth=Math.max(0,depth-1);if(!depth)offerFile(false);});
  drop.addEventListener('drop',e=>{
    e.preventDefault();clearOffering();
    if(busy)return;
    if(e.dataTransfer.files.length!==1){
      $('error').textContent='請一次放入一份 .xlsx 班表。';$('error').hidden=false;
      $('bubble').textContent='一次一份，我才接得住。';
      $('announcement').textContent='未完成讀取，請一次放入一份班表。';
      stopPet();stopTrain();updatePetControl();return;
    }
    process(e.dataTransfer.files[0]);
  });
}
