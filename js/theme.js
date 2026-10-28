(() => {
  const key = 'railway-overtime-theme';
  const root = document.documentElement;
  const sizeKey = 'railway-overtime-font-size';
  const sizes = ['12', '13', '14'];
  let fontSize = '14';
  const systemTheme = window.matchMedia('(prefers-color-scheme: dark)');
  let theme = systemTheme.matches ? 'dark' : 'light';
  let chosenTheme = false;
  try {
    const saved = localStorage.getItem(key);
    if (saved === 'dark' || saved === 'light') { theme = saved; chosenTheme = true; }
  } catch { /* Theme switching also works when browser storage is unavailable. */ }
  root.dataset.theme = theme;
  root.dataset.screen = 'landing';
  try {
    const saved = localStorage.getItem(sizeKey);
    if (sizes.includes(saved)) fontSize = saved;
  } catch { /* Appearance controls remain usable without storage. */ }
  root.dataset.fontSize = fontSize;
  const motionKey = 'railway-overtime-motion';
  const reducedMotion = window.matchMedia('(prefers-reduced-motion: reduce)');
  let motion = 'on';
  try { if (localStorage.getItem(motionKey) === 'off') motion = 'off'; } catch { /* Use the current tab preference. */ }
  const applyMotion = () => { root.dataset.motion = reducedMotion.matches ? 'off' : motion; };
  applyMotion();
  document.addEventListener('DOMContentLoaded', () => {
    const motionButton = document.getElementById('motion-toggle');
    const updateMotion = () => {
      applyMotion();
      motionButton.disabled = reducedMotion.matches;
      motionButton.textContent = reducedMotion.matches ? '動畫：關（依系統）' : `動畫：${motion === 'on' ? '開' : '關'}`;
      motionButton.setAttribute('aria-pressed', String(root.dataset.motion === 'on'));
    };
    updateMotion();
    motionButton.addEventListener('click', () => {
      motion = motion === 'on' ? 'off' : 'on';
      updateMotion();
      try { localStorage.setItem(motionKey, motion); } catch { /* Keep the current tab usable. */ }
    });
    reducedMotion.addEventListener('change', updateMotion);
    let railVisible = false;
    const pauseMotion = () => { root.dataset.motionPaused = String(document.hidden || !railVisible); };
    pauseMotion();
    const rail = document.getElementById('rail-art');
    new IntersectionObserver(entries => {
      railVisible = entries[0].isIntersecting;
      pauseMotion();
    }).observe(rail);
    document.addEventListener('visibilitychange', pauseMotion);
    const select = document.getElementById('font-size');
    select.value = fontSize;
    select.addEventListener('change', () => {
      if (!sizes.includes(select.value)) { select.value = fontSize; return; }
      fontSize = select.value;
      root.dataset.fontSize = fontSize;
      try { localStorage.setItem(sizeKey, fontSize); } catch { /* Keep the current tab usable. */ }
    });
    const button = document.getElementById('theme-toggle');
    const updateLabel = () => { button.textContent = theme === 'dark' ? '切換淺色' : '切換暗色'; };
    updateLabel();
    systemTheme.addEventListener('change', () => {
      if (chosenTheme) return;
      theme = systemTheme.matches ? 'dark' : 'light';
      root.dataset.theme = theme;
      updateLabel();
    });
    button.addEventListener('click', () => {
      chosenTheme = true;
      theme = theme === 'dark' ? 'light' : 'dark';
      root.dataset.theme = theme;
      updateLabel();
      try { localStorage.setItem(key, theme); } catch { /* Keep the current tab usable. */ }
    });
  });
})();
