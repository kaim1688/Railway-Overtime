(() => {
  const key = 'railway-overtime-theme';
  const root = document.documentElement;
  let theme = 'light';
  try {
    if (localStorage.getItem(key) === 'dark') theme = 'dark';
  } catch { /* Theme switching also works when browser storage is unavailable. */ }
  root.dataset.theme = theme;
  document.addEventListener('DOMContentLoaded', () => {
    const button = document.getElementById('theme-toggle');
    const updateLabel = () => { button.textContent = theme === 'dark' ? '切換淺色' : '切換暗色'; };
    updateLabel();
    button.addEventListener('click', () => {
      theme = theme === 'dark' ? 'light' : 'dark';
      root.dataset.theme = theme;
      updateLabel();
      try { localStorage.setItem(key, theme); } catch { /* Keep the current tab usable. */ }
    });
  });
})();
