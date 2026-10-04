/* All branch and profile links remain readable without JavaScript. */
(() => {
  const form = document.querySelector('[data-tc-filters]');
  if (!form) return;
  const cards = [...document.querySelectorAll('[data-tc-card]')];
  const controls = Object.fromEntries(['q', 'region', 'approach'].map(key => [key, form.elements.namedItem(key)]));
  const status = document.querySelector('[data-tc-status]');
  const empty = document.querySelector('[data-tc-empty]');
  const normalize = value => value.normalize('NFKC').toLocaleLowerCase('ko-KR').replace(/\s+/g, ' ').trim();
  const sync = () => {
    const params = new URL(location.href).searchParams;
    for (const [key, field] of Object.entries(controls)) {
      const value = params.get(key) || '';
      field.value = key === 'q' ? value.slice(0, 100) : [...field.options].some(option => option.value === value) ? value : '';
    }
  };
  const apply = mode => {
    const tokens = normalize(controls.q.value).split(' ').filter(Boolean);
    let count = 0;
    for (const card of cards) {
      const matches = tokens.every(token => normalize(card.dataset.search).includes(token)) &&
        (!controls.region.value || card.dataset.region === controls.region.value) &&
        (!controls.approach.value || JSON.parse(card.dataset.approaches).includes(controls.approach.value));
      card.hidden = !matches;
      if (matches) count++;
    }
    status.textContent = `전체 ${cards.length}개 지점 중 ${count}개 지점`;
    empty.hidden = count !== 0;
    if (mode) {
      const url = new URL(location.href);
      for (const [key, field] of Object.entries(controls)) field.value.trim() ? url.searchParams.set(key, field.value.trim()) : url.searchParams.delete(key);
      if (url.href !== location.href) history[mode === 'push' ? 'pushState' : 'replaceState'](null, '', url);
    }
  };
  sync(); apply(); form.hidden = false;
  controls.q.addEventListener('input', () => apply('replace'));
  controls.region.addEventListener('change', () => apply('push'));
  controls.approach.addEventListener('change', () => apply('push'));
  form.addEventListener('submit', event => { event.preventDefault(); apply('push'); });
  document.querySelectorAll('[data-tc-reset]').forEach(button => button.addEventListener('click', () => {
    for (const field of Object.values(controls)) field.value = '';
    apply('push'); controls.q.focus();
  }));
  window.addEventListener('popstate', () => { sync(); apply(); });
})();
