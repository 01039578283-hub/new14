/* Static articles and fallback directory links remain usable without JavaScript. */
(() => {
  const form = document.querySelector('[data-ei-filters]');
  if (form) {
    const cards = [...document.querySelectorAll('[data-ei-card]')];
    const educationCount = cards.filter(card => card.dataset.kind === '교육정보').length;
    const guideCount = cards.filter(card => card.dataset.kind === '학습가이드').length;
    const fields = Object.fromEntries(['q', 'kind', 'category'].map(key => [key, form.elements.namedItem(key)]));
    const status = document.querySelector('[data-ei-status]');
    const empty = document.querySelector('[data-ei-empty]');
    const normalize = value => value.normalize('NFKC').toLocaleLowerCase('ko-KR').replace(/\s+/g, ' ').trim();
    const sync = () => {
      const params = new URL(location.href).searchParams;
      for (const [key, field] of Object.entries(fields)) {
        const value = params.get(key) || '';
        field.value = key === 'q' ? value.slice(0, 100) : [...field.options].some(option => option.value === value) ? value : '';
      }
    };
    const apply = mode => {
      const tokens = normalize(fields.q.value).split(' ').filter(Boolean);
      let count = 0;
      for (const card of cards) {
        const match = tokens.every(token => normalize(card.dataset.search).includes(token)) &&
          (!fields.kind.value || card.dataset.kind === fields.kind.value) &&
          (!fields.category.value || card.dataset.category === fields.category.value);
        card.hidden = !match;
        if (match) count++;
      }
      status.textContent = `전체 ${cards.length}편 중 ${count}편 · 교육정보 ${educationCount}편과 학습가이드 ${guideCount}편`;
      empty.hidden = count !== 0;
      if (mode) {
        const url = new URL(location.href);
        for (const [key, field] of Object.entries(fields)) field.value.trim() ? url.searchParams.set(key, field.value.trim()) : url.searchParams.delete(key);
        if (url.href !== location.href) history[mode === 'push' ? 'pushState' : 'replaceState'](null, '', url);
      }
    };
    sync(); apply(); form.hidden = false;
    fields.q.addEventListener('input', () => apply('replace'));
    fields.kind.addEventListener('change', () => apply('push'));
    fields.category.addEventListener('change', () => apply('push'));
    form.addEventListener('submit', event => { event.preventDefault(); apply('push'); });
    document.querySelectorAll('[data-ei-reset]').forEach(button => button.addEventListener('click', () => {
      for (const field of Object.values(fields)) field.value = '';
      apply('push'); fields.q.focus();
    }));
    window.addEventListener('popstate', () => { sync(); apply(); });
  }
  const local = document.querySelector('[data-ei-local]');
  if (local) {
    const select = local.querySelector('#ei-area');
    const center = local.querySelector('#ei-center');
    const areaLink = local.querySelector('[data-ei-area-link]');
    const centerLink = local.querySelector('[data-ei-center-link]');
    const status = local.querySelector('[data-ei-local-status]');
    const defaults = {areaHref: areaLink.getAttribute('href'), centerHref: centerLink.getAttribute('href'), areaText: areaLink.textContent, centerText: centerLink.textContent};
    const apply = write => {
      const option = select.selectedOptions[0];
      const branch = center.selectedOptions[0];
      const selected = !!option?.value;
      const selectedBranch = !!branch?.value;
      areaLink.setAttribute('href', selected ? option.dataset.area : defaults.areaHref);
      centerLink.setAttribute('href', selected ? option.dataset.center : selectedBranch ? branch.dataset.center : defaults.centerHref);
      areaLink.textContent = selected ? `${option.dataset.name} 학습·수강 안내` : defaults.areaText;
      centerLink.textContent = selected ? `${option.dataset.branch} 지점 안내` : selectedBranch ? `${branch.value} 지점 안내` : defaults.centerText;
      status.textContent = selected ? `${option.textContent}의 동네 안내와 ${option.dataset.branch} 정보를 연결했습니다.` : selectedBranch ? `${branch.textContent}의 지점 안내를 연결했습니다. 동네를 선택하면 해당 동네 안내도 볼 수 있습니다.` : '동네나 지점을 고르면 해당 안내로 바로 이동할 수 있습니다.';
      if (write) {
        const url = new URL(location.href);
        selected ? url.searchParams.set('area', option.value) : url.searchParams.delete('area');
        !selected && selectedBranch ? url.searchParams.set('center', branch.value) : url.searchParams.delete('center');
        if (url.href !== location.href) history.pushState(null, '', url);
      }
    };
    const sync = () => {
      const params = new URL(location.href).searchParams;
      const value = params.get('area') || '';
      select.value = [...select.options].some(option => option.value === value) ? value : '';
      const branchValue = select.value ? select.selectedOptions[0].dataset.branch : params.get('center') || '';
      center.value = [...center.options].some(option => option.value === branchValue) ? branchValue : '';
      apply(false);
    };
    sync(); select.addEventListener('change', () => { center.value = select.value ? select.selectedOptions[0].dataset.branch : ''; apply(true); });
    center.addEventListener('change', () => { select.value = ''; apply(true); });
    window.addEventListener('popstate', sync);
    local.querySelector('[data-ei-local-controls]').hidden = false;
  }
})();
