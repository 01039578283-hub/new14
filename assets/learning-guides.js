/* Static guide links work without JavaScript. Record contents stay in this page. */
(() => {
  const form = document.querySelector('[data-lg-filters]');
  if (form) {
    const cards = [...document.querySelectorAll('[data-lg-card]')];
    const groups = [...document.querySelectorAll('[data-lg-category]')];
    const status = document.querySelector('[data-lg-count]');
    const empty = document.querySelector('[data-lg-empty]');
    const controls = Object.fromEntries(['q', 'category', 'stage', 'audience'].map(key => [key, form.elements.namedItem(key)]));
    const normalize = value => value.normalize('NFKC').toLocaleLowerCase('ko-KR').replace(/\s+/g, ' ').trim();
    const sync = () => {
      const params = new URL(location.href).searchParams;
      for (const [key, control] of Object.entries(controls)) {
        const value = params.get(key) || '';
        control.value = key === 'q' ? value.slice(0, 100) : [...control.options].some(option => option.value === value) ? value : '';
      }
    };
    const apply = (mode, hash) => {
      const tokens = normalize(controls.q.value).split(' ').filter(Boolean);
      for (const card of cards) {
        const matches = tokens.every(token => normalize(card.dataset.search).includes(token)) &&
          (!controls.category.value || card.dataset.category === controls.category.value) &&
          (!controls.stage.value || card.dataset.stages.split(' ').includes(controls.stage.value)) &&
          (!controls.audience.value || card.dataset.audiences.split(' ').includes(controls.audience.value));
        card.hidden = !matches;
      }
      const count = cards.filter(card => !card.hidden).length;
      status.textContent = `전체 ${cards.length}개 중 ${count}개의 가이드`;
      empty.hidden = count !== 0;
      for (const group of groups) {
        const visible = [...group.querySelectorAll('[data-lg-card]')].filter(card => !card.hidden).length;
        group.hidden = visible === 0;
        group.querySelector('[data-lg-category-count]').textContent = `${visible}개`;
      }
      if (mode) {
        const url = new URL(location.href);
        if (hash) url.hash = hash;
        for (const [key, control] of Object.entries(controls)) control.value.trim() ? url.searchParams.set(key, control.value.trim()) : url.searchParams.delete(key);
        if (url.href !== location.href) history[mode === 'push' ? 'pushState' : 'replaceState'](null, '', url);
      }
    };
    sync(); apply(); form.hidden = false;
    controls.q.addEventListener('input', () => apply('replace'));
    for (const key of ['category', 'stage', 'audience']) controls[key].addEventListener('change', () => apply('push'));
    form.addEventListener('submit', event => { event.preventDefault(); apply('push'); });
    document.querySelectorAll('[data-lg-reset]').forEach(button => button.addEventListener('click', () => {
      for (const control of Object.values(controls)) control.value = '';
      apply('push'); controls.q.focus();
    }));
    document.querySelectorAll('[data-lg-jump]').forEach(link => link.addEventListener('click', event => {
      event.preventDefault();
      for (const control of Object.values(controls)) control.value = '';
      controls.category.value = link.dataset.lgJump;
      apply('push', link.hash);
      document.querySelector(link.hash).scrollIntoView({block: 'start'});
    }));
    window.addEventListener('popstate', () => { sync(); apply(); });
  }
  const record = document.querySelector('[data-lg-record]');
  if (!record) return;
  const fields = [...record.querySelectorAll('textarea')];
  const status = record.querySelector('[data-lg-record-status]');
  const save = record.querySelector('[data-lg-save]');
  const clear = record.querySelector('[data-lg-clear]');
  const print = record.querySelector('[data-lg-print]');
  const title = document.querySelector('h1').textContent.trim();
  const update = () => {
    const count = fields.filter(field => field.value.trim()).length;
    save.disabled = count === 0; clear.disabled = count === 0;
    status.textContent = `${fields.length}개 항목 중 ${count}개 작성 · 저장하지 않고 페이지를 떠나면 지워집니다.`;
    fields.forEach(field => { record.querySelector(`[data-print-for="${field.id}"]`).textContent = field.value || '(미작성)'; });
  };
  fields.forEach(field => field.addEventListener('input', update));
  save.addEventListener('click', () => {
    const lines = [title, document.querySelector('link[rel="canonical"]').href, '', '작성 기록', ...fields.flatMap(field => ['[' + field.dataset.label + ']', field.value.trim() || '(미작성)', ''])];
    const blob = new Blob(['\uFEFF', lines.join('\r\n')], {type: 'text/plain;charset=utf-8'});
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url; link.download = record.dataset.slug + '-작성기록.txt';
    document.body.append(link); link.click(); link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
    status.textContent = '기록 TXT 저장을 요청했습니다. 다운로드한 파일을 확인하세요.';
  });
  clear.addEventListener('click', () => { fields.forEach(field => {field.value = '';}); update(); fields[0].focus(); });
  print.addEventListener('click', () => { update(); window.print(); });
  let printDetails;
  window.addEventListener('beforeprint', () => {
    update();
    printDetails = [...document.querySelectorAll('.lg-faq details')].map(detail => [detail, detail.open]);
    printDetails.forEach(([detail]) => { detail.open = true; });
  });
  window.addEventListener('afterprint', () => {
    printDetails?.forEach(([detail, open]) => { detail.open = open; });
    printDetails = undefined;
  });
  const discard = () => { fields.forEach(field => { field.value = ''; }); update(); };
  window.addEventListener('pagehide', discard);
  window.addEventListener('pageshow', event => { if (event.persisted) discard(); });
  record.querySelector('[data-lg-record-actions]').hidden = false;
  update();
})();
