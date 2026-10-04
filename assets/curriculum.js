/* Static curriculum links remain available when JavaScript is disabled. */
(() => {
 'use strict';
 const form = document.querySelector('[data-cu-filters]');
 if (!form) return;
 const cards = [...document.querySelectorAll('[data-cu-card]')];
 const fields = Object.fromEntries(['q','stage','grade','subject'].map(k => [k, form.elements.namedItem(k)]));
 const status = document.querySelector('[data-cu-status]');
 const empty = document.querySelector('[data-cu-empty]');
 const norm = s => (s || '').normalize('NFKC').toLocaleLowerCase('ko').replace(/\s+/g, ' ').trim();
 const allowed = (field, value) => !value || [...field.options].some(o => o.value === value);
 function reconcile() {
  for (const option of fields.grade.options) {
   const valid = !option.value || !fields.stage.value || option.dataset.stage === fields.stage.value;
   option.disabled = !valid;
   option.hidden = !valid;
  }
  if (fields.grade.selectedOptions[0]?.disabled) fields.grade.value = '';
  if (fields.stage.value === '초등' && fields.subject.value === '역사') fields.subject.value = '';
  const history = [...fields.subject.options].find(o => o.value === '역사');
  if (history) history.disabled = fields.stage.value === '초등';
 }
 function render() {
  reconcile();
  const tokens = norm(fields.q.value).split(' ').filter(Boolean);
  let count = 0;
  for (const card of cards) {
   const visible = ['stage','grade','subject'].every(k => !fields[k].value || card.dataset[k] === fields[k].value) && tokens.every(t => norm(card.dataset.search).includes(t));
   card.hidden = !visible;
   if (visible) count++;
  }
  status.textContent = `조건에 맞는 학년·과목 ${count}개 / 전체 ${cards.length}개`;
  empty.hidden = count > 0;
 }
 function writeURL() {
  const url = new URL(location.href);
  for (const [k, field] of Object.entries(fields)) {
   const value = field.value.trim();
   if (value) url.searchParams.set(k, value); else url.searchParams.delete(k);
  }
  if (url.href !== location.href) history.pushState(null, '', url.href);
 }
 function readURL() {
  const params = new URLSearchParams(location.search);
  for (const [k, field] of Object.entries(fields)) {
   const value = params.get(k) || '';
   field.value = k === 'q' ? value.slice(0, 100) : allowed(field, value) ? value : '';
  }
  render();
 }
 form.addEventListener('submit', event => {event.preventDefault(); render(); writeURL();});
 for (const [k,field] of Object.entries(fields)) if (k !== 'q') field.addEventListener('change', () => {render();writeURL();});
 document.querySelectorAll('[data-cu-reset]').forEach(button => button.addEventListener('click', () => {
  form.reset(); render(); writeURL(); fields.q.focus();
 }));
 window.addEventListener('popstate', readURL);
 readURL(); form.hidden = false;
})();
