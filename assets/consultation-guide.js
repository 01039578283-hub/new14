/* Preserve the chosen public center/course context; no request is submitted here. */
(() => {
  'use strict';
  const root = document.getElementById('consultation-guide');
  if (!root) return;
  const center = document.getElementById('consult-center');
  const area = document.getElementById('consult-area');
  const school = document.getElementById('consult-school');
  const schoolNote = document.getElementById('consult-school-note');
  const grade = document.getElementById('consult-grade');
  const subjects = [...root.querySelectorAll('input[name="consult-subject"]')];
  const memo = document.getElementById('consult-memo');
  const availability = document.getElementById('consult-availability');
  const branch = document.getElementById('consult-branch');
  const copy = document.getElementById('consult-copy');
  const feedback = document.getElementById('consult-copy-status');
  const transfer = document.getElementById('consult-next-step');
  const message = document.getElementById('consult-message');
  const visit = document.getElementById('consult-visit');
  const visitTitle = document.getElementById('consult-visit-title');
  const visitName = document.getElementById('consult-registration');
  const visitAddress = document.getElementById('consult-address');
  const visitTipRow = document.getElementById('consult-visit-tip-row');
  const visitTip = document.getElementById('consult-visit-tip');
  const visitTipNote = document.getElementById('consult-visit-tip-note');
  const visitWarning = document.getElementById('consult-visit-warning');
  const visitMap = document.getElementById('consult-map');
  const stages = { '초': '초등', '중': '중등', '고': '고등' };
  const gradePlaceholder = grade.options[0].textContent;
  const validGrades = new Set([...grade.options].map(option => option.value));
  const validCenters = new Set([...center.options].map(option => option.value));
  let stage = '';
  let transferRevision = 0;
  let copyBusy = false;

  function fitMemo() {
    // Measure wrapped text again so both longer and shorter memos fit the box.
    memo.style.height = '0px';
    memo.style.height = (memo.scrollHeight + memo.offsetHeight - memo.clientHeight) + 'px';
  }
  memo.style.resize = 'none';
  memo.style.overflowY = 'hidden';

  function updateAreas(requested = '') {
    area.replaceChildren(new Option('동네 선택', ''));
    const option = center.selectedOptions[0];
    const names = JSON.parse(option?.dataset.areas || '[]');
    names.forEach(name => area.add(new Option(name, name)));
    if (names.includes(requested)) area.value = requested;
  }

  function updateSchools(requested = '') {
    const option = center.selectedOptions[0];
    const byArea = JSON.parse(option?.dataset.areaSchools || '{}');
    const names = area.value ? (byArea[area.value] || []) : JSON.parse(option?.dataset.schools || '[]');
    school.replaceChildren(new Option('학교 선택 안 함', ''));
    names.forEach(name => school.add(new Option(name, name)));
    if (names.includes(requested)) school.value = requested;
    school.disabled = !names.length;
    schoolNote.textContent = !center.value ? '지점을 선택하면 자료에 있는 학교를 살펴볼 수 있습니다.'
      : !names.length ? '선택한 지점·동네의 학교명 자료는 확인이 필요합니다. 상담할 때 학교명을 별도로 알려 주세요.'
      : '학교명은 상담 참고 정보입니다. 학교 제휴·통학 지원을 뜻하지 않으며 학생 학년은 별도로 선택해 주세요.';
  }

  function render() {
    transferRevision += 1;
    if (message) message.hidden = true;
    grade.options[0].textContent = stage ? stages[stage] + ' · 학년 선택' : gradePlaceholder;
    feedback.textContent = '';
    availability.replaceChildren();
    const selected = subjects.filter(input => input.checked).map(input => input.value);
    const option = center.selectedOptions[0];
    const courses = JSON.parse(option?.dataset.courses || '{}');
    if (visit) {
      visit.hidden = !center.value;
      visitTitle.textContent = center.value ? center.value + ' 방문 정보' : '';
      visitName.textContent = option?.dataset.registeredName || '';
      visitAddress.textContent = option?.dataset.address || '';
      visitTip.textContent = option?.dataset.visitTip || '';
      visitTipRow.hidden = !visitTip.textContent;
      visitTipNote.hidden = !visitTip.textContent;
      visitWarning.textContent = option?.dataset.visitWarning || '';
      visitWarning.hidden = !visitWarning.textContent;
      if (center.value) {
        visitMap.setAttribute('href', option.dataset.mapLink);
        visitMap.setAttribute('aria-label', center.value + ' 주소를 네이버 지도에서 확인');
      } else {
        visitMap.removeAttribute('href');
        visitMap.removeAttribute('aria-label');
      }
    }
    branch.hidden = !center.value;
    if (center.value) branch.setAttribute('href', option.dataset.path + '#courses');
    else branch.removeAttribute('href');
    copy.disabled = !center.value || !selected.length || copyBusy;
    if (transfer) transfer.hidden = !center.value || !selected.length;
    if (!center.value) {
      memo.value = '희망 지점과 과목을 선택하면 상담용 메모가 만들어집니다.';
      fitMemo();
      return;
    }
    const notice = document.createElement('p');
    notice.className = 'sw-note';
    notice.textContent = grade.value && selected.length
      ? '아래는 수강 학년표 기준입니다. 선택과목·수강 제한·현재 모집 상태는 지점 안내와 상담에서 함께 확인해 주세요.'
      : '학년과 과목을 선택하면 센터 자료의 안내 범위와 비교할 수 있습니다.';
    availability.append(notice);
    if (grade.value && selected.length) {
      const list = document.createElement('ul');
      for (const subject of selected) {
        const item = document.createElement('li');
        const listed = (courses[subject] || []).includes(grade.value);
        item.textContent = subject + ' · ' + grade.value + ': ' + (listed
          ? '수강 학년표에 기재됨 · 세부 조건과 시간표 확인 필요'
          : '자료에서 해당 학년 확인이 필요함 · 개설 여부 문의');
        item.dataset.listed = String(listed);
        list.append(item);
      }
      availability.append(list);
    }
    const lines = [center.value + ' 수강 상담을 희망합니다.'];
    if (area.value) lines.push('참고한 지역: ' + area.value);
    if (school.value) lines.push('상담 참고 학교: ' + school.value);
    lines.push('학생 학년: ' + (grade.value || (stage ? stages[stage] + ' · 세부 학년 상담 시 전달' : '상담 시 전달')));
    lines.push('희망 과목: ' + (selected.join('·') || '선택 예정'));
    lines.push('확인할 내용: 현재 개설 여부, 선택과목과 수강 조건, 가능한 시간표, 교습비·교재비, 방문 예약');
    if (grade.value && selected.some(subject => !(courses[subject] || []).includes(grade.value))) {
      lines.push('참고: 학년표에 없는 과정이 있어 개설 여부부터 확인하고 싶습니다.');
    }
    memo.value = lines.join('\n');
    fitMemo();
  }

  function writeURL(mode) {
    const url = new URL(window.location.href);
    const values = {
      center: center.value,
      area: area.value,
      school: school.value,
      grade: grade.value,
      subjects: subjects.filter(input => input.checked).map(input => input.value).join(','),
      stage
    };
    for (const [key, value] of Object.entries(values)) {
      if (value) url.searchParams.set(key, value);
      else url.searchParams.delete(key);
    }
    const href = url.pathname + url.search + url.hash;
    if (href === window.location.pathname + window.location.search + window.location.hash) return;
    window.history[mode === 'push' ? 'pushState' : 'replaceState'](window.history.state, '', href);
  }

  function restore() {
    const params = new URLSearchParams(window.location.search);
    const requestedCenter = params.get('center') || '';
    center.value = validCenters.has(requestedCenter) ? requestedCenter : '';
    updateAreas(params.get('area') || '');
    updateSchools(params.get('school') || '');
    const requestedGrade = params.get('grade') || '';
    grade.value = validGrades.has(requestedGrade) ? requestedGrade : '';
    stage = grade.value ? grade.value.slice(0, 1)
      : (Object.hasOwn(stages, params.get('stage')) ? params.get('stage') : '');
    const selected = new Set((params.get('subjects') || '').split(','));
    subjects.forEach(input => { input.checked = selected.has(input.value); });
    render();
    writeURL('replace');
  }

  function changed() {
    render();
    writeURL('push');
  }
  center.addEventListener('change', () => { updateAreas(); updateSchools(); changed(); });
  area.addEventListener('change', () => { updateSchools(school.value); changed(); });
  school.addEventListener('change', changed);
  grade.addEventListener('change', () => { stage = grade.value.slice(0, 1); changed(); });
  subjects.forEach(input => input.addEventListener('change', changed));
  window.addEventListener('popstate', restore);
  window.addEventListener('pageshow', restore);
  window.addEventListener('resize', fitMemo);
  document.fonts?.ready.then(fitMemo);
  restore();
  copy.addEventListener('click', async () => {
    if (copy.disabled) return;
    const text = memo.value;
    const revision = transferRevision;
    copyBusy = true;
    copy.disabled = true;
    if (message) message.hidden = true;
    feedback.textContent = '메모를 복사하는 중입니다.';
    let copied = false;
    try {
      await navigator.clipboard.writeText(text);
      copied = true;
    } catch {
      // Manual selection remains available when automatic copying is unavailable.
    }
    copyBusy = false;
    copy.disabled = !center.value || !subjects.some(input => input.checked);
    if (revision !== transferRevision) {
      feedback.textContent = '선택 조건이 바뀌었습니다. 현재 메모를 다시 확인하고 복사해 주세요.';
      return;
    }
    if (copied) {
      feedback.textContent = '상담용 메모를 복사했습니다. 문자 문의 화면을 열어 입력란에 붙여넣고 내용을 확인해 주세요.';
    } else {
      memo.focus();
      memo.select();
      feedback.textContent = '자동으로 복사하지 못했습니다. 선택된 메모를 직접 복사한 뒤 문자 문의 화면을 열어 붙여넣어 주세요.';
    }
    if (message) message.hidden = false;
  });
})();
