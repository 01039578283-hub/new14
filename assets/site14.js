(() => {
  const current = document.body.dataset.page;
  document.querySelectorAll('[data-nav]').forEach((link) => {
    if (link.dataset.nav === current) link.setAttribute('aria-current', 'page');
  });

  document.querySelectorAll('a[target="_blank"]').forEach((link) => {
    const rel = new Set((link.getAttribute('rel') || '').split(/\s+/).filter(Boolean));
    rel.add('noopener');
    link.setAttribute('rel', [...rel].join(' '));
  });

  const directory = document.querySelector('.directory-list');
  const search = document.querySelector('[data-local-search]');
  if (directory && search) {
    const cards = [...directory.querySelectorAll('.directory-card')];
    const schoolMatches = new Map(cards.map(card => {
      const panel = card.querySelector('[data-sw-search-match]');
      return [card, { panel, schools: JSON.parse(panel?.dataset.swSearchSchools || '[]') }];
    }));
    const regions = [...directory.querySelectorAll('.directory-region')];
    const count = document.querySelector('[data-directory-count]');
    const empty = document.querySelector('[data-directory-empty]');
    const filters = [...document.querySelectorAll('[data-region-filter]')];
    let activeRegion = 'all';

    const updateRecovery = () => {
      const summary = empty?.querySelector('[data-sw-empty-conditions]');
      if (!summary) return;
      const active = [];
      if (search.value.trim()) active.push('검색어: ' + search.value.trim());
      if (activeRegion !== 'all') active.push('지역: ' + activeRegion);
      summary.textContent = active.length ? active.join(' · ') : '추가 검색 조건 없음';
      empty.querySelector('[data-sw-clear-filter="q"]').hidden = !search.value.trim();
      empty.querySelector('[data-sw-clear-filter="q"]').setAttribute('aria-label', '검색어 ' + search.value.trim() + ' 조건 해제');
      empty.querySelector('[data-sw-clear-filter="region"]').hidden = activeRegion === 'all';
      empty.querySelector('[data-sw-clear-filter="region"]').setAttribute('aria-label', '지역 ' + activeRegion + ' 조건 해제');
      empty.querySelector('[data-sw-clear-all]').hidden = !active.length;
    };

    const normalize = (value) => value.toLocaleLowerCase('ko-KR').replace(/\s+/g, '');
    const updateSchoolLink = (card) => {
      const link = card.querySelector('[data-sw-consult-card]');
      if (!link) return;
      const url = new URL(link.href, window.location.href);
      const school = card.querySelector('[data-sw-school-choice]')?.value;
      if (school) url.searchParams.set('school', school); else url.searchParams.delete('school');
      link.setAttribute('href', url.pathname + url.search + url.hash);
    };
    cards.forEach(card => schoolMatches.get(card).panel?.querySelector('[data-sw-school-choice]').addEventListener('change', () => updateSchoolLink(card)));
    const applyFilter = () => {
      const query = normalize(search.value);
      let visibleCount = 0;
      regions.forEach((region) => {
        const regionMatch = activeRegion === 'all' || region.dataset.region === activeRegion;
        let regionCount = 0;
        region.querySelectorAll('.directory-district').forEach((district) => {
          let districtCount = 0;
          district.querySelectorAll('.directory-card').forEach((card) => {
            const matches = regionMatch && (!query || normalize(card.dataset.locality || card.textContent).includes(query));
            card.hidden = !matches;
            const { panel, schools } = schoolMatches.get(card);
            if (panel) {
              const matched = matches && query ? schools.filter(name => normalize(name).includes(query)) : [];
              panel.hidden = !matched.length;
              panel.querySelector('[data-sw-search-match-label]').textContent = matched.length ? '검색과 일치한 학교: ' + matched.join(' · ') : '';
              const choice = panel.querySelector('[data-sw-school-choice]');
              const selected = choice.value;
              choice.replaceChildren(new Option('학교를 선택해 주세요', ''));
              matched.forEach(name => choice.add(new Option(name, name)));
              if (matched.includes(selected)) choice.value = selected;
              updateSchoolLink(card);
            }
            if (matches) districtCount += 1;
          });
          district.hidden = districtCount === 0;
          const districtLabel = district.querySelector('.directory-district-head span');
          if (districtLabel) districtLabel.textContent = `${districtCount}개 지역`;
          regionCount += districtCount;
        });
        region.hidden = regionCount === 0;
        const regionLabel = region.querySelector('summary small');
        if (regionLabel) regionLabel.textContent = `${regionCount}개 지역`;
        if (query && regionCount) region.open = true;
        visibleCount += regionCount;
      });
      if (count) count.textContent = query || activeRegion !== 'all' ? `${visibleCount}개 지역 검색됨` : `전체 ${cards.length}개 지역`;
      if (empty) empty.hidden = visibleCount !== 0;
      updateRecovery();
    };

    const setRegion = (value) => {
      activeRegion = filters.some(item => item.dataset.regionFilter === value) ? value : 'all';
      filters.forEach((item) => {
        const active = item.dataset.regionFilter === activeRegion;
        item.classList.toggle('is-active', active);
        item.setAttribute('aria-pressed', String(active));
      });
    };
    const openSelectedRegion = () => {
      const firstVisible = regions.find((region) => !region.hidden);
      if (firstVisible && activeRegion !== 'all') firstVisible.open = true;
    };
    const restore = () => {
      const parameters = new URLSearchParams(window.location.search);
      search.value = parameters.get('q') || '';
      setRegion(parameters.get('region') || 'all');
      applyFilter();
      openSelectedRegion();
    };
    const writeURL = (mode) => {
      const url = new URL(window.location.href);
      if (search.value.trim()) url.searchParams.set('q', search.value); else url.searchParams.delete('q');
      if (activeRegion !== 'all') url.searchParams.set('region', activeRegion); else url.searchParams.delete('region');
      const href = url.pathname + url.search + url.hash;
      if (href === window.location.pathname + window.location.search + window.location.hash) return;
      window.history[mode === 'push' ? 'pushState' : 'replaceState'](window.history.state, '', href);
    };
    search.addEventListener('input', () => { applyFilter(); writeURL('replace'); });
    filters.forEach((filter) => filter.addEventListener('click', () => {
      setRegion(filter.dataset.regionFilter || 'all');
      applyFilter();
      openSelectedRegion();
      writeURL('push');
    }));
    document.querySelector('[data-expand-all]')?.addEventListener('click', () => regions.filter((region) => !region.hidden).forEach((region) => { region.open = true; }));
    document.querySelector('[data-collapse-all]')?.addEventListener('click', () => regions.forEach((region) => { region.open = false; }));
    const clearFilters = (key) => {
      if (key === 'q' || key === 'all') search.value = '';
      if (key === 'region' || key === 'all') setRegion('all');
      applyFilter();
      openSelectedRegion();
      writeURL('push');
      if (key === 'region') filters.find(item => item.dataset.regionFilter === 'all')?.focus();
      else search.focus();
    };
    document.querySelector('[data-reset-directory]')?.addEventListener('click', () => clearFilters('all'));
    empty?.querySelectorAll('[data-sw-clear-filter]').forEach(button => button.addEventListener('click', () => clearFilters(button.dataset.swClearFilter)));
    empty?.querySelector('[data-sw-clear-all]')?.addEventListener('click', () => clearFilters('all'));
    window.addEventListener('popstate', restore);
    window.addEventListener('pageshow', restore);
    restore();
    writeURL('replace');
  }
})();
