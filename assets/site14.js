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
    const regions = [...directory.querySelectorAll('.directory-region')];
    const count = document.querySelector('[data-directory-count]');
    const empty = document.querySelector('[data-directory-empty]');
    const filters = [...document.querySelectorAll('[data-region-filter]')];
    let activeRegion = 'all';

    const normalize = (value) => value.toLocaleLowerCase('ko-KR').replace(/\s+/g, '');
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
            if (matches) districtCount += 1;
          });
          district.hidden = districtCount === 0;
          regionCount += districtCount;
        });
        region.hidden = regionCount === 0;
        if (query && regionCount) region.open = true;
        visibleCount += regionCount;
      });
      if (count) count.textContent = query || activeRegion !== 'all' ? `${visibleCount}개 지역 검색됨` : `전체 ${cards.length}개 지역`;
      if (empty) empty.hidden = visibleCount !== 0;
    };

    search.addEventListener('input', applyFilter);
    filters.forEach((filter) => filter.addEventListener('click', () => {
      activeRegion = filter.dataset.regionFilter || 'all';
      filters.forEach((item) => item.classList.toggle('is-active', item === filter));
      applyFilter();
      const firstVisible = regions.find((region) => !region.hidden);
      if (firstVisible && activeRegion !== 'all') firstVisible.open = true;
    }));
    document.querySelector('[data-expand-all]')?.addEventListener('click', () => regions.filter((region) => !region.hidden).forEach((region) => { region.open = true; }));
    document.querySelector('[data-collapse-all]')?.addEventListener('click', () => regions.forEach((region) => { region.open = false; }));
  }
})();
