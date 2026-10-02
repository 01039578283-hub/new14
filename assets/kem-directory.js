(() => {
  document.querySelectorAll('[data-kd-directory]').forEach(root => {
    const query=root.querySelector('[data-kd-query]');
    const region=root.querySelector('[data-kd-region]');
    const subject=root.querySelector('[data-kd-subject]');
    const grade=root.querySelector('[data-kd-grade]');
    const empty=root.querySelector('[data-kd-empty]');
    const recoveryFields=[[query,'q','검색어'],[region,'region','지역'],[subject,'subject','과목'],[grade,'grade','학년']];
    const updateRecovery=()=>{
      const summary=empty.querySelector('[data-sw-empty-conditions]');if(!summary)return;
      const active=[];
      recoveryFields.forEach(([field,key,label])=>{
        const value=field.value.trim();const text=field.tagName==='SELECT'?field.selectedOptions[0].textContent:value;
        if(value)active.push(label+': '+text);
        const button=empty.querySelector('[data-sw-clear-filter="'+key+'"]');
        button.hidden=!value;button.setAttribute('aria-label',label+' '+text+' 조건 해제');
      });
      summary.textContent=active.length?active.join(' · '):'추가 검색 조건 없음';
      empty.querySelector('[data-sw-clear-all]').hidden=!active.length;
    };
    const cards=[...root.querySelectorAll('[data-kd-center]')].map(node=>{const schoolMatch=node.querySelector('[data-sw-search-match]');return {node,courses:JSON.parse(node.dataset.courses),comparison:node.querySelector('[data-sw-comparison]'),schoolMatch,schools:JSON.parse(schoolMatch?.dataset.swSearchSchools||'[]')};});
    const consultLinks=[...document.querySelectorAll('[data-sw-consult-filter], [data-sw-consult-card]')];
    const updateConsultation=()=>{
      consultLinks.forEach(link=>{
        const url=new URL(link.href,window.location.href);
        for(const key of ['subjects','grade','stage','school'])url.searchParams.delete(key);
        if(subject.value)url.searchParams.set('subjects',subject.value);
        if(grade.value){url.searchParams.set('grade',grade.value);url.searchParams.set('stage',grade.value.slice(0,1));}
        const school=link.closest('[data-kd-center]')?.querySelector('[data-sw-school-choice]')?.value;
        if(school)url.searchParams.set('school',school);
        link.setAttribute('href',url.pathname+url.search+'#consultation-guide');
      });
    };
    const stages={초:'초등',중:'중등',고:'고등'};
    const updateFees=panel=>{
      const fee=panel.querySelector('[data-sw-fee-comparison]');if(!fee)return;
      const scope=JSON.parse(fee.dataset.swFeeScope);
      // Common reference tables never become a confirmed branch fee.
      if(scope.mode==='regional-reference')return;
      let text='자료에 '+scope.total+'개 과정이 기재되어 있습니다. 적용 과목·학년·납부 기간은 지점 상세에서 확인해 주세요.';
      if(subject.value||grade.value){
        const requested=subject.value?[subject.value]:['국어','영어','수학','과학','사회'];
        const count=scope.individual.filter(row=>requested.some(s=>(row[s]||[]).some(g=>!grade.value||g===grade.value))).length;
        text=count?'선택한 과목·학년과 관련된 개별 과목 자료 '+count+'건이 있습니다. 과정별 적용 범위와 추가 비용은 지점 상세에서 확인해 주세요.':'선택 조건의 개별 과목 금액은 바로 확인하기 어렵습니다. 묶음 과정·별도 프로그램·적용 조건을 지점 상세에서 확인해 주세요.';
      }
      fee.querySelector('[data-sw-fee-summary]').textContent=text;
    };
    const updateComparison=panel=>{
      const scope=JSON.parse(panel.dataset.swWeekend);
      const stage=grade.value.slice(0,1);
      const missing=subject.value&&scope.subjects.length&&!scope.subjects.includes(subject.value);
      const mismatch=stage&&scope.stages.length&&!scope.stages.includes(stage);
      let status='주말 과목·학년 확인 필요';const warnings=[];
      if(scope.closed)status='주말 정규 수업 미운영 안내';
      else {
        if(missing)warnings.push('자료에 명시된 주말 과목은 '+scope.subjects.join('·')+'입니다. '+subject.value+'의 주말 개설 여부는 따로 확인해 주세요.');
        if(mismatch)warnings.push('자료에 명시된 주말 학년은 '+scope.stages.map(s=>stages[s]).join('·')+'입니다. '+stages[stage]+' 주말 개설 여부는 따로 확인해 주세요.');
        if(!scope.subjects.length)warnings.push('주말 과목별 구분이 명확하지 않아 희망 과목·학년의 수업 여부를 확인해야 합니다.');
        if(scope.special)status='주말 보강·특강 등 안내';
        else if(missing)status=subject.value+' 주말 개설 확인 필요';
        else if(mismatch)status=stages[stage]+' 주말 개설 확인 필요';
        else if(scope.subjects.length)status='주말 과목 안내: '+scope.subjects.join('·');
      }
      panel.querySelector('[data-sw-weekend-status]').textContent=status;
      const warning=panel.querySelector('[data-sw-weekend-warning]');warning.textContent=warnings.join(' ');warning.hidden=!warnings.length;
      const token=(subject.value||'*')+':'+(stage||'*');let total=0;
      panel.querySelectorAll('[data-sw-condition-scopes]').forEach(row=>{row.hidden=!row.dataset.swConditionScopes.split(' ').includes(token);if(!row.hidden)total++;});
      panel.querySelector('[data-sw-condition-count]').textContent=String(total);
      panel.querySelector('[data-sw-condition-empty]').hidden=total!==0;
      updateFees(panel);
    };
    const normalize=value=>value.toLocaleLowerCase('ko-KR').replace(/\s+/g,'');
    const updateLocalGuides=node=>{
      const panel=node.querySelector('[data-sw-local-guides]');if(!panel)return;
      const stage=grade.value.slice(0,1);let total=0;
      panel.querySelectorAll('[data-sw-local-topic]').forEach(row=>{
        const subjects=JSON.parse(row.dataset.swLocalSubjects);
        row.hidden=Boolean((subject.value&&!subjects.includes(subject.value))||(stage&&row.dataset.swLocalStage&&row.dataset.swLocalStage!==stage));
        if(!row.hidden)total++;
      });
      panel.querySelectorAll('[data-sw-local-group]').forEach(group=>{group.hidden=![...group.querySelectorAll('[data-sw-local-topic]')].some(row=>!row.hidden);});
      panel.querySelectorAll('[data-sw-local-area]').forEach(area=>{area.hidden=![...area.querySelectorAll('[data-sw-local-group]')].some(group=>!group.hidden);});
      panel.querySelector('[data-sw-local-count]').textContent=String(total);
      panel.querySelector('[data-sw-local-empty]').hidden=total!==0;
    };
    const apply=()=>{
      const q=normalize(query.value);let count=0;
      cards.forEach(({node,courses,comparison,schoolMatch,schools})=>{
        const offered=subject.value?(courses[subject.value]||[]):Object.values(courses).flat();
        const matches=(!q||normalize(node.dataset.kdCenter).includes(q))&&(!region.value||node.dataset.region===region.value)&&(!subject.value||offered.length>0)&&(!grade.value||offered.includes(grade.value));
        node.hidden=!matches;if(matches)count++;
        if(schoolMatch){
          const matched=matches&&q?schools.filter(name=>normalize(name).includes(q)):[];
          schoolMatch.hidden=!matched.length;
          schoolMatch.querySelector('[data-sw-search-match-label]').textContent=matched.length?'검색과 일치한 학교: '+matched.join(' · '):'';
          const choice=schoolMatch.querySelector('[data-sw-school-choice]');
          const selected=choice.value;
          choice.replaceChildren(new Option('학교를 선택해 주세요',''));
          matched.forEach(name=>choice.add(new Option(name,name)));
          if(matched.includes(selected))choice.value=selected;
        }
        if(comparison)updateComparison(comparison);
        updateLocalGuides(node);
      });
      root.querySelector('[data-kd-count]').textContent=count+'개 지점';
      empty.hidden=count!==0;
      updateRecovery();
      updateConsultation();
    };
    const fields=[[region,'region'],[subject,'subject'],[grade,'grade']];
    const restore=()=>{
      const parameters=new URLSearchParams(window.location.search);
      query.value=parameters.get('q')||'';
      fields.forEach(([field,key])=>{
        const value=parameters.get(key)||'';
        field.value=[...field.options].some(option=>option.value===value)?value:'';
      });
      apply();
    };
    const writeURL=mode=>{
      const url=new URL(window.location.href);
      for(const [key,value] of [['q',query.value],...fields.map(([field,key])=>[key,field.value])]) {
        if(value.trim())url.searchParams.set(key,value);else url.searchParams.delete(key);
      }
      const href=url.pathname+url.search+url.hash;
      if(href===window.location.pathname+window.location.search+window.location.hash)return;
      window.history[mode==='push'?'pushState':'replaceState'](window.history.state,'',href);
    };
    // Keep one history entry while typing; discrete filter changes remain reversible.
    query.addEventListener('input',()=>{apply();writeURL('replace');});
    cards.forEach(({schoolMatch})=>schoolMatch?.querySelector('[data-sw-school-choice]').addEventListener('change',updateConsultation));
    [region,subject,grade].forEach(field=>field.addEventListener('change',()=>{apply();writeURL('push');}));
    const clearFilters=key=>{
      recoveryFields.forEach(([field,name])=>{if(key==='all'||key===name)field.value='';});
      apply();writeURL('push');
      (key==='all'?query:recoveryFields.find(([,name])=>name===key)[0]).focus();
    };
    root.querySelector('[data-kd-reset]').addEventListener('click',()=>clearFilters('all'));
    empty.querySelectorAll('[data-sw-clear-filter]').forEach(button=>button.addEventListener('click',()=>clearFilters(button.dataset.swClearFilter)));
    empty.querySelector('[data-sw-clear-all]')?.addEventListener('click',()=>clearFilters('all'));
    window.addEventListener('popstate',restore);
    window.addEventListener('pageshow',restore);
    root.querySelector('[data-kd-controls]').hidden=false;
    restore();
    writeURL('replace');
  });
})();
