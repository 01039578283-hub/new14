(() => {
  document.querySelectorAll('[data-kd-directory]').forEach(root => {
    const query=root.querySelector('[data-kd-query]');
    const region=root.querySelector('[data-kd-region]');
    const subject=root.querySelector('[data-kd-subject]');
    const grade=root.querySelector('[data-kd-grade]');
    const cards=[...root.querySelectorAll('[data-kd-center]')].map(node=>({node,courses:JSON.parse(node.dataset.courses)}));
    const normalize=value=>value.toLocaleLowerCase('ko-KR').replace(/\s+/g,'');
    const apply=()=>{
      const q=normalize(query.value);let count=0;
      cards.forEach(({node,courses})=>{
        const offered=subject.value?(courses[subject.value]||[]):Object.values(courses).flat();
        const matches=(!q||normalize(node.dataset.kdCenter).includes(q))&&(!region.value||node.dataset.region===region.value)&&(!subject.value||offered.length>0)&&(!grade.value||offered.includes(grade.value));
        node.hidden=!matches;if(matches)count++;
      });
      root.querySelector('[data-kd-count]').textContent=count+'개 지점';
      root.querySelector('[data-kd-empty]').hidden=count!==0;
    };
    query.addEventListener('input',apply);
    [region,subject,grade].forEach(field=>field.addEventListener('change',apply));
    root.querySelector('[data-kd-reset]').addEventListener('click',()=>{query.value='';region.value='';subject.value='';grade.value='';apply();query.focus();});
    root.querySelector('[data-kd-controls]').hidden=false;
  });
})();
