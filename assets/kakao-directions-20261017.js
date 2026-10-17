/* Keep the original map images visible; automatically start one map near the viewport. */
(()=>{'use strict';
 const frames=[...document.querySelectorAll('iframe[data-kakao-src]')];
 const safePath=value=>{try{const u=new URL(value,location.href);return u.origin===location.origin&&/^\/assets\/kakao-maps\/\d+\.htm$/.test(u.pathname)?u.pathname+u.search:null;}catch{return null;}};
 function load(frame){
  const path=safePath(frame.dataset.kakaoSrc);if(!path||frame.dataset.kakaoLoaded===path)return;
  const shell=frame.closest('.kd-map-shell');shell.removeAttribute('data-map-ready');
  shell.querySelector('.kd-map-status').textContent='위치 지도를 불러오는 중입니다.';
  frame.dataset.kakaoLoaded=path;frame.src=path;
  setTimeout(()=>{if(frame.dataset.kakaoLoaded===path&&!shell.hasAttribute('data-map-ready'))shell.querySelector('.kd-map-status').textContent='지도가 늦게 도착하고 있습니다. 기존 위치 안내 이미지와 아래 주소를 함께 확인해 주세요.';},15000);
 }
 window.addEventListener('message',event=>{
  if(event.origin!==location.origin||event.data?.type!=='wawa-kakao-ready')return;
  const id=String(event.data.id);if(!/^\d+$/.test(id))return;
  const f=frames.find(f=>f.contentWindow===event.source&&new URL(f.dataset.kakaoSrc,location.href).pathname==='/assets/kakao-maps/'+id+'.htm');
  if(f)f.closest('.kd-map-shell').dataset.mapReady='true';
 });
 if('IntersectionObserver' in window){
  const observer=new IntersectionObserver(entries=>{for(const e of entries)if(e.isIntersecting){load(e.target);observer.unobserve(e.target);}},{rootMargin:'300px 0px'});
  frames.forEach(f=>observer.observe(f));
 }else frames.forEach(load);
 function trafficList(groups){
  const list=document.createElement('ul');list.className='kd-traffic';
  for(const group of groups){
   const li=document.createElement('li'),h=document.createElement('h3');h.textContent=group.title;li.append(h);
   if(!group.items.length){const p=document.createElement('p');p.className='kd-empty';p.textContent='이 약도에 등록된 정보가 없습니다.';li.append(p);list.append(li);continue;}
   const items=document.createElement(group.className==='kd-stops'?'ol':'ul');items.className=group.className;
   group.items.forEach((value,index)=>{
    const item=document.createElement('li');
    if(group.className==='kd-stops'){const n=document.createElement('span');n.className='kd-stop-number';n.setAttribute('aria-hidden','true');n.textContent=String(index+1);item.append(n);}
    if(typeof value==='string'){const text=document.createElement('span');text.textContent=value;item.append(text);}
    else{const type=document.createElement('span');type.className='kd-bus-type kd-bus-'+value.type;type.textContent=value.label;const nums=document.createElement('span');nums.className='kd-bus-numbers';value.names.forEach(n=>{const text=document.createElement('span');text.textContent=n;nums.append(text);});item.append(type,nums);}
    items.append(item);
   });li.append(items);list.append(li);
  }return list;
 }
 for(const select of document.querySelectorAll('[data-kakao-select]'))select.addEventListener('change',async()=>{
  const section=select.closest('.kakao-directions'),option=select.selectedOptions[0],frame=section.querySelector('iframe');
  const path=safePath(option.value);if(!path)return;
  section.dataset.kakaoCenter=option.textContent;
  section.querySelector('.kd-intro').textContent=option.textContent+'의 위치와 주변 교통편을 확인하세요.';
  section.dataset.selection=path;frame.title=option.textContent+' 위치 지도';frame.dataset.kakaoSrc=path;load(frame);
  section.querySelector('.kd-address dd').textContent=option.dataset.address;
  const link=section.querySelector('.kd-hub-link'),u=new URL(option.dataset.link,location.href);
  if(u.origin===location.origin){link.href=u.pathname+u.hash;link.textContent=option.textContent+' 지점 안내';}
  const holder=section.querySelector('.kd-traffic-holder');holder.textContent='주변 교통정보를 불러오는 중입니다.';
  try{
   const url=new URL(option.dataset.traffic,location.href);
   if(url.origin!==location.origin||!/^\/assets\/kakao-traffic\/\d+\.json$/.test(url.pathname))throw Error('Invalid traffic path');
   const response=await fetch(url);if(!response.ok)throw Error('Traffic unavailable');const groups=await response.json();
   if(section.dataset.selection===path)holder.replaceChildren(trafficList(groups));
  }catch{if(section.dataset.selection===path)holder.textContent='주변 교통정보를 불러오지 못했습니다. 위 방문 주소를 확인해 주세요.';}
 });
})();
