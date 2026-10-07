/** Reviewed local information and progressive image enlargement; source pages stay intact. */
import fs from 'node:fs';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {parseHTML} from './image-order-build.mjs';
const esc=s=>String(s??'').replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('"','&quot;');
const plain=s=>s.replace(/<[^>]+>/g,' ').replace(/\s+/g,' ').trim();
const ancestry=n=>{const result=[];for(let p=n.parent;p;p=p.parent)result.push(p);return result;};
const attr=(tag,key,value)=>{const r=new RegExp('\\s'+key+'(?:\\s*=\\s*(?:"[^"]*"|\'[^\']*\'|[^\\s>]+))?','i');return r.test(tag)?tag.replace(r,' '+key+'="'+esc(value)+'"'):tag.replace(/\s*\/?>$/,' '+key+'="'+esc(value)+'">');};
export const crawlerContext=root=>JSON.parse(fs.readFileSync(path.join(root,'crawler-content-review.json'),'utf8'));
function schoolPreparation(info){
 const area=info.area||info.center;
 let task=info.subjects.length===1&&info.subjects[0]==='수학'?'수학은 첫 식을 세운 이유와 막힌 풀이 줄을 남기고, 같은 문제를 해설 없이 다시 푼 결과를 준비해 주세요.':info.subjects.length===1&&info.subjects[0]==='영어'?'영어는 모르는 단어와 문장 구조 때문에 막힌 부분을 따로 표시하고, 학교 교과서의 단원명·본문·평가 범위를 준비해 주세요.':'희망 과목마다 최근 답안 한 장에서 스스로 설명할 수 있는 부분과 도움이 필요한 부분을 표시해 주세요.';
 const stage=info.stage==='초'?'초등학생은 학교 알림과 사용 교재의 단원명을 함께 적고, 혼자 읽거나 풀 때 어디서 멈추는지 확인하세요.':info.stage==='중'?'중학생은 시험 범위 안내와 수행평가 기준을 가져오고, 서술형 답안에서 생략한 근거를 표시하세요.':info.stage==='고'?'고등학생은 학년만 적지 말고 학교에서 이수하는 정확한 과목명과 선택과목, 시험 범위를 알려 주세요.':'학교의 평가 범위 안내와 현재 교재의 단원명을 대조하고, 재학 학교·학년·이수 과목을 적어 주세요.';
 return '<div class="local-preparation"><h3>'+esc(area)+' 학생의 학교 자료를 상담에 쓰는 방법</h3><ol><li>'+esc(stage)+'</li><li>'+esc(task)+'</li><li>'+esc(info.center)+'에서 학교 진도와 보완할 범위를 어떻게 함께 계획할지 질문하세요. '+esc(area)+'에서 출발하는 하교 시각과 실제 이동 시간을 적어 희망 요일과 대조하면 좋습니다.</li></ol><p>학교별 반 개설이나 현재 모집 여부는 학교 이름만으로 판단할 수 없습니다. 위 자료를 기준으로 실제 수강 조건을 확인해 주세요.</p></div>';
}
export function enhanceCrawlerHTML(source,name,context){
 const info=context.pages[name];if(!info)return {html:source,changed:false};
 if(source.includes('data-readable-media="20261007"'))return {html:source,changed:false};
 let html=source;let nodes=parseHTML(html);let edits=[];
 for(const image of nodes.filter(n=>n.tag==='img'&&['body','map'].includes(n.attrs['data-media-role']))){
  const role=image.attrs['data-media-role'];let url;
  try{url=decodeURIComponent(new URL(image.attrs.src,'https://xn--3e0bz50b1zcyxat54c.com/'+name).pathname).replace(/^\//,'');}catch{throw Error('Invalid media URL '+name);}
  const derivative=role==='map'?context.maps[url]:undefined;
  let tag=html.slice(image.start,image.end);let full='/'+url;
  if(derivative){tag=attr(attr(attr(tag,'src',derivative.src),'width',derivative.width),'height',derivative.height);full=derivative.src;}
  tag=attr(tag,'decoding','async');tag=attr(tag,'loading','lazy');
  if(ancestry(image).some(n=>n.tag==='a'))throw Error('Existing linked media needs separate review '+name);
  const label=image.attrs.alt||info.center+(role==='map'?' 지도':' 본문');
  edits.push({start:image.start,end:image.end,value:'<a class="readable-media" href="'+esc(full)+'" data-image-zoom aria-label="'+esc(label+' 확대 보기')+'" title="이미지를 누르면 원래 크기로 확대할 수 있습니다">'+tag+'</a>'});
 }
 const school=nodes.find(n=>n.tag==='section'&&n.attrs.id==='schools');if(school)edits.push({start:school.closeStart,end:school.closeStart,value:schoolPreparation(info)});
 const teacher=nodes.find(n=>n.tag==='section'&&n.attrs.id==='teacher-introductions');
 if(teacher&&info.teachers.length){
  const paragraph=teacher.children.find(n=>n.tag==='p');const value='<p>'+esc(info.center)+' 선생님 소개에서 학생의 풀이와 학습 기록을 살피는 방법을 비교해 보세요. 담당 학년·과목과 실제 상담 가능 일정은 지점에서 확인해 주세요.</p><ul class="local-teachers">'+info.teachers.map(t=>'<li><a href="'+esc(t.path)+'">'+esc(t.name)+' 선생님 · '+esc(t.approaches.join(' / '))+'</a></li>').join('')+'</ul>';
  if(paragraph)edits.push({start:paragraph.start,end:paragraph.end,value});
 }
 const curriculum=nodes.find(n=>n.tag==='section'&&n.attrs.id==='learning-curriculum');
 if(curriculum&&info.curriculum.length){
  const value='<div class="local-curriculum"><h3>안내 학년의 공부 내용을 먼저 확인하세요</h3><p>'+esc(info.center)+' 수강 안내표에 기재된 학년·과목과 연결되는 학습 자료입니다. 학교 진도와 현재 수준을 대조하는 참고용이며 지점의 확정 진도표는 아닙니다.</p><ul>'+info.curriculum.map(x=>'<li><a href="'+esc(x.path)+'">'+esc(x.label)+'</a></li>').join('')+'</ul></div>';
  edits.push({start:curriculum.closeStart,end:curriculum.closeStart,value});
 }
 // Related preparation material previously occupied four consecutive top-level sections.
 // Keep every paragraph/link/anchor and make their common purpose one section.
 if(info.kind==='branch'){
  const ids=['learning','error-review','study-priorities','feedback'];const group=ids.map(id=>nodes.find(n=>n.tag==='section'&&n.attrs.id===id));
  if(group.every(Boolean)){
   for(const section of group){
    edits.push({start:section.start,end:section.openEnd,value:html.slice(section.start,section.openEnd).replace(/^<section\b/,'<div')});
    edits.push({start:section.closeStart,end:section.end,value:'</div>'});
    const heading=nodes.find(n=>n.tag==='h2'&&ancestry(n).includes(section));
    if(heading){edits.push({start:heading.start,end:heading.openEnd,value:html.slice(heading.start,heading.openEnd).replace(/^<h2\b/,'<h3')});edits.push({start:heading.closeStart,end:heading.end,value:'</h3>'});}
   }
   edits.push({start:group[0].start,end:group[0].start,value:'<section class="local-consultation-group" aria-labelledby="local-consultation-title"><h2 id="local-consultation-title">'+esc(info.center)+' 상담 자료·오답·피드백을 함께 준비하세요</h2>'});
   edits.push({start:group.at(-1).end,end:group.at(-1).end,value:'</section>'});
  }
 }
 const sequence=nodes.find(n=>n.attrs['data-image-order']==='sequence-v1');
 if(sequence)edits.push({start:sequence.closeStart,end:sequence.closeStart,value:'<p class="image-readability-note">본문과 지도는 바로 볼 수 있습니다. 글씨가 작으면 이미지를 눌러 원래 크기로 확대하고 좌우·아래로 이동해 읽어 주세요. 지도 이미지와 현재 방문 주소를 함께 확인하세요.</p>'});
 edits.sort((a,b)=>b.start-a.start||b.end-a.end);let last=html.length;
 for(const edit of edits){if(edit.end>last)throw Error('Overlapping reviewed changes '+name);html=html.slice(0,edit.start)+edit.value+html.slice(edit.end);last=edit.start;}
 html=html.replace(/<script\b[^>]*type=["']application\/ld\+json["'][^>]*>([\s\S]*?)<\/script>/gi,(all,json)=>{
  const value=JSON.parse(json);const visit=v=>{if(!v||typeof v!=='object')return;if(Array.isArray(v)){v.forEach(visit);return;}if(v.dateModified)v.dateModified=context.date;Object.values(v).forEach(visit);};visit(value);return all.replace(json,JSON.stringify(value));
 });
 html=html.replace(/<meta\b[^>]*(?:property|name)=["']article:modified_time["'][^>]*>/gi,tag=>attr(tag,'content',context.date));
 html=html.replace(/<\/head\s*>/i,'<link rel="stylesheet" href="'+context.style+'" data-readable-media="20261007"><script defer src="'+context.script+'"></script>$&');
 return {html,changed:html!==source};
}
if(process.argv[1]&&path.resolve(process.argv[1])===fileURLToPath(import.meta.url)){
 const root=path.dirname(fileURLToPath(import.meta.url));const context=crawlerContext(root);let changed=0;
 for(const name of Object.keys(context.pages)){const file=path.join(root,'.public-release',name);const r=enhanceCrawlerHTML(fs.readFileSync(file,'utf8'),name,context);if(r.changed){fs.writeFileSync(file,r.html);changed++;}}
 console.log(JSON.stringify({localPages:Object.keys(context.pages).length,changed,originalBodyPixelsAndAssetsRetained:true}));
}
