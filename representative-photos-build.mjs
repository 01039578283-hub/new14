/** Pin verified branch photos at their existing lower-body positions. */
import fs from 'node:fs';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {createHash} from 'node:crypto';
import {spawnSync} from 'node:child_process';
import {parseHTML} from './image-order-build.mjs';
import {verifyPublicOutput} from './release-public-verify.mjs';
const esc=s=>String(s).replaceAll('&','&amp;').replaceAll('"','&quot;').replaceAll('<','&lt;');
const attr=(tag,key,value)=>{const r=new RegExp('\\s'+key+'(?:\\s*=\\s*(?:"[^"]*"|\'[^\']*\'|[^\\s>]+))?','i');return r.test(tag)?tag.replace(r,' '+key+'="'+esc(value)+'"'):tag.replace(/\s*\/?>$/,' '+key+'="'+esc(value)+'">');};
const ancestry=n=>{const a=[];for(let p=n.parent;p;p=p.parent)a.push(p);return a;};
const digest=b=>createHash('sha256').update(b).digest('hex');
const canonicalPath=u=>decodeURIComponent(new URL(u).pathname).replace(/\/$/,'')||'/';
const STYLE='<style data-branch-photo-style="20261009">[data-branch-photo-representative] img{display:block;width:100%;height:auto;object-fit:contain;max-height:none}[data-branch-photo-representative]{min-width:0;max-width:918px;margin:24px auto}[data-branch-photo-representative] figure{margin:0}[data-branch-photo-representative] figcaption{padding:12px 0;color:#46536a;font-size:14px;line-height:1.8}.branch-photo-heading{font-size:1.5rem;margin:0 0 18px}</style>';
export function representativePhotoContext(root){
 const c=JSON.parse(fs.readFileSync(path.join(root,'representative-photos-review.json'),'utf8'));
 for(const p of Object.values(c.photos)){
  if(!p.visuallyReviewed||p.width<=150||p.height<=150||p.bytes<5000||p.width/p.height>3)throw Error('Unreviewed or unsuitable branch photo '+p.center);
  if(digest(fs.readFileSync(path.join(root,p.src.replace(/^\//,''))))!==p.sha256)throw Error('Verified branch photo changed '+p.center);
 }
 return c;
}
// The existing favicon pass has this exact result; no logo or icon is changed here.
export function normalizeFinalFavicon(html){
 if(!/<html\b/i.test(html)||!/<\/head>/i.test(html))return html;
 const at=html.search(/<\/head>/i),head=html.slice(0,at),tail=html.slice(at);
 const clean=head.replace(/<link\b[^>]*>/gi,tag=>{const r=tag.match(/\brel\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s>]+))/i);return /^(?:icon|shortcut icon|apple-touch-icon|apple-touch-icon-precomposed)$/i.test((r?.[1]??r?.[2]??r?.[3]??'').trim())?'':tag;});
 return clean+'<link rel="icon" type="image/png" sizes="96x96" href="/assets/site-favicon.png"><link rel="apple-touch-icon" sizes="180x180" href="/assets/site-apple-touch-icon.png">'+tail;
}
export function enhanceRepresentativePhotoHTML(source,name,context){
 const info=context.pages[name];if(!info)return {html:source,changed:false};
 if(source.includes('data-branch-photo-representative="20261009"'))return {html:source,changed:false};
 const nodes=parseHTML(source),edits=[],main=nodes.find(n=>n.tag==='main');if(!main)throw Error('No main on photo target '+name);
 const localPath='/'+name.replace(/index\.html$/,'');
 const imagePath=n=>decodeURIComponent(new URL(n.attrs.src,context.origin+localPath).pathname);
 const hidden=nodes.filter(n=>n.tag==='img'&&(n.attrs['data-media-role']==='representative'||((n.attrs.hidden!==undefined||n.attrs['aria-hidden']==='true'||/display\s*:\s*none/i.test(n.attrs.style||''))&&/\/(?:representative\/|rep-)/.test(imagePath(n)))));
 if(hidden.length!==info.hiddenCount)throw Error('Hidden representative differs from reviewed target '+name);
 for(const n of hidden)edits.push({start:n.start,end:n.end,value:''});
 const gallery=nodes.find(n=>n.attrs.id==='center-photos');
 const selected=nodes.filter(n=>n.tag==='img'&&ancestry(n).includes(main)&&imagePath(n)===info.src);
 if(selected.length!==(info.existingBottom?1:0))throw Error('Representative photo count differs from review '+name);
 if(info.existingBottom){
  const img=selected[0];if(!gallery||!ancestry(img).includes(gallery))throw Error('Existing photo is not in the reviewed lower gallery '+name);
  let tag=source.slice(img.start,img.end);
  for(const [k,v] of Object.entries({src:info.absolute,alt:info.alt,width:info.width,height:info.height,loading:'lazy',decoding:'async','data-branch-photo-selected':'true'}))tag=attr(tag,k,v);
  edits.push({start:img.start,end:img.end,value:tag});
  const figure=ancestry(img).find(n=>n.tag==='figure');const caption=figure?.children.find(n=>n.tag==='figcaption');
  if(caption)edits.push({start:caption.openEnd,end:caption.closeStart,value:esc(info.caption)});
  edits.push({start:gallery.start,end:gallery.openEnd,value:attr(source.slice(gallery.start,gallery.openEnd),'data-branch-photo-representative','20261009')});
 }else{
  const heading=esc(info.center+' 사진');
  const block='<section class="site-shell kd-section" data-branch-photo-representative="20261009" aria-label="'+heading+'"><h2 class="branch-photo-heading">'+heading+'</h2><figure><a href="'+esc(info.absolute)+'" target="_blank" rel="noopener" aria-label="'+esc(info.alt+' 크게 보기')+'"><img data-branch-photo-selected="true" src="'+esc(info.absolute)+'" alt="'+esc(info.alt)+'" width="'+info.width+'" height="'+info.height+'" loading="lazy" decoding="async"></a><figcaption>'+esc(info.caption)+'</figcaption></figure></section>';
  edits.push({start:main.closeStart,end:main.closeStart,value:block});
 }
 // Unfold only this lower photo gallery. Its images, notes, order and anchors remain.
 if(gallery){
  const gates=nodes.filter(n=>n.tag==='details'&&ancestry(n).includes(gallery));
  for(const n of gates){
   edits.push({start:n.start,end:n.openEnd,value:source.slice(n.start,n.openEnd).replace(/^<details\b/i,'<div').replace(/\sopen(?:="[^"]*")?/i,'')});
   edits.push({start:n.closeStart,end:n.end,value:'</div>'});
   for(const summary of n.children.filter(n=>n.tag==='summary'))edits.push({start:summary.start,end:summary.end,value:''});
  }
 }
 edits.sort((a,b)=>b.start-a.start||b.end-a.end);let last=source.length,html=source;
 for(const e of edits){if(e.end>last)throw Error('Overlapping photo changes '+name);html=html.slice(0,e.start)+e.value+html.slice(e.end);last=e.start;}
 const metas='<meta property="og:image" content="'+esc(info.absolute)+'"><meta property="og:image:secure_url" content="'+esc(info.absolute)+'"><meta property="og:image:type" content="'+(info.format==='JPEG'?'image/jpeg':'image/'+info.format.toLowerCase())+'"><meta property="og:image:width" content="'+info.width+'"><meta property="og:image:height" content="'+info.height+'"><meta property="og:image:alt" content="'+esc(info.alt)+'"><meta name="twitter:image" content="'+esc(info.absolute)+'"><meta name="twitter:image:alt" content="'+esc(info.alt)+'">';
 html=html.replace(/<head\b[^>]*>[\s\S]*?<\/head>/i,head=>head.replace(/<meta\b[^>]*>/gi,tag=>{const k=tag.match(/\b(?:property|name)\s*=\s*["']([^"']+)["']/i)?.[1]||'';return /^(?:og:image(?::[\w_]+)?|twitter:image(?::alt)?)$/i.test(k)?'':tag;}).replace(/<\/head>/i,metas+STYLE+'$&'));
 const canonical=html.match(/<link\b[^>]*rel=["']canonical["'][^>]*href=["']([^"']+)["']/i)?.[1];if(!canonical)throw Error('Photo target lacks canonical '+name);
 html=html.replace(/<script\b[^>]*type=["']application\/ld\+json["'][^>]*>([\s\S]*?)<\/script>/gi,(all,json)=>{
  const value=JSON.parse(json);let changed=false;
  function visit(v){
   if(!v||typeof v!=='object')return;if(Array.isArray(v)){v.forEach(visit);return;}
   const types=Array.isArray(v['@type'])?v['@type']:[v['@type']];
   const same=[v.url,v['@id'],v.mainEntityOfPage?.['@id'],typeof v.mainEntityOfPage==='string'?v.mainEntityOfPage:null].some(u=>{try{return !!u&&canonicalPath(u)===canonicalPath(canonical)}catch{return false}});
   if(same&&types.some(t=>['WebPage','Article','BlogPosting','AboutPage'].includes(t))){
    if(v.primaryImageOfPage!==undefined){v.primaryImageOfPage={'@type':'ImageObject',url:info.absolute,width:info.width,height:info.height,caption:info.alt};changed=true;}
    if(v.image!==undefined){v.image=info.absolute;changed=true;}
   }
   if(types.includes('LocalBusiness')&&v.address&&v.image!==undefined){
    const old=Array.isArray(v.image)?v.image:[v.image];
    const remaining=old.filter(u=>typeof u==='string'&&!/\/(?:representative\/|rep-)/.test(u)&&u!==info.absolute);
    v.image=[info.absolute,...remaining];changed=true;
   }
   Object.values(v).forEach(visit);
  }
  visit(value);return changed?all.replace(json,JSON.stringify(value)):all;
 });
 return {html,changed:html!==source,existingBottom:info.existingBottom,hiddenRemoved:hidden.length};
}
if(process.argv[1]&&path.resolve(process.argv[1])===fileURLToPath(import.meta.url)){
 const root=path.dirname(fileURLToPath(import.meta.url));
 if(!process.argv.includes('--apply-only')){const r=spawnSync(process.execPath,['favicon-build.mjs'],{cwd:root,stdio:'inherit'});if(r.error)throw r.error;if(r.status!==0)process.exit(r.status??1);}
 const context=representativePhotoContext(root);let changed=0,existing=0,added=0,removed=0;
 for(const name of Object.keys(context.pages)){
  const f=path.join(root,'.public-release',name),source=fs.readFileSync(f,'utf8'),result=enhanceRepresentativePhotoHTML(source,name,context);
  if(result.changed){fs.writeFileSync(f,result.html);changed++;existing+=Number(result.existingBottom);added+=Number(!result.existingBottom);removed+=result.hiddenRemoved;}
 }
 console.log(JSON.stringify({photoPages:changed,existingBottom:existing,newBottom:added,hiddenImagesRemoved:removed,verifiedCenters:Object.keys(context.photos).length,heldPages:context.holds.length}));
 console.log(JSON.stringify(verifyPublicOutput(root,path.join(root,'.public-release'),{representativePhotos:true,finalFavicon:true})));
}
