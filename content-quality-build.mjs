/** Reviewed content additions: keep page identity, source prose and all media in place. */
import fs from 'node:fs';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {createHash} from 'node:crypto';
import {spawnSync} from 'node:child_process';
import {verifyPublicOutput} from './release-public-verify.mjs';
const hash=s=>createHash('sha256').update(s).digest('hex');
export const contentQualityContext=root=>JSON.parse(fs.readFileSync(path.join(root,'content-quality-review.json'),'utf8'));
const pageTypes=new Set(['WebPage','Article','BlogPosting','ContactPage','AboutPage','CollectionPage']);
export function enhanceContentQualityHTML(source,name,context){
 const review=context.pages[name];if(!review)return {html:source,changed:false};
 const before=source.replaceAll('\r\n','\n');
 if(before.includes('data-content-quality="'+context.version+'"'))return {html:source,changed:false};
 if(hash(before)!==review.inputSha256)throw Error('Content input differs from reviewed page: '+name);
 let html=before,last=before.length;
 for(const op of review.operations){
  if(op.end>last||before.slice(op.start,op.end)!==op.before)throw Error('Content operation differs from reviewed page: '+name);
  html=html.slice(0,op.start)+op.after+html.slice(op.end);last=op.start;
 }
 if(review.grammar)html=html.replaceAll('과학는','과학은');
 const canonical=html.match(/<link\b[^>]*rel=["']canonical["'][^>]*href=["']([^"']+)["']/i)?.[1];
 if(!canonical)throw Error('Content target has no canonical: '+name);
 const same=u=>{try{return !!u&&new URL(u,canonical).pathname===new URL(canonical).pathname}catch{return false}};
 html=html.replace(/<script\b[^>]*type=["']application\/ld\+json["'][^>]*>([\s\S]*?)<\/script>/gi,(all,json)=>{
  const value=JSON.parse(json);let changed=false;
  function visit(v){
   if(!v||typeof v!=='object')return;if(Array.isArray(v)){v.forEach(visit);return;}
   const types=Array.isArray(v['@type'])?v['@type']:[v['@type']];
   if(types.some(t=>pageTypes.has(t))&&[v.url,v['@id'],typeof v.mainEntityOfPage==='string'?v.mainEntityOfPage:v.mainEntityOfPage?.['@id']].some(same)){
    if(v.dateModified!==context.date){v.dateModified=context.date;changed=true;}
   }
   Object.values(v).forEach(visit);
  }
  visit(value);return changed?all.replace(json,JSON.stringify(value)):all;
 });
 html=html.replace(/<meta\b[^>]*(?:property|name)=["']article:modified_time["'][^>]*>/gi,tag=>tag.replace(/\bcontent=["'][^"']*["']/i,'content="'+context.date+'"'));
 html=html.replace(/<\/head\s*>/i,'<link rel="stylesheet" href="'+context.style+'" data-content-quality="'+context.version+'">$&');
 return {html,changed:html!==source};
}
export function enhanceContentQualitySitemap(source,context){
 return source.replace(/<url>\s*<loc>([^<]+)<\/loc>[\s\S]*?<\/url>/g,(block,loc)=>{
  const name=decodeURIComponent(new URL(loc.replaceAll('&amp;','&')).pathname).replace(/^\//,'')+'index.html';
  if(!context.pages[name])return block;
  return /<lastmod>/.test(block)?block.replace(/<lastmod>[^<]*<\/lastmod>/,'<lastmod>'+context.date+'</lastmod>'):block.replace('</loc>','</loc><lastmod>'+context.date+'</lastmod>');
 });
}
if(process.argv[1]&&path.resolve(process.argv[1])===fileURLToPath(import.meta.url)){
 const root=path.dirname(fileURLToPath(import.meta.url));
 if(!process.argv.includes('--apply-only')){
  const r=spawnSync(process.execPath,['representative-photos-build.mjs'],{cwd:root,stdio:'inherit'});if(r.error)throw r.error;if(r.status!==0)process.exit(r.status??1);
 }
 const context=contentQualityContext(root);let changed=0;
 for(const name of Object.keys(context.pages)){
  const f=path.join(root,'.public-release',name),result=enhanceContentQualityHTML(fs.readFileSync(f,'utf8'),name,context);
  if(result.changed){fs.writeFileSync(f,result.html);changed++;}
 }
 const sitemap=path.join(root,'.public-release','sitemap.xml');fs.writeFileSync(sitemap,enhanceContentQualitySitemap(fs.readFileSync(sitemap,'utf8'),context));
 console.log(JSON.stringify({contentPages:changed,reviewedDate:context.date,existingUrlsAndImagesPreserved:true}));
 console.log(JSON.stringify(verifyPublicOutput(root,path.join(root,'.public-release'),{representativePhotos:true,finalFavicon:true,contentQuality:true})));
}
