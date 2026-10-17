/** Add reviewed branch directions after the existing map, preserving the earlier release passes. */
import fs from 'node:fs';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {createHash} from 'node:crypto';
import {spawnSync} from 'node:child_process';
import {verifyPublicOutput} from './release-public-verify.mjs';
export const kakaoDirectionsContext=root=>JSON.parse(fs.readFileSync(path.join(root,'kakao-directions-review.json'),'utf8'));
export function enhanceKakaoDirectionsHTML(source,name,context){
 const info=context.pages[name];if(!info)return {html:source,changed:false};
 const text=source.replaceAll('\r\n','\n');
 if(text.includes('data-kakao-directions="'+context.version+'"'))return {html:source,changed:false};
 if(createHash('sha256').update(text).digest('hex')!==info.inputSha256)throw Error('Kakao directions input differs from reviewed page: '+name);
 if(text.includes('id="kakao-directions"'))throw Error('Duplicate directions section: '+name);
 const section=context.sections[info.section];if(!section)throw Error('Missing reviewed map section: '+name);
 let html=text.slice(0,info.insertAt)+section+text.slice(info.insertAt);
 const canonical=html.match(/<link\b[^>]*rel=["']canonical["'][^>]*href=["']([^"']+)["']/i)?.[1];
 const same=u=>{try{return !!u&&new URL(u,canonical).origin===new URL(canonical).origin&&new URL(u,canonical).pathname===new URL(canonical).pathname;}catch{return false;}};
 const pageTypes=new Set(['WebPage','Article','BlogPosting','ContactPage','AboutPage','CollectionPage']);
 html=html.replace(/<script\b[^>]*type=["']application\/ld\+json["'][^>]*>([\s\S]*?)<\/script>/gi,(all,json)=>{
  const value=JSON.parse(json);let changed=false;
  function visit(v){if(!v||typeof v!=='object')return;if(Array.isArray(v)){v.forEach(visit);return;}
   const types=Array.isArray(v['@type'])?v['@type']:[v['@type']];
   if(types.some(t=>pageTypes.has(t))&&[v.url,v['@id'],typeof v.mainEntityOfPage==='string'?v.mainEntityOfPage:v.mainEntityOfPage?.['@id']].some(same)&&v.dateModified!==context.date){v.dateModified=context.date;changed=true;}
   Object.values(v).forEach(visit);
  }visit(value);return changed?all.replace(json,JSON.stringify(value)):all;
 });
 html=html.replace(/<meta\b[^>]*(?:property|name)=["']article:modified_time["'][^>]*>/gi,tag=>tag.replace(/\bcontent=["'][^"']*["']/i,'content="'+context.date+'"'));
 html=html.replace(/<\/head\s*>/i,'<link rel="stylesheet" href="'+context.style+'" data-kakao-directions="'+context.version+'"><script defer src="'+context.script+'" data-kakao-directions="'+context.version+'"></script>$&');
 return {html,changed:html!==source};
}
export function enhanceKakaoDirectionsSitemap(source,context){
 return source.replace(/<url>\s*<loc>([^<]+)<\/loc>[\s\S]*?<\/url>/g,(block,loc)=>{
  const name=decodeURIComponent(new URL(loc.replaceAll('&amp;','&')).pathname).replace(/^\//,'')+'index.html';
  if(!context.pages[name])return block;
  return /<lastmod>/.test(block)?block.replace(/<lastmod>[^<]*<\/lastmod>/,'<lastmod>'+context.date+'</lastmod>'):block.replace('</loc>','</loc><lastmod>'+context.date+'</lastmod>');
 });
}
if(process.argv[1]&&path.resolve(process.argv[1])===fileURLToPath(import.meta.url)){
 const root=path.dirname(fileURLToPath(import.meta.url));
 if(!process.argv.includes('--apply-only')){const r=spawnSync(process.execPath,['content-quality-build.mjs'],{cwd:root,stdio:'inherit'});if(r.error)throw r.error;if(r.status!==0)process.exit(r.status??1);}
 const context=kakaoDirectionsContext(root);let changed=0;
 for(const name of Object.keys(context.pages)){
  const file=path.join(root,'.public-release',name),result=enhanceKakaoDirectionsHTML(fs.readFileSync(file,'utf8'),name,context);
  if(result.changed){fs.writeFileSync(file,result.html);changed++;}
 }
 const sitemap=path.join(root,'.public-release','sitemap.xml');fs.writeFileSync(sitemap,enhanceKakaoDirectionsSitemap(fs.readFileSync(sitemap,'utf8'),context));
 console.log(JSON.stringify({kakaoDirectionsPages:changed,centers:Object.keys(context.maps).length,held:context.holds.length}));
 console.log(JSON.stringify(verifyPublicOutput(root,path.join(root,'.public-release'),{representativePhotos:true,finalFavicon:true,contentQuality:true,kakaoDirections:true})));
}
