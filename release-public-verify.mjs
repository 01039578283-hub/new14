/** Verify reviewed build inputs and the exact public result after all HTML passes. */
import fs from 'node:fs';
import path from 'node:path';
import {createHash} from 'node:crypto';
import {fileURLToPath} from 'node:url';

export const BUILD_INPUTS=['release-public-build.mjs','wawa-analytics-build.mjs','seo-descriptions.mjs','release-public-verify.mjs','seo-descriptions.json','vercel.json'];
const digest=bytes=>createHash('sha256').update(bytes).digest('hex');
const normalize=bytes=>Buffer.from(bytes.toString('utf8').replaceAll('\r\n','\n'));
const readJSON=file=>JSON.parse(fs.readFileSync(file,'utf8').replace(/^\uFEFF/,''));
// Deployment providers may reserialize JSON. Preserve every config value and
// array order while ignoring only whitespace and object-key ordering.
const orderedJSON=value=>Array.isArray(value)?value.map(orderedJSON):value&&typeof value==='object'?Object.fromEntries(Object.keys(value).sort().map(key=>[key,orderedJSON(value[key])])):value;
export const configDigest=bytes=>digest(Buffer.from(JSON.stringify(orderedJSON(JSON.parse(bytes.toString('utf8').replace(/^\uFEFF/,''))))));
const siteOrigin='https://xn--3e0bz50b1zcyxat54c.com';
const tracker='<script defer src="https://wawa-visit-collector.clean-peach-8202.chatgpt.site/tracker.js" data-site="wawa-15" crossorigin="anonymous" referrerpolicy="no-referrer"></script>';

export function verifyBuildInputs(root,manifest=readJSON(path.join(root,'release-public-manifest.json'))) {
  const reviewed=manifest.reviewedBuildInputs;
  if(!reviewed||Object.keys(reviewed).sort().join('\n')!==[...BUILD_INPUTS].sort().join('\n'))throw Error('Reviewed build inputs are missing or incomplete; regenerate and review the release manifest');
  for(const name of BUILD_INPUTS) {
    const bytes=fs.readFileSync(path.join(root,name));
    if(!Buffer.from(bytes.toString('utf8')).equals(bytes))throw Error('Non-UTF8 reviewed build input: '+name);
    if(digest(normalize(bytes))===reviewed[name])continue;
    if(name==='vercel.json'&&manifest.reviewedVercelConfigSha256&&configDigest(bytes)===manifest.reviewedVercelConfigSha256)continue;
    throw Error('Reviewed build input changed: '+name);
  }
  return BUILD_INPUTS.length;
}

const decode=s=>s.replace(/&(?:amp|quot|apos|lt|gt|#39|#x[\da-f]+|#\d+);/gi,v=>{
  const named={'&amp;':'&','&quot;':'"','&apos;':"'",'&#39;':"'",'&lt;':'<','&gt;':'>'};
  if(named[v])return named[v];
  return String.fromCodePoint(v.toLowerCase().startsWith('&#x')?parseInt(v.slice(3,-1),16):parseInt(v.slice(2,-1),10));
});
const attrs=tag=>Object.fromEntries([...tag.matchAll(/([\w:-]+)\s*=\s*(["'])([\s\S]*?)\2/g)].map(m=>[m[1].toLowerCase(),decode(m[3])]));
const pageTypes=new Set(['WebPage','CollectionPage','Article','BlogPosting','AboutPage','ContactPage']);
function samePage(value,canonical) {
  if(typeof value!=='string')return false;
  try {const a=new URL(value,canonical),b=new URL(canonical);return a.origin===b.origin&&a.pathname.replace(/\/$/,'')===b.pathname.replace(/\/$/,'');}catch{return false;}
}
function pageNodes(value,canonical,result=[]) {
  if(!value||typeof value!=='object')return result;
  if(Array.isArray(value)){value.forEach(v=>pageNodes(v,canonical,result));return result;}
  const types=Array.isArray(value['@type'])?value['@type']:[value['@type']];
  const ids=[value.url,value['@id'],typeof value.mainEntityOfPage==='string'?value.mainEntityOfPage:value.mainEntityOfPage?.['@id']];
  if(types.some(t=>pageTypes.has(t))&&ids.some(id=>samePage(id,canonical)))result.push(value);
  Object.values(value).forEach(v=>pageNodes(v,canonical,result));return result;
}

export function verifyPublicOutput(root,output=path.join(root,'.public-release')) {
  const manifest=readJSON(path.join(root,'release-public-manifest.json'));
  const buildInputsChecked=verifyBuildInputs(root,manifest);
  const selected=new Set(Object.keys(manifest.files)),actual=[];
  function walk(dir) {
    for(const entry of fs.readdirSync(dir,{withFileTypes:true})) {
      if(entry.isSymbolicLink())throw Error('Unexpected output symlink');
      const file=path.join(dir,entry.name);
      if(entry.isDirectory())walk(file);else actual.push(path.relative(output,file).replaceAll('\\','/'));
    }
  }
  walk(output);
  if(actual.length!==selected.size||actual.some(n=>!selected.has(n)))throw Error('Public output inventory differs from the reviewed manifest');
  const protectedInputs=new Set([...BUILD_INPUTS,'release-public-manifest.json']);
  if([...selected].some(n=>protectedInputs.has(n)||n.startsWith('tools/')||n.startsWith('.')))throw Error('Build or authoring input must not be a public asset');
  let htmlFilesChecked=0;
  for(const name of selected) {
    const source=fs.readFileSync(path.join(root,name)),built=fs.readFileSync(path.join(output,name));
    let expected=source;
    if(name.endsWith('.html')) {
      const html=source.toString('utf8');
      if(!Buffer.from(html).equals(source))throw Error('Non-UTF8 reviewed HTML: '+name);
      if(/<html[\s>]/i.test(html)&&/<\/head\s*>/i.test(html)) {
        const tags=html.match(/<script\b[^>]*wawa-visit-collector[^>]*>[\s\S]*?<\/script>/gi)||[];
        if(tags.length>1||(tags.length===1&&tags[0]!==tracker))throw Error('Unexpected source tracker: '+name);
        if(!tags.length)expected=Buffer.from(html.replace(/<\/head\s*>/i,m=>tracker+m));
      }
      htmlFilesChecked++;
    }
    if(!expected.equals(built))throw Error('Unreviewed final output change: '+name);
  }
  const config=readJSON(path.join(root,'seo-descriptions.json'));
  const map=fs.readFileSync(path.join(output,'sitemap.xml'),'utf8');
  const urls=[...map.matchAll(/<loc>([^<]+)<\/loc>/g)].map(m=>new URL(decode(m[1])));
  if(urls.length!==manifest.sitemapPages||new Set(urls.map(u=>u.href)).size!==urls.length)throw Error('Reviewed sitemap coverage mismatch');
  const titles=new Set(),descriptions=new Set();let schemaPagesChecked=0;
  for(const url of urls) {
    if(url.origin!==siteOrigin||!url.pathname.endsWith('/'))throw Error('Unexpected canonical URL');
    const name=decodeURIComponent(url.pathname).replace(/^\//,'')+'index.html';
    if(!selected.has(name))throw Error('Sitemap page is missing: '+name);
    const html=fs.readFileSync(path.join(output,name),'utf8'),head=html.match(/<head\b[^>]*>[\s\S]*?<\/head>/i)?.[0];
    if(!head)throw Error('Missing head: '+name);
    const canonicals=[...head.matchAll(/<link\b[^>]*>/gi)].map(m=>attrs(m[0])).filter(a=>a.rel==='canonical');
    if(canonicals.length!==1||new URL(canonicals[0].href).href!==url.href)throw Error('Canonical mismatch: '+name);
    const titleMatches=[...head.matchAll(/<title\b[^>]*>([\s\S]*?)<\/title>/gi)];
    const title=titleMatches.length===1?decode(titleMatches[0][1]).trim():'';
    if(!title||titles.has(title))throw Error('Missing or duplicate title: '+name);titles.add(title);
    const key=decodeURIComponent(url.pathname).replace(/\/$/,'')||'/';
    const description=config.pages[key]?.description;
    if(!description||[...description].length>80||!/[.!?]$/.test(description)||descriptions.has(description))throw Error('Invalid or duplicate reviewed description: '+name);descriptions.add(description);
    const tags=[...head.matchAll(/<meta\b[^>]*>/gi)].map(m=>attrs(m[0]));
    for(const field of ['description','og:description','twitter:description']) {
      const found=tags.filter(a=>(a.name||a.property||'').toLowerCase()===field);
      if(found.length!==1||found[0].content!==description)throw Error('Final description mismatch: '+name+' '+field);
    }
    const nodes=[];
    for(const match of html.matchAll(/<script\b[^>]*type=["']application\/ld\+json["'][^>]*>([\s\S]*?)<\/script>/gi))pageNodes(JSON.parse(match[1]),url.href,nodes);
    if(!nodes.length||nodes.some(n=>n.description!==description))throw Error('Final page schema description mismatch: '+name);schemaPagesChecked++;
  }
  return {publicFiles:actual.length,htmlFilesChecked,sitemapPages:urls.length,uniqueTitles:titles.size,uniqueDescriptions:descriptions.size,schemaPagesChecked,buildInputsChecked,unreviewedOutputChanges:0,privateInputsPublished:false};
}

if(process.argv[1]&&path.resolve(process.argv[1])===fileURLToPath(import.meta.url)) {
  const root=path.dirname(fileURLToPath(import.meta.url));
  console.log(JSON.stringify(verifyPublicOutput(root)));
}
