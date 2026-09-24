/** Publish only canonical pages, public media and discovery files. */
import fs from 'node:fs';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
const root=path.dirname(fileURLToPath(import.meta.url));
const output=path.resolve(root,'.public-release');
if(path.dirname(output)!==root)throw Error('Invalid artifact root');
if(fs.existsSync(output)&&fs.lstatSync(output).isSymbolicLink())throw Error('Artifact cannot be a symlink');
fs.mkdirSync(output,{recursive:true});
const expected=new Set();let pages=0;
function copy(relative){
  const source=path.resolve(root,relative),target=path.resolve(output,relative);
  if(!source.startsWith(root+path.sep)||!target.startsWith(output+path.sep)||fs.lstatSync(source).isSymbolicLink())throw Error('Unsafe path: '+relative);
  fs.mkdirSync(path.dirname(target),{recursive:true});fs.copyFileSync(source,target);expected.add(path.relative(output,target));
}
const sitemap=fs.readFileSync(path.join(root,'sitemap.xml'),'utf8');
const urls=[...sitemap.matchAll(/<(?:\w+:)?loc>([^<]+)<\/(?:\w+:)?loc>/g)].map(m=>new URL(m[1].replaceAll('&amp;','&')));
const seen=new Set();
for(const u of urls){
  if(u.origin!=='https://xn--3e0bz50b1zcyxat54c.com'||!u.pathname.endsWith('/'))throw Error('Invalid public URL '+u.href);
  const relative=path.join(decodeURIComponent(u.pathname).replace(/^\//,''),'index.html');
  if(seen.has(relative))throw Error('Duplicate sitemap URL');seen.add(relative);copy(relative);pages++;
}
const extensions=new Set(['.css','.js','.jpg','.jpeg','.png','.gif','.webp','.avif','.svg','.ico','.woff','.woff2','.ttf','.mp4','.webm']);
function assets(dir){for(const e of fs.readdirSync(path.join(root,dir),{withFileTypes:true})){
  if(e.isSymbolicLink()||e.name.startsWith('.'))continue;
  const rel=path.join(dir,e.name);
  if(e.isDirectory())assets(rel);else if(extensions.has(path.extname(e.name).toLowerCase()))copy(rel);
}}
assets('assets');
for(const name of ['404.html','sitemap.xml','rss.xml','robots.txt','llms.txt'])copy(name);
for(const name of fs.readdirSync(root))if(/^(google[a-z0-9]+|naver[a-z0-9]+)\.html$/i.test(name))copy(name);
function check(dir){for(const e of fs.readdirSync(dir,{withFileTypes:true})){
  const f=path.join(dir,e.name);if(e.isSymbolicLink())throw Error('Unexpected artifact symlink');
  if(e.isDirectory())check(f);else if(!expected.has(path.relative(output,f)))throw Error('Unrecognized stale artifact: '+f);
}}
check(output);
console.log(JSON.stringify({output:'.public-release',pages,files:expected.size,sourceFilesPublished:false}));
