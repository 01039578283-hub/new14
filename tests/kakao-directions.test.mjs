import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import {createHash} from 'node:crypto';
import {fileURLToPath} from 'node:url';
import {enhanceKakaoDirectionsHTML,enhanceKakaoDirectionsSitemap,kakaoDirectionsContext} from '../kakao-directions-build.mjs';
const root=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..');
const review=kakaoDirectionsContext(root),hash=s=>createHash('sha256').update(s).digest('hex');
const name='지점안내/서울/명일점/index.html';
const source='<!doctype html>\n<html><head><link rel="canonical" href="https://xn--3e0bz50b1zcyxat54c.com/지점안내/서울/명일점/"><title>기존 제목</title></head><body><main><p>기존 주소와 원고</p><figure><img src="/body.jpg" width="800" height="1200"></figure><figure><img src="/map.jpg" width="661" height="1475" loading="lazy"></figure><p>기존 지도 뒤 안내</p></main></body></html>';
const fixture=()=>({...review,pages:{[name]:{center:'명일점',section:'명일점',inputSha256:hash(source),insertAt:source.indexOf('<p>기존 지도 뒤')}}});
test('directions follow the existing map without replacing or moving old body content',()=>{
 const result=enhanceKakaoDirectionsHTML(source,name,fixture()).html;
 assert.ok(result.includes('</figure><section class="kakao-directions"'));
 const body=s=>s.match(/<body>([\s\S]*)<\/body>/)[1];
 assert.equal(body(result).replace(/<section class="kakao-directions"[\s\S]*?<\/section>/,''),body(source));
 assert.ok(result.includes('<title>기존 제목</title>'));
});
test('the reviewed transformation is repeatable across LF/CRLF and rejects changed input',()=>{
 const ctx=fixture(),a=enhanceKakaoDirectionsHTML(source,name,ctx).html;
 assert.equal(enhanceKakaoDirectionsHTML(source.replaceAll('\n','\r\n'),name,ctx).html,a);
 assert.equal(enhanceKakaoDirectionsHTML(a,name,ctx).html,a);
 assert.throws(()=>enhanceKakaoDirectionsHTML(source.replace('기존 주소','다른 주소'),name,ctx),/differs from reviewed/);
 assert.equal(enhanceKakaoDirectionsHTML(source,'학습가이드/index.html',ctx).html,source);
});
test('sitemap dates change only for reviewed pages and URL order remains intact',()=>{
 const ctx=fixture(),map='<urlset><url><loc>https://xn--3e0bz50b1zcyxat54c.com/지점안내/서울/명일점/</loc><lastmod>2026-10-14</lastmod></url><url><loc>https://xn--3e0bz50b1zcyxat54c.com/</loc><lastmod>2026-10-01</lastmod></url></urlset>';
 const next=enhanceKakaoDirectionsSitemap(map,ctx);assert.ok(next.includes('<lastmod>2026-10-17</lastmod>'));assert.ok(next.endsWith('<lastmod>2026-10-01</lastmod></url></urlset>'));
 assert.deepEqual([...next.matchAll(/<loc>(.*?)<\/loc>/g)].map(x=>x[1]),[...map.matchAll(/<loc>(.*?)<\/loc>/g)].map(x=>x[1]));
});
test('each reviewed center uses its pinned source key, address and non-linked traffic information',()=>{
 assert.equal(Object.keys(review.maps).length,193);
 for(const [center,map] of Object.entries(review.maps)){
  const section=review.sections[center],frame=fs.readFileSync(path.join(root,new URL(map.frame,'https://site.invalid').pathname),'utf8');
  const address=section.match(/<dd>([\s\S]*?)<\/dd>/)[1].replaceAll('&quot;','"').replaceAll('&#x27;',"'").replaceAll('&lt;','<').replaceAll('&gt;','>').replaceAll('&amp;','&');
  assert.equal(address,map.address,center);assert.ok(section.includes(map.timestamp));
  assert.ok(frame.includes('"timestamp":"'+map.timestamp+'"'));assert.ok(frame.includes('"key":"'+map.key+'"'));
  assert.ok(frame.includes('noindex,follow'));assert.ok(!frame.includes('wawa-visit-collector'));
  assert.ok(!/<a\b/.test(section));assert.ok(!/(?:tel:|0\d{1,2}-\d{3,4}-\d{4})/.test(section));
  for(const heading of ['주변 지하철','주변 정류장','주변 버스'])assert.ok(section.includes(heading));
  assert.deepEqual(JSON.parse(fs.readFileSync(path.join(root,map.traffic),'utf8')),map.transport);
 }
});
