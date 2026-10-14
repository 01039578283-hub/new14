import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {execFileSync,spawnSync} from 'node:child_process';
import {BUILD_INPUTS,inputDigest,configDigest,verifyReviewedInputs} from '../release-inputs.mjs';
import {checkRelease,reviewInput,checkPush} from '../release-preflight.mjs';
import {installGuard} from '../tools/release/install-pre-push.mjs';

const source=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..');
const fixture=t=>{
  const root=fs.mkdtempSync(path.join(os.tmpdir(),'new14-release-test-'));
  t.after(()=>{const resolved=path.resolve(root);if(path.dirname(resolved)!==path.resolve(os.tmpdir())||!path.basename(resolved).startsWith('new14-release-test-'))throw Error('Unsafe fixture cleanup');fs.rmSync(resolved,{recursive:true,force:true});});
  const config={buildCommand:'node content-quality-build.mjs',outputDirectory:'.public-release',trailingSlash:true,redirects:[{source:'/old/',destination:'/new/',permanent:true},{source:'/second/',destination:'/new/',permanent:true}],headers:[{source:'/assets/(.*)',headers:[{key:'Cache-Control',value:'public, max-age=0, must-revalidate'}]}]};
  for(const n of BUILD_INPUTS)fs.writeFileSync(path.join(root,n),n==='vercel.json'?JSON.stringify(config,null,2)+'\n':'// fixture '+n+'\n');
  const manifest={files:{'index.html':'public-snapshot-must-remain-unchanged'},reviewedBuildInputs:Object.fromEntries(BUILD_INPUTS.map(n=>[n,inputDigest(fs.readFileSync(path.join(root,n)),n)])),reviewedVercelConfigSha256:configDigest(fs.readFileSync(path.join(root,'vercel.json')))};
  fs.writeFileSync(path.join(root,'release-public-manifest.json'),JSON.stringify(manifest,null,2)+'\n');
  return {root,config,manifest};
};
const writeConfig=(root,c)=>fs.writeFileSync(path.join(root,'vercel.json'),JSON.stringify(c,null,2)+'\n');
const readManifest=root=>JSON.parse(fs.readFileSync(path.join(root,'release-public-manifest.json'),'utf8'));
const runGit=(root,...args)=>execFileSync('git',['-C',root,...args],{encoding:'utf8',stdio:['pipe','pipe','pipe']}).trim();
function initGit(root){runGit(root,'init','-q');runGit(root,'config','user.name','Release test');runGit(root,'config','user.email','release-test@example.invalid');runGit(root,'config','core.autocrlf','false');runGit(root,'config','commit.gpgsign','false');}
const commit=(root,message)=>{runGit(root,'add','--all');runGit(root,'commit','-qm',message);return runGit(root,'rev-parse','HEAD');};

test('LF, CRLF and JSON whitespace/key ordering retain the exact semantic contract',t=>{
  const {root,config}=fixture(t);assert.equal(checkRelease(root).buildInputsChecked,BUILD_INPUTS.length);
  for(const n of BUILD_INPUTS)fs.writeFileSync(path.join(root,n),fs.readFileSync(path.join(root,n),'utf8').replaceAll('\n','\r\n'));
  fs.writeFileSync(path.join(root,'vercel.json'),'\uFEFF'+JSON.stringify(Object.fromEntries(Object.entries(config).reverse())));
  assert.equal(checkRelease(root).buildInputsChecked,BUILD_INPUTS.length);
});
for(const [name,mutate] of [
  ['unrecorded redirect addition',c=>c.redirects.push({source:'/teacher+name/',destination:'/teacher/',permanent:true})],
  ['redirect destination change',c=>c.redirects[0].destination='/wrong/'],
  ['redirect array order change',c=>c.redirects.reverse()],
  ['build command change',c=>c.buildCommand='node other.mjs'],
  ['output directory change',c=>c.outputDirectory='.'],
  ['header removal',c=>delete c.headers],
  ['new configuration key',c=>c.cleanUrls=true]
])test(name+' is rejected without updating the manifest',t=>{
  const {root,config}=fixture(t),before=fs.readFileSync(path.join(root,'release-public-manifest.json'));
  mutate(config);writeConfig(root,config);
  assert.throws(()=>checkRelease(root),/Reviewed build input changed: vercel.json/);
  assert.deepEqual(fs.readFileSync(path.join(root,'release-public-manifest.json')),before);
});
test('changed scripts, missing inputs and invalid UTF-8 are blocked',t=>{
  const {root,manifest}=fixture(t);
  fs.appendFileSync(path.join(root,'content-quality-build.mjs'),'// changed');
  assert.throws(()=>checkRelease(root),/content-quality-build.mjs/);
  fs.writeFileSync(path.join(root,'content-quality-build.mjs'),Buffer.from([255]));
  assert.throws(()=>checkRelease(root),/Non-UTF8/);
  delete manifest.reviewedBuildInputs['release-inputs.mjs'];
  assert.throws(()=>verifyReviewedInputs(n=>fs.readFileSync(path.join(root,n)),manifest),/incomplete/);
});
test('a stale semantic configuration digest is rejected even if the byte digest matches',t=>{
  const {root,manifest}=fixture(t);
  manifest.reviewedVercelConfigSha256='0'.repeat(64);
  fs.writeFileSync(path.join(root,'release-public-manifest.json'),JSON.stringify(manifest));
  assert.throws(()=>checkRelease(root),/vercel.json/);
  delete manifest.reviewedVercelConfigSha256;
  fs.writeFileSync(path.join(root,'release-public-manifest.json'),JSON.stringify(manifest));
  assert.throws(()=>checkRelease(root),/vercel.json/);
});
test('explicit review updates only the named input and its semantic digest',t=>{
  const {root,config,manifest}=fixture(t);config.redirects.push({source:'/added/',destination:'/new/',permanent:true});writeConfig(root,config);
  fs.appendFileSync(path.join(root,'seo-descriptions.mjs'),'// still unreviewed');
  reviewInput(root,'vercel.json');const after=readManifest(root);
  assert.deepEqual(after.files,manifest.files);
  for(const n of BUILD_INPUTS.filter(n=>n!=='vercel.json'))assert.equal(after.reviewedBuildInputs[n],manifest.reviewedBuildInputs[n]);
  assert.notEqual(after.reviewedVercelConfigSha256,manifest.reviewedVercelConfigSha256);
  assert.throws(()=>checkRelease(root),/seo-descriptions.mjs/);
  assert.throws(()=>reviewInput(root,'../outside'),/allowlisted/);
  const before=fs.readFileSync(path.join(root,'release-public-manifest.json'));
  fs.writeFileSync(path.join(root,'vercel.json'),'{');assert.throws(()=>reviewInput(root,'vercel.json'));
  assert.deepEqual(fs.readFileSync(path.join(root,'release-public-manifest.json')),before);
});
test('checks outgoing commits rather than dirty or staged working files, including multi-ref pushes',t=>{
  const {root,config}=fixture(t);initGit(root);const good=commit(root,'reviewed');
  config.redirects.push({source:'/forgotten/',destination:'/new/',permanent:true});writeConfig(root,config);const bad=commit(root,'forgot manifest');
  reviewInput(root,'vercel.json');runGit(root,'add','release-public-manifest.json');
  assert.doesNotThrow(()=>checkRelease(root));assert.throws(()=>checkRelease(root,bad),/vercel.json/);
  assert.doesNotThrow(()=>checkRelease(root,good));
  const zero='0'.repeat(40);
  assert.throws(()=>checkPush(root,`refs/heads/good ${good} refs/heads/main ${zero}\nrefs/heads/bad ${bad} refs/heads/preview ${zero}\n`),/vercel.json/);
  assert.deepEqual(checkPush(root,`(delete) ${zero} refs/heads/old ${good}\n`),[]);
  const fixed=commit(root,'record reviewed configuration');assert.equal(checkPush(root,`refs/heads/fixed ${fixed} refs/heads/main ${good}\n`)[0].commit,fixed);
});
test('installed hook blocks an actual push, preserves previous hook stdin, and works from another worktree',t=>{
  const {root,config}=fixture(t);initGit(root);const good=commit(root,'reviewed');
  const bare=path.join(root,'remote.git');runGit(root,'init','--bare','-q',bare);runGit(root,'remote','add','origin',bare);
  const hooks=path.join(root,'.git','hooks'),previous=Buffer.from('#!/bin/sh\ncat > previous-hook-input.txt\nprintf "%s\\n" "$@" > previous-hook-args.txt\n');
  fs.writeFileSync(path.join(hooks,'pre-push'),previous,{mode:0o755});
  installGuard(root);installGuard(root);
  assert.deepEqual(fs.readFileSync(path.join(hooks,'new14-release-guard','previous-pre-push')),previous);
  runGit(root,'push','origin',good+':refs/heads/main');
  assert.match(fs.readFileSync(path.join(root,'previous-hook-input.txt'),'utf8'),new RegExp(good));
  assert.equal(fs.readFileSync(path.join(root,'previous-hook-args.txt'),'utf8').split('\n')[0],'origin');
  config.buildCommand='node unreviewed.mjs';writeConfig(root,config);runGit(root,'add','vercel.json');runGit(root,'commit','-qm','bad');const bad=runGit(root,'rev-parse','HEAD');
  // Fix only the working manifest: the outgoing commit must still fail.
  reviewInput(root,'vercel.json');
  const result=spawnSync('git',['-C',root,'push','origin',bad+':refs/heads/main'],{encoding:'utf8'});
  assert.notEqual(result.status,0);assert.match(result.stderr,/Release preflight blocked/);
  assert.equal(runGit(root,'--git-dir',bare,'rev-parse','refs/heads/main'),good);
  const other=path.join(root,'other-worktree');runGit(root,'worktree','add','--detach',other,good);
  runGit(other,'push','origin',good+':refs/heads/other');
  assert.match(fs.readFileSync(path.join(other,'previous-hook-input.txt'),'utf8'),/refs\/heads\/other/);
});
test('custom hooksPath is retained and never silently replaced',t=>{
  const {root}=fixture(t);initGit(root);runGit(root,'config','core.hooksPath','custom-hooks');
  assert.throws(()=>installGuard(root),/Custom core.hooksPath/);assert.equal(runGit(root,'config','--get','core.hooksPath'),'custom-hooks');
});
