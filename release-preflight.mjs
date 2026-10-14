/** Check working files or the exact outgoing Git commit without executing it. */
import fs from 'node:fs';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {execFileSync} from 'node:child_process';
import {BUILD_INPUTS,inputDigest,configDigest,verifyReviewedInputs} from './release-inputs.mjs';

const parse=bytes=>JSON.parse(bytes.toString('utf8').replace(/^\uFEFF/,''));
const git=(root,args)=>execFileSync('git',['-C',root,...args],{maxBuffer:64*1024*1024,stdio:['pipe','pipe','pipe']});
export function checkRelease(root,ref=null) {
  let commit=null;
  if(ref) {
    if(ref.startsWith('-')||/[\r\n\0]/.test(ref))throw Error('Invalid Git ref');
    commit=git(root,['rev-parse','--verify',ref+'^{commit}']).toString().trim();
  }
  const read=name=>commit?git(root,['show',commit+':'+name]):fs.readFileSync(path.join(root,name));
  const manifest=parse(read('release-public-manifest.json'));
  const buildInputsChecked=verifyReviewedInputs(read,manifest);
  return {commit,buildInputsChecked,mode:commit?'committed-files':'working-files'};
}
export function reviewInput(root,name) {
  if(!BUILD_INPUTS.includes(name))throw Error('Not an allowlisted build input: '+name);
  const file=path.join(root,'release-public-manifest.json');
  const manifest=parse(fs.readFileSync(file));
  if(!manifest.reviewedBuildInputs)throw Error('Reviewed input inventory is missing');
  const bytes=fs.readFileSync(path.join(root,name));
  const before=manifest.reviewedBuildInputs[name],after=inputDigest(bytes,name);
  // Parse configuration before writing anything; invalid JSON must not be blessed.
  const config=name==='vercel.json'?configDigest(bytes):null;
  manifest.reviewedBuildInputs[name]=after;
  if(config)manifest.reviewedVercelConfigSha256=config;
  fs.writeFileSync(file,JSON.stringify(manifest,null,2)+'\n');
  return {reviewedInput:name,before,after,configSha256:config};
}
export function checkPush(root,input) {
  const commits=new Set();
  for(const line of input.trim().split(/\r?\n/).filter(Boolean)) {
    const fields=line.trim().split(/\s+/);
    if(fields.length!==4)throw Error('Malformed pre-push ref update');
    const [,sha,remoteRef]=fields;
    if(!remoteRef.startsWith('refs/heads/')||/^0+$/.test(sha))continue;
    if(!/^[a-f0-9]{40,64}$/.test(sha))throw Error('Invalid outgoing commit');
    commits.add(sha);
  }
  return [...commits].map(sha=>checkRelease(root,sha));
}
function main(args) {
  const root=process.cwd();
  if(!args.length)return checkRelease(root);
  if(args.length===2&&args[0]==='--ref')return checkRelease(root,args[1]);
  if(args.length===2&&args[0]==='--review-input')return reviewInput(root,args[1]);
  if(args.length===1&&args[0]==='--pre-push')return {outgoingCommits:checkPush(root,fs.readFileSync(0,'utf8'))};
  throw Error('Usage: node release-preflight.mjs [--ref <commit> | --review-input <file> | --pre-push]');
}
if(process.argv[1]&&path.resolve(process.argv[1])===fileURLToPath(import.meta.url)) {
  try {console.log(JSON.stringify(main(process.argv.slice(2))));}
  catch(error) {console.error('Release preflight blocked: '+error.message);process.exitCode=1;}
}
