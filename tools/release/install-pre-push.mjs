/** Install a repository-wide guard, preserving and chaining the existing hook. */
import fs from 'node:fs';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {execFileSync} from 'node:child_process';

const source=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'../..');
const marker='# new14-reviewed-release-guard-v1';
const quote=s=>"'"+s.replaceAll("'","'\\''")+"'";
export function installGuard(root,node=process.execPath) {
  const git=args=>execFileSync('git',['-C',root,...args],{encoding:'utf8',stdio:['pipe','pipe','pipe']}).trim();
  let configured='';try {configured=git(['config','--get','core.hooksPath']);}catch(error){if(error.status!==1)throw error;}
  if(configured)throw Error('Custom core.hooksPath exists; integrate the guard there without overwriting it: '+configured);
  const common=git(['rev-parse','--path-format=absolute','--git-common-dir']);
  const hooks=path.join(common,'hooks'),dir=path.join(hooks,'new14-release-guard');
  const hook=path.join(hooks,'pre-push'),backup=path.join(dir,'previous-pre-push');
  fs.mkdirSync(dir,{recursive:true});
  const previous=fs.existsSync(hook)?fs.readFileSync(hook):null;
  const managed=previous?.toString('utf8').includes(marker);
  if(previous&&!managed) {
    if(fs.existsSync(backup)&&!fs.readFileSync(backup).equals(previous))throw Error('Existing hook differs from the saved backup; refusing to overwrite either');
    fs.writeFileSync(backup,previous,{mode:0o755});
  }
  for(const name of ['release-inputs.mjs','release-preflight.mjs'])fs.copyFileSync(path.join(source,name),path.join(dir,name));
  const runner=`import fs from 'node:fs';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {spawnSync} from 'node:child_process';
const dir=path.dirname(fileURLToPath(import.meta.url)),input=fs.readFileSync(0);
function run(command,args){const r=spawnSync(command,args,{input,stdio:['pipe','inherit','inherit']});if(r.error){console.error(r.error.message);process.exit(1);}if(r.status!==0)process.exit(r.status??1);}
run(process.execPath,[path.join(dir,'release-preflight.mjs'),'--pre-push']);
const previous=path.join(dir,'previous-pre-push');
if(fs.existsSync(previous))run('sh',[previous,...process.argv.slice(2)]);
`;
  fs.writeFileSync(path.join(dir,'runner.mjs'),runner);
  // Commit objects are read from the invoking checkout; no dependency on this worktree remaining present.
  fs.writeFileSync(hook,'#!/bin/sh\n'+marker+'\nexec '+quote(node.replaceAll('\\','/'))+' '+quote(path.join(dir,'runner.mjs').replaceAll('\\','/'))+' "$@"\n',{mode:0o755});
  fs.chmodSync(hook,0o755);
  return {hook,previousHookPreserved:fs.existsSync(backup),scope:'all worktrees in this local repository'};
}
if(process.argv[1]&&path.resolve(process.argv[1])===fileURLToPath(import.meta.url)) {
  try {if(process.argv.length>2)throw Error('Run without arguments from the intended Git repository');console.log(JSON.stringify(installGuard(process.cwd())));}
  catch(error){console.error(error.message);process.exitCode=1;}
}
