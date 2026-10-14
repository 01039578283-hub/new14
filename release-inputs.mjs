/** One fail-closed input contract shared by the build and pre-push checks. */
import {createHash} from 'node:crypto';

export const BUILD_INPUTS=['release-public-build.mjs','wawa-analytics-build.mjs','seo-descriptions.mjs','release-public-verify.mjs','seo-descriptions.json','vercel.json','image-order-build.mjs','image-order-review.json','crawler-content-build.mjs','crawler-content-review.json','favicon-build.mjs','favicon-config.json','representative-photos-build.mjs','representative-photos-review.json','content-quality-build.mjs','content-quality-review.json','release-inputs.mjs'];
export const digest=bytes=>createHash('sha256').update(bytes).digest('hex');
export function utf8(bytes,name='input') {
  const text=bytes.toString('utf8');
  if(!Buffer.from(text).equals(bytes))throw Error('Non-UTF8 reviewed build input: '+name);
  return text;
}
export const inputDigest=(bytes,name)=>digest(Buffer.from(utf8(bytes,name).replaceAll('\r\n','\n')));
// Ignore formatting and object-key order, but preserve values and array order.
const ordered=value=>Array.isArray(value)?value.map(ordered):value&&typeof value==='object'?Object.fromEntries(Object.keys(value).sort().map(key=>[key,ordered(value[key])])):value;
export const configDigest=bytes=>digest(Buffer.from(JSON.stringify(ordered(JSON.parse(utf8(bytes,'vercel.json').replace(/^\uFEFF/,''))))));
export function verifyReviewedInputs(read,manifest) {
  const reviewed=manifest.reviewedBuildInputs;
  if(!reviewed||Object.keys(reviewed).sort().join('\n')!==[...BUILD_INPUTS].sort().join('\n'))throw Error('Reviewed build inputs are missing or incomplete; review the input inventory before release');
  const changed=[];
  for(const name of BUILD_INPUTS) {
    const bytes=read(name);
    if(name==='vercel.json') {
      // Always check the semantic hash, even when the byte hash matches. This
      // catches a half-updated manifest before provider JSON reserialization.
      if(configDigest(bytes)===manifest.reviewedVercelConfigSha256)continue;
      changed.push(name);continue;
    }
    if(inputDigest(bytes,name)===reviewed[name])continue;
    changed.push(name);
  }
  if(changed.length)throw Error('Reviewed build input changed: '+changed.join(', ')+
    '\nReview the actual diff, then explicitly record each intended change with: node release-preflight.mjs --review-input <file>.'+
    '\nCommit the input and release-public-manifest.json together. Build and push checks never update review hashes automatically.');
  return BUILD_INPUTS.length;
}
