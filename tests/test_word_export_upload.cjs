const assert=require('node:assert/strict');
const fs=require('node:fs');
const path=require('node:path');
const vm=require('node:vm');
const source=fs.readFileSync(path.join(__dirname,'..','word-export.js'),'utf8').split('function downloadWordBlob')[0];
const headers={'X-Export-Token':'test-token'};
async function run(fetch,file){
  const context=vm.createContext({fetch,Blob,TypeError,Error});
  vm.runInContext(source,context);
  return context.uploadWordAsset('图/照片.png',file,headers);
}
(async()=>{
  let requests=[];
  const file=new Blob(['image bytes']);
  const result=await run(async(url,options)=>{
    requests.push([url,options]);
    if(requests.length===1)throw new TypeError('Failed to fetch');
    if(url==='/api/session')return {ok:true,json:async()=>({token:'test-token'})};
    return {ok:true};
  },file);
  assert.equal(result.ok,true);
  assert.equal(requests.length,3);
  assert.equal(await requests[2][1].body.text(),'image bytes');
  assert.notEqual(requests[2][1].body,file);
  assert.equal(requests[0][1].headers,requests[2][1].headers);
  await assert.rejects(run(async()=>{throw new TypeError('Failed to fetch');},file),/本地服务连接中断/);
  await assert.rejects(run(async url=>{
    if(url==='/api/session')return {ok:true,json:async()=>({token:'test-token'})};
    throw new TypeError('Failed to fetch');
  },{arrayBuffer:async()=>{throw new Error('file changed');}}),/无法读取素材.*照片.png/);
  await assert.rejects(run(async url=>{
    if(url==='/api/session')return {ok:true,json:async()=>({token:'new-token'})};
    throw new TypeError('Failed to fetch');
  },file),/服务已重启/);
  requests=[];
  await run(async(url,options)=>{requests.push([url,options]);return {ok:true};},file);
  assert.equal(requests.length,1);
  console.log('Word upload recovery checks passed');
})().catch(error=>{console.error(error);process.exitCode=1;});
