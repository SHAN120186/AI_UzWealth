// Offline checks for exported Code nodes. This is not an n8n runtime.
// No HTTP request, credential, database or account is used.
import fs from 'node:fs';
import vm from 'node:vm';
import assert from 'node:assert/strict';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
const dir=path.dirname(fileURLToPath(import.meta.url));
const read=name=>JSON.parse(fs.readFileSync(path.join(dir,name),'utf8'));
const smoke=read('uzwealth_f100_tests.json');
const backend=read('uzwealth_backend_integration.json');
const invoke=(workflow,nodeName,items,extras={})=>{
  const node=workflow.nodes.find(n=>n.name===nodeName);
  assert(node?.type==='n8n-nodes-base.code',nodeName);
  const input=JSON.parse(JSON.stringify(items));
  const ctx={$input:{all:()=>input,first:()=>input[0]},...extras};
  return JSON.parse(JSON.stringify(vm.runInNewContext('(function(){'+node.parameters.jsCode+'\n})()',ctx,{timeout:1000})));
};
let items=[{json:{}}];
for(const name of ['Generate fixtures','Validate deduplicate calculate','Assert expected results'])items=invoke(smoke,name,items);
assert.equal(items[0].json.tests_passed,true);
assert.deepEqual(items[0].json.counts,{accepted:1,duplicate:1,rejected:1});
assert.equal(items[0].json.allocation.cash_tiyin,'200000520');
assert.equal(items[0].json.allocation.btc.quantity_btc,'0.01485148');
for(const input of [
 {event_id:'bad',capital_tiyin:'-1',horizon_months:12,effective_risk:'moderate',confirmed:true},
 {event_id:'bad',capital_tiyin:'1e10',horizon_months:12,effective_risk:'moderate',confirmed:true},
 {event_id:'bad',capital_tiyin:10000000000,horizon_months:12,effective_risk:'moderate',confirmed:true},
 {event_id:'bad',capital_tiyin:'10000000000',horizon_months:12,effective_risk:'moderate',confirmed:false},
])assert.equal(invoke(smoke,'Validate deduplicate calculate',[{json:input}])[0].json.reason_code,'VALIDATION_ERROR');
const valid={event_id:'F100-001',capital_tiyin:'10000000000',horizon_months:12,effective_risk:'moderate',confirmed:true};
assert.equal(invoke(smoke,'Validate deduplicate calculate',[{json:valid},{json:valid}])[1].json.status,'duplicate');
assert.equal(invoke(smoke,'Validate deduplicate calculate',[{json:valid}])[0].json.status,'accepted');
assert.equal(invoke(smoke,'Validate deduplicate calculate',[{json:{...valid,horizon_months:24}}])[0].json.reason_code,'UNSUPPORTED_DEMO_PROFILE');
const missing=invoke(backend,'Validate config and request',[{json:{backend_base_url:'',event_id:''}}])[0].json;
assert.equal(missing.config_ok,false);
assert.equal(missing.error_code,'CONFIG_REQUIRED');
const invalidKey=invoke(backend,'Validate config and request',[{json:{backend_base_url:'https://example.invalid',event_id:'not-a-uuid'}}])[0].json;
assert.equal(invalidKey.config_ok,false);
const eventId='A914FBAB-8720-4310-9D53-44B5F2ED03DC';
const configured=invoke(backend,'Validate config and request',[{json:{backend_base_url:'https://example.invalid',event_id:eventId}}])[0].json;
assert.equal(configured.config_ok,true);
assert.equal(configured.idempotency_key,eventId.toLowerCase());
assert.equal(configured.request_body.input.crypto_opt_in,false);
const init={job_id:'11111111-1111-4111-8111-111111111111',idempotency_key:eventId.toLowerCase(),deadline_ms:Date.now()+45000};
const mocks={ $:name=>({first:()=>({json:init})}),$runIndex:0 };
const statusCase=(body,expected)=>assert.equal(invoke(backend,'Evaluate job status',[{json:{statusCode:200,body}}],mocks)[0].json.error_code,expected);
statusCase({job_id:'wrong',status:'completed',result:{}},'BACKEND_RESPONSE_INVALID');
statusCase({job_id:init.job_id,status:'completed',result:null},'BACKEND_RESPONSE_INVALID');
const completed=invoke(backend,'Evaluate job status',[{json:{statusCode:200,body:{job_id:init.job_id,status:'completed',result:{allocation_id:'fixture'}}}}],mocks)[0].json;
assert.equal(completed.status,'completed');
assert.equal(invoke(backend,'Evaluate job status',[{json:{error:'synthetic timeout'}}],mocks)[0].json.error_code,'NETWORK_ERROR');
const late={...mocks,$:name=>({first:()=>({json:{...init,deadline_ms:Date.now()-1}})})};
assert.equal(invoke(backend,'Evaluate job status',[{json:{statusCode:200,body:{job_id:init.job_id,status:'completed',result:{}}}}],late)[0].json.error_code,'POLL_TIMEOUT');
for(const w of [smoke,backend])assert.equal(w.active,false);
const wait=backend.nodes.find(n=>n.name==='Wait 5s');
assert.equal(wait.parameters.amount,5);assert.equal(wait.parameters.unit,'seconds');
for(const n of backend.nodes.filter(n=>n.type==='n8n-nodes-base.httpRequest')){
  assert.equal(n.onError,'continueRegularOutput');
  assert.equal(n.parameters.options.redirect.redirect.followRedirects,false);
  assert(!n.credentials || Object.keys(n.credentials).length===0);
}
console.log(JSON.stringify({offline_checks_passed:true,f100:items[0].json,backend_runtime_executed:false},null,2));
