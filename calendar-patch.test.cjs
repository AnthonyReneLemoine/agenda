// Run with: node calendar-patch.test.cjs
process.env.TZ = 'Europe/Paris';
const fs = require('node:fs');
const vm = require('node:vm');
const assert = require('node:assert/strict');
const source = fs.readFileSync(process.argv[2] || __dirname + '/app.js', 'utf8');
const elements = new Map();
const element = id => {
  if (!elements.has(id)) elements.set(id, {value:'', checked:false, style:{}, classList:{add(){},remove(){}}, addEventListener(){}, focus(){}});
  return elements.get(id);
};
const context = vm.createContext({console, Date, Intl, URLSearchParams, setTimeout(){}, setInterval(){},
  ResizeObserver:class { observe(){} }, requestAnimationFrame(){},
  document:{getElementById:element, addEventListener(){}}, window:{addEventListener(){}}});
vm.runInContext(source, context);
const run = code => vm.runInContext(code, context);
run(`calendars=[{id:'calendar',name:'Test'}]; showToast=(msg,type)=>{ if(type==='error') throw new Error(msg); }; setStatus=()=>{}; loadEvents=async()=>{};`);
let sent;
context.gcalFetch = async (url,options) => { sent = {url,...options,body:JSON.parse(options.body)}; return {}; };
function merge(base, patch) {
  const result = {...base};
  for (const [key,value] of Object.entries(patch)) {
    if (value === null) delete result[key];
    else result[key] = value && typeof value === 'object' && !Array.isArray(value) ? merge(result[key],value) : value;
  }
  return result;
}
function validTimeFields(event) {
  for(const key of ['start','end']) assert.equal(Number(!!event[key].date)+Number(!!event[key].dateTime),1, key+' must have exactly one date representation');
}
(async()=>{
  // Imported midnight-to-midnight event, displayed as all-day in the modal.
  run(`openModal(null,true,{id:'imported',calendarId:'calendar',title:'CCAS',start:'2026-09-21T00:00:00+02:00',end:'2026-09-23T00:00:00+02:00',allDay:true,apiAllDay:false});`);
  assert.equal(element('m-de').value,'2026-09-22');
  element('m-ts').value='19:00'; element('m-te').value='20:00';
  await run('saveEvent()');
  assert.equal(sent.method,'PATCH');
  const old = {start:{dateTime:'2026-09-21T00:00:00+02:00',timeZone:'Europe/Paris'},end:{dateTime:'2026-09-23T00:00:00+02:00',timeZone:'Europe/Paris'},attendees:[{email:'guest@example.org'}]};
  const converted=merge(old,sent.body);
  validTimeFields(converted);
  assert.deepEqual(converted.start,{date:'2026-09-21'});
  assert.deepEqual(converted.end,{date:'2026-09-23'});
  assert.deepEqual(converted.attendees,old.attendees);
  const allDay = JSON.parse(JSON.stringify(sent.body));
  // Reverse conversion and drag/resize updates use the same wrapper.
  await run(`gcalUpdateEvent('calendar','imported',{start:{dateTime:'2026-09-21T17:00:00Z',timeZone:'Europe/Paris'},end:{dateTime:'2026-09-21T18:00:00Z',timeZone:'Europe/Paris'}})`);
  validTimeFields(merge(converted,sent.body));
  assert.equal(sent.body.start.date,null);
  assert.equal(sent.body.end.date,null);
  // Same-type edits remain valid; title-only edits do not touch event times.
  validTimeFields(merge(old,sent.body));
  validTimeFields(merge(converted,allDay));
  await run(`gcalUpdateEvent('calendar','imported',{summary:'Renamed'})`);
  assert.deepEqual(sent.body,{summary:'Renamed'});
  // POST remains unchanged, with no patch-only clearing fields.
  await run(`gcalCreateEvent('calendar',buildGCalBody({title:'New',allDay:true,startDate:'2026-09-30',endDate:'2026-09-30'}))`);
  assert.equal(sent.method,'POST');
  assert.deepEqual(sent.body.start,{date:'2026-09-30'});
  assert.deepEqual(sent.body.end,{date:'2026-10-01'});
  // The caller's payload is never mutated.
  await run(`(async()=>{const body={start:{date:'2026-09-21'},end:{date:'2026-09-22'}}; const before=JSON.stringify(body); await gcalUpdateEvent('calendar','id',body); if(JSON.stringify(body)!==before) throw new Error('mutated');})()`);
  console.log('PASS: imported event modal/save, both conversions, same-type edits, partial edits, POST and payload immutability');
})().catch(e=>{console.error(e);process.exitCode=1;});
