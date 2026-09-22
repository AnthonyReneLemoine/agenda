const fs=require('node:fs'),vm=require('node:vm'),assert=require('node:assert/strict');
const source=fs.readFileSync(__dirname+'/app.js','utf8');
let height=600,cssValue='',frames=[],renders=0;
const ctx=vm.createContext({document:{getElementById:()=>({clientHeight:height}),documentElement:{style:{setProperty:(k,v)=>cssValue=v}}},requestAnimationFrame:f=>{frames.push(f);return frames.length;},renderAll:()=>renders++});
vm.runInContext('const H_START=9,H_END=23;let SLOT_H=24,listView=false;'+source.slice(source.indexOf('    // Mesure'),source.indexOf('    function parseLocalDate')),ctx);
for(height of [250,399,600,850]){vm.runInContext('fitCalendarHeight()',ctx);assert(Math.abs(parseFloat(cssValue)*28-(height-1))<0.001);}
vm.runInContext('listView=true;fitCalendarHeight()',ctx);assert.equal(parseFloat(cssValue)*28,849);
vm.runInContext('listView=false',ctx);height=0;vm.runInContext('fitCalendarHeight()',ctx);assert.equal(parseFloat(cssValue)*28,849);
vm.runInContext('scheduleCalendarResize();scheduleCalendarResize();scheduleCalendarResize()',ctx);assert.equal(frames.length,1);frames[0]();assert.equal(renders,1);
vm.runInContext('scheduleCalendarResize()',ctx);assert.equal(frames.length,2);
console.log('PASS: adaptive height, hidden/list guard, resize coalescing and repeat resizing');
