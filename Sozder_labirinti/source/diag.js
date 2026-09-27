const { chromium } = require(require('child_process').execSync('npm root -g').toString().trim()+'/playwright');
(async()=>{ const b=await chromium.launch(); const p=await b.newPage({deviceScaleFactor:2, viewport:{width:1300,height:700}});
 await p.goto('file://'+require('path').resolve(__dirname,'diagrams.html'));
 await (await p.$('#d1')).screenshot({path:'figs/f8_method.png'}); await (await p.$('#d2')).screenshot({path:'figs/f9_flow.png'}); await b.close(); })();
