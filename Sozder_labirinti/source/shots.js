// Ойын беттерінің скриншоттары (жоба жұмысындағы суреттер үшін)
const { chromium } = require(require('child_process').execSync('npm root -g').toString().trim()+'/playwright');
const URL = 'file://' + require('path').resolve(__dirname, '../oiyn/index.html');
(async()=>{
  const b = await chromium.launch();
  const p = await b.newPage({viewport:{width:1100,height:760}, deviceScaleFactor:2, colorScheme:'light'});
  await p.goto(URL); await p.waitForTimeout(500);
  await p.fill('#pcode','3А-01');
  await p.screenshot({path:'figs/f1_home.png'});
  await p.click('[data-go=levels]'); await p.waitForTimeout(200);
  await p.screenshot({path:'figs/f2_levels.png'});
  await p.click('[data-l="3"]'); await p.waitForTimeout(300);
  // бірнеше әріп жинап, бірінші қақпаға дейін жүру
  await p.evaluate(async()=>{
    const sleep=ms=>new Promise(r=>setTimeout(r,ms));
    const dirTo=(a,b)=>{const n=G.n; const d=b-a; return d===-n?0:d===1?1:d===n?2:3;};
    const cells=[...G.letters.keys()].sort((a,b)=>bfsPath(G.w,G.n,0,a).length-bfsPath(G.w,G.n,0,b).length).slice(0,2);
    for(const t of cells){ let k=0; while(G.pos!==t && k++<200){ const nx=bfsPath(G.w,G.n,G.pos,t)[1];
      if(G.gates.has(nx)){ G.gates.delete(nx); } move(dirTo(G.pos,nx)); await sleep(140);} }
  });
  await p.waitForTimeout(300);
  await p.screenshot({path:'figs/f3_game.png'});
  // қақпа сұрағы
  await p.evaluate(()=>{ const g=[...G.gates.keys()][0]; askGate(g); });
  await p.waitForTimeout(200);
  await p.evaluate(()=>{ const Q=G.L.q.find(x=>x.q===document.querySelector('#qt').textContent); const w=[...document.querySelectorAll('.opt')].find(o=>o.dataset.o!==Q.a[0]); w.click(); });
  await p.waitForTimeout(450);
  await p.screenshot({path:'figs/f4_question.png'});
  await p.evaluate(()=>{ modalRoot.innerHTML=''; G.got=[...G.L.word].map((_,i)=>i); anagram(); });
  await p.waitForTimeout(200);
  await p.screenshot({path:'figs/f5_anagram.png'});
  await p.evaluate(()=>{ S.log={levels:[],answers:[],tests:[],survey:[]}; save(); stats(); });
  await p.waitForTimeout(300);
  await p.screenshot({path:'figs/f6_stats.png'});
  await p.evaluate(()=>{ testIntro(); }); await p.click('[data-k=pre]'); await p.waitForTimeout(200);
  await p.screenshot({path:'figs/f7_test.png'});
  await b.close();
})();
