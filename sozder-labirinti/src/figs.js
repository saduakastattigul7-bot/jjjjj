// Ғылыми жұмыс пен слайдтарға арналған суреттер: ойын скриншоттары, схема, диаграмма
const { chromium } = require('/opt/node22/lib/node_modules/playwright');
const path = require('path');
const GAME = 'file://' + path.resolve(__dirname, '../oiyn/index.html');
const OUT = path.join(__dirname, 'figs');
(async () => {
  const b = await chromium.launch();
  const p = await b.newPage({ viewport: { width: 400, height: 760 }, deviceScaleFactor: 2 });
  await p.goto(GAME);
  await p.evaluate(() => { try { localStorage.clear(); } catch (e) {} });
  await p.reload();
  await p.screenshot({ path: `${OUT}/s_home.png` });
  // Лабиринт: «Жануарлар» тақырыбы, бірінші сөзде 2 әріп жиналған күй
  await p.click('[data-go=themes]');
  await p.screenshot({ path: `${OUT}/s_themes.png` });
  await p.click('.theme-card >> nth=0');
  await p.evaluate(() => {
    J.words[0] = ["қоян","🐰","заяц"]; startRound();
    function bfs(from,to){const prev={};prev[from]=-1;const q=[from];while(q.length){const c=q.shift();if(c===to)break;for(const [d,m] of [['n',c-R.w],['s',c+R.w],['e',c+1],['w',c-1]]){if(!R.maze[c][d]&&prev[m]===undefined){prev[m]=c;q.push(m);}}}const path=[];let c=to;while(c!==from){path.unshift(c);c=prev[c];}return path;}
    function go(to){for(const c of bfs(R.pos,to)){const d=c-R.pos;move(d===1?1:d===-1?-1:0,d===R.w?1:d===-R.w?-1:0);}}
    for (let k=0;k<2;k++){ const need=R.word[R.next]; for(const [i,t] of R.tiles) if(t.word&&!t.taken&&t.ch===need){go(i);break;} }
    msg("Дұрыс! Келесі әріпті ізде.","good");
  });
  await p.waitForTimeout(300);
  await p.screenshot({ path: `${OUT}/s_maze.png` });
  await p.evaluate(() => {
    function bfs(from,to){const prev={};prev[from]=-1;const q=[from];while(q.length){const c=q.shift();if(c===to)break;for(const [d,m] of [['n',c-R.w],['s',c+R.w],['e',c+1],['w',c-1]]){if(!R.maze[c][d]&&prev[m]===undefined){prev[m]=c;q.push(m);}}}const path=[];let c=to;while(c!==from){path.unshift(c);c=prev[c];}return path;}
    function go(to){for(const c of bfs(R.pos,to)){const d=c-R.pos;move(d===1?1:d===-1?-1:0,d===R.w?1:d===-R.w?-1:0);}}
    while(R.next<R.word.length){const need=R.word[R.next];for(const [i,t] of R.tiles) if(t.word&&!t.taken&&t.ch===need){go(i);break;}}
    go(R.w*R.h-1);
  });
  await p.waitForTimeout(400);
  await p.screenshot({ path: `${OUT}/s_card.png` });
  await p.click('#ovNext'); await p.click('#homeLink');
  await p.click('[data-go=dict]'); await p.click('#dictTabs button >> nth=2');
  await p.screenshot({ path: `${OUT}/s_dict.png` });
  await p.click('#homeLink'); await p.click('[data-go=test]'); await p.click('[data-stage="бастапқы"]');
  await p.screenshot({ path: `${OUT}/s_test.png` });
  await p.evaluate(() => { try { localStorage.clear(); } catch (e) {} });

  // Схема және диаграмма
  const q = await b.newPage({ viewport: { width: 1000, height: 600 }, deviceScaleFactor: 2 });
  for (const f of ['scheme', 'letters']) {
    await q.goto('file://' + path.join(__dirname, `${f}.html`));
    const el = await q.$('#fig');
    await el.screenshot({ path: `${OUT}/${f}.png` });
  }
  await b.close();
})();
