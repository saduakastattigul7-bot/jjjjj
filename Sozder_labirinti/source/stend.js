// Көрме стендінің панельдері (9-қосымша): сол 30×80 см, орта 60×80 см, оң 30×80 см
const { chromium } = require(require('child_process').execSync('npm root -g').toString().trim()+'/playwright');
const path = require('path');
(async()=>{
  const b = await chromium.launch(); const p = await b.newPage();
  await p.goto('file://' + path.resolve(__dirname, 'stend.html')); await p.waitForTimeout(500);
  for (const [id, w] of [['left',300],['center',600],['right',300]]) {
    await p.evaluate(id => document.querySelectorAll('.panel').forEach(s => s.style.display = s.id === id ? 'flex' : 'none'), id);
    await p.pdf({ path: `stend_${id}.pdf`, width: `${w}mm`, height: '800mm', printBackground: true, pageRanges: '1' });
    await p.evaluate(() => document.querySelectorAll('.panel').forEach(s => s.style.display = 'flex'));
  }
  await b.close();
  // үш панельді бір PDF-ке біріктіру
  require('child_process').execSync(`python3 -c "from pypdf import PdfWriter; w=PdfWriter(); [w.append(f'stend_{x}.pdf') for x in ('left','center','right')]; w.write('../11_Stend_panelderi.pdf')"`);
  for (const x of ['left','center','right']) require('fs').unlinkSync(`stend_${x}.pdf`);
})();
