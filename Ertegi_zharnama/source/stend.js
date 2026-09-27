// Renders the exhibition stand layout (Appendix 9: 120 x 80 cm) to PDF and a PNG preview.
const path = require("path");
const { chromium } = require("/opt/node22/lib/node_modules/playwright");
(async () => {
  const out = process.argv[2] || ".";
  const browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width: 1200 * 3.78, height: 800 * 3.78 } });
  await page.goto("file://" + path.join(__dirname, "stend.html"));
  await page.pdf({ path: path.join(out, "07_Stend_maketi_120x80.pdf"), width: "1200mm", height: "800mm", printBackground: true });
  await page.screenshot({ path: path.join(out, "07_Stend_maketi_preview.png"), scale: "css", fullPage: false });
  await browser.close();
})();
