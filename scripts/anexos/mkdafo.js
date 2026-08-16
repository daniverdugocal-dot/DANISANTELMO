const { chromium } = require('playwright');
(async () => {
  const b = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium-1194/chrome-linux/chrome', args: ['--no-sandbox'] });
  const p = await b.newPage({ viewport: { width: 1123, height: 794 }, deviceScaleFactor: 2 });
  await p.goto('file://' + __dirname + '/dafo.html', { waitUntil: 'load' });
  await p.pdf({ path: '/home/user/DANISANTELMO/LYDES-dafo-PCC.pdf', format: 'A4', landscape: true,
    printBackground: true, margin: { top: '14mm', bottom: '14mm', left: '14mm', right: '14mm' } });
  await p.screenshot({ path: 'dafo_preview.png', fullPage: true });
  await b.close(); console.log('ok');
})();
