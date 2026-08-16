const { chromium } = require('playwright');
(async () => {
  const b = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium-1194/chrome-linux/chrome', args: ['--no-sandbox'] });
  const p = await b.newPage({ viewport: { width: 794, height: 1123 }, deviceScaleFactor: 2 });
  await p.goto('file://' + __dirname + '/anexos4.html', { waitUntil: 'load' });
  await p.pdf({ path: '/home/user/DANISANTELMO/LYDES-anexos-PCC.pdf', format: 'A4', printBackground: true,
    margin: { top: '20mm', bottom: '16mm', left: '22mm', right: '22mm' } });
  await p.screenshot({ path: 'anexos_preview.png', fullPage: true });
  await b.close(); console.log('ok');
})();
