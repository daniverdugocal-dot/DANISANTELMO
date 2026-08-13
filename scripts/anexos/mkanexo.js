const { chromium } = require('playwright');
(async () => {
  const b = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium-1194/chrome-linux/chrome', args: ['--no-sandbox'] });
  const p = await b.newPage({ viewport: { width: 794, height: 1123 }, deviceScaleFactor: 2 });
  await p.goto('file://' + __dirname + '/anexos.html', { waitUntil: 'load' });
  await p.pdf({ path: '/home/user/DANISANTELMO/LYDES-anexo1-PCC.pdf', format: 'A4', printBackground: true,
    margin: { top: '22mm', bottom: '18mm', left: '25mm', right: '25mm' } });
  await p.screenshot({ path: 'anexo_preview.png', fullPage: true });
  // chart alone, for embedding in Word
  const fig = await p.$('figure');
  await fig.screenshot({ path: 'chart.png' });
  await b.close(); console.log('ok');
})();
