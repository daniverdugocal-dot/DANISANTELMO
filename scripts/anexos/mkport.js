const { chromium } = require('playwright');
(async () => {
  const b = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium-1194/chrome-linux/chrome', args: ['--no-sandbox'] });
  const p = await b.newPage({ viewport: { width: 794, height: 1123 }, deviceScaleFactor: 2 });
  await p.goto('file://' + __dirname + '/portada.html', { waitUntil: 'load' });
  await p.pdf({ path: '/home/user/DANISANTELMO/LYDES-portada-indice.pdf', format: 'A4', printBackground: true, margin: {top:'0',bottom:'0',left:'0',right:'0'} });
  await p.screenshot({ path: 'portada_preview.png', fullPage: true });
  await b.close(); console.log('ok');
})();
