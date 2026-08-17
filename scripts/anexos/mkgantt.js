const { chromium } = require('playwright');
(async () => {
  const b = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium-1194/chrome-linux/chrome', args: ['--no-sandbox'] });
  const p = await b.newPage({ viewport: { width: 960, height: 400 }, deviceScaleFactor: 3 });
  await p.goto('file://' + __dirname + '/cronograma.html', { waitUntil: 'load' });
  const el = await p.$('#wrap');
  await el.screenshot({ path: '/home/user/DANISANTELMO/LYDES-cronograma.png' });
  await b.close(); console.log('png ok');
})();
