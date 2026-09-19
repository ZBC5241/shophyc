// check_khjx.cjs —— 专查「考核机型」板块在看板上的真实渲染结果
const { chromium } = require('playwright');
const path = require('path');

const CHROME = '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome';
const FILE = 'file://' + path.resolve(__dirname, 'index.html');
const sleep = ms => new Promise(r => setTimeout(r, ms));

(async () => {
  const browser = await chromium.launch({
    headless: true, executablePath: CHROME,
    args: ['--no-sandbox', '--disable-dev-shm-usage', '--disable-blink-features=AutomationControlled'],
  });
  const page = await browser.newPage({ viewport: { width: 1440, height: 2400 } });
  const errs = [];
  page.on('pageerror', e => errs.push('PAGEERROR: ' + e.message));
  page.on('console', m => { if (m.type() === 'error') errs.push('CONSOLE: ' + m.text().slice(0, 200)); });
  await page.goto(FILE, { waitUntil: 'load', timeout: 60000 });
  await sleep(4000);

  const out = await page.evaluate(() => {
    const secs = [...document.querySelectorAll('.sec')];
    const found = secs
      .filter(s => (s.innerText || '').includes('考核机型'))
      .map(s => (s.innerText || '').replace(/\n+/g, ' | ').trim());
    const allText = document.body.innerText || '';
    let data = null;
    try {
      data = {
        storeQcs: (DATA.store.qcs || {})['考核机型'] || null,
        people: Object.fromEntries(Object.entries(DATA.people).map(
          ([n, p]) => [n, (p.qcs || {})['考核机型'] || null])),
        employees: DATA.meta.employees,
      };
    } catch (e) { data = 'ERR ' + e.message; }
    return {
      khjxSections: found,
      hasWord: allText.includes('考核机型'),
      hasBiancheng: allText.includes('华阳城'),
      data,
    };
  });

  console.log(JSON.stringify({ out, errs }, null, 1));
  await browser.close();
})().catch(e => { console.error('FATAL', e); process.exit(2); });
