// e2e_hyc.cjs —— 华阳城看板真机 E2E：加载单文件 index.html，抓控制台错误、断言关键数字、截图
const { chromium } = require('playwright');
const path = require('path');

const CHROME = '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome';
const FILE = 'file://' + path.resolve(__dirname, 'index.html');
const SHOT = process.env.SHOT || '/tmp/hyc_board.png';
const sleep = ms => new Promise(r => setTimeout(r, ms));

(async () => {
  const browser = await chromium.launch({
    headless: true, executablePath: CHROME,
    args: ['--no-sandbox', '--disable-dev-shm-usage', '--disable-blink-features=AutomationControlled'],
  });
  const page = await browser.newPage({ viewport: { width: 1440, height: 2200 } });
  const errs = [], warns = [];
  page.on('pageerror', e => errs.push('PAGEERROR: ' + e.message));
  page.on('console', m => {
    if (m.type() === 'error') errs.push('CONSOLE: ' + m.text().slice(0, 200));
    if (m.type() === 'warning') warns.push(m.text().slice(0, 120));
  });
  await page.goto(FILE, { waitUntil: 'load', timeout: 60000 });
  await sleep(4000);

  const info = await page.evaluate(() => {
    const t = document.body.innerText || '';
    const el = document.getElementById('storeName');
    return {
      title: document.title,
      storeName: el ? el.textContent.trim() : null,
      chars: t.length,
      hasHuaYang: t.includes('华阳城'),
      hasLiJia: t.includes('李家村'),
      hasShao: t.includes('邵乐乐'),
      hasZhangMei: t.includes('张梅B'),
      hasTianRui: t.includes('田蕊'),
      has207: /207[,.]?26/.test(t),
      has1660: /1[,.]?66|166\s*台|1660/.test(t),
    };
  });

  await page.screenshot({ path: SHOT, fullPage: true });
  console.log(JSON.stringify({ info, errs, warnCount: warns.length }, null, 1));
  console.log('screenshot ->', SHOT);
  await browser.close();
  process.exit(errs.length ? 1 : 0);
})().catch(e => { console.error('FATAL', e); process.exit(2); });
