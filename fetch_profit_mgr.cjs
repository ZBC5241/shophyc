// 打开毛利明细报表 → 导出Excel(原始数据) → 保存到 /tmp/profit_dl
const { chromium } = require('playwright');
const { execSync } = require('child_process');
const fs=require('fs'), path=require('path');
const CHROME="/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";
const YY="https://c3.yonyoucloud.com";
const ACCT="18591910491";
const RID="a76e21a0-fe9b-4366-9b8e-2c9327c15ab9";  // 门店毛利明细表-华为终端
const DL=process.env.PROFIT_DL||"/tmp/profit_dl";
const sleep=(ms)=>new Promise(r=>setTimeout(r,ms));
function getPwd(){return execSync(`security find-generic-password -s yonyou-mgr -w`,{encoding:'utf8'}).trim();}
function frameUrl(f){try{return typeof f.url==='function'?f.url():(f.url||'');}catch{return '';}}
async function doLogin(page){
  await page.goto(YY,{waitUntil:'domcontentloaded',timeout:30000}); await sleep(3000);
  try{const b=await page.$('.button_accept'); if(b){await b.click(); await sleep(1000);}}catch{}
  let lf=null;
  for(let i=0;i<6;i++){for(const f of page.frames()) if(frameUrl(f).includes('euc.yonyoucloud.com')){lf=f;break;} if(lf)break; await sleep(3000);}
  if(!lf){console.log("NO_LOGIN_FRAME");return false;}
  await lf.fill('#username',ACCT); await sleep(300); await lf.fill('#password',getPwd()); await sleep(300);
  await lf.click('#submit_btn_login');
  for(let i=0;i<15;i++){await sleep(2000);const u=page.url().toLowerCase();if(!u.includes('login')&&!u.includes('cas'))break;}
  await sleep(3000); console.log("LOGGED_IN"); return true;
}
(async()=>{
  fs.mkdirSync(DL,{recursive:true});
  const browser=await chromium.launch({headless:true,executablePath:CHROME,args:['--no-sandbox','--disable-dev-shm-usage','--disable-blink-features=AutomationControlled']});
  const ctx=await browser.newContext({acceptDownloads:true});
  const page=await ctx.newPage();
  if(!await doLogin(page)){await browser.close();return;}
  await page.goto(YY,{waitUntil:'domcontentloaded'}); await sleep(5000);
  for(let i=0;i<15;i++){const ok=await page.evaluate((rid)=>!!document.getElementById('recent-'+rid)||!!document.getElementById('favor-'+rid),RID); if(ok)break; await sleep(2000);}
  await page.evaluate((rid)=>{const el=document.getElementById('recent-'+rid)||document.getElementById('favor-'+rid); if(el)el.click();},RID);
  let ready=false;
  for(let i=0;i<20;i++){await sleep(2000); const c=await page.evaluate(()=>document.querySelectorAll('button.wui-dropdown-trigger.ana-header-dropdown-list').length); if(c>=3){ready=true;console.log("面板OK",c);break;}}
  if(!ready){console.log("面板未加载"); await browser.close(); return;}
  await sleep(5000);
  // 点击导出(第3个下拉)
  const triggers=await page.$$('button.wui-dropdown-trigger.ana-header-dropdown-list');
  console.log("下拉按钮数:",triggers.length);
  await triggers[2].click(); await sleep(1200);
  const has=await page.$('li[fieldid="analysis|toolbar|export|export_excel"]');
  console.log("导出Excel项:",!!has);
  if(has){ await page.click('li[fieldid="analysis|toolbar|export|export_excel"]'); await sleep(2000); }
  // 选原始数据
  await page.evaluate(()=>{
    document.querySelectorAll('input[type=radio]').forEach(r=>{
      const lb=r.parentElement?r.parentElement.textContent.trim():'';
      if(lb.includes('原始数据')) r.click();
    });
  });
  await sleep(800);
  console.log("准备下载...");
  try{
    const [dl]=await Promise.all([
      page.waitForEvent('download',{timeout:300000}),
      (async()=>{
        await page.evaluate(()=>{
          document.querySelectorAll('button,input[type=button]').forEach(b=>{
            if((b.textContent||'').trim()==='确定' && b.offsetParent!==null) b.click();
          });
        });
      })()
    ]);
    const fn=dl.suggestedFilename()||'profit.xlsx';
    const out=path.join(DL,fn);
    await dl.saveAs(out);
    console.log("✅ 下载:",out, fs.statSync(out).size);
  }catch(e){ console.log("下载失败/超时:",e.message); }
  await browser.close();
})().catch(e=>{console.error("FATAL",e);process.exit(1);});
