const { chromium } = require('playwright');
const { spawn } = require('child_process');
const assert = require('assert');

const port = 18765;
const server = spawn('.venv/bin/python', ['-m', 'uvicorn', 'app.main:app', '--host', '127.0.0.1', '--port', String(port)], { stdio: 'ignore' });
const rows = [
  {sku:'TEST-SET-001', category:'女士两件套', title:'Set', description_en:'Set', price:'200.00', supplier_code:'SET', attr_39:'低弹', attr_160:'棉', images:{display:[{url:'https://example.com/a.jpg'},{url:'https://example.com/b.jpg'}]}},
  {sku:'TEST-DRESS-001', category:'女士短连衣裙', title:'Dress', description_en:'Dress', price:'200.00'},
  {sku:'TEST-JUMP-001', category:'女士连体裤', title:'Jump', description_en:'Jump', price:'200.00'}
];
async function ready() {
  for (let i = 0; i < 50; i++) {
    try { const response = await fetch(`http://127.0.0.1:${port}/`); if (response.ok) return; } catch (_) {}
    await new Promise(resolve => setTimeout(resolve, 100));
  }
  throw new Error('FastAPI did not start');
}
(async () => {
  let browser;
  try {
    await ready();
    browser = await chromium.launch({headless:true, executablePath:'/usr/bin/chromium', args:['--no-sandbox']});
    const page = await browser.newPage();
    const errors = [];
    page.on('pageerror', error => errors.push(error.message));
    let validateRequests = 0;
    page.on('request', request => { if (request.url().includes('/api/validate')) validateRequests++; });
    await page.route('**/api/drafts', route => route.fulfill({status:200, contentType:'application/json', body:JSON.stringify({products:rows})}));
    await page.goto(`http://127.0.0.1:${port}/`);
    assert.equal(errors.length, 0, errors.join('\n'));
    await page.getByRole('button', {name:'检查并预览'}).click();
    assert.equal(validateRequests, 0);
    assert.match(await page.locator('#result').innerText(), /请先上传/);
    await page.getByRole('button', {name:'加载草稿'}).click();
    await page.locator('.product-row').first().waitFor();
    assert.equal(await page.locator('.product-row').count(), 3);
    await page.locator('input[data-select="0"]').check();
    const responsePromise = page.waitForResponse(r => r.url().endsWith('/api/validate-draft'));
    await page.getByRole('button', {name:'检查并预览'}).click();
    const response = await responsePromise;
    assert.equal(response.status(), 200);
    const checked = await response.json();
    assert.equal(checked.count, 1);
    assert.equal(checked.preview[0].sku, 'TEST-SET-001');
    assert.equal(await page.locator('.product-row').count(), 3);
    await page.locator('.sku').first().click();
    assert.equal(await page.locator('#imageManager .thumb').count(), 2);
    await page.locator('input[data-select="0"]').uncheck();
    await page.locator('#products').setInputFiles({name:'bad.csv', mimeType:'text/csv', buffer:Buffer.from('category,title,description_en,sku,price\n女士两件套,,Set,NEW-001,200.00\n')});
    const csvResponse = page.waitForResponse(r => r.url().endsWith('/api/validate'));
    await page.getByRole('button', {name:'检查并预览'}).click();
    assert.equal((await csvResponse).status(), 200);
    assert.equal(await page.locator('.product-row').count(), 3);
    assert.match(await page.locator('#result').innerText(), /必填项为空/);
    await page.locator('#products').setInputFiles({name:'good.csv', mimeType:'text/csv', buffer:Buffer.from('category,title,description_en,sku,price\n女士两件套,Set,Set,NEW-001,200.00\n')});
    const validCsvResponse = page.waitForResponse(r => r.url().endsWith('/api/validate'));
    await page.getByRole('button', {name:'检查并预览'}).click();
    assert.equal((await validCsvResponse).status(), 200);
    assert.equal(await page.locator('.product-row').count(), 1);
    assert.equal(await page.locator('.sku').first().inputValue(), 'NEW-001');
    assert.equal(errors.length, 0, errors.join('\n'));
    console.log('Browser: no JS error; no undefined upload; draft load; one-SKU validation; failed CSV keeps 3 rows; valid CSV loads; 2 thumbnails visible');
  } finally {
    if (browser) await browser.close();
    server.kill('SIGTERM');
  }
})().catch(error => { console.error(error); process.exitCode = 1; });
