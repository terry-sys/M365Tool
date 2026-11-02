// Run with Node.js and Playwright; uses an isolated headless browser and local fixtures only.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { chromium } = require('playwright');

const repository = path.resolve(__dirname, '..');
const source = fs.readFileSync(path.join(repository, 'src/M365Tool.UI/MyAccountLoginVerificationForm.cs'), 'utf8');
const constant = source.split('internal const string SignedInAccountScript =')[1].split('private void CompleteVerification')[0];
const script = [...constant.matchAll(/"(?:\\.|[^"\\])*"/g)].map(match => JSON.parse(match[0])).join('');
const profile = fs.readFileSync(path.join(__dirname, 'M365Tool.Application.Tests/Fixtures/portal-profile-cn.html'), 'utf8');
const captureConstant = source.split('internal const string SignInAccountCaptureScript =')[1].split('private Task<string> ReadSignedInAccountAsync')[0];
const captureScript = [...captureConstant.split('internal const string SignedInAccountScript =')[0].matchAll(/"(?:\\.|[^"\\])*"/g)].map(match => JSON.parse(match[0])).join('');
const surfaceConstant = source.split('internal const string SignInSurfaceReadyScript =')[1].split('internal const string SignInAccountCaptureScript =')[0];
const surfaceScript = [...surfaceConstant.matchAll(/"(?:\\.|[^"\\])*"/g)].map(match => JSON.parse(match[0])).join('');
const account = 'actual@tenant.onmschina.cn';
const portal = 'https://myaccount.windowsazure.cn/';
const cases = [
  ['Chinese profile card', profile, account],
  ['English profile card', profile.replace('从所有位置注销', 'Sign out everywhere'), account],
  ['No sign-out action', profile.replace('<button id="signout-everywhere">从所有位置注销</button>', ''), ''],
  ['No profile photo', profile.replace('id="profilePhoto"', 'id="unrelated-photo"'), ''],
  ['Hidden profile card', profile.replace('class="ms-Card"', 'class="ms-Card" style="display:none"'), ''],
  ['Invisible profile photo', profile.replace('id="profilePhoto"', 'id="profilePhoto" style="visibility:hidden"'), ''],
  ['Display name without account', profile.replace(/<a href="mailto:[^"]+">[^<]+<\/a>/, ''), ''],
  ['Hidden account email', profile.replace('href="mailto:', 'style="display:none" href="mailto:'), ''],
  ['Sign-out action outside profile card', profile.replace('<p><button id="signout-everywhere">从所有位置注销</button></p>', '') + '<button>从所有位置注销</button>', ''],
  ['Typed email is not sign-in proof', `<input type="email" value="${account}"><p>${account}</p>`, ''],
  ['Error page is not sign-in proof', `<h1>登录失败</h1><p>${account}</p>`, ''],
  ['Wrong portal host', profile, '', 'https://example.invalid/'],
  ['HTTP portal is rejected', profile, '', 'http://myaccount.windowsazure.cn/']
];

(async () => {
  const browser = await chromium.launch({ headless: true, executablePath: process.env.M365_BROWSER_EXECUTABLE || undefined });
  try {
    for (const [name, html, expected, url = portal] of cases) {
      const page = await browser.newPage();
      await page.route('**/*', route => route.fulfill({ contentType: 'text/html; charset=utf-8', body: html }));
      await page.goto(url);
      assert.equal(await page.evaluate(script), expected, name);
      await page.close();
      console.log('PASS: ' + name);
    }
    const page = await browser.newPage();
    const menu = `<button id="mectrl_main_trigger" aria-expanded="false" onclick="document.querySelector('#account-menu').style.display='block';this.setAttribute('aria-expanded','true')">Account</button>
      <div id="account-menu" style="display:none"><span id="mectrl_currentAccount_secondary">${account}</span><button id="mectrl_body_signOut">Sign out</button></div>`;
    await page.route('**/*', route => route.fulfill({ contentType: 'text/html; charset=utf-8', body: menu }));
    await page.goto(portal);
    assert.equal(await page.evaluate(script), '', 'Closed account menu waits for the signed-in identity');
    assert.equal(await page.evaluate(script), account, 'Modern account menu still verifies');
    await page.close();
    console.log('PASS: Modern account menu');

    const readinessCases = [
      ['Visible email field', '<input type="email">', '1'],
      ['Hidden input', '<input type="email" style="display:none">', ''],
      ['Transparent input', '<input type="email" style="opacity:0">', ''],
      ['Disabled field', '<input type="email" disabled>', ''],
      ['Enabled continue button', '<button>Continue</button>', '1'],
      ['Empty app shell', '<div id="root"></div>', '']
    ];
    for (const [name, html, expected] of readinessCases) {
      const readyPage = await browser.newPage();
      await readyPage.route('**/*', route => route.fulfill({ contentType: 'text/html', body: html }));
      await readyPage.goto('https://login.partner.microsoftonline.cn/');
      assert.equal(await readyPage.evaluate(surfaceScript), expected, name);
      await readyPage.close();
      console.log('PASS: Sign-in surface: ' + name);
    }
    const delayed = await browser.newPage();
    let releaseImage;
    const imageGate = new Promise(resolve => { releaseImage = resolve; });
    await delayed.route('**/*', async route => {
      if (route.request().url().endsWith('/slow.png')) {
        await imageGate;
        return route.fulfill({ contentType: 'image/png', body: Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9ZlS8AAAAASUVORK5CYII=', 'base64') });
      }
      return route.fulfill({ contentType: 'text/html', body: `<img src="/slow.png"><script>setTimeout(()=>document.body.insertAdjacentHTML('beforeend','<input type="email">'),300)</script>` });
    });
    await delayed.goto('https://login.partner.microsoftonline.cn/', { waitUntil: 'domcontentloaded' });
    assert.equal(await delayed.evaluate(surfaceScript), '', 'App shell is not ready immediately after DOMContentLoaded');
    await delayed.locator('input[type=email]').waitFor({ state: 'visible' });
    assert.equal(await delayed.evaluate(surfaceScript), '1', 'Late-rendered form is usable before all resources finish');
    assert.notEqual(await delayed.evaluate(() => document.readyState), 'complete', 'A slow resource still prevents navigation completion');
    releaseImage();
    await delayed.close();
    console.log('PASS: Late-rendered sign-in form is ready before navigation completion');

    const signIn = await browser.newPage();
    const messages = [];
    await signIn.exposeFunction('recordAccountForTest', message => messages.push(message));
    await signIn.addInitScript(() => {
      window.chrome = window.chrome || {};
      window.chrome.webview = { postMessage: message => window.recordAccountForTest(message) };
    });
    await signIn.addInitScript(captureScript);
    await signIn.route('**/*', route => {
      const url = new URL(route.request().url());
      if (url.pathname === '/complete') return route.fulfill({ status: 302, headers: {
        location: portal, 'set-cookie': 'ESTSAUTH=fixture-session-only; Path=/; Secure; HttpOnly'
      }});
      if (url.host === 'myaccount.windowsazure.cn') return route.fulfill({ contentType: 'text/html', body: '<title>Loading profile</title>' });
      return route.fulfill({ contentType: 'text/html', body: '<input type="email" name="loginfmt"><input type="password"><a href="/complete">Continue</a>' });
    });
    await signIn.goto('https://login.partner.microsoftonline.cn/');
    await signIn.locator('input[type=email]').fill(account);
    await signIn.locator('input[type=password]').fill('TEST_ONLY_PASSWORD_DO_NOT_RECORD');
    await signIn.locator('a').click();
    await signIn.waitForURL(portal);
    assert.ok(messages.some(message => message.account === account), 'Account is captured before the fast redirect');
    assert.ok(messages.every(message => message.account === account), 'Only the email is recorded, never the password');
    const cookies = await signIn.context().cookies('https://login.partner.microsoftonline.cn/');
    assert.ok(cookies.some(cookie => cookie.name === 'ESTSAUTH' && cookie.httpOnly), 'Server-established session is available after the return');
    assert.equal(await signIn.locator('#profilePhoto').count(), 0, 'Authentication does not require profile rendering');
    await signIn.close();
    console.log('PASS: Fast authentication return captures the account and session before profile rendering');
    console.log(`${cases.length + 2 + readinessCases.length + 1} portal browser regression cases passed; all requests were fulfilled from local fixtures.`);
  } finally { await browser.close(); }
})().catch(error => { console.error(error); process.exitCode = 1; });
