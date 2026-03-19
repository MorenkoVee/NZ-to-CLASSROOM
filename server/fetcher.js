import puppeteer from 'puppeteer-extra';
import StealthPlugin from 'puppeteer-extra-plugin-stealth';
import { existsSync } from 'fs';

puppeteer.use(StealthPlugin());

const chromePaths = [
  'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
  'C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe',
  process.env.PUPPETEER_EXECUTABLE_PATH
].filter(Boolean);

function getChromePath() {
  for (const p of chromePaths) {
    if (p && existsSync(p)) return p;
  }
  return undefined;
}

const cookieCache = new Map();
const CACHE_TTL = 30 * 60 * 1000;

async function loginToNz(page, login, password, fromRedirect = false) {
  if (!fromRedirect) {
    await page.goto('https://nz.ua/', { waitUntil: 'load', timeout: 20000 });
    await new Promise(r => setTimeout(r, 1500));
  }
  const loginSelectors = [
    'input[name="login"]',
    'input[name="username"]',
    'input[name="email"]',
    'input[name="user"]',
    'input[type="text"]',
    'input[autocomplete="username"]'
  ];
  const passSelectors = [
    'input[name="password"]',
    'input[name="pass"]',
    'input[type="password"]'
  ];
  let loginEl = null;
  let passEl = null;
  for (const sel of loginSelectors) {
    try {
      const el = await page.$(sel);
      if (el) {
        const vis = await el.evaluate(e => e.offsetParent !== null);
        if (vis) { loginEl = el; break; }
      }
    } catch (_) {}
  }
  for (const sel of passSelectors) {
    try {
      const el = await page.$(sel);
      if (el) {
        const vis = await el.evaluate(e => e.offsetParent !== null);
        if (vis) { passEl = el; break; }
      }
    } catch (_) {}
  }
  if (!loginEl || !passEl) {
    if (!fromRedirect) {
      const href = await page.evaluate(() => {
        const a = document.querySelector('a[href*="login"], a[href*="auth"], a[href*="vhid"]');
        return a ? a.href : null;
      });
      if (href) {
        await page.goto(href, { waitUntil: 'networkidle2', timeout: 15000 });
        await new Promise(r => setTimeout(r, 1500));
        return loginToNz(page, login, password, true);
      }
    }
    throw new Error('Не знайдено форму входу на nz.ua');
  }
  await loginEl.type(login, { delay: 50 });
  await passEl.type(password, { delay: 50 });
  await Promise.all([
    page.waitForNavigation({ waitUntil: 'networkidle2', timeout: 15000 }).catch(() => {}),
    page.evaluate(() => {
      const form = document.querySelector('form');
      if (form) form.submit();
      else {
        const btn = document.querySelector('button[type="submit"], input[type="submit"], .btn-login, [type="submit"]');
        if (btn) btn.click();
      }
    })
  ]);
  await new Promise(r => setTimeout(r, 1500));
  const cookies = await page.cookies();
  return cookies;
}

export async function fetchWithPuppeteer(url, options = {}) {
  const { login, password } = options;
  let browser;
  const execPath = getChromePath();
  const launchOpts = {
    headless: process.env.HEADLESS !== 'false' ? 'new' : false,
    args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-dev-shm-usage', '--disable-gpu', '--disable-blink-features=AutomationControlled']
  };
  if (execPath) launchOpts.executablePath = execPath;
  try {
    browser = await puppeteer.launch(launchOpts);
    const page = await browser.newPage();
    await page.setUserAgent('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36');
    if (login && password) {
      const cacheKey = login;
      const cached = cookieCache.get(cacheKey);
      if (cached && Date.now() - cached.ts <= CACHE_TTL) {
        await page.goto('https://nz.ua/', { waitUntil: 'domcontentloaded', timeout: 15000 });
        await page.setCookie(...cached.cookies);
      } else {
        await loginToNz(page, login, password);
        const cookies = await page.cookies();
        cookieCache.set(cacheKey, { cookies, ts: Date.now() });
      }
    }
    await page.goto(url, { waitUntil: 'load', timeout: 25000 });
    await page.waitForSelector('table.journal-choose, table', { timeout: 8000 }).catch(() => {});
    await new Promise(r => setTimeout(r, 1500));
    const html = await page.content();
    return html;
  } finally {
    if (browser) await browser.close();
  }
}
