// login.js
import { chromium } from "playwright";

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

/**
 * @param {object} config
 * @returns {Promise<{ browser, page, headers }>}
 */
export async function createAuthedContext(config) {
  const {
    BASE,
    LOGIN_URL,
    TOP_URL,
    USER,
    PASS,
    TIMEOUTS = { action: 2500, nav: 12000 },
    headless = true,
    blockResources = ["image", "font"],
  } = config;

  const browser = await chromium.launch({ headless });
  const page = await browser.newPage();

  page.setDefaultTimeout(TIMEOUTS.action);
  page.setDefaultNavigationTimeout(TIMEOUTS.nav);

  await page.route("**/*", (route) => {
    const t = route.request().resourceType();
    if (blockResources.includes(t)) return route.abort();
    route.continue();
  });

  // 1) 로그인
  await page.goto(LOGIN_URL, { waitUntil: "domcontentloaded", timeout: TIMEOUTS.nav });

  await page.locator("#USERNAME").waitFor({ state: "visible", timeout: 5000 });
  await page.locator("#PASSWORD_T").waitFor({ state: "visible", timeout: 5000 });
  await page.locator('button[name="LoginButton"]').waitFor({ state: "visible", timeout: 5000 });

  await page.locator('select[name="domainname"]').selectOption("localhost").catch(() => {});

  await page.locator("#USERNAME").click();
  await page.locator("#USERNAME").fill("");
  await page.keyboard.type(USER, { delay: 15 });

  await page.locator("#PASSWORD_T").click();
  await page.locator("#PASSWORD_T").fill("");
  await page.keyboard.type(PASS, { delay: 15 });

  await Promise.allSettled([
    page.waitForNavigation({ waitUntil: "commit", timeout: 7000 }),
    page.locator('button[name="LoginButton"]').click({ timeout: 1500 }),
  ]);

  // 2) TOP (컨텍스트 생성)
  await page.goto(TOP_URL, { waitUntil: "domcontentloaded", timeout: TIMEOUTS.nav });
  await sleep(500);

  const headers = { Referer: TOP_URL, Origin: BASE };
  return { browser, page, headers };
}

export async function safeClose(browser) {
  try {
    if (browser) await browser.close();
  } catch {}
}