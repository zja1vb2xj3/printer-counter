// capture.js  (수동 1클릭으로 실제 카운터 호출 URL 캡처)
import { chromium } from "playwright";

const BASE = "http://10.100.1.15:8000";
const LOGIN_URL = `${BASE}/rps/`;

const USER = "Administrator";
const PASS = "7654321";

const browser = await chromium.launch({ headless: false }); // 화면 띄움
const page = await browser.newPage();

page.on("request", (req) => {
  const u = req.url();
  if (u.includes("dcounter.cgi") || u.includes("CorePGTAG=")) {
    console.log("[CAPTURE] request:", u);
  }
});
page.on("response", async (res) => {
  const u = res.url();
  if (u.includes("dcounter.cgi")) {
    console.log("[CAPTURE] response:", res.status(), u);
  }
});

await page.goto(LOGIN_URL, { waitUntil: "domcontentloaded" });

await page.locator("#USERNAME").fill(USER);
await page.locator("#PASSWORD_T").fill(PASS);
await page.locator('button[name="LoginButton"]').click();

// 여기서부터는 네가 브라우저에서 “카운터 메뉴” 1번만 클릭하면 됨.
// 콘솔에 [CAPTURE] 로 URL이 찍힘.
console.log("브라우저에서 카운터 메뉴를 클릭하세요. (URL 캡처 중)");
await page.waitForTimeout(60000);

await browser.close();
