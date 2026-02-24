// countManager.js
import * as cheerio from "cheerio";

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

export function isCannotOpen(html) {
  return (
    html.includes("Remote UI : Cannot open this page.") ||
    html.includes("Cannot display the specified page.") ||
    html.includes("authentication information has been updated") ||
    html.includes("session has expired")
  );
}

function tdVisibleText($, tdEl) {
  const c = $(tdEl).clone();
  c.find("script").remove();
  return c.text().replace(/\s+/g, " ").trim();
 }

export function parseCounterRows(html) {
  const $ = cheerio.load(html);
  const rows = [];

  const trs = $("div.ItemListComponent tbody tr");
  trs.each((_, tr) => {
    const tds = $(tr).find("td");
    const td0 = tds.get(0);
    const td1 = tds.get(1);

    const kindText = tdVisibleText($, td0);
    const valueText = tdVisibleText($, td1);

    const trHtml = $.html(tr);

    const idx =
      trHtml.match(/write_index\("(\d+)"\)/)?.[1] ||
      kindText.match(/^(\d+)\s*:/)?.[1] ||
      trHtml.match(/write_value\("(\d+)"\s*,\s*\d+\)/)?.[1] ||
      null;

    let value = null;
    const mWrite = trHtml.match(/write_value\("\d+"\s*,\s*([0-9,]+)\s*\)/);
    if (mWrite) value = Number(mWrite[1].replace(/,/g, ""));
    if (value === null) {
      const mTxt = valueText.replace(/,/g, "").match(/(\d+)/);
      if (mTxt) value = Number(mTxt[1]);
    }

    rows.push({ idx, kind: kindText, value });
  });

  return rows;
}

/**
 * rows[2] = 흑백, rows[3] = 컬러로 포맷해서 출력 라인 생성
 * @param {string} comment
 * @param {Array<{idx:string|null, kind:string, value:number|null}>} rows
 * @returns {{ ok:boolean, lines:string[], warn?:string }}
 */
export function formatBwColorLines(comment, rows) {
  const bw = rows?.[2];
  const color = rows?.[3];

  // rows가 예상과 다르면: 그래도 뭔가 출력은 되게(경고 + 원본 나열)
  if (!bw || !color) {
    const lines = [];
    lines.push(`${comment}\tWARN\trows length=${rows?.length ?? 0} (expected 4)`);
    for (const r of rows ?? []) {
      lines.push(`${comment}\t${r.kind}\t${r.value ?? ""}`);
    }
    return { ok: false, lines, warn: "unexpected_rows_length" };
  }

  return {
    ok: true,
    lines: [
      `${comment}\t흑백\t${bw.value ?? ""}`,
      `${comment}\t컬러\t${color.value ?? ""}`,
    ],
  };
}

/**
 * @param {import('playwright').Page} page
 * @param {object} config
 * @param {object} headers
 * @param {string} comment  장소명
 */
export async function fetchCounters(page, config, headers, comment = "") {
  const { BASE, TOP_URL, TIMEOUTS = { action: 2500, nav: 12000 } } = config;

  await page
    .goto(TOP_URL, { waitUntil: "domcontentloaded", timeout: TIMEOUTS.nav })
    .catch(() => {});
  await sleep(200);

  const nativetopUrl = `${BASE}/rps/nativetop.cgi?RUIPNxBundle=&CorePGTAG=PGTAG_JOB_PRT_STAT&Dummy=${Date.now()}`;
  const jstatpriUrl  = `${BASE}/rps/jstatpri.cgi?Flag=Init_Data&CorePGTAG=1&FromTopPage=1&Dummy=${Date.now() + 1}`;
  const dcounterUrl  = `${BASE}/rps/dcounter.cgi?CorePGTAG=14&Dummy=${Date.now() + 2}`;

  const r1 = await page.request.get(nativetopUrl, { headers });
  const r2 = await page.request.get(jstatpriUrl,  { headers });
  const r3 = await page.request.get(dcounterUrl,  { headers });

  const html = await r3.text();

  const debug = {
    nativetop: { status: r1.status(), url: nativetopUrl },
    jstatpri:  { status: r2.status(), url: jstatpriUrl },
    dcounter:  { status: r3.status(), url: dcounterUrl },
    title: html.match(/<title>(.*?)<\/title>/i)?.[1] || "N/A",
  };

  if (isCannotOpen(html)) {
    return { ok: false, html, rows: [], lines: [], debug };
  }

  const rows = parseCounterRows(html);
  const formatted = formatBwColorLines(comment, rows);

  return { ok: true, html, rows, lines: formatted.lines, debug };
}