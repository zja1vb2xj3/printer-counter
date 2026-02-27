// main.js
import { createAuthedContext, safeClose } from "./login.js";
import { fetchCounters } from "./countManager.js";
import {
  saveSummaryToExcel,
  getSheetNameYYYYMM,
  parseSummaryLines,
  pivotSummaryRows,
} from "./excelManager.js";

import dotenv from "dotenv";

dotenv.config();

const USER = process.env.PRINTER_USER;
const PASS = process.env.PRINTER_PASS;

const TIMEOUTS = { action: 2500, nav: 12000 };

const printers = [
  { base: "http://10.10.21.11:8000", comment: "SCM자재사무실" },
  { base: "http://10.100.1.11:8000", comment: "관리동 1층 1" },
  { base: "http://10.100.1.12:8000", comment: "관리동 1층 2" },
  { base: "http://10.100.1.21:8000", comment: "경영관리 2층" },
  { base: "http://10.20.21.11:8000", comment: "금형관리실" },
  { base: "http://10.100.1.14:8000", comment: "설계실" },
  { base: "http://10.30.11.220:8000", comment: "스탬핑동" },
  { base: "http://10.100.1.15:8000", comment: "시험측정실" },
  { base: "http://10.10.31.11:8000", comment: "열처리동" },
  { base: "http://10.10.11.11:8000", comment: "사출동" },
];

function buildConfig(base) {
  return {
    BASE: base,
    LOGIN_URL: `${base}/rps/`,
    TOP_URL: `${base}/rps/_top.htm`,
    USER,
    PASS,
    TIMEOUTS,
    headless: true,
  };
}

function formatOneLineSummary(comment, lines) {
  // lines: ["장소\t흑백\t45252", "장소\t컬러\t36860", ...]
  const longRows = parseSummaryLines(lines);
  const wideRows = pivotSummaryRows(longRows);

  // 정상이라면 해당 프린터는 place 1개만 나와야 함(그래도 방어)
  const row = wideRows[0];
  const bw = row?.bw ?? 0;
  const color = row?.color ?? 0;

  // 원하는 콘솔 형태: "관리동 1층 1 28280 113805"
  // (구분문자 최소화, \n, 탭 제거)
  return `${comment} ${bw} ${color}`;
}

async function runOnePrinter(printer) {
  const config = buildConfig(printer.base);

  let browser;
  try {
    const { browser: b, page, headers } = await createAuthedContext(config);
    browser = b;

    const result = await fetchCounters(page, config, headers, printer.comment);

    return { base: printer.base, comment: printer.comment, ...result };
  } catch (e) {
    return {
      base: printer.base,
      comment: printer.comment,
      ok: false,
      error: e?.message || String(e),
      debug: null,
      lines: [],
    };
  } finally {
    await safeClose(browser);
  }
}

const results = [];
for (const p of printers) {
  const r = await runOnePrinter(p);
  results.push(r);

  console.log("====================================");
  console.log(`[${r.comment}] ${r.base}`);

  if (r.debug) {
    console.log("[DEBUG] nativetop:", r.debug.nativetop.status);
    console.log("[DEBUG] jstatpri :", r.debug.jstatpri.status);
    console.log("[DEBUG] dcounter :", r.debug.dcounter.status, r.debug.dcounter.url);
  }

  if (!r.ok) {
    console.log("[RESULT] FAIL");
    if (r.error) console.log("[ERROR]", r.error);
    if (r.debug?.title) console.log("[DEBUG] title:", r.debug.title);

    console.log("프로그램 종료 (수동 재실행 필요)");
    process.exit(1); // 🔴 즉시 종료
  } else {
    console.log("[RESULT] SUCCESS");

    // ✅ 콘솔은 1줄 요약만 출력
    console.log(formatOneLineSummary(r.comment, r.lines));
  }

  console.log("====================================");
}

// SUMMARY 출력 + 엑셀 저장
console.log("\n==== SUMMARY (BW/COLOR) ====");

// 엑셀 저장용: 세로형 원본 라인만 수집
const summaryLines = [];
for (const r of results) {
  if (!r.ok) continue;
  for (const line of r.lines) {
    summaryLines.push(line);
  }
}

// ✅ SUMMARY 콘솔도 동일한 1줄 형식으로 전체 출력(원하면 이 블록 삭제 가능)
const allLong = parseSummaryLines(summaryLines);
const allWide = pivotSummaryRows(allLong);

for (const row of allWide) {
  const bw = row.bw ?? 0;
  const color = row.color ?? 0;
  console.log(`${row.place} ${bw} ${color}`);
}

// 엑셀 저장
const sheetName = getSheetNameYYYYMM(new Date()); // 예: 2026.02
const filePath = "./printer_counters.xlsx";

const saved = await saveSummaryToExcel(summaryLines, { filePath, sheetName, mode: "replace" });
console.log(`[SAVED] ${saved.filePath} (sheet: ${saved.sheetName}, rows: ${saved.rowCount})`);