// excelManager.js
import ExcelJS from "exceljs";
import fs from "fs";

/**
 * "장소\t구분\t출력매수" 라인 파싱
 * @param {string[]} lines
 * @returns {Array<{place:string, type:string, count:number}>}
 */
export function parseSummaryLines(lines) {
  const rows = [];

  for (const line of lines) {
    const parts = String(line).trim().split(/\t+/);
    if (parts.length < 3) continue;

    const place = parts[0]?.trim();
    const type = parts[1]?.trim(); // 흑백/컬러
    const count = Number(String(parts[2]).replace(/,/g, "").trim());

    if (!place || !type || Number.isNaN(count)) continue;
    rows.push({ place, type, count });
  }

  return rows;
}

/**
 * 현재 날짜 기준 "YYYY.MM" 시트명 생성(로컬 시간 기준)
 * @param {Date} [d]
 */
export function getSheetNameYYYYMM(d = new Date()) {
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  return `${yyyy}.${mm}`;
}

/**
 * 엑셀 파일에 "장소/구분/출력매수" 형태로 저장
 * - 파일이 없으면 생성
 * - 동일 시트명이 있으면 기본은 덮어쓰기 (mode="replace")
 * - mode="append"면 기존 시트 아래에 이어쓰기 (헤더는 유지)
 *
 * @param {string[]} summaryLines   예: ["경영관리 2층\t흑백\t50277", ...]
 * @param {object} opts
 * @param {string} opts.filePath    예: "./printer_counters.xlsx"
 * @param {string} opts.sheetName   예: "2026.02"
 * @param {"replace"|"append"} [opts.mode="replace"]
 * @returns {Promise<{filePath:string, sheetName:string, rowCount:number}>}
 */
export async function saveSummaryToExcel(summaryLines, opts) {
  const { filePath, sheetName, mode = "replace" } = opts;

  const dataRows = parseSummaryLines(summaryLines);

  const wb = new ExcelJS.Workbook();
  if (fs.existsSync(filePath)) {
    await wb.xlsx.readFile(filePath);
  }

  let ws = wb.getWorksheet(sheetName);

  if (ws && mode === "replace") {
    wb.removeWorksheet(ws.id);
    ws = null;
  }

  if (!ws) {
    ws = wb.addWorksheet(sheetName);

    ws.columns = [
      { header: "장소", key: "place", width: 22 },
      { header: "구분", key: "type", width: 10 },
      { header: "출력매수", key: "count", width: 12 },
    ];

    ws.getRow(1).font = { bold: true };
    ws.views = [{ state: "frozen", ySplit: 1 }];
    ws.getColumn("count").numFmt = "#,##0";
  } else {
    // append 모드일 때, 시트가 이미 있으면 헤더 여부 확인 후 정리
    // (첫 행이 헤더가 아니면 헤더 추가)
    const a1 = String(ws.getCell("A1").value ?? "").trim();
    const b1 = String(ws.getCell("B1").value ?? "").trim();
    const c1 = String(ws.getCell("C1").value ?? "").trim();
    const hasHeader = a1 === "장소" && b1 === "구분" && c1 === "출력매수";
    if (!hasHeader) {
      ws.spliceRows(1, 0, ["장소", "구분", "출력매수"]);
      ws.getRow(1).font = { bold: true };
      ws.views = [{ state: "frozen", ySplit: 1 }];
    }
    // 숫자 서식 보장
    ws.getColumn(3).numFmt = "#,##0";
  }

  // 데이터 입력
  for (const r of dataRows) {
    ws.addRow(r);
  }

  await wb.xlsx.writeFile(filePath);

  return { filePath, sheetName, rowCount: dataRows.length };
}