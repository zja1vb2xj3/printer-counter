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
    const typeRaw = parts[1]?.trim(); // 흑백/컬러
    const count = Number(String(parts[2]).replace(/,/g, "").trim());

    if (!place || !typeRaw || Number.isNaN(count)) continue;

    // 표기 흔들림 방지(예: "컬러", "칼라" 등 들어오면 컬러로 통일하고 싶으면 여기서 정규화)
    let type = typeRaw;
    if (type === "칼라") type = "컬러";

    rows.push({ place, type, count });
  }

  return rows;
}

/**
 * long 형식(장소/구분/출력매수)을 wide 형식(장소/흑백/컬러)으로 변환
 * - 같은 장소에 동일 구분이 여러 번 나오면 합산
 * @param {Array<{place:string, type:string, count:number}>} longRows
 * @returns {Array<{place:string, bw:number|null, color:number|null}>}
 */
export function pivotSummaryRows(longRows) {
  const map = new Map(); // place -> { place, bw, color }

  for (const r of longRows) {
    const key = r.place;
    if (!map.has(key)) {
      map.set(key, { place: key, bw: null, color: null });
    }
    const obj = map.get(key);

    if (r.type === "흑백") {
      obj.bw = (obj.bw ?? 0) + r.count;
    } else if (r.type === "컬러") {
      obj.color = (obj.color ?? 0) + r.count;
    } else {
      // 흑백/컬러 외 구분이 들어오면 일단 무시(원하면 컬럼 추가로 확장 가능)
      // console.warn(`[WARN] Unknown type "${r.type}" for place "${r.place}"`);
    }
  }

  return Array.from(map.values());
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
 * 엑셀 파일에 "장소/흑백/컬러" 형태로 저장 (두번째 사진 형식)
 * - 입력은 기존처럼 "장소\t구분\t출력매수" (long)
 * - 저장은 장소별 피벗(wide)
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

  const longRows = parseSummaryLines(summaryLines);
  const dataRows = pivotSummaryRows(longRows); // <-- 핵심: wide로 변환

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
      { header: "흑백", key: "bw", width: 12 },
      { header: "컬러", key: "color", width: 12 },
    ];

    ws.getRow(1).font = { bold: true };
    ws.views = [{ state: "frozen", ySplit: 1 }];

    // 숫자 서식
    ws.getColumn("bw").numFmt = "#,##0";
    ws.getColumn("color").numFmt = "#,##0";
  } else {
    // append 모드일 때, 시트가 이미 있으면 헤더 여부 확인 후 정리
    const a1 = String(ws.getCell("A1").value ?? "").trim();
    const b1 = String(ws.getCell("B1").value ?? "").trim();
    const c1 = String(ws.getCell("C1").value ?? "").trim();
    const hasHeader = a1 === "장소" && b1 === "흑백" && c1 === "컬러";

    if (!hasHeader) {
      ws.spliceRows(1, 0, ["장소", "흑백", "컬러"]);
      ws.getRow(1).font = { bold: true };
      ws.views = [{ state: "frozen", ySplit: 1 }];
    }

    // 숫자 서식 보장
    ws.getColumn(2).numFmt = "#,##0";
    ws.getColumn(3).numFmt = "#,##0";
  }

  // 데이터 입력 (place, bw, color)
  for (const r of dataRows) {
    ws.addRow(r);
  }

  await wb.xlsx.writeFile(filePath);

  return { filePath, sheetName, rowCount: dataRows.length };
}