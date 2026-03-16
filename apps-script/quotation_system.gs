/**
 * 홈페이지 제작 견적서 발행/이력관리 스크립트
 *
 * 전제:
 * - 원본 견적서는 Google Spreadsheet 1개를 템플릿으로 사용
 * - 발행 시 템플릿 파일을 통째로 복사하여 고객사별 개별 링크를 발급
 * - 이력은 관리자 Spreadsheet의 HISTORY_SHEET_NAME에 기록
 */

const CONFIG = {
  // 관리자 스프레드시트(이 스크립트가 바인딩된 문서)에서 사용하는 시트명
  REQUEST_SHEET_NAME: '견적요청',
  HISTORY_SHEET_NAME: '견적이력',

  // 견적서 템플릿 파일 ID (질문에 첨부된 시트 ID로 초기값 세팅)
  TEMPLATE_SPREADSHEET_ID: '1dfeo-MDrKdmzAh3SCJ1xI7_7bjjsgYci',

  // 견적 사본이 저장될 Google Drive 폴더 ID
  // 비워두면 템플릿과 같은 위치에 생성
  OUTPUT_FOLDER_ID: '',

  // 고객에게 보여줄 시트 gid (필요 시 수정)
  // 질문에 첨부된 gid로 초기값 세팅
  TARGET_GID: '902297916',

  // 생성 파일명 규칙
  FILE_NAME_PREFIX: '[견적서]',
};

/**
 * 관리자 메뉴 생성
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('견적서 자동화')
    .addItem('선택 행 견적서 발행', 'issueQuotationFromActiveRow')
    .addItem('요청 시트 전체 일괄 발행', 'issueQuotationsBatch')
    .addToUi();
}

/**
 * 활성 행 1건 발행
 * REQUEST_SHEET_NAME 기준 컬럼
 * A: 고객사명
 * B: 담당자(선택)
 * C: 메모(선택)
 */
function issueQuotationFromActiveRow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reqSheet = ss.getSheetByName(CONFIG.REQUEST_SHEET_NAME);
  if (!reqSheet) throw new Error(`시트가 없습니다: ${CONFIG.REQUEST_SHEET_NAME}`);

  const row = reqSheet.getActiveCell().getRow();
  if (row < 2) throw new Error('헤더 아래의 데이터 행을 선택하세요.');

  const customerName = String(reqSheet.getRange(row, 1).getValue()).trim();
  const manager = String(reqSheet.getRange(row, 2).getValue()).trim();
  const memo = String(reqSheet.getRange(row, 3).getValue()).trim();

  if (!customerName) throw new Error('고객사명이 비어 있습니다.');

  const result = issueQuotation({ customerName, manager, memo });

  // 요청 시트에 즉시 결과 반영 (D~G)
  reqSheet.getRange(row, 4).setValue(result.token);
  reqSheet.getRange(row, 5).setValue(result.webUrl);
  reqSheet.getRange(row, 6).setValue(result.pdfUrl);
  reqSheet.getRange(row, 7).setValue(result.createdAt);

  SpreadsheetApp.getUi().alert(`견적서 발행 완료\n고객사: ${customerName}`);
}

/**
 * 요청 시트 전체 일괄 발행
 * token(D열)이 비어 있는 행만 처리
 */
function issueQuotationsBatch() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reqSheet = ss.getSheetByName(CONFIG.REQUEST_SHEET_NAME);
  if (!reqSheet) throw new Error(`시트가 없습니다: ${CONFIG.REQUEST_SHEET_NAME}`);

  const lastRow = reqSheet.getLastRow();
  if (lastRow < 2) return;

  const rows = reqSheet.getRange(2, 1, lastRow - 1, 7).getValues();

  rows.forEach((r, idx) => {
    const rowNum = idx + 2;
    const customerName = String(r[0]).trim();
    const manager = String(r[1]).trim();
    const memo = String(r[2]).trim();
    const existingToken = String(r[3]).trim();

    if (!customerName || existingToken) return;

    const result = issueQuotation({ customerName, manager, memo });
    reqSheet.getRange(rowNum, 4).setValue(result.token);
    reqSheet.getRange(rowNum, 5).setValue(result.webUrl);
    reqSheet.getRange(rowNum, 6).setValue(result.pdfUrl);
    reqSheet.getRange(rowNum, 7).setValue(result.createdAt);
  });

  SpreadsheetApp.getUi().alert('일괄 발행이 완료되었습니다.');
}

/**
 * 단건 견적서 발행 핵심 함수
 */
function issueQuotation({ customerName, manager, memo }) {
  const now = new Date();
  const token = generateToken(customerName, now);
  const createdAt = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

  // 1) 템플릿 스프레드시트 파일 복사
  const templateFile = DriveApp.getFileById(CONFIG.TEMPLATE_SPREADSHEET_ID);
  const outputName = `${CONFIG.FILE_NAME_PREFIX} ${customerName} ${Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd-HHmm')}`;

  let copiedFile;
  if (CONFIG.OUTPUT_FOLDER_ID) {
    const folder = DriveApp.getFolderById(CONFIG.OUTPUT_FOLDER_ID);
    copiedFile = templateFile.makeCopy(outputName, folder);
  } else {
    copiedFile = templateFile.makeCopy(outputName);
  }

  const copiedSpreadsheetId = copiedFile.getId();

  // 2) 공유 권한 설정 (링크가 있는 모든 사용자 읽기)
  copiedFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // 3) 고객 발송 URL / PDF URL 생성
  const webUrl = `https://docs.google.com/spreadsheets/d/${copiedSpreadsheetId}/edit#gid=${CONFIG.TARGET_GID}`;
  const pdfUrl = buildPdfExportUrl(copiedSpreadsheetId, CONFIG.TARGET_GID);

  // 4) 이력 기록
  appendHistoryRow({
    createdAt,
    customerName,
    manager,
    memo,
    token,
    spreadsheetId: copiedSpreadsheetId,
    webUrl,
    pdfUrl,
  });

  return { token, webUrl, pdfUrl, createdAt };
}

/**
 * 이력 시트에 기록
 */
function appendHistoryRow(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let historySheet = ss.getSheetByName(CONFIG.HISTORY_SHEET_NAME);

  if (!historySheet) {
    historySheet = ss.insertSheet(CONFIG.HISTORY_SHEET_NAME);
    historySheet
      .getRange(1, 1, 1, 9)
      .setValues([['발행일시', '고객사명', '담당자', '메모', '토큰', '스프레드시트ID', '웹URL', 'PDF URL', '생성자']]);
  }

  historySheet.appendRow([
    data.createdAt,
    data.customerName,
    data.manager,
    data.memo,
    data.token,
    data.spreadsheetId,
    data.webUrl,
    data.pdfUrl,
    Session.getActiveUser().getEmail() || '',
  ]);
}

/**
 * PDF Export URL 생성
 */
function buildPdfExportUrl(spreadsheetId, gid) {
  const params = [
    'format=pdf',
    'portrait=true',
    'size=A4',
    'fitw=true',
    'sheetnames=false',
    'printtitle=false',
    'pagenumbers=false',
    'gridlines=false',
    'fzr=false',
    `gid=${encodeURIComponent(gid)}`,
  ].join('&');

  return `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?${params}`;
}

/**
 * 토큰 생성 (고객사 + 타임스탬프 기반)
 */
function generateToken(customerName, date) {
  const normalized = customerName
    .toLowerCase()
    .replace(/\s+/g, '-')
    .replace(/[^a-z0-9\-가-힣]/g, '')
    .slice(0, 20);

  const ts = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyyMMddHHmmss');
  return `${normalized}-${ts}`;
}
