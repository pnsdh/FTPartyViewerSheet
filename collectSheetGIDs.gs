/**
 * ===== collectSheetGIDs.gs =====
 * 로컬 스프레드시트의 GID를 수집해서 SheetNames 시트 13행에 기록
 */
function collectSheetGids() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetNamesSheet = ss.getSheetByName(SHEET_COLLECTION_CONFIG.TARGET_SHEET_NAME);
  
  // 현재 스프레드시트의 모든 시트 정보 수집
  const allSheets = ss.getSheets();
  const sheetGidMap = new Map();
  
  allSheets.forEach(sheet => {
    sheetGidMap.set(sheet.getName(), sheet.getSheetId());
  });
  
  // 12행에서 시트명들 가져오기
  const startColumn = 2; // B열부터
  const lastColumn = sheetNamesSheet.getLastColumn();
  const sheetNames = sheetNamesSheet.getRange(12, startColumn, 1, lastColumn - startColumn + 1).getValues()[0];
  
  // 13행에 기록할 GID 배열 준비
  const gidValues = sheetNames.map(sheetName => {
    if (!sheetName || sheetName.toString().trim() === '') {
      return '';
    }
    return sheetGidMap.get(sheetName.toString().trim()) || '';
  });
  
  // 13행에 GID들 기록
  sheetNamesSheet.getRange(13, startColumn, 1, gidValues.length).setValues([gidValues]);
  
  console.log('GID 수집 완료');
}