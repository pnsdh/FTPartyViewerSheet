/**
 * ===== summarySheet.gs =====
 * 요약 시트 생성 메인 기능
 * 수집된 시트명들을 분석하여 정리된 요약 시트를 생성합니다.
 */

/**
 * 요약 시트 생성 메인 함수 (외부 호출용)
 */
function createSummarySheet() {
  try {
    logMessage_('=== 시트명 분석 시작 ===');
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = getSummarySourceSheet_(spreadsheet);
    const summarySheet = getSummaryResultSheet_(spreadsheet);
    
    // 데이터 수집 및 분석
    logMessage_('데이터 수집 중...');
    const analyzedSheetData = collectAndAnalyzeSheetData_(sourceSheet);
    
    // 데이터 정렬
    logMessage_('데이터 정렬 중...');
    const sortedData = sortAnalyzedSheetData_(analyzedSheetData);
    
    // 중복 제거 (설정에 따라)
    logMessage_('중복 제거 중...');
    const finalData = SUMMARY_CONFIG.ENABLE_DUPLICATE_REMOVAL ? 
                      removeDuplicateSheetNames_(sortedData) : sortedData;
    
    // 기존 데이터 삭제 후 새 데이터 기록
    logMessage_('기존 데이터 삭제 및 새 데이터 기록 중...');
    clearSummarySheetData_(summarySheet);
    writeSummaryDataToSheet_(summarySheet, finalData);

    // 필터 적용
    applySummarySheetFilters_(summarySheet);
    
    logMessage_(`=== 분석 완료: ${finalData.length}개 시트명 처리 ===`);
    
  } catch (error) {
    logMessage_(`요약 시트 생성 오류: ${error.message}`, 'error');
    throw error;
  }
}

/**
 * ===== 내부 함수들 =====
 */

function collectAndAnalyzeSheetData_(sourceSheet) {
  const maxColumn = sourceSheet.getLastColumn() - SUMMARY_CONFIG.SOURCE_DATA_START_COLUMN + 1;
  
  // 기본 데이터 추출
  const urls = sourceSheet.getRange(SUMMARY_CONFIG.SOURCE_URL_ROW, SUMMARY_CONFIG.SOURCE_DATA_START_COLUMN, 1, maxColumn).getValues()[0];
  const spreadsheetNames = sourceSheet.getRange(SUMMARY_CONFIG.SOURCE_SPREADSHEET_NAME_ROW, SUMMARY_CONFIG.SOURCE_DATA_START_COLUMN, 1, maxColumn).getValues()[0];
  const typeNames = sourceSheet.getRange(SUMMARY_CONFIG.SOURCE_TYPE_NAME_ROW, SUMMARY_CONFIG.SOURCE_DATA_START_COLUMN, 1, maxColumn).getValues()[0];
  const sheetNamesData = sourceSheet.getRange(SUMMARY_CONFIG.SOURCE_SHEET_NAMES_START_ROW, SUMMARY_CONFIG.SOURCE_DATA_START_COLUMN, SUMMARY_CONFIG.SOURCE_SHEET_NAMES_END_ROW - SUMMARY_CONFIG.SOURCE_SHEET_NAMES_START_ROW + 1, maxColumn).getValues();
  const gidData = sourceSheet.getRange(SUMMARY_CONFIG.SOURCE_SHEET_GIDS_START_ROW, SUMMARY_CONFIG.SOURCE_DATA_START_COLUMN, SUMMARY_CONFIG.SOURCE_SHEET_GIDS_END_ROW - SUMMARY_CONFIG.SOURCE_SHEET_GIDS_START_ROW + 1, maxColumn).getValues();
  
  // 참조 월 추정
  const allSheetNames = [];
  for (let col = 0; col < spreadsheetNames.length; col++) {
    if (!isValidStringValue_(spreadsheetNames[col])) continue;
    for (let row = 0; row < sheetNamesData.length; row++) {
      const sheetName = sheetNamesData[row][col];
      if (isValidStringValue_(sheetName)) {
        allSheetNames.push(sheetName.toString());
      }
    }
  }
  const referenceMonth = estimateReferenceMonth_(allSheetNames);
  
  // 개별 시트 분석
  const analyzedSheets = [];
  for (let col = 0; col < spreadsheetNames.length; col++) {
    const url = urls[col];
    const spreadsheetName = spreadsheetNames[col];
    const typeName = typeNames[col];
    
    if (!isValidStringValue_(spreadsheetName)) continue;
    
    for (let row = 0; row < sheetNamesData.length; row++) {
      const sheetName = sheetNamesData[row][col];
      if (!isValidStringValue_(sheetName)) continue;
      
      const gid = row < gidData.length ? gidData[row][col] : '';
      
      const analyzedData = analyzeIndividualSheetName_(
        sheetName.toString(), 
        spreadsheetName.toString(), 
        referenceMonth, 
        typeName ? typeName.toString() : ''
      );
      
      analyzedData.url = url ? url.toString() : '';
      analyzedData.gid = gid ? gid.toString() : '';
      analyzedData.sheetNumber = row + 1;
      
      analyzedSheets.push(analyzedData);
    }
  }
  
  return analyzedSheets;
}

function sortAnalyzedSheetData_(data) {
  return data.sort((a, b) => {
    // 1. 진도 우선순위
    const progressOrderA = SUMMARY_CONFIG.PROGRESS_SORT_ORDER.indexOf(a.progress);
    const progressOrderB = SUMMARY_CONFIG.PROGRESS_SORT_ORDER.indexOf(b.progress);
    if (progressOrderA !== progressOrderB) {
      return progressOrderA - progressOrderB;
    }
    
    // 2. 날짜 정렬
    if (a.date && b.date) {
      const dateA = new Date(`${SUMMARY_CONFIG.CURRENT_YEAR}/${a.date}`);
      const dateB = new Date(`${SUMMARY_CONFIG.CURRENT_YEAR}/${b.date}`);
      if (dateA.getTime() !== dateB.getTime()) {
        return dateA - dateB;
      }
    } else if (a.date && !b.date) return -1;
    else if (!a.date && b.date) return 1;
    
    // 3. 시간 정렬
    if (a.time && b.time) {
      const timeCompare = a.time.localeCompare(b.time);
      if (timeCompare !== 0) return timeCompare;
    } else if (a.time && !b.time) return -1;
    else if (!a.time && b.time) return 1;
    
    // 4. 시트명 정렬
    return a.sheetName.localeCompare(b.sheetName);
  });
}

function removeDuplicateSheetNames_(data) {
  const seen = new Set();
  const uniqueData = [];
  
  for (const item of data) {
    const key = item.sheetName.trim();
    if (!seen.has(key)) {
      seen.add(key);
      uniqueData.push(item);
    } else {
      logMessage_(`중복 제거: "${key}" (스프레드시트: ${item.spreadsheetName})`);
    }
  }
  
  return uniqueData;
}

function clearSummarySheetData_(summarySheet) {
  const lastRow = summarySheet.getLastRow();
  if (lastRow < SUMMARY_CONFIG.SUMMARY_DATA_START_ROW) return;
  
  const rowsToClean = lastRow - SUMMARY_CONFIG.SUMMARY_DATA_START_ROW + 1;
  
  // 설정에서 정의된 자동 생성 열들만 삭제 (수동 입력 열은 보존)
  SUMMARY_CONFIG.AUTO_GENERATED_COLUMNS.forEach(columnKey => {
    const columnIndex = SUMMARY_CONFIG.SUMMARY_COLUMNS[columnKey];
    try {
      summarySheet.getRange(SUMMARY_CONFIG.SUMMARY_DATA_START_ROW, columnIndex, rowsToClean, 1).clearContent();
    } catch (error) {
      logMessage_(`${columnIndex}열 삭제 중 오류: ${error.message}`, 'warn');
    }
  });
}

function writeSummaryDataToSheet_(summarySheet, data) {
  if (data.length === 0) {
    const firstColumn = Math.min(...Object.values(SUMMARY_CONFIG.SUMMARY_COLUMNS));
    summarySheet.getRange(SUMMARY_CONFIG.SUMMARY_DATA_START_ROW, firstColumn).setValue('분석할 데이터가 없습니다.');
    return;
  }
  
  // 설정에서 정의된 자동 생성 열들만 기록
  SUMMARY_CONFIG.AUTO_GENERATED_COLUMNS.forEach(columnKey => {
    const columnIndex = SUMMARY_CONFIG.SUMMARY_COLUMNS[columnKey];
    let values;
    
    // 열별 데이터 매핑
    switch (columnKey) {
      case 'PROGRESS':
        values = data.map(item => [item.progress]);
        break;
      case 'DATE':
        values = data.map(item => [item.date]);
        break;
      case 'DAY_OF_WEEK':
        values = data.map(item => [item.dayOfWeek]);
        break;
      case 'TIME':
        values = data.map(item => [item.time]);
        break;
      case 'LOCAL_SHEET_NAME':
        values = data.map(item => [item.spreadsheetName]); // 타입명(별칭) 또는 스프레드시트명
        break;
      case 'ORIGINAL_SHEET_NAME':
        values = data.map(item => [String(item.sheetName)]);
        break;
      case 'ORIGINAL_SPREADSHEET_URL':
        values = data.map(item => [item.url]);
        break;
      case 'ORIGINAL_SHEET_GID':
        values = data.map(item => [item.gid]);
        break;
      case 'LOCAL_SHEET_ITEM_NUMBER':
        values = data.map(item => [item.sheetNumber]);
        break;
      default:
        return; // 정의되지 않은 열은 건너뛰기
    }
    
    // 데이터 기록
    try {
      const range = summarySheet.getRange(SUMMARY_CONFIG.SUMMARY_DATA_START_ROW, columnIndex, values.length, 1);
      range.setValues(values);
      
      // ORIGINAL_SHEET_NAME 열은 텍스트 서식 강제 적용
      if (columnKey === 'ORIGINAL_SHEET_NAME') {
        range.setNumberFormat('@');
      }
    } catch (error) {
      logMessage_(`${columnIndex}열 기록 중 오류: ${error.message}`, 'warn');
    }
  });
}

function applySummarySheetFilters_(summarySheet) {
  try {
    const lastRow = summarySheet.getLastRow();
    if (lastRow < SUMMARY_CONFIG.SUMMARY_DATA_START_ROW) {
      logMessage_('필터링할 데이터가 없습니다.');
      return;
    }
    
    // 필터 기준 날짜 계산
    const today = new Date();
    const cutoffDays = SUMMARY_CONFIG.FILTER_SETTINGS.HIDE_OLD_DAYS;
    const cutoffDate = new Date(today.getTime() - (cutoffDays * 24 * 60 * 60 * 1000));
    
    // 필터 조건 분석 및 로깅
    const filterStats = analyzeFilterConditions_(summarySheet, lastRow, cutoffDate);
    
    // 필터 적용
    const dataRange = summarySheet.getRange(1, 1, lastRow, summarySheet.getLastColumn());
    
    let filter = dataRange.getFilter();
    if (!filter) {
      filter = dataRange.createFilter();
    }
    
    // 인원수 필터 (48명 숨김)
    if (SUMMARY_CONFIG.SUMMARY_COLUMNS.MEMBER_COUNT) {
      const memberCountCriteria = SpreadsheetApp.newFilterCriteria()
        .whenTextDoesNotContain(SUMMARY_CONFIG.FILTER_SETTINGS.HIDE_MEMBER_COUNT_PATTERN)
        .build();
      filter.setColumnFilterCriteria(SUMMARY_CONFIG.SUMMARY_COLUMNS.MEMBER_COUNT, memberCountCriteria);
    }
    
    // 날짜 필터 (오래된 날짜 숨김)
    const year = cutoffDate.getFullYear();
    const month = cutoffDate.getMonth() + 1;
    const day = cutoffDate.getDate();
    const dateCriteria = SpreadsheetApp.newFilterCriteria()
      .whenFormulaSatisfied(`=OR(B:B="", B:B>DATE(${year},${month},${day}))`)
      .build();
    filter.setColumnFilterCriteria(SUMMARY_CONFIG.SUMMARY_COLUMNS.DATE, dateCriteria);
    
    // 상태 필터 (빈 값 숨김)
    if (SUMMARY_CONFIG.SUMMARY_COLUMNS.ORIGINAL_SHEET_LINK && SUMMARY_CONFIG.FILTER_SETTINGS.HIDE_EMPTY_STATUS) {
      const statusCriteria = SpreadsheetApp.newFilterCriteria()
        .whenFormulaSatisfied('=M:M<>""')
        .build();
      filter.setColumnFilterCriteria(SUMMARY_CONFIG.SUMMARY_COLUMNS.ORIGINAL_SHEET_LINK, statusCriteria);
    }
    
    // 필터링 결과 로깅
    logMessage_('=== 필터링 완료 ===');
    
    if (filterStats.memberCountFilterCount > 0) {
      logMessage_(`인원수 필터: ${filterStats.memberCountFilterCount}개 행 (${SUMMARY_CONFIG.FILTER_SETTINGS.HIDE_MEMBER_COUNT_PATTERN}명 포함)`);
    }
    
    if (filterStats.dateFilterCount > 0) {
      logMessage_(`날짜 필터: ${filterStats.dateFilterCount}개 행 (${SUMMARY_CONFIG.FILTER_SETTINGS.HIDE_OLD_DAYS}일 초과)`);
    }
    
    if (filterStats.statusFilterCount > 0) {
      logMessage_(`상태 필터: ${filterStats.statusFilterCount}개 행 (빈 값)`);
    }
    
    if (filterStats.memberCountFilterCount === 0 && filterStats.dateFilterCount === 0 && filterStats.statusFilterCount === 0) {
      logMessage_('필터 조건에 해당하는 항목이 없습니다.');
    }
    
  } catch (error) {
    logMessage_(`필터링 중 오류: ${error.message}`, 'error');
  }
}

function analyzeFilterConditions_(summarySheet, lastRow, cutoffDate) {
  const dataRowCount = lastRow - SUMMARY_CONFIG.SUMMARY_DATA_START_ROW + 1;
  
  // 각 열의 데이터 가져오기
  const memberCountData = SUMMARY_CONFIG.SUMMARY_COLUMNS.MEMBER_COUNT ? 
    summarySheet.getRange(SUMMARY_CONFIG.SUMMARY_DATA_START_ROW, SUMMARY_CONFIG.SUMMARY_COLUMNS.MEMBER_COUNT, dataRowCount, 1).getValues() :
    Array(dataRowCount).fill(['']);
    
  const dateData = summarySheet.getRange(SUMMARY_CONFIG.SUMMARY_DATA_START_ROW, SUMMARY_CONFIG.SUMMARY_COLUMNS.DATE, dataRowCount, 1).getValues();
  const sheetNameData = summarySheet.getRange(SUMMARY_CONFIG.SUMMARY_DATA_START_ROW, SUMMARY_CONFIG.SUMMARY_COLUMNS.ORIGINAL_SHEET_NAME, dataRowCount, 1).getValues();
  
  const statusData = SUMMARY_CONFIG.SUMMARY_COLUMNS.ORIGINAL_SHEET_LINK ?
    summarySheet.getRange(SUMMARY_CONFIG.SUMMARY_DATA_START_ROW, SUMMARY_CONFIG.SUMMARY_COLUMNS.ORIGINAL_SHEET_LINK, dataRowCount, 1).getValues() :
    Array(dataRowCount).fill(['']);
  
  let memberCountFilterCount = 0;
  let dateFilterCount = 0;
  let statusFilterCount = 0;
  
  // 각 행별로 필터 조건 확인
  for (let i = 0; i < dataRowCount; i++) {
    const rowNumber = SUMMARY_CONFIG.SUMMARY_DATA_START_ROW + i;
    const filterReasons = [];
    
    // 인원수 필터 확인 (개별 로그 출력)
    const memberCount = memberCountData[i][0];
    const hidePattern = SUMMARY_CONFIG.FILTER_SETTINGS.HIDE_MEMBER_COUNT_PATTERN;
    if (memberCount && memberCount.toString().includes(hidePattern)) {
      filterReasons.push(`인원수: ${memberCount}`);
      memberCountFilterCount++;
    }
    
    // 날짜 필터 확인 (개별 로그 출력)
    const dateValue = dateData[i][0];
    if (dateValue instanceof Date && dateValue < cutoffDate) {
      filterReasons.push(`날짜: ${dateValue.toLocaleDateString('ko-KR')}`);
      dateFilterCount++;
    }
    
    // 상태 필터 확인 (카운트만 하고 로그는 제외)
    if (SUMMARY_CONFIG.FILTER_SETTINGS.HIDE_EMPTY_STATUS) {
      const statusValue = statusData[i][0];
      if (!statusValue || statusValue.toString().trim() === '') {
        statusFilterCount++;
      }
    }
    
    // 필터 조건에 걸린 행만 로깅 (인원수, 날짜 필터 대상)
    if (filterReasons.length > 0) {
      const sheetName = sheetNameData[i][0];
      logMessage_(`[필터 대상] 행 ${rowNumber}: "${sheetName}" | ${filterReasons.join(', ')}`);
    }
  }
  
  return {
    memberCountFilterCount,
    dateFilterCount,
    statusFilterCount
  };
}