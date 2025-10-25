/**
 * ===== coreUtilities.gs =====
 * 핵심 유틸리티 함수들
 * 내부에서만 사용되는 기본 기능들을 제공합니다.
 */

/**
 * ===== 시트 관련 유틸리티 =====
 */

function getSheetCollectionTargetSheet_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(SHEET_COLLECTION_CONFIG.TARGET_SHEET_NAME);
  
  if (!sheet) {
    throw new Error(`${SHEET_COLLECTION_CONFIG.TARGET_SHEET_NAME} 시트를 찾을 수 없습니다.`);
  }
  
  return sheet;
}

function getSummarySourceSheet_(spreadsheet) {
  const sheet = spreadsheet.getSheetByName(SUMMARY_CONFIG.SOURCE_SHEET_NAME);
  if (!sheet) {
    throw new Error(`${SUMMARY_CONFIG.SOURCE_SHEET_NAME} 시트를 찾을 수 없습니다.`);
  }
  return sheet;
}

function getSummaryResultSheet_(spreadsheet) {
  const sheet = spreadsheet.getSheetByName(SUMMARY_CONFIG.SUMMARY_SHEET_NAME);
  if (!sheet) {
    throw new Error(`${SUMMARY_CONFIG.SUMMARY_SHEET_NAME} 시트를 찾을 수 없습니다. 먼저 시트를 생성하고 헤더를 설정해주세요.`);
  }
  return sheet;
}

function setExecutionStatus_(isRunning) {
  try {
    const monitoringSheet = SpreadsheetApp.getActiveSpreadsheet()
                                          .getSheetByName(SHEET_COLLECTION_CONFIG.MONITORING_SHEET_NAME);
    if (monitoringSheet) {
      monitoringSheet.getRange('B1').setValue(isRunning);
    }
  } catch (error) {
    logMessage_(`실행 상태 설정 실패: ${error.message}`, 'warn');
  }
}

/**
 * ===== 데이터 추출 유틸리티 =====
 */

function extractSpreadsheetUrls_(sheet) {
  const urlRange = sheet.getRange(
    SHEET_COLLECTION_CONFIG.URL_ROW, 
    SHEET_COLLECTION_CONFIG.DATA_START_COLUMN, 
    1, 
    sheet.getLastColumn() - SHEET_COLLECTION_CONFIG.DATA_START_COLUMN + 1
  );
  return urlRange.getValues()[0];
}

function extractSpreadsheetId_(url) {
  if (!url || typeof url !== 'string') {
    return null;
  }
  
  const urlString = url.toString().trim();
  
  for (const pattern of SYSTEM_CONFIG.SPREADSHEET_ID_PATTERNS) {
    const match = urlString.match(pattern);
    if (match && match[1] && match[1].length >= SYSTEM_CONFIG.VALIDATION.MIN_SPREADSHEET_ID_LENGTH) {
      return match[1];
    }
  }
  
  return null;
}

function extractExcludeStringsForColumn_(sheet, columnIndex) {
  const actualColumn = columnIndex + SHEET_COLLECTION_CONFIG.DATA_START_COLUMN;
  const excludeStrings = [];
  
  try {
    const excludeRange = sheet.getRange(
      SHEET_COLLECTION_CONFIG.EXCLUDE_STRINGS_START_ROW,
      actualColumn,
      SHEET_COLLECTION_CONFIG.EXCLUDE_STRINGS_END_ROW - SHEET_COLLECTION_CONFIG.EXCLUDE_STRINGS_START_ROW + 1,
      1
    );
    const excludeValues = excludeRange.getValues();
    
    for (const [value] of excludeValues) {
      if (isValidStringValue_(value)) {
        excludeStrings.push(value.toString().trim());
      }
    }
  } catch (error) {
    logMessage_(`제외 문자열 추출 실패 (열 ${columnIndex + 1}): ${error.message}`, 'warn');
  }
  
  return excludeStrings;
}

/**
 * ===== 배치 처리 유틸리티 =====
 */

function createBatches_(items, batchSize) {
  const batches = [];
  
  for (let i = 0; i < items.length; i += batchSize) {
    batches.push(items.slice(i, i + batchSize));
  }
  
  return batches;
}

/**
 * ===== 데이터 검증 유틸리티 =====
 */

function isValidStringValue_(value) {
  if (!value) return false;
  
  const stringValue = value.toString().trim();
  return stringValue !== '' && !SYSTEM_CONFIG.VALIDATION.EMPTY_VALUE_INDICATORS.includes(stringValue);
}

/**
 * ===== 로깅 유틸리티 =====
 */

function logMessage_(message, level = 'info') {
  if (SHEET_COLLECTION_CONFIG.ENABLE_DETAILED_LOGGING) {
    const timestamp = new Date().toLocaleTimeString();
    const formattedMessage = `[${timestamp}] ${message}`;
    
    switch (level) {
      case 'error':
        console.error(formattedMessage);
        break;
      case 'warn':
        console.warn(formattedMessage);
        break;
      default:
        console.log(formattedMessage);
        break;
    }
  } else if (level === 'error') {
    console.error(message);
  }
}

/**
 * 시트를 제외해야 하는지 판단합니다
 * @param {string} sheetName - 시트 이름
 * @param {string} excludeStr - 제외 문자열 (접두사 포함 가능)
 * @returns {boolean} 제외해야 하면 true
 */
function shouldExcludeSheet_(sheetName, excludeStr) {
  const exactMatchPrefix = SHEET_COLLECTION_CONFIG.EXCLUDE_STRING_EXACT_MATCH_PREFIX;
  
  // 접두사로 시작하는지 확인
  if (excludeStr.startsWith(exactMatchPrefix)) {
    // 접두사 제거하고 정확히 일치하는지 확인 (대소문자 무시)
    const exactPattern = excludeStr.substring(exactMatchPrefix.length);
    return sheetName.toLowerCase() === exactPattern.toLowerCase();
  }
  
  // 접두사가 없으면 포함 여부로 확인 (대소문자 무시)
  return sheetName.toLowerCase().includes(excludeStr.toLowerCase());
}