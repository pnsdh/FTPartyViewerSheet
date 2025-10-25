/**
 * ===== sheetCollection.gs =====
 * 시트명 수집 메인 기능
 * 외부 스프레드시트들의 시트명을 수집하여 정리합니다.
 */

/**
 * 시트명 수집 메인 함수 (외부 호출용)
 * 설정된 URL들에서 시트명을 수집하고 요약 시트를 생성합니다.
 */
function collectSheetNames() {
  const startTime = new Date();
  logMessage_('=== 시트명 수집 시작 ===');
  
  setExecutionStatus_(true);
  
  let successCount = 0;
  let errorCount = 0;
  const errorDetails = [];
  
  try {
    const targetSheet = getSheetCollectionTargetSheet_();
    const spreadsheetUrls = extractSpreadsheetUrls_(targetSheet);
    
    // 유효한 URL 필터링
    const urlsToProcess = spreadsheetUrls.map((url, index) => ({ url, index }))
                                       .filter(item => item.url && item.url.toString().trim() !== '');
    
    logMessage_(`처리할 항목: ${urlsToProcess.length}개`);
    
    // 배치별로 처리
    const batches = createBatches_(urlsToProcess, SHEET_COLLECTION_CONFIG.BATCH_SIZE);
    
    for (let batchIndex = 0; batchIndex < batches.length; batchIndex++) {
      const batch = batches[batchIndex];
      logMessage_(`배치 ${batchIndex + 1}/${batches.length} 처리 중 (${batch.length}개)`);
      
      const batchResults = processBatch_(targetSheet, batch);
      
      // 결과 집계
      for (const result of batchResults) {
        if (result.success) {
          successCount++;
        } else {
          errorCount++;
          const errorMsg = `${result.columnIndex + 1}번째 열 오류: ${result.error}`;
          errorDetails.push(errorMsg);
        }
      }
      
      // 배치 간 대기
      if (batchIndex < batches.length - 1) {
        Utilities.sleep(SHEET_COLLECTION_CONFIG.INTER_BATCH_DELAY_MS);
      }
    }
    
  } catch (error) {
    logMessage_(`전체 처리 중 치명적 오류: ${error.message}`, 'error');
    errorCount++;
  } finally {
    setExecutionStatus_(false);

    // 마지막 완료 시간 기록
    try {
      const monitoringSheet = SpreadsheetApp.getActiveSpreadsheet()
                                            .getSheetByName(SHEET_COLLECTION_CONFIG.MONITORING_SHEET_NAME);
      if (monitoringSheet) {
        monitoringSheet.getRange('B4').setValue(new Date().toLocaleString('ko-KR'));
      }
    } catch (error) {
      // 무시
    }
  }
  
  // 실행 통계 출력
  const duration = Math.round((new Date() - startTime) / 1000);
  logMessage_('=== 처리 완료 ===');
  logMessage_(`${duration}초, 성공 ${successCount}개, 오류 ${errorCount}개`);
  if (errorCount > 0 && errorDetails.length > 0) {
    logMessage_(`오류 목록: ${errorDetails.join(', ')}`, 'error');
  }
  
  // 요약 시트 생성
  try {
    logMessage_('요약 시트 생성 시작...');
    createSummarySheet();
    logMessage_('요약 시트 생성 완료');
  } catch (summaryError) {
    logMessage_(`요약 시트 생성 실패: ${summaryError.message}`, 'error');
  }  
}
/**
 * ===== 내부 함수들 =====
 */

function processBatch_(sheetNamesSheet, batch) {
  const results = [];
  const excludeStringsMap = {};
  
  // 제외 문자열 사전 준비
  for (const { index } of batch) {
    excludeStringsMap[index] = extractExcludeStringsForColumn_(sheetNamesSheet, index);
  }
  
  for (let i = 0; i < batch.length; i++) {
    const { url, index } = batch[i];
    
    try {
      if (i > 0) {
        Utilities.sleep(SHEET_COLLECTION_CONFIG.INTER_REQUEST_DELAY_MS);
      }
      
      clearColumnData_(sheetNamesSheet, index);
      
      const result = processIndividualSpreadsheet_(
        sheetNamesSheet, 
        index, 
        url, 
        excludeStringsMap[index]
      );
      
      results.push({
        success: true,
        columnIndex: index,
        title: result.title,
        sheetCount: result.sheetCount
      });
      
    } catch (error) {
      logMessage_(`배치 내 오류 (${index + 1}번째 열): ${error.message}`, 'error');
      
      // 오류 기록
      try {
        const actualColumn = index + SHEET_COLLECTION_CONFIG.DATA_START_COLUMN;
        const shortErrorMessage = error.message.substring(0, SHEET_COLLECTION_CONFIG.ERROR_MESSAGE_MAX_LENGTH);
        sheetNamesSheet.getRange(SHEET_COLLECTION_CONFIG.SPREADSHEET_NAME_ROW, actualColumn)
                      .setValue(`오류: ${shortErrorMessage}`);
      } catch (writeError) {
        // 오류 기록 실패는 무시
      }
      
      results.push({
        success: false,
        columnIndex: index,
        error: error.message
      });
    }
  }
  
  return results;
}

function processIndividualSpreadsheet_(targetSheet, columnIndex, url, excludeStrings) {
  const actualColumn = columnIndex + SHEET_COLLECTION_CONFIG.DATA_START_COLUMN;
  
  const spreadsheetId = extractSpreadsheetId_(url.toString());
  if (!spreadsheetId) {
    throw new Error('유효하지 않은 URL');
  }
  
  // 재시도 로직
  let targetSpreadsheet = null;
  let lastError = null;
  
  for (let attempt = 0; attempt < SHEET_COLLECTION_CONFIG.RETRY_ATTEMPTS; attempt++) {
    try {
      targetSpreadsheet = SpreadsheetApp.openById(spreadsheetId);
      break;
    } catch (error) {
      lastError = error;
      if (attempt < SHEET_COLLECTION_CONFIG.RETRY_ATTEMPTS - 1) {
        Utilities.sleep(SHEET_COLLECTION_CONFIG.RETRY_DELAY_MS);
      }
    }
  }
  
  if (!targetSpreadsheet) {
    throw new Error(`접근 실패 (${SHEET_COLLECTION_CONFIG.RETRY_ATTEMPTS}회 시도): ${lastError.message}`);
  }
  
  const spreadsheetTitle = targetSpreadsheet.getName();
  
  // 시트 정보 수집
  let sheetsData = [];
  try {
    const sheets = targetSpreadsheet.getSheets();
    sheetsData = sheets.map(sheet => ({
      name: sheet.getName(),
      gid: sheet.getSheetId(),
      isHidden: sheet.isSheetHidden()
    }));
  } catch (error) {
    logMessage_(`시트 정보 수집 실패: ${error.message}`, 'warn');
    sheetsData = [];
  }
  
  // 데이터 기록
  writeSheetDataToColumn_(targetSheet, actualColumn, { title: spreadsheetTitle, sheets: sheetsData }, excludeStrings);
  
  return {
    title: spreadsheetTitle,
    sheetCount: sheetsData.length
  };
}

function clearColumnData_(sheet, columnIndex) {
  const actualColumn = columnIndex + SHEET_COLLECTION_CONFIG.DATA_START_COLUMN;
  
  const clearRanges = [
    [SHEET_COLLECTION_CONFIG.SPREADSHEET_NAME_ROW, actualColumn, 1, 1],
    [SHEET_COLLECTION_CONFIG.SHEET_NAMES_START_ROW, actualColumn, 
     SHEET_COLLECTION_CONFIG.SHEET_NAMES_END_ROW - SHEET_COLLECTION_CONFIG.SHEET_NAMES_START_ROW + 1, 1],
    [SHEET_COLLECTION_CONFIG.SHEET_GIDS_START_ROW, actualColumn, 
     SHEET_COLLECTION_CONFIG.SHEET_GIDS_END_ROW - SHEET_COLLECTION_CONFIG.SHEET_GIDS_START_ROW + 1, 1]
  ];
  
  clearRanges.forEach(([row, col, numRows, numCols]) => {
    try {
      sheet.getRange(row, col, numRows, numCols).clearContent();
    } catch (error) {
      logMessage_(`데이터 삭제 실패 (행${row}, 열${col}): ${error.message}`, 'warn');
    }
  });
}

function writeSheetDataToColumn_(sheet, column, data, excludeStrings) {
  sheet.getRange(SHEET_COLLECTION_CONFIG.SPREADSHEET_NAME_ROW, column).setValue(data.title);
  
  // 유효한 시트 필터링
  const validSheets = data.sheets.filter(sheetInfo => {
    if (sheetInfo.isHidden) return false;
    
    for (const excludeStr of excludeStrings) {
      if (shouldExcludeSheet_(sheetInfo.name, excludeStr)) {
        return false;
      }
    }
    
    return true;
  });
  
  // 데이터 준비
  const maxSheetRows = SHEET_COLLECTION_CONFIG.SHEET_NAMES_END_ROW - SHEET_COLLECTION_CONFIG.SHEET_NAMES_START_ROW + 1;
  const maxGidRows = SHEET_COLLECTION_CONFIG.SHEET_GIDS_END_ROW - SHEET_COLLECTION_CONFIG.SHEET_GIDS_START_ROW + 1;
  
  const sheetNameValues = [];
  const gidValues = [];
  
  for (let i = 0; i < maxSheetRows; i++) {
    if (i < validSheets.length) {
      sheetNameValues.push([String(validSheets[i].name)]);
      if (i < maxGidRows) {
        gidValues.push([validSheets[i].gid]);
      }
    } else {
      sheetNameValues.push(["-"]);
    }
  }
  
  // 데이터 기록
  try {
    if (sheetNameValues.length > 0) {
      const range = sheet.getRange(SHEET_COLLECTION_CONFIG.SHEET_NAMES_START_ROW, column, sheetNameValues.length, 1);
      range.setValues(sheetNameValues);
      range.setNumberFormat('@');  // ← 텍스트 서식 강제 적용
    }
    
    if (gidValues.length > 0) {
      sheet.getRange(SHEET_COLLECTION_CONFIG.SHEET_GIDS_START_ROW, column, gidValues.length, 1)
           .setValues(gidValues);
    }
  } catch (error) {
    logMessage_(`데이터 기록 실패 (열 ${column}): ${error.message}`, 'error');
  }
}