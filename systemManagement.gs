/**
 * ===== systemManagement.gs =====
 * 시스템 운영 관리 기능
 * 24시간 자동 실행 설정과 트리거 관리를 담당합니다.
 */

/**
 * 트리거 설정
 */
function setupTrigger() {
  logMessage_('=== 시트명 수집 서비스 시작 ===');

  try {
    // 기존 트리거 정리
    const cleanedCount = cleanupTriggers_();
    if (cleanedCount > 0) {
      logMessage_(`기존 트리거 ${cleanedCount}개 정리 완료`);
    }
    
    // 메인 실행 트리거 생성
    const intervalMinutes = SHEET_COLLECTION_CONFIG.MAIN_EXECUTION_INTERVAL_MINUTES;
    ScriptApp.newTrigger('collectSheetNames')
      .timeBased()
      .everyMinutes(intervalMinutes)
      .create();
    logMessage_(`메인 실행 트리거 생성 완료: ${intervalMinutes}분마다 실행`);
    
    // 운영 정보 로깅
    const expectedDailyRuns = Math.round(24 * 60 / intervalMinutes);
    
    logMessage_(`=== 운영 설정 정보 ===`);
    logMessage_(`실행 주기: ${intervalMinutes}분`);
    logMessage_(`예상 일일 실행 횟수: ${expectedDailyRuns}회 (자동)`);
    logMessage_(`배치 크기: ${SHEET_COLLECTION_CONFIG.BATCH_SIZE}개`);
    logMessage_(`재시도 횟수: ${SHEET_COLLECTION_CONFIG.RETRY_ATTEMPTS}회`);
    
    // 시스템 상태 업데이트
    updateSystemStatus_('running', '서비스 시작됨');
    
    logMessage_('서비스 시작 완료');
    
  } catch (error) {
    logMessage_(`서비스 시작 실패: ${error.message}`, 'error');
    updateSystemStatus_('error', `시작 실패: ${error.message}`);
    throw error;
  }
}

/**
 * 트리거 중단
 */
function teardownTrigger() {
  logMessage_('=== 시트명 수집 서비스 중단 ===');
  
  try {
    const deletedCount = cleanupTrigger_();
    logMessage_(`총 ${deletedCount}개 트리거 삭제 완료`);
    
    updateSystemStatus_('stopped', '서비스 중단됨');
    
    logMessage_('서비스 중단 완료');
    
  } catch (error) {
    logMessage_(`서비스 중단 중 오류: ${error.message}`, 'error');
    updateSystemStatus_('error', `중단 중 오류: ${error.message}`);
  }
}

/**
 * ===== 내부 함수들 =====
 */

function cleanupTrigger_() {
  const triggers = ScriptApp.getProjectTriggers();
  let deletedCount = 0;
  
  const managedFunctions = ['collectSheetNames'];
  
  triggers.forEach(trigger => {
    const functionName = trigger.getHandlerFunction();
    if (managedFunctions.includes(functionName)) {
      ScriptApp.deleteTrigger(trigger);
      deletedCount++;
      logMessage_(`트리거 삭제: ${functionName}`);
    }
  });
  
  return deletedCount;
}

function updateSystemStatus_(status, message) {
  try {
    const monitoringSheet = SpreadsheetApp.getActiveSpreadsheet()
                                          .getSheetByName(SHEET_COLLECTION_CONFIG.MONITORING_SHEET_NAME);
                                          
    if (!monitoringSheet) {
      logMessage_('모니터링 시트를 찾을 수 없습니다.', 'warn');
      return;
    }
    
    const currentTime = new Date().toLocaleString('ko-KR');
    const isRunning = status === 'running';
    
    monitoringSheet.getRange('B1').setValue(isRunning);
    monitoringSheet.getRange('B2').setValue(message);
    monitoringSheet.getRange('B3').setValue(currentTime);
    
    logMessage_(`시스템 상태 업데이트: ${status} - ${message}`);
    
  } catch (error) {
    logMessage_(`시스템 상태 업데이트 실패: ${error.message}`, 'warn');
  }
}