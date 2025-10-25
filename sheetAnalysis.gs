/**
 * ===== sheetAnalysis.gs =====
 * 시트명 분석 로직 (내부 함수들)
 * 시트명에서 진도, 날짜, 시간 등의 정보를 추출하고 분석합니다.
 */

function analyzeIndividualSheetName_(sheetName, spreadsheetName, referenceMonth, typeName) {
  const progress = analyzeProgress_(sheetName);
  const dateInfo = analyzeDate_(sheetName, referenceMonth);
  const time = analyzeTime_(sheetName);
  
  // 즉시 출발 키워드 처리
  const immediateResult = handleInstantRun_(sheetName, dateInfo, time);
  
  return {
    progress: progress,
    date: immediateResult.date,
    dayOfWeek: immediateResult.dayOfWeek,
    time: immediateResult.time,
    spreadsheetName: typeName || spreadsheetName,
    sheetName: sheetName
  };
}

/**
 * 즉시 출발 키워드를 처리하고 최종 날짜/시간을 반환합니다
 */
function handleInstantRun_(sheetName, dateInfo, parsedTime) {
  const settings = SUMMARY_CONFIG.INSTANT_RUN_SETTINGS;
  
  // 기능 비활성화 또는 키워드 없음 → 원본 그대로
  if (!settings.ENABLED || !hasInstantRunKeyword_(sheetName, settings.KEYWORDS)) {
    return {
      date: dateInfo.date,
      dayOfWeek: dateInfo.dayOfWeek,
      time: parsedTime
    };
  }  
  const now = new Date();
  
  // 날짜가 파싱된 경우
  if (dateInfo.date) {
    const isToday = isSameDateAsToday_(dateInfo.date, now);
    
    if (isToday) {
      // 오늘이면 현재 시간 사용
      logMessage_(`즉시 출발 키워드 감지: "${sheetName}": 오늘 날짜 감지 - 현재 시간 사용`);
      return {
        date: dateInfo.date,
        dayOfWeek: dateInfo.dayOfWeek,
        time: formatCurrentTime_(now)
      };
    } else {
      // 오늘이 아니면 시간 공란
      logMessage_(`즉시 출발 키워드 감지: "${sheetName}": 날짜가 오늘이 아님 - 시간 공란`);
      return {
        date: dateInfo.date,
        dayOfWeek: dateInfo.dayOfWeek,
        time: ''
      };
    }
  }
  
  // 날짜가 없으면 (시간 유무 관계없이) 현재 날짜/시간 사용
  return {
    date: formatCurrentDate_(now),
    dayOfWeek: SYSTEM_CONFIG.WEEKDAYS[now.getDay()],
    time: formatCurrentTime_(now)
  };
}

/**
 * 즉시 출발 키워드가 있는지 확인
 */
function hasInstantRunKeyword_(sheetName, keywords) {
  return keywords.some(keyword => sheetName.includes(keyword));
}

/**
 * 파싱된 날짜가 오늘인지 확인
 */
function isSameDateAsToday_(parsedDate, now) {
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const [month, day] = parsedDate.split('/').map(Number);
  const parsedDateObj = new Date(now.getFullYear(), month - 1, day);
  
  return parsedDateObj.getTime() === today.getTime();
}

/**
 * 현재 날짜를 MM/DD 형식으로 반환
 */
function formatCurrentDate_(date) {
  const month = (date.getMonth() + 1).toString().padStart(2, '0');
  const day = date.getDate().toString().padStart(2, '0');
  return `${month}/${day}`;
}

/**
 * 현재 시간을 HH:MM 형식으로 반환
 */
function formatCurrentTime_(date) {
  const hour = date.getHours().toString().padStart(2, '0');
  const minute = date.getMinutes().toString().padStart(2, '0');
  return `${hour}:${minute}`;
}

/**
 * 시트명에서 진도 정보를 분석합니다
 * 설정된 패턴을 우선순위에 따라 매칭하여 진도를 결정합니다
 */
function analyzeProgress_(sheetName) {
  // 우선순위 순으로 패턴 검사 (먼저 매칭되는 것이 적용됨)
  for (const { keywords, result } of SUMMARY_CONFIG.PROGRESS_ANALYSIS_PATTERNS) {
    for (const keyword of keywords) {
      if (sheetName.includes(keyword)) {
        return result;
      }
    }
  }
  return '?'; // 매칭되는 패턴이 없으면 미분류
}

/**
 * 시트명에서 시간 정보를 분석합니다
 * 다양한 시간 표기 형식을 인식하여 HH:MM 형식으로 변환합니다
 */
function analyzeTime_(sheetName) {
  // 복잡한 패턴부터 단순한 패턴 순으로 검사
  for (const { pattern, type } of SUMMARY_CONFIG.TIME_ANALYSIS_PATTERNS) {
    const match = sheetName.match(pattern);
    if (match) {
      return convertTime_(match, type);
    }
  }
  return '';
}

/**
 * 정규식 매치 결과를 시간 타입에 따라 HH:MM 형식으로 변환합니다
 * 다양한 시간 표기법 (한국어, 영어, 12시간/24시간제)을 처리합니다
 */
function convertTime_(match, type) {
  let hour, minute;
  
  switch (type) {
    case 'midnight':
      return '00:00';
    
    case 'noon':
      return '12:00';

    case 'time_range':
      hour = parseInt(match[1]);
      return `${hour.toString().padStart(2, '0')}:00`;

    case 'time_range_detailed':
      hour = parseInt(match[1]);
      minute = parseInt(match[2]);
      return `${hour.toString().padStart(2, '0')}:${minute.toString().padStart(2, '0')}`;

    // 기존 케이스들
    case 'korean_time_with_prefix':
      const timePrefix = match[1];
      hour = parseInt(match[2]);
      const mins = match[3] ? parseInt(match[3]) : (match[4] ? 30 : 0);
      
      if (timePrefix === '오후') {
        hour = hour === 12 ? 12 : hour + 12;
      } else if ((timePrefix === '오전' || timePrefix === '새벽') && hour === 12) {
        hour = 0;
      }
      
      return `${hour.toString().padStart(2, '0')}:${mins.toString().padStart(2, '0')}`;

    case 'korean_time_detailed':
      hour = parseInt(match[1]);
      const minutes = match[2] ? parseInt(match[2]) : (match[3] ? 30 : 0);
      return `${hour.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}`;

    case 'time_ampm':
      const ampmBefore = match[1];
      hour = parseInt(match[2]);
      minute = parseInt(match[3]);
      const ampmAfter = match[4];
      
      const isPM = (ampmBefore && ampmBefore.toLowerCase() === 'p') || 
                  (ampmAfter && ampmAfter.toLowerCase() === 'p');
      const isAM = (ampmBefore && ampmBefore.toLowerCase() === 'a') || 
                  (ampmAfter && ampmAfter.toLowerCase() === 'a');
      
      // AM/PM 표시가 없으면 24시간제로 처리
      if (!isPM && !isAM) {
        return `${hour.toString().padStart(2, '0')}:${minute.toString().padStart(2, '0')}`;
      }
      
      if (isPM) {
        hour = hour === 12 ? 12 : hour + 12;
      } else {
        hour = hour === 12 ? 0 : hour;
      }
      
      return `${hour.toString().padStart(2, '0')}:${minute.toString().padStart(2, '0')}`;

    case 'ampm_suffix':
      hour = parseInt(match[1]);
      const ampmSuffix = match[2];
      
      if (ampmSuffix && ampmSuffix.toLowerCase() === 'p') {
        hour = hour === 12 ? 12 : hour + 12;
      } else {
        hour = hour === 12 ? 0 : hour;
      }
      
      return `${hour.toString().padStart(2, '0')}:00`;

    case 'ampm_prefix':
      hour = parseInt(match[2]);
      const ampmPrefix = match[1];
      
      if (ampmPrefix && ampmPrefix.toLowerCase() === 'p') {
        hour = hour === 12 ? 12 : hour + 12;
      } else {
        hour = hour === 12 ? 0 : hour;
      }
      
      return `${hour.toString().padStart(2, '0')}:00`;

    case 'simple_hour':
      hour = parseInt(match[1]);
      return `${hour.toString().padStart(2, '0')}:00`;

    default:
      return '';
  }
}

/**
 * 시트명에서 날짜 정보를 분석합니다
 * 스마트 월 할당: 월 정보가 없는 경우 추정된 참조 월을 사용합니다
 */
function analyzeDate_(sheetName, referenceMonth) {
  // 다양한 날짜 형식 패턴을 순서대로 검사
  for (const { pattern, format } of SUMMARY_CONFIG.DATE_ANALYSIS_PATTERNS) {
    const match = sheetName.match(pattern);
    if (match) {
      const dateResult = parseDate_(match, format, referenceMonth);
      if (dateResult.date) {
        return {
          date: dateResult.date,
          // 요일 정보: 계산된 요일 우선, 없으면 텍스트에서 추출
          dayOfWeek: calculateDayOfWeek_(dateResult.month, dateResult.day) || extractDayOfWeek_(sheetName)
        };
      }
    }
  }
  
  return { 
    date: '', 
    dayOfWeek: extractDayOfWeek_(sheetName) // 날짜 없어도 요일만 있을 수 있음
  };
}

/**
 * 정규식 매치 결과를 날짜 형식에 따라 파싱합니다
 * 월 정보가 없는 경우 참조 월을 사용합니다 (스마트 월 할당)
 */
function parseDate_(match, format, referenceMonth) {

  let month, day;
  
  switch (format) {
    case 'YYYY_MM_DD':
    case 'YY_MM_DD':
    case 'YYMMDD_6DIGIT':
    case 'YYYYMMDD_8DIGIT':
      // 예: "2024/12/15", "24/12/15", "241215", "20241215"
      month = parseInt(match[2]);
      day = parseInt(match[3]);
      break;

    case 'KOREAN_MM_DD':
    case 'MM_DD':
    case 'MMDD_4DIGIT':
      // 예: "12/15", "1215"
      month = parseInt(match[1]);
      day = parseInt(match[2]);
      break;

    case 'DD_ONLY':
    case 'DD_WITH_DAY':
      // 예: "15일", "15(월)" - 월 정보가 없으므로 참조 월 사용
      day = parseInt(match[1]);
      month = getSmartMonthAssignment_(day, referenceMonth);
      break;
      
    default:
      return { date: '', month: null, day: null };
  }
  
  // 날짜 유효성 검사
  if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
    return {
      date: `${month.toString().padStart(2, '0')}/${day.toString().padStart(2, '0')}`,
      month: month,
      day: day
    };
  }
  
  return { date: '', month: null, day: null };
}

function calculateDayOfWeek_(month, day) {
  try {
    const date = new Date(SUMMARY_CONFIG.CURRENT_YEAR, month - 1, day);
    return SYSTEM_CONFIG.WEEKDAYS[date.getDay()];
  } catch (error) {
    return '';
  }
}

function extractDayOfWeek_(text) {
  const dayChars = ['월', '화', '수', '목', '금', '토', '일'];
  return dayChars.find(day => text.includes(day)) || '';
}

/**
 * 모든 시트명을 분석하여 스마트 월을 추정합니다
 * 월/일 형태의 날짜들을 분석하여 가장 빈도가 높은 월을 참조 월로 결정합니다
 * 월 경계 패턴이 감지되면 스마트 할당을 통해 월말/월초를 적절히 분리합니다
 */
function estimateReferenceMonth_(allSheetNames) {
  const monthCounts = {};
  const daysWithMonths = [];
  const daysWithoutMonth = [];
  
  // 1단계: 월/일 정보 분류
  classifyDateInformation_(allSheetNames, monthCounts, daysWithMonths, daysWithoutMonth);
  
  // 2단계: 월 경계 처리 및 최종 월 결정
  const monthBoundaryInfo = determineFinalReferenceMonth_(daysWithMonths, daysWithoutMonth, monthCounts);
  
  logReferenceMonthAnalysis_(monthBoundaryInfo.referenceMonth, monthCounts, daysWithMonths, monthBoundaryInfo);
  
  // 3단계: 월 경계 정보를 전역적으로 저장 (parseDate_에서 사용)
  if (typeof globalThis === 'undefined') {
    this.monthBoundaryInfo = monthBoundaryInfo;
  } else {
    globalThis.monthBoundaryInfo = monthBoundaryInfo;
  }
  
  return monthBoundaryInfo.referenceMonth;
}

/**
 * 시트명들에서 날짜 정보를 분류합니다
 * 기존의 parseDate_ 함수를 직접 재사용하여 완전한 일관성을 유지합니다
 */
function classifyDateInformation_(allSheetNames, monthCounts, daysWithMonths, daysWithoutMonth) {
  // 임시 참조 월 (실제 값은 중요하지 않음, parseDate_ 호출용)
  const tempReferenceMonth = 9;
  
  for (const sheetName of allSheetNames) {
    let foundMonthDay = false;
    
    // 기존 DATE_ANALYSIS_PATTERNS를 순서대로 시도
    for (const { pattern, format } of SUMMARY_CONFIG.DATE_ANALYSIS_PATTERNS) {
      const match = sheetName.match(pattern);
      if (match) {
        // parseDate_ 함수를 직접 사용
        const dateResult = parseDate_(match, format, tempReferenceMonth);
        
        if (dateResult.date && dateResult.month && dateResult.day) {
          // DD_ONLY, DD_WITH_DAY 형식이 아닌 경우만 월/일 정보로 카운트
          if (format !== 'DD_ONLY' && format !== 'DD_WITH_DAY') {
            monthCounts[dateResult.month] = (monthCounts[dateResult.month] || 0) + 1;
            daysWithMonths.push({ month: dateResult.month, day: dateResult.day });
            foundMonthDay = true;
          } else {
            // DD_ONLY, DD_WITH_DAY는 일만 있는 정보로 처리
            daysWithoutMonth.push({ day: dateResult.day });
            foundMonthDay = true;
          }
          break; // 첫 번째 매칭된 패턴만 사용
        }
      }
    }
  }
}

/**
 * 최종 참조 월을 결정하고 월 경계 정보를 생성합니다
 */
function determineFinalReferenceMonth_(daysWithMonths, daysWithoutMonth, monthCounts) {
  if (Object.keys(monthCounts).length === 0) {
    return {
      referenceMonth: new Date().getMonth() + 1,
      isMonthBoundary: false,
      nextMonth: null
    };
  }
  
  // 가장 많이 나타난 월 찾기
  let maxCount = 0;
  let referenceMonth = new Date().getMonth() + 1;
  
  for (const [month, count] of Object.entries(monthCounts)) {
    if (count > maxCount) {
      maxCount = count;
      referenceMonth = parseInt(month);
    }
  }
  
  // 월 경계 패턴 감지
  const boundaryInfo = detectMonthBoundaryPattern_(daysWithMonths, daysWithoutMonth, referenceMonth);
  
  return {
    referenceMonth: referenceMonth,
    isMonthBoundary: boundaryInfo.detected,
    nextMonth: boundaryInfo.nextMonth,
    earlyDayCount: boundaryInfo.earlyDayCount,
    lateDayCount: boundaryInfo.lateDayCount
  };
}

/**
 * 월 경계 패턴을 감지하고 다음 월을 계산합니다
 * 실제 날짜 분포를 분석하여 월 경계가 진짜 발생했는지 판단합니다
 */
function detectMonthBoundaryPattern_(daysWithMonths, daysWithoutMonth, referenceMonth) {
  const settings = SUMMARY_CONFIG.MONTH_BOUNDARY_SETTINGS;
  
  if (daysWithoutMonth.length === 0 || daysWithMonths.length === 0) {
    return { detected: false, nextMonth: null, earlyDayCount: 0, lateDayCount: 0 };
  }
  
  // 월별 날짜 개수 계산 - 참조 월이 압도적으로 많으면 월 경계가 아님
  const monthDayCounts = {};
  daysWithMonths.forEach(d => {
    monthDayCounts[d.month] = (monthDayCounts[d.month] || 0) + 1;
  });
  
  // 월/일 정보가 있는 날짜들의 일(day)만 추출하고 정렬
  const days = daysWithMonths.map(d => d.day).sort((a, b) => a - b);
  
  // 중복 제거
  const uniqueDays = [...new Set(days)];
  
  if (uniqueDays.length < 2) {
    return { detected: false, nextMonth: null, earlyDayCount: 0, lateDayCount: 0 };
  }
  
  // 날짜 간 갭 분석 (순환 고려)
  let hasMonthBoundary = false;
  let maxGap = 0;
  
  for (let i = 1; i < uniqueDays.length; i++) {
    const gap = uniqueDays[i] - uniqueDays[i-1];
    maxGap = Math.max(maxGap, gap);
  }
  
  // 월 경계 판단 조건:
  // 1. 25일 이상의 날짜가 있고
  // 2. 5일 이하의 날짜가 있고
  // 3. 두 그룹 사이에 큰 갭(15일 이상)이 있음
  const hasLateDays = uniqueDays.some(d => d >= settings.LATE_DAY_THRESHOLD);
  const hasEarlyDays = uniqueDays.some(d => d <= settings.EARLY_DAY_THRESHOLD);
  const hasLargeGap = maxGap >= 15;
  
  // 순환 패턴 체크: 마지막 날짜가 크고(25+) 첫 날짜가 작으면(5-) 월 경계
  const lastDay = uniqueDays[uniqueDays.length - 1];
  const firstDay = uniqueDays[0];
  const isCircularPattern = lastDay >= settings.LATE_DAY_THRESHOLD && 
                           firstDay <= settings.EARLY_DAY_THRESHOLD;
  
  hasMonthBoundary = (hasLateDays && hasEarlyDays && hasLargeGap) || isCircularPattern;
  
  if (hasMonthBoundary) {
    // 월 정보 없는 날짜들 확인
    const earlyDaysWithoutMonth = daysWithoutMonth.filter(d => d.day <= settings.EARLY_DAY_THRESHOLD);
    
    // 다음 월 계산 (12월 → 1월 처리)
    const nextMonth = referenceMonth === 12 ? 1 : referenceMonth + 1;
    
    return {
      detected: true,
      nextMonth: nextMonth,
      earlyDayCount: earlyDaysWithoutMonth.length,
      lateDayCount: uniqueDays.filter(d => d >= settings.LATE_DAY_THRESHOLD).length
    };
  }
  
  return { detected: false, nextMonth: null, earlyDayCount: 0, lateDayCount: 0 };
}

/**
 * 스마트 월 할당 - 월 경계 패턴에 따라 적절한 월을 결정합니다
 */
function getSmartMonthAssignment_(day, referenceMonth) {
  // 월 경계 정보 가져오기
  const boundaryInfo = (typeof globalThis !== 'undefined' ? globalThis.monthBoundaryInfo : this.monthBoundaryInfo) || 
                       { isMonthBoundary: false };
  
  const settings = SUMMARY_CONFIG.MONTH_BOUNDARY_SETTINGS;
  
  // 스마트 할당이 비활성화되어 있거나 월 경계가 아니면 기본 월 사용
  if (!settings.ENABLE_SMART_ASSIGNMENT || !boundaryInfo.isMonthBoundary) {
    return referenceMonth;
  }
  
  // 월말 날짜 → 이전 달
  if (day >= settings.LATE_DAY_THRESHOLD) {
    const prevMonth = referenceMonth === 1 ? 12 : referenceMonth - 1;
    logMessage_(`스마트 월 할당: ${day}일 → ${prevMonth}월 (월말로 판단)`);
    return prevMonth;
  }
  
  // 월초 날짜 → 참조 월 그대로 (nextMonth 아님!)
  if (day <= settings.EARLY_DAY_THRESHOLD) {
    logMessage_(`스마트 월 할당: ${day}일 → ${referenceMonth}월 (월초로 판단)`);
    return referenceMonth;
  }
  
  // 그 외는 참조 월 사용
  return referenceMonth;
}

/**
 * 참조 월 분석 결과를 로그로 출력합니다
 */
function logReferenceMonthAnalysis_(finalMonth, monthCounts, daysWithMonths, boundaryInfo) {
  logMessage_(`참조 월 추정 결과: ${finalMonth}월 (${Object.values(monthCounts).reduce((a, b) => a + b, 0)}개 항목 기준)`);
  logMessage_(`월별 분포: ${JSON.stringify(monthCounts)}`);
  
  if (daysWithMonths.length > 0) {
    logMessage_(`월 정보가 있는 날짜 개수: ${daysWithMonths.length}개`);
  }
  
  if (boundaryInfo.isMonthBoundary) {
    logMessage_(`월 경계 패턴 감지: 월말 ${boundaryInfo.lateDayCount}개, 월초 ${boundaryInfo.earlyDayCount}개`);
    logMessage_(`스마트 할당: ${SUMMARY_CONFIG.MONTH_BOUNDARY_SETTINGS.EARLY_DAY_THRESHOLD}일 이하 → ${boundaryInfo.nextMonth}월`);
  }
}