/**
 * ===== config.gs =====
 * 전체 프로젝트 통합 설정 파일
 * 모든 설정값과 상수를 여기서 관리합니다.
 */

// ===== 시트명 수집 관련 설정 =====
const SHEET_COLLECTION_CONFIG = {
  // 시트 이름
  TARGET_SHEET_NAME: 'SheetNames',          // 시트명을 수집할 대상 시트
  MONITORING_SHEET_NAME: 'M',               // 실행 상태를 모니터링하는 시트
  EXCLUDE_STRING_EXACT_MATCH_PREFIX: '_P_', // 정확히 일치할 때만 제외하는 접두사
  
  // 행 번호 설정 (SheetNames 시트 기준)
  URL_ROW: 1,                               // 스프레드시트 URL이 위치한 행
  EXCLUDE_STRINGS_START_ROW: 2,             // 제외할 시트명 패턴 시작 행
  EXCLUDE_STRINGS_END_ROW: 11,              // 제외할 시트명 패턴 끝 행
  SPREADSHEET_NAME_ROW: 14,                 // 스프레드시트명을 기록할 행
  SHEET_NAMES_START_ROW: 15,                // 시트명 목록 시작 행
  SHEET_NAMES_END_ROW: 34,                  // 시트명 목록 끝 행
  SHEET_GIDS_START_ROW: 35,                 // 시트 GID 목록 시작 행
  SHEET_GIDS_END_ROW: 54,                   // 시트 GID 목록 끝 행
  
  // 열 번호 설정
  DATA_START_COLUMN: 2,                     // 데이터가 시작되는 열 (B열)
  
  // 성능 및 안정성 설정
  BATCH_SIZE: 4,                            // 한 번에 처리할 스프레드시트 개수
  RETRY_ATTEMPTS: 2,                        // 실패 시 재시도 횟수
  INTER_BATCH_DELAY_MS: 500,                // 배치 간 대기시간 (밀리초)
  INTER_REQUEST_DELAY_MS: 100,              // 요청 간 대기시간 (밀리초)
  RETRY_DELAY_MS: 500,                      // 재시도 시 대기시간 (밀리초)
  
  // 24시간 운영 설정
  MAIN_EXECUTION_INTERVAL_MINUTES: 10,      // 메인 실행 주기 (분)
  
  // 로깅 설정
  ENABLE_DETAILED_LOGGING: true,            // 상세 로그 활성화 (false로 하면 성능 향상)
  ERROR_MESSAGE_MAX_LENGTH: 30              // 오류 메시지 최대 길이
};

// ===== 요약 시트 관련 설정 =====
const SUMMARY_CONFIG = {
  // 시스템 설정
  CURRENT_YEAR: new Date().getFullYear(),   // 현재 연도 (자동 계산)
  ENABLE_DUPLICATE_REMOVAL: true,           // 중복 시트명 제거 활성화
  
  // 시트 이름
  SOURCE_SHEET_NAME: 'SheetNames',          // 원본 데이터 시트명
  SUMMARY_SHEET_NAME: '요약',                // 요약 결과 시트명
  
  // 원본 시트의 행 번호 설정 (SheetNames 시트 기준)
  SOURCE_URL_ROW: 1,                        // URL 행
  SOURCE_TYPE_NAME_ROW: 12,                 // 타입명 행
  SOURCE_SPREADSHEET_NAME_ROW: 14,          // 스프레드시트명 행
  SOURCE_SHEET_NAMES_START_ROW: 15,         // 시트명 시작 행
  SOURCE_SHEET_NAMES_END_ROW: 34,           // 시트명 끝 행
  SOURCE_SHEET_GIDS_START_ROW: 35,          // GID 시작 행
  SOURCE_SHEET_GIDS_END_ROW: 54,            // GID 끝 행
  SOURCE_DATA_START_COLUMN: 2,              // 데이터 시작 열
  
  // 요약 시트의 행/열 설정
  SUMMARY_DATA_START_ROW: 2,                // 데이터 시작 행
  
  // 요약 시트 열 배치 (각 데이터가 기록될 열 번호)
  SUMMARY_COLUMNS: {
    PROGRESS: 1,                            // A열: 진도
    DATE: 2,                                // B열: 날짜
    DAY_OF_WEEK: 3,                         // C열: 요일
    TIME: 4,                                // D열: 시간
    MEMBER_COUNT: 5,                        // E열: 인원 수
    LOCAL_SHEET_NAME: 6,                    // F열: 로컬 시트 이름
    ORIGINAL_SHEET_NAME: 7,                 // G열: 원본 시트 이름
    LOCAL_SHEET_LINK: 8,                    // H열: 로컬 시트 링크
    ORIGINAL_SPREADSHEET_URL: 9,            // I열: 원본 스프레드시트 URL
    ORIGINAL_SHEET_GID: 10,                 // J열: 원본 시트 GID
    ORIGINAL_SHEET_URL: 11,                 // K열: 원본 시트 URL
    MEMBER_COUNT_POSITION: 12,              // L열: 인원 수 항목 위치
    ORIGINAL_SHEET_LINK: 13,                // M열: 원본 시트 링크
    LOCAL_SHEET_ITEM_NUMBER: 14,            // N열: 로컬 시트 항목 번호
    LOCAL_SHEET_ROW_POSITION: 15            // O열: 로컬 시트 행 위치
  },
  
  // 자동 생성되는 열들 (나머지는 수동 입력 열로 보존됨)
  AUTO_GENERATED_COLUMNS: [
    'PROGRESS',                             // A열: 진도
    'DATE',                                 // B열: 날짜
    'DAY_OF_WEEK',                          // C열: 요일
    'TIME',                                 // D열: 시간
    'LOCAL_SHEET_NAME',                     // F열: 로컬 시트 이름 (12행 타입명 참조)
    'ORIGINAL_SHEET_NAME',                  // G열: 원본 시트 이름
    'ORIGINAL_SPREADSHEET_URL',             // I열: 원본 스프레드시트 URL
    'ORIGINAL_SHEET_GID',                   // J열: 원본 시트 GID
    'LOCAL_SHEET_ITEM_NUMBER'               // N열: 로컬 시트 항목 번호
  ],
  
  // 진도 분석 설정
  PROGRESS_SORT_ORDER: ['1넴', '2넴', '3넴', '4넴', '파밍', '주회', '?'], // 정렬 우선순위
  PROGRESS_ANALYSIS_PATTERNS: [
    // 우선순위 순으로 배열 (먼저 매칭되는 것이 적용됨)
    { keywords: ['주회', '고속', '고주'], result: '주회' },
    { keywords: ['1넴', '1네임드', '초행', '랜진', '진도무관', '진도 무관'], result: '1넴' },
    { keywords: ['2넴', '2네임드'], result: '2넴' },
    { keywords: ['3넴', '3네임드'], result: '3넴' },
    { keywords: ['4넴', '4네임드', '클목', '클리어', '성불', '48탱커'], result: '4넴' },
    { keywords: ['1클', '2클', '3클', '4클', '파밍', '저속'], result: '파밍' }
  ],
  
  // 시간 분석 설정 (정규식 패턴과 변환 타입)
  TIME_ANALYSIS_PATTERNS: [
    { pattern: /자정/, type: 'midnight' },
    { pattern: /정오/, type: 'noon' },
    { pattern: /(\d{1,2})\s*:\s*(\d{1,2})\s*~\s*\d{1,2}\s*:\s*\d{1,2}/, type: 'time_range_detailed' },
    { pattern: /(\d{1,2})(?:시)?\s*~\s*\d{1,2}시?(?![0-9가-힣])/, type: 'time_range' },
    { pattern: /(오후|오전|새벽)\s*(\d{1,2})(?:시)?\s*(?:(\d{1,2})분|(반))?(?![가-힣])/, type: 'korean_time_with_prefix' },
    { pattern: /(\d{1,2})시\s*(?:(\d{1,2})분|(반))(?![가-힣])/, type: 'korean_time_detailed' },
    { pattern: /(?:([ap])\.?m\.?\s*?)?(\d{1,2})\s*:\s*(\d{1,2})(?:\s*([ap])\.?m\.?)?/i, type: 'time_ampm' },
    { pattern: /\b(\d{1,2})\s*([ap])\.?m\.?\b/i, type: 'ampm_suffix' },
    { pattern: /\b([ap])\.?m\.?\s*(\d{1,2})\b/i, type: 'ampm_prefix' },
    { pattern: /(\d{1,2})(?:시|출발?)(?!\d)/, type: 'simple_hour' }
  ],
  
  // 날짜 분석 설정 (정규식 패턴과 형식)
  DATE_ANALYSIS_PATTERNS: [
    { pattern: /(\d{1,2})월\s*(\d{1,2})일/, format: 'KOREAN_MM_DD' },
    { pattern: /(\d{4})[\/.:-]+(\d{1,2})[\/.:-]+(\d{1,2})/, format: 'YYYY_MM_DD' },
    { pattern: /(\d{2})[\/.:-]+(\d{1,2})[\/.:-]+(\d{1,2})/, format: 'YY_MM_DD' },
    { pattern: /(\d{1,2})[\/.:-]+(\d{1,2})(?![넴클])/, format: 'MM_DD' },
    { pattern: /(\d{4})(\d{2})(\d{2})(?![넴클])/, format: 'YYYYMMDD_8DIGIT' },
    { pattern: /(\d{2})(\d{2})(\d{2})(?![넴클])/, format: 'YYMMDD_6DIGIT' },
    { pattern: /(\d{2})(\d{2})(?![넴클])/, format: 'MMDD_4DIGIT' },
    { pattern: /(\d{1,2})\([월화수목금토일]\)/, format: 'DD_WITH_DAY' },
    { pattern: /(\d{1,2})일/, format: 'DD_ONLY' }
  ],
  
  // 월 경계 스마트 할당 설정
  MONTH_BOUNDARY_SETTINGS: {
    LATE_DAY_THRESHOLD: 25,                 // 월말로 간주할 일 기준 (25일 이후)
    EARLY_DAY_THRESHOLD: 5,                 // 월초로 간주할 일 기준 (5일 이전)
    ENABLE_SMART_ASSIGNMENT: true           // 스마트 월 할당 활성화
  },
  
  // 필터링 설정
  FILTER_SETTINGS: {
    HIDE_MEMBER_COUNT_PATTERN: '50',        // 숨길 인원수 패턴 (48명) - 50으로 바꿔서 모두 보이게 해둠
    HIDE_OLD_DAYS: 2,                       // 숨길 오래된 날짜 (일)
    HIDE_EMPTY_STATUS: true                 // 빈 상태값 숨김
  },

  // 즉시 출발 키워드 설정
  INSTANT_RUN_SETTINGS: {
    ENABLED: true,                           // 기능 활성화/비활성화
    KEYWORDS: ['모출', '번개', '즉시'],        // 감지할 키워드들
  }
};

// ===== 시스템 공통 설정 =====
const SYSTEM_CONFIG = {
  // 요일 배열 (일요일=0부터 시작)
  WEEKDAYS: ['일', '월', '화', '수', '목', '금', '토'],
  
  // 스프레드시트 ID 추출 패턴들
  SPREADSHEET_ID_PATTERNS: [
    /\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/,  // 전체 URL 형식
    /\/d\/([a-zA-Z0-9-_]+)/,                // 단축 URL 형식
    /^([a-zA-Z0-9-_]+)$/                    // ID만 있는 형식
  ],
  
  // 유효성 검사 설정
  VALIDATION: {
    EMPTY_VALUE_INDICATORS: ['-', ''],      // 빈 값으로 간주할 문자들
    MIN_SPREADSHEET_ID_LENGTH: 20           // 최소 스프레드시트 ID 길이
  }
};