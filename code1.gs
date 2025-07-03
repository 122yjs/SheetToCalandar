/**
 * @OnlyCurrentDoc
 * 
 * 구글 시트 데이터를 기반으로 구글 캘린더와 구글 Tasks를 동기화하는 스크립트입니다.
 * 모든 설정은 스프레드시트 내의 '설정' 시트에서 관리됩니다.
 */

// --- 전역 설정 변수 ---
// 이 변수들은 loadSettings 함수를 통해 '설정' 시트에서 값을 읽어와 채워집니다.
// 코드에는 기본값이 없으며, 사용자가 '설정' 시트를 통해 입력해야 합니다.
let CALENDAR_ID;
let TASK_LIST_ID;
let SHEET_NAME;
let HEADER_ROW;
let DATA_START_ROW;
let COLOR_ID_GRADE_3; // (레거시) 3학년 캘린더 색상 ID
let CALENDAR_COLOR_RULES = []; // 색상 규칙을 저장할 배열
let COLOR_RULES_SHEET_NAME_SETTING = '색상 규칙 시트 이름'; // '설정' 시트에서 규칙 시트 이름을 가져올 때 사용할 키 값

// 데이터 시트의 열 번호 설정
let COL_TYPE;
let COL_TITLE;
let COL_START_DATE;
let COL_START_TIME;
let COL_END_DATE;
let COL_END_TIME;
let COL_DUE_DATE;
let COL_GRADE;
let COL_DESC;
let COL_STATUS;
let COL_SYNC_ID;
let COL_SYNC_RESULT;

// 고정된 시트 이름 정의
const SETTINGS_SHEET_NAME = '설정';
const SYNC_HISTORY_SHEET_NAME = '동기화이력'; // 동기화 이력을 추적할 시트 이름 (자동 생성 및 숨김 처리됨)

/**
 * '설정' 시트에서 스크립트 설정을 불러와 전역 변수에 할당합니다.
 * @return {boolean} 설정 로드 성공 여부
 */
function loadSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
  if (!settingsSheet) {
    SpreadsheetApp.getUi().alert(`'${SETTINGS_SHEET_NAME}' 시트를 찾을 수 없습니다. 스크립트 설정을 불러올 수 없습니다. 매뉴얼을 확인해주세요.`);
    return false;
  }

  // 설정 시트의 A열(키)과 B열(값)을 읽어옵니다.
  const settingsData = settingsSheet.getRange('A1:B30').getValues(); // 범위는 설정 항목 수에 맞게 조정
  const settings = {}; // 임시로 설정을 담을 객체
  settingsData.forEach(row => {
    if (row[0] && row[1] !== undefined && row[1] !== null && row[1] !== '') {
      settings[row[0].toString().trim()] = row[1].toString().trim();
    } else if (row[0]) {
      Logger.log(`정보: '${SETTINGS_SHEET_NAME}' 시트의 '${row[0]}' 항목 값이 비어있습니다.`);
    }
  });

  // --- 필수 기본 설정 값 확인 ---
  const requiredBasicSettings = ['캘린더 ID', 'Task List ID', '데이터 시트 이름', '헤더 행 번호', '데이터 시작 행 번호',
                                '구분 열 번호', '제목 열 번호', 'Sync ID 열 번호', '동기화 결과 열 번호'];
  for (const key of requiredBasicSettings) {
    if (!settings[key]) {
      SpreadsheetApp.getUi().alert(`필수 설정 누락: '${SETTINGS_SHEET_NAME}' 시트에서 '${key}' 항목을 설정해주세요.`);
      return false;
    }
  }

  // --- 전역 변수에 기본 설정 할당 ---
  CALENDAR_ID = settings['캘린더 ID'];
  TASK_LIST_ID = settings['Task List ID'];
  SHEET_NAME = settings['데이터 시트 이름'];
  HEADER_ROW = parseInt(settings['헤더 행 번호'], 10);
  DATA_START_ROW = parseInt(settings['데이터 시작 행 번호'], 10);
  COLOR_ID_GRADE_3 = settings['3학년 캘린더 색상 ID'] || null; // 선택 사항

  // 열 번호 할당 (parseInt를 사용하여 숫자로 변환)
  COL_TYPE = parseInt(settings['구분 열 번호'], 10);
  COL_TITLE = parseInt(settings['제목 열 번호'], 10);
  COL_START_DATE = parseInt(settings['시작일 열 번호'], 10);
  COL_START_TIME = parseInt(settings['시작 시간 열 번호'], 10);
  COL_END_DATE = parseInt(settings['종료일 열 번호'], 10);
  COL_END_TIME = parseInt(settings['종료 시간 열 번호'], 10);
  COL_DUE_DATE = parseInt(settings['마감일 열 번호 (Tasks)'], 10);
  COL_GRADE = parseInt(settings['학년 열 번호'], 10);
  COL_DESC = parseInt(settings['설명 열 번호'], 10);
  COL_STATUS = parseInt(settings['상태 열 번호'], 10);
  COL_SYNC_ID = parseInt(settings['Sync ID 열 번호'], 10);
  COL_SYNC_RESULT = parseInt(settings['동기화 결과 열 번호'], 10);

  // --- 설정 값 유효성 검사 (숫자 여부 등) ---
  const numericColsToValidate = {
    HEADER_ROW, DATA_START_ROW, COL_TYPE, COL_TITLE, 
    COL_SYNC_ID, COL_SYNC_RESULT
  };
  // 선택적 숫자 열 (값이 있을 경우에만 숫자여야 함)
  const optionalNumericCols = {
    COL_START_DATE, COL_START_TIME, COL_END_DATE, COL_END_TIME, 
    COL_DUE_DATE, COL_GRADE, COL_DESC, COL_STATUS
  };

  // 필수 숫자 값 검사
  for (const key in numericColsToValidate) {
    if (isNaN(numericColsToValidate[key])) {
      SpreadsheetApp.getUi().alert(`설정 오류: '${key}'에 해당하는 설정 항목이 숫자가 아닙니다. '설정' 시트를 확인해주세요.`);
      return false;
    }
  }
  // 선택적 숫자 값 검사
  for (const key in optionalNumericCols) {
    const settingName = key.replace(/_/g, ' ').replace(/\bCOL\b/g, '열').replace('ROW', '행 번호').trim(); // 설정 시트의 실제 이름과 매칭 시도
    // 설정 시트에 해당 항목 이름이 있고, 값이 비어있지 않으며, 변환된 숫자가 NaN인 경우 오류 처리
    if (settings[settingName] && settings[settingName] !== '' && isNaN(optionalNumericCols[key])) {
       SpreadsheetApp.getUi().alert(`설정 오류: '${settingName}' 관련 값이 있지만 숫자가 아닙니다. 비워두거나 올바른 열 번호를 입력하세요.`);
       return false;
    }
  }

  // 행 번호 논리 검사
  if (HEADER_ROW < 1 || DATA_START_ROW < 1 || DATA_START_ROW <= HEADER_ROW) {
    SpreadsheetApp.getUi().alert(`설정 오류: '헤더 행 번호'(${HEADER_ROW}) 또는 '데이터 시작 행 번호'(${DATA_START_ROW})가 올바르지 않습니다. (헤더 행 >= 1, 데이터 시작 행 > 헤더 행)`);
    return false;
  }

  // --- 캘린더 색상 규칙 로드 (별도 시트 방식) ---
  CALENDAR_COLOR_RULES = []; // 초기화
  const colorRulesSheetNameFromSettings = settings[COLOR_RULES_SHEET_NAME_SETTING]; // '설정' 시트에서 규칙 시트 이름 가져오기

  if (colorRulesSheetNameFromSettings) {
    const colorRulesSheet = ss.getSheetByName(colorRulesSheetNameFromSettings.toString().trim());
    if (colorRulesSheet) {
      // 규칙 시트의 데이터는 2행부터 시작한다고 가정 (1행은 헤더)
      const rulesSheetDataStartRow = 2;
      const numRowsInRulesSheet = colorRulesSheet.getLastRow() - rulesSheetDataStartRow + 1;
      const numColsInRulesSheet = 3; // 대상열설정이름, 검색값, 색상ID (A, B, C열)

      if (numRowsInRulesSheet > 0) {
        const rulesData = colorRulesSheet.getRange(rulesSheetDataStartRow, 1, numRowsInRulesSheet, numColsInRulesSheet).getValues();

        rulesData.forEach((ruleRow, index) => {
          const ruleRowIndexForLog = rulesSheetDataStartRow + index; // 실제 시트 행 번호 (로그용)
          const targetColSettingName = ruleRow[0] ? ruleRow[0].toString().trim() : null; // A열
          const searchValue = ruleRow[1] ? ruleRow[1].toString().trim() : null;          // B열
          const colorId = ruleRow[2] ? ruleRow[2].toString().trim() : null;            // C열

          if (targetColSettingName && searchValue && colorId) {
            // 'settings' 객체에서 targetColSettingName에 해당하는 '실제 열 번호'를 조회
            const targetColNumberFromSettings = settings[targetColSettingName];
            if (!targetColNumberFromSettings || isNaN(parseInt(targetColNumberFromSettings, 10))) {
              Logger.log(`경고: '${colorRulesSheetNameFromSettings}' 시트 ${ruleRowIndexForLog}행 규칙: '대상 열 설정 이름'('${targetColSettingName}')에 해당하는 열 번호 설정을 '${SETTINGS_SHEET_NAME}' 시트에서 찾을 수 없거나 숫자가 아닙니다. 이 규칙을 건너뜁니다.`);
              return; // forEach에서 다음 요소로 넘어감 (continue)
            }
            const targetActualColNumber = parseInt(targetColNumberFromSettings, 10);

            CALENDAR_COLOR_RULES.push({
              targetColNumber: targetActualColNumber, // 실제 데이터 시트의 열 번호 (1-based)
              searchValue: searchValue,
              colorId: colorId // Google Calendar 색상 ID (문자열 '1'-'11')
            });
          } else if (targetColSettingName || searchValue || colorId) { // 행에 일부 데이터만 있는 경우 (모두 비어있지 않으면)
             Logger.log(`경고: '${colorRulesSheetNameFromSettings}' 시트 ${ruleRowIndexForLog}행 규칙에 필수 값(대상 열 설정 이름, 검색 값, 색상 ID 중 하나 이상)이 누락되었습니다. 이 규칙을 건너뜁니다.`);
          }
          // 모든 값이 비어있는 행은 자동으로 무시됨
        });
        Logger.log(`'${colorRulesSheetNameFromSettings}' 시트로부터 총 ${CALENDAR_COLOR_RULES.length}개의 캘린더 색상 규칙이 로드되었습니다.`);
      } else {
        Logger.log(`정보: '${colorRulesSheetNameFromSettings}' 시트에 정의된 색상 규칙이 없습니다 (헤더 제외).`);
      }
    } else {
      Logger.log(`경고: '${SETTINGS_SHEET_NAME}' 시트에 설정된 색상 규칙 시트 '${colorRulesSheetNameFromSettings}'을(를) 찾을 수 없습니다.`);
      SpreadsheetApp.getUi().alert(`경고: 색상 규칙 시트 '${colorRulesSheetNameFromSettings}'을(를) 찾을 수 없습니다. 캘린더 색상 규칙이 적용되지 않습니다.`);
      // 색상 규칙 시트가 없어도 기본 동기화는 진행되도록 return false는 하지 않음.
    }
  } else {
    Logger.log(`정보: '${COLOR_RULES_SHEET_NAME_SETTING}'이(가) '${SETTINGS_SHEET_NAME}' 시트에 설정되지 않았습니다. 캘린더 색상 규칙을 로드하지 않습니다.`);
  }

  return true;
}

// ================================================================================
// 메뉴 및 실행 함수들
// ================================================================================

/**
 * 스크립트가 열릴 때 맞춤 메뉴를 생성합니다.
 */
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('동기화 관리')
      .addItem('캘린더/Tasks와 동기화 실행', 'runSyncWithPreview')
      .addSeparator()
      .addItem('Task List ID 확인', 'getTaskLists')
      .addItem('설정 다시 불러오기', 'reloadSettingsShowAlert')
      .addToUi();
}

/**
 * 설정을 다시 불러오고 사용자에게 알림을 표시합니다.
 */
function reloadSettingsShowAlert(){
  if (loadSettings()){
    SpreadsheetApp.getUi().alert("'설정' 시트에서 설정을 성공적으로 다시 불러왔습니다.");
  }
  // loadSettings 내부에서 오류 발생 시 이미 알림을 띄움
}

/**
 * 사용자의 Task List ID를 로그에 출력합니다. (설정 도우미 기능)
 */
function getTaskLists() {
  try {
    // Google Tasks API 고급 서비스 활성화 확인
    if (!Tasks || !Tasks.Tasklists) { 
       const message = "Google Tasks API 고급 서비스가 활성화되지 않았습니다. 스크립트 편집기의 '서비스' 메뉴(+)에서 'Google Tasks API'를 추가하고 활성화해주세요.";
       Logger.log(message);
       SpreadsheetApp.getUi().alert(message);
       return;
    }

    const taskLists = Tasks.Tasklists.list({maxResults: 100}); // 최대 100개 목록 가져오기
    if (!taskLists.items || taskLists.items.length === 0) {
      Logger.log('사용 가능한 Google Tasks 목록이 없습니다.');
      SpreadsheetApp.getUi().alert('사용 가능한 Google Tasks 목록이 없습니다. Google Tasks에서 목록을 먼저 생성해주세요.');
      return;
    }

    let message = '사용 가능한 Task Lists (ID는 로그에 기록됨):\n';
    Logger.log('사용 가능한 Task Lists:');
    taskLists.items.forEach(function(taskList) {
      Logger.log(` - 제목: ${taskList.title}, ID: ${taskList.id}`);
      message += `- 제목: ${taskList.title}\n`;
    });
    SpreadsheetApp.getUi().alert('Task List ID가 Apps Script 로그에 기록되었습니다. (실행 -> 실행 기록)\n\n' + message + "\n\n로그에서 원하는 목록의 ID를 복사하여 '설정' 시트의 'Task List ID' 값으로 입력하세요.");

  } catch (e) {
    Logger.log(`Task List ID를 가져오는 중 오류 발생: ${e.toString()}\nStack: ${e.stack}`);
    SpreadsheetApp.getUi().alert(`Task List ID를 가져오는 중 오류가 발생했습니다. 로그를 확인하세요. (고급 Google 서비스에서 Tasks API가 활성화되었는지 확인하세요.)`);
  }
}

/**
 * 동기화 실행 전 미리보기를 제공하고 사용자 확인을 받습니다.
 */
function runSyncWithPreview() {
  if (!loadSettings()) return; // 설정 로드 실패 시 중단

  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    ui.alert(`데이터 시트 '${SHEET_NAME}'를 찾을 수 없습니다. '설정' 시트의 '데이터 시트 이름'을 확인하세요.`);
    return;
  }

  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();

  if (values.length <= HEADER_ROW) {
      ui.alert("시트에 동기화할 데이터가 없습니다 (헤더 제외).");
      return;
  }

  // 미리보기 요약 및 상세 로그 준비
  let previewSummary = { create: 0, update: 0, delete: 0, noChange: 0, error: 0, skip: 0 };
  let previewDetailsLog = [`--- 동기화 미리보기 (시트: ${SHEET_NAME}) ---`];

  const numDataRows = values.length - HEADER_ROW;
  // Sync ID 열이 설정되어 있는지 확인
  if (isNaN(COL_SYNC_ID) || COL_SYNC_ID < 1) {
    ui.alert("설정 오류: 'Sync ID 열 번호'가 올바르게 설정되지 않았습니다.");
    return;
  }
  const syncIds = sheet.getRange(DATA_START_ROW, COL_SYNC_ID, numDataRows, 1).getValues();

  // 현재 시트에 있는 모든 Sync ID 수집
  const currentSyncIds = new Set();
  
  // 각 행을 순회하며 예상 작업 분류
  for (let i = HEADER_ROW; i < values.length; i++) {
    const row = values[i];
    const rowIndexInSheet = i + 1;
    const dataRowIndex = i - HEADER_ROW;

    const title = row[COL_TITLE -1] ? row[COL_TITLE -1].toString().trim() : '';
    if (!title) {
      previewSummary.skip++;
      previewDetailsLog.push(`행 ${rowIndexInSheet}: 제목 없음 - 건너뜀`);
      continue;
    }

    const type = row[COL_TYPE -1] ? row[COL_TYPE -1].toString().trim().toLowerCase() : '';
    const status = row[COL_STATUS -1] ? row[COL_STATUS -1].toString().trim().toLowerCase() : '';
    const currentSyncId = syncIds[dataRowIndex][0] ? syncIds[dataRowIndex][0].toString().trim() : null;

    let action = "정보: 변경 없음 또는 처리 대상 아님";
    if (type === 'calendar') {
      if (status === '취소됨') {
        action = currentSyncId ? "캘린더 이벤트 삭제 예정" : "정보: 삭제할 이벤트 없음 (ID 없음)";
        if (currentSyncId) previewSummary.delete++; else previewSummary.noChange++;
      } else if (currentSyncId) {
        action = "캘린더 이벤트 업데이트 확인 필요";
        previewSummary.update++;
      } else {
        action = "캘린더 이벤트 생성 예정";
        previewSummary.create++;
      }
    } else if (type === 'tasks') {
      if (status === '취소됨' || status === '운영완료') { // Task는 '취소됨' 또는 '운영완료' 시 완료 처리
         action = currentSyncId ? "Task 완료 처리 또는 업데이트 예정" : "Task 생성 및 완료 처리 예정";
         if (currentSyncId) previewSummary.update++; else previewSummary.create++;
      } else if (currentSyncId) {
        action = "Task 업데이트 확인 필요";
        previewSummary.update++;
      } else {
        action = "Task 생성 예정";
        previewSummary.create++;
      }
    } else {
      action = `경고: 알 수 없는 구분 값(${type}) - 건너뜀`;
      previewSummary.skip++;
    }
    previewDetailsLog.push(`행 ${rowIndexInSheet} (제목: ${title}): ${action} (SyncID: ${currentSyncId || '없음'})`);
    
    if (currentSyncId) {
      currentSyncIds.add(currentSyncId);
    }
  }

  // 삭제된 행 확인을 위한 이력 비교 ('동기화이력' 시트 활용)
  const deletedItemsCount = checkDeletedItems(currentSyncIds);
  if (deletedItemsCount > 0) {
    previewSummary.delete += deletedItemsCount;
    previewDetailsLog.push(`\n삭제된 행: ${deletedItemsCount}개의 캘린더/Task 항목이 시트에서 제거되어 삭제 예정`);
  }

  // 미리보기 로그 기록
  previewDetailsLog.push("--- 동기화 미리보기 끝 ---");
  Logger.log(previewDetailsLog.join('\n'));

  // 사용자에게 요약 알림 및 확인
  let summaryMessage = `동기화 미리보기:\n` +
                       `- 생성 예정: ${previewSummary.create} 건\n` +
                       `- 업데이트/완료 예정: ${previewSummary.update} 건\n` +
                       `- 삭제 예정: ${previewSummary.delete} 건\n` +
                       `- 변경 없음/건너뜀/오류 예상: ${previewSummary.noChange + previewSummary.skip + previewSummary.error} 건\n\n` +
                       `자세한 내용은 스크립트 로그(실행 -> 실행 기록)를 확인하세요.\n\n` +
                       `동기화를 진행하시겠습니까?`;

  const response = ui.alert('동기화 확인', summaryMessage, ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES) {
    syncSheetToApis(); // 사용자가 동의하면 실제 동기화 실행
  } else {
    ui.alert('동기화가 취소되었습니다.');
  }
}

// ================================================================================
// 핵심 동기화 로직
// ================================================================================

/**
 * 구글 시트 데이터를 구글 캘린더 및 구글 Tasks와 동기화합니다. (핵심 함수)
 */
function syncSheetToApis() {
  // 설정은 runSyncWithPreview에서 이미 로드됨.
  
  // 1. API 및 ID 유효성 검사
  // Tasks API 고급 서비스 활성화 확인
  if (typeof Tasks === 'undefined' || !Tasks.Tasks || !Tasks.Tasklists) {
    const msg = "Google Tasks API 고급 서비스가 활성화되지 않았습니다. 'Task List ID 확인' 메뉴를 먼저 실행하거나, 스크립트 편집기 '서비스' 메뉴에서 활성화해주세요.";
    Logger.log(msg);
    SpreadsheetApp.getUi().alert(msg);
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  // sheet 존재 유무는 runSyncWithPreview에서 확인됨

  // 캘린더 접근 확인
  const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
  if (!calendar) {
    const msg = `캘린더 ID '${CALENDAR_ID}'를 찾을 수 없거나 접근 권한이 없습니다. '설정' 시트의 '캘린더 ID'를 확인하세요.`;
    Logger.log(msg);
    SpreadsheetApp.getUi().alert(msg);
    return;
  }

  // Task List 접근 확인
  try { 
    Tasks.Tasklists.get(TASK_LIST_ID); // Task List 존재 확인 시도
    Logger.log("Task List ID '%s' 확인 완료.", TASK_LIST_ID);
  } catch (e) {
    const msg = `Task List ID '${TASK_LIST_ID}'에 접근할 수 없습니다. ID가 정확한지, Tasks API가 활성화되었는지, 권한이 부여되었는지 확인하세요. 오류: ${e.toString()}`;
    Logger.log(msg);
    SpreadsheetApp.getUi().alert(`${msg} 자세한 내용은 로그를 확인하세요.`);
    return;
  }

  // 2. 데이터 및 동기화 정보 범위 설정
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  const numDataRows = values.length - HEADER_ROW;

  // Sync ID 및 결과 메시지 기록을 위한 열 번호 유효성 확인
  if (isNaN(COL_SYNC_ID) || COL_SYNC_ID < 1 || isNaN(COL_SYNC_RESULT) || COL_SYNC_RESULT < 1) {
      SpreadsheetApp.getUi().alert("설정 오류: 'Sync ID 열 번호' 또는 '동기화 결과 열 번호'가 올바르게 설정되지 않았습니다.");
      return;
  }
  
  // Sync ID와 결과 메시지를 한 번에 읽고 쓰기 위한 범위 설정
  const firstCol = Math.min(COL_SYNC_ID, COL_SYNC_RESULT);
  const lastCol = Math.max(COL_SYNC_ID, COL_SYNC_RESULT);
  const numColsForSyncData = lastCol - firstCol + 1;
  
  const syncDataRange = sheet.getRange(DATA_START_ROW, firstCol, numDataRows, numColsForSyncData);
  const syncDataValues = syncDataRange.getValues(); // Sync ID와 결과를 담을 배열

  // 배열 내에서의 인덱스 계산
  const syncIdColIndexInArray = COL_SYNC_ID - firstCol;
  const syncResultColIndexInArray = COL_SYNC_RESULT - firstCol;

  // 3. 동기화 이력 준비
  const syncHistory = getSyncHistory();
  const processedSyncIds = new Set(); // 현재 시트에 처리된 Sync ID 기록
  const currentItems = []; // '동기화이력' 시트 업데이트용 데이터

  Logger.log(`동기화 시작: 총 ${numDataRows}개 행 처리 예정 (시트: ${SHEET_NAME})`);
  SpreadsheetApp.getActiveSpreadsheet().toast(`동기화 시작... 총 ${numDataRows}개 항목 처리 중`, "동기화 진행", -1);

  // 4. 각 행 데이터 처리
  for (let i = HEADER_ROW; i < values.length; i++) {
    const rowData = values[i]; // 시트의 현재 행 전체 데이터
    const rowIndexInSheet = i + 1;
    const dataRowIndex = i - HEADER_ROW; // syncDataValues 배열의 인덱스

    // 배열 범위 확인 및 초기화 (안전 장치)
    if (!syncDataValues[dataRowIndex]) {
        Logger.log(`오류: 행 ${rowIndexInSheet}의 syncDataValues 인덱스 접근 불가.`);
        syncDataValues[dataRowIndex] = new Array(numColsForSyncData).fill(''); 
    }

    let currentSyncId = syncDataValues[dataRowIndex][syncIdColIndexInArray] ? syncDataValues[dataRowIndex][syncIdColIndexInArray].toString().trim() : null;
    let resultMessage = "";

    const title = rowData[COL_TITLE - 1] ? rowData[COL_TITLE - 1].toString().trim() : '';
    if (!title) {
      resultMessage = "오류: 제목 없음";
      Logger.log(`행 ${rowIndexInSheet}: ${resultMessage}. 건너뜁니다.`);
      syncDataValues[dataRowIndex][syncResultColIndexInArray] = resultMessage;
      continue;
    }

    const type = rowData[COL_TYPE - 1] ? rowData[COL_TYPE - 1].toString().trim().toLowerCase() : '';
    Logger.log(`행 ${rowIndexInSheet} 처리: 구분='${type}', 제목='${title}', SyncID='${currentSyncId || '없음'}'`);

    // 현재 시트에 있는 Sync ID 추가
    if (currentSyncId) {
      processedSyncIds.add(currentSyncId);
    }

    try {
      let syncResult = { id: currentSyncId, message: "정보: 처리되지 않음" }; // 기본값

      if (type === 'calendar') {
        // 캘린더 이벤트 처리
        syncResult = processCalendarEvent(calendar, rowData, currentSyncId, ss.getSpreadsheetTimeZone(), rowIndexInSheet);
      } else if (type === 'tasks') {
        // Tasks 할 일 처리
        syncResult = processTask(TASK_LIST_ID, rowData, currentSyncId, ss.getSpreadsheetTimeZone(), rowIndexInSheet);
      } else {
        // 알 수 없는 구분 값 처리
        resultMessage = `경고: 알 수 없는 구분 값 '${rowData[COL_TYPE - 1]}'`;
        Logger.log(`행 ${rowIndexInSheet}: ${resultMessage}. 건너뜁니다.`);
        syncResult = { id: null, message: resultMessage }; // 알 수 없는 타입이면 ID 클리어
      }

      // 처리 결과를 배열에 기록
      syncDataValues[dataRowIndex][syncIdColIndexInArray] = syncResult.id;
      syncDataValues[dataRowIndex][syncResultColIndexInArray] = syncResult.message;

      // 동기화 이력 업데이트를 위한 정보 수집
      if (syncResult.id) {
        currentItems.push({
          syncId: syncResult.id,
          type: type,
          title: title
        });
        // 새로 생성되거나 업데이트된 ID도 처리 목록에 추가
        if (syncResult.id !== currentSyncId) {
          processedSyncIds.add(syncResult.id);
        }
      }

    } catch (e) {
      // 예측하지 못한 오류 발생 시
      resultMessage = `심각한 오류: ${e.message.substring(0, 100)}... (로그 확인)`;
      Logger.log(`행 ${rowIndexInSheet} 처리 중 심각한 오류: ${e.toString()}`);
      syncDataValues[dataRowIndex][syncResultColIndexInArray] = resultMessage;
      // 오류 발생 시 Sync ID는 변경하지 않음 (기존 ID 유지)
    }
  }

  // 5. 시트에서 삭제된 항목들 처리 (동기화 이력 비교)
  Logger.log("삭제된 항목 확인 중...");
  for (const [syncId, itemInfo] of Object.entries(syncHistory)) {
    // 이력에는 있지만 현재 시트 처리 목록(processedSyncIds)에는 없는 경우 -> 삭제된 것으로 간주
    if (!processedSyncIds.has(syncId)) {
      Logger.log(`삭제 감지: ${itemInfo.type} - ${itemInfo.title} (ID: ${syncId})`);
      
      try {
        if (itemInfo.type === 'calendar') {
          const event = calendar.getEventById(syncId);
          if (event) {
            event.deleteEvent();
            Logger.log(`캘린더 이벤트 삭제 완료: ${itemInfo.title} (ID: ${syncId})`);
          }
        } else if (itemInfo.type === 'tasks') {
          try {
            const task = Tasks.Tasks.get(TASK_LIST_ID, syncId);
            if (task) {
              Tasks.Tasks.remove(TASK_LIST_ID, syncId);
              Logger.log(`Task 삭제 완료: ${itemInfo.title} (ID: ${syncId})`);
            }
          } catch (e) {
            // 이미 삭제된 경우(Not Found)는 무시하고, 다른 오류는 기록
            if (!e.message.toLowerCase().includes('not found')) {
              throw e; 
            }
          }
        }
      } catch (e) {
        Logger.log(`삭제 중 오류 발생: ${itemInfo.type} - ${itemInfo.title} (ID: ${syncId}): ${e.toString()}`);
      }
    }
  }

  // 6. 결과 반영 및 마무리
  // 동기화 이력 업데이트
  updateSyncHistory(currentItems);

  // Sync ID와 결과 메시지를 시트에 한 번에 기록
  syncDataRange.setValues(syncDataValues);

  Logger.log("동기화 완료.");
  SpreadsheetApp.getActiveSpreadsheet().toast("동기화가 완료되었습니다.", "완료", 5);
  SpreadsheetApp.getUi().alert('캘린더/Tasks 동기화가 완료되었습니다. 각 행의 "동기화 결과" 열과 스크립트 로그를 확인하세요.');
}

// ================================================================================
// 캘린더 이벤트 처리 함수
// ================================================================================

/**
 * 단일 캘린더 이벤트를 처리하고, 결과 ID와 메시지를 반환합니다.
 * @param {GoogleAppsScript.Calendar.Calendar} calendar 대상 캘린더 객체
 * @param {Array} rowData 시트의 행 데이터 배열 (0-indexed)
 * @param {string|null} syncId 현재 저장된 이벤트 ID
 * @param {string} timeZone 스프레드시트의 시간대
 * @param {number} rowIndexForLog 로깅을 위한 실제 시트 행 번호 (1-indexed)
 * @return {{id: string|null, message: string}} 처리 후 이벤트 ID와 결과 메시지
 */
function processCalendarEvent(calendar, rowData, syncId, timeZone, rowIndexForLog) {
  let eventIdToReturn = syncId;
  let message = "";

  // --- 1. 데이터 추출 및 기본 검증 ---
  const title = rowData[COL_TITLE - 1] ? rowData[COL_TITLE - 1].toString().trim() : '';
  if (!title) {
    return { id: syncId, message: "오류: 제목 없음. 처리 불가." };
  }

  const status = rowData[COL_STATUS - 1] ? rowData[COL_STATUS - 1].toString().trim().toLowerCase() : '';
  let event = null;

  // --- 2. 기존 이벤트 조회 ---
  if (syncId) {
    try {
      event = calendar.getEventById(syncId);
      if (event === null) { // 이벤트가 삭제되었거나 ID가 잘못된 경우
        Logger.log(`캘린더 (행 ${rowIndexForLog}, ID ${syncId}): 이벤트 ID로 조회했으나 찾을 수 없음 (삭제된 것으로 간주). ID 초기화.`);
        message = "정보: 이전 이벤트 못 찾음. ID 초기화됨.";
        // event가 null이므로 아래에서 새로 생성됨
      }
    } catch (e) {
      Logger.log(`캘린더 (행 ${rowIndexForLog}, ID ${syncId}): 기존 이벤트 조회 중 오류 - ${e.toString()}. ID 초기화.`);
      message = `오류: 기존 이벤트(${syncId}) 조회 실패. ID 초기화됨. (${e.message.substring(0,30)})`;
      event = null; // 오류 발생 시 event를 null로 설정하여 새로 생성 유도
    }
  }

  // --- 3. '취소됨' 상태 처리 (삭제) ---
  if (status === '취소됨') {
    if (event) {
      try {
        event.deleteEvent();
        Logger.log(`캘린더 (행 ${rowIndexForLog}, ID ${syncId}): 이벤트 삭제 완료 (상태: 취소됨).`);
        return { id: null, message: "성공: 이벤트 삭제됨 (취소)" };
      } catch (deleteError) {
        Logger.log(`캘린더 (행 ${rowIndexForLog}, ID ${syncId}): 이벤트 삭제 실패 - ${deleteError.toString()}`);
        return { id: syncId, message: `오류: 이벤트 삭제 실패 (${deleteError.message.substring(0,50)}...)` }; // 기존 ID 유지
      }
    } else {
      Logger.log(`캘린더 (행 ${rowIndexForLog}): 상태 '취소됨'이나 동기화된 이벤트 없음. 작업 없음.`);
      return { id: null, message: "정보: 취소됨 (삭제할 기존 이벤트 없음)" };
    }
  }

  // --- 4. 이벤트 데이터 준비 (날짜, 시간, 설명 등) ---
  
  // 날짜/시간 문자열 추출
  const startDateStr = rowData[COL_START_DATE - 1] instanceof Date ? Utilities.formatDate(rowData[COL_START_DATE - 1], timeZone, "yyyy-MM-dd") : (rowData[COL_START_DATE - 1] ? rowData[COL_START_DATE - 1].toString().trim() : '');
  let startTimeStr = rowData[COL_START_TIME - 1] instanceof Date ? Utilities.formatDate(rowData[COL_START_TIME - 1], timeZone, "HH:mm") : (rowData[COL_START_TIME - 1] ? rowData[COL_START_TIME - 1].toString().trim() : '');
  const endDateStr = rowData[COL_END_DATE - 1] instanceof Date ? Utilities.formatDate(rowData[COL_END_DATE - 1], timeZone, "yyyy-MM-dd") : (rowData[COL_END_DATE - 1] ? rowData[COL_END_DATE - 1].toString().trim() : '');
  let endTimeStr = rowData[COL_END_TIME - 1] instanceof Date ? Utilities.formatDate(rowData[COL_END_TIME - 1], timeZone, "HH:mm") : (rowData[COL_END_TIME - 1] ? rowData[COL_END_TIME - 1].toString().trim() : '');

  if (!startDateStr) {
    return { id: syncId, message: "오류: 시작일 없음. 처리 불가." };
  }

  const isAllDay = !startTimeStr && !endTimeStr; // 시트 기준 종일 이벤트 여부

  // 설명 구성
  const grade = COL_GRADE > 0 && rowData[COL_GRADE - 1] ? rowData[COL_GRADE - 1].toString().trim() : '';
  const description = COL_DESC > 0 && rowData[COL_DESC - 1] ? rowData[COL_DESC - 1].toString().trim() : '';
  const eventDescription = formatDescription(grade, description, rowData[COL_STATUS -1] ? rowData[COL_STATUS -1].toString().trim() : '');
  
  // --- 색상 결정 로직 (CALENDAR_COLOR_RULES 사용) ---
  let targetColorId = null;
  for (const rule of CALENDAR_COLOR_RULES) {
    if (rule.targetColNumber > 0 && rule.targetColNumber <= rowData.length) {
      const cellValue = rowData[rule.targetColNumber - 1] ? rowData[rule.targetColNumber - 1].toString().trim() : "";
      if (rule.searchValue && cellValue.includes(rule.searchValue)) {
        targetColorId = rule.colorId;
        Logger.log(`행 ${rowIndexForLog}: 색상 규칙 적용됨. 색상 ID: ${targetColorId}`);
        break; // 첫 번째 일치하는 규칙 사용
      }
    }
  }
  
  // 규칙이 없거나 일치하는 규칙이 없는 경우 기존 로직 (3학년 색상 ID) 사용 (하위 호환성 유지)
  if (!targetColorId && COLOR_ID_GRADE_3 && grade && grade.includes('3학년')) {
    targetColorId = COLOR_ID_GRADE_3;
    Logger.log(`행 ${rowIndexForLog}: 규칙이 없어 기본 3학년 색상 규칙 적용됨. 색상 ID: ${targetColorId}`);
  }

  // --- 날짜/시간 객체 생성 ---
  let sheetStartDateTime, sheetEndDateTime; // 시트 데이터를 기반으로 계산된 Date 객체

  try {
    if (isAllDay) {
      // 종일 이벤트: Date.UTC를 사용하여 날짜의 시작(자정)을 명확히 표현
      const dateRegex = /^\d{4}-\d{2}-\d{2}$/;
      if (!dateRegex.test(startDateStr)) throw new Error(`시작일(${startDateStr}) 형식 오류. 'yyyy-MM-dd' 필요.`);
      let parts = startDateStr.split('-');
      sheetStartDateTime = new Date(Date.UTC(parseInt(parts[0],10), parseInt(parts[1],10) - 1, parseInt(parts[2],10)));

      if (!endDateStr) { // 종료일 없으면 시작일과 동일 (하루짜리 이벤트)
        sheetEndDateTime = new Date(sheetStartDateTime.getTime());
      } else {
        if (!dateRegex.test(endDateStr)) throw new Error(`종료일(${endDateStr}) 형식 오류. 'yyyy-MM-dd' 필요.`);
        parts = endDateStr.split('-');
        sheetEndDateTime = new Date(Date.UTC(parseInt(parts[0],10), parseInt(parts[1],10) - 1, parseInt(parts[2],10)));
      }

      if (sheetEndDateTime.getTime() < sheetStartDateTime.getTime()) {
        Logger.log(`캘린더 (행 ${rowIndexForLog}): 종일 이벤트의 시트상 종료일이 시작일보다 이전. 종료일을 시작일로 조정.`);
        sheetEndDateTime = new Date(sheetStartDateTime.getTime());
      }
      
    } else {
      // 시간 지정 이벤트
      if (!startTimeStr) { // 시작 시간은 필수, 없으면 기본값 또는 오류
          Logger.log(`캘린더 (행 ${rowIndexForLog}): 시간 지정 이벤트에 시작 시간이 없음. 00:00으로 가정.`);
          startTimeStr = "00:00";
      }
      sheetStartDateTime = parseDateTimeInternal(startDateStr, startTimeStr, timeZone);

      if (!endDateStr && !endTimeStr) { // 종료 날짜/시간 모두 없으면 시작 시간 + 1시간
        sheetEndDateTime = new Date(sheetStartDateTime.getTime() + 60 * 60 * 1000);
      } else {
        const effectiveEndDateStr = endDateStr || startDateStr; // 종료일 없으면 시작일 사용
        const effectiveEndTimeStr = endTimeStr || Utilities.formatDate(new Date(sheetStartDateTime.getTime() + 60 * 60 * 1000), timeZone, "HH:mm"); // 종료시간 없으면 시작시간+1시간
        sheetEndDateTime = parseDateTimeInternal(effectiveEndDateStr, effectiveEndTimeStr, timeZone);
      }

      if (sheetEndDateTime.getTime() <= sheetStartDateTime.getTime()) {
        Logger.log(`캘린더 (행 ${rowIndexForLog}): 계산된 종료시간이 시작시간보다 이전/같음. 시작시간+1시간으로 강제 조정.`);
        sheetEndDateTime = new Date(sheetStartDateTime.getTime() + 60 * 60 * 1000);
      }
    }
    
    // 디버깅용 로그
    Logger.log(`행 ${rowIndexForLog}: Start=${Utilities.formatDate(sheetStartDateTime, timeZone, "yyyy-MM-dd HH:mm")}, End=${Utilities.formatDate(sheetEndDateTime, timeZone, "yyyy-MM-dd HH:mm")}`);

  } catch (e) {
    Logger.log(`캘린더 (행 ${rowIndexForLog}, 제목 ${title}): 날짜/시간 파싱 오류 - ${e.toString()}`);
    return { id: syncId, message: `오류: 날짜/시간 형식 (${e.message.substring(0, 70)}...)` };
  }

  // --- 5. 이벤트 업데이트 또는 생성 ---
  if (event) { // 기존 이벤트 업데이트
    let needsUpdate = false;
    
    // 제목, 설명 변경 감지 및 업데이트
    if (event.getTitle() !== title) { event.setTitle(title); needsUpdate = true; Logger.log(` - 제목 변경됨`);}
    if (event.getDescription() !== eventDescription) { event.setDescription(eventDescription); needsUpdate = true; Logger.log(` - 설명 변경됨`);}

    // 색상 변경 감지 및 업데이트
    const currentColorId = (typeof event.getColor === 'function') ? event.getColor() : null;
    if (targetColorId && currentColorId !== targetColorId) {
      if (typeof event.setColor === 'function') { event.setColor(targetColorId); needsUpdate = true; Logger.log(` - 색상 변경됨: ${targetColorId}`);}
    }
    
    // 시간/날짜 업데이트 (종일 이벤트 <-> 시간 지정 이벤트 타입 변경 포함)
    try {
      let timeChanged = false;
      const existingEventTypeIsAllDay = event.isAllDayEvent();

      if (isAllDay) { // 시트가 종일 이벤트를 원함
        let needsAllDayUpdate = false;

        // 기존 이벤트가 시간 지정이었거나, 종일 이벤트라도 날짜가 다른 경우 업데이트 필요
        if (!existingEventTypeIsAllDay) {
          needsAllDayUpdate = true;
        } else {
          // 날짜 비교 로직 (UTC 기준으로 비교)
          const existingStartStr = Utilities.formatDate(event.getAllDayStartDate(), "UTC", "yyyy-MM-dd");
          // getAllDayEndDate는 종료일 다음날 자정을 반환하므로 하루를 빼고 비교해야 할 수 있음.
          // 여기서는 시트 데이터 생성 로직과 맞춰야 함.
          
          const sheetStartStr = Utilities.formatDate(sheetStartDateTime, "UTC", "yyyy-MM-dd");
          const sheetEndStr = Utilities.formatDate(sheetEndDateTime, "UTC", "yyyy-MM-dd");

          // 복잡한 종일 이벤트 날짜 비교 로직은 생략하고, 단순화하여 시트의 날짜와 다르면 업데이트 시도
          // (Calendar API의 setAllDayDates는 매우 민감하므로 주의 필요)
          needsAllDayUpdate = true; // 안전하게 항상 업데이트 시도 (세부 비교 로직 복잡성 회피)
        }

        if (needsAllDayUpdate) {
          try {
            // 하루짜리 종일 이벤트는 특수 처리 (Calendar API에서 종료일은 exclusive)
            if (sheetStartDateTime.getTime() === sheetEndDateTime.getTime()) {
              // 종료일을 다음날로 설정 (하루 추가)
              const nextDayDate = new Date(sheetEndDateTime.getTime());
              nextDayDate.setUTCDate(nextDayDate.getUTCDate() + 1);
              
              Logger.log(`행 ${rowIndexForLog}: 하루짜리 종일 이벤트 날짜 업데이트 시도: ${sheetStartDateTime.toISOString()} to ${nextDayDate.toISOString()}`);
              event.setAllDayDates(sheetStartDateTime, nextDayDate);
            } else {
              // 여러 날짜에 걸친 종일 이벤트 (시트의 종료일이 포함되려면 +1일 필요할 수 있음)
              // 현재 코드 구조상 sheetEndDateTime은 포함되어야 할 마지막 날짜의 자정으로 가정.
              const exclusiveEndDate = new Date(sheetEndDateTime.getTime());
              exclusiveEndDate.setUTCDate(exclusiveEndDate.getUTCDate() + 1);

              Logger.log(`행 ${rowIndexForLog}: 여러 날짜 종일 이벤트 날짜 업데이트 시도: ${sheetStartDateTime.toISOString()} to ${exclusiveEndDate.toISOString()}`);
              event.setAllDayDates(sheetStartDateTime, exclusiveEndDate);
            }
            timeChanged = true;
            Logger.log(` - 종일 날짜/시간으로 변경/업데이트됨`);
          } catch (allDayError) {
            Logger.log(`행 ${rowIndexForLog}: 종일 이벤트 날짜 업데이트 오류 - ${allDayError.toString()}`);
            message += `경고: 종일 이벤트 날짜 업데이트 실패 (${allDayError.message.substring(0,30)}). `;
          }
        }
      } else {
        // 시간 지정 이벤트 업데이트
        // 기존 이벤트가 종일 이벤트였거나, 시간 지정 이벤트라도 시간이 다른 경우 업데이트 필요
        if (existingEventTypeIsAllDay || 
            event.getStartTime().getTime() !== sheetStartDateTime.getTime() || 
            event.getEndTime().getTime() !== sheetEndDateTime.getTime()) {
          
          event.setTime(sheetStartDateTime, sheetEndDateTime);
          timeChanged = true;
          Logger.log(` - 시간 지정 이벤트로 변경/업데이트됨`);
        }
      }
      
      if (timeChanged) needsUpdate = true;

    } catch (timeError) {
      Logger.log(`캘린더 (행 ${rowIndexForLog}, ID ${syncId}): 시간 업데이트 중 오류 - ${timeError.toString()}`);
      // 부분 성공 가능성: 메시지에 오류 추가하고 ID는 유지
      message += `경고: 시간 업데이트 실패 (${timeError.message.substring(0,30)}). `;
    }

    if (needsUpdate) {
      message = "성공: 이벤트 업데이트됨. " + message;
      Logger.log(`캘린더 (행 ${rowIndexForLog}, ID ${syncId}): 이벤트 업데이트 완료.`);
    } else {
      message = "정보: 변경 사항 없음. " + message;
      Logger.log(`캘린더 (행 ${rowIndexForLog}, ID ${syncId}): 시트 내용과 이벤트가 동일하여 변경 사항 없음.`);
    }
    eventIdToReturn = syncId; // 업데이트 성공/실패 여부와 관계없이 기존 ID 반환

  } else { // 새 이벤트 생성
    try {
      const eventOptions = { description: eventDescription };
      let createdEvent; 

      if (isAllDay) {
        // 종일 이벤트 생성
        if (sheetStartDateTime.getTime() === sheetEndDateTime.getTime()) {
            // 하루짜리 종일 이벤트
            Logger.log(` -> 하루짜리 종일 이벤트 생성 시도: ${title}, ${Utilities.formatDate(sheetStartDateTime, "UTC", "yyyy-MM-dd")}`);
            createdEvent = calendar.createAllDayEvent(title, sheetStartDateTime, eventOptions);
        } else {
            // 여러 날짜 종일 이벤트 (시트의 종료일이 포함되어야 함)
            const exclusiveEndDate = new Date(sheetEndDateTime.getTime());
            exclusiveEndDate.setUTCDate(exclusiveEndDate.getUTCDate() + 1);

            Logger.log(` -> 여러 날짜 종일 이벤트 생성 시도: ${title}, Start: ${Utilities.formatDate(sheetStartDateTime, "UTC", "yyyy-MM-dd")}, End (Exclusive): ${Utilities.formatDate(exclusiveEndDate, "UTC", "yyyy-MM-dd")}`);
            createdEvent = calendar.createAllDayEvent(title, sheetStartDateTime, exclusiveEndDate, eventOptions);
        }
      } else {
        // 시간 지정 이벤트 생성
        Logger.log(` -> 시간 지정 이벤트 생성 시도: ${title}, ${Utilities.formatDate(sheetStartDateTime, timeZone, "yyyy-MM-dd HH:mm")} - ${Utilities.formatDate(sheetEndDateTime, timeZone, "yyyy-MM-dd HH:mm")}`);
        createdEvent = calendar.createEvent(title, sheetStartDateTime, sheetEndDateTime, eventOptions);
      }

      // 생성된 이벤트 확인 및 후처리 (색상 적용)
      if (createdEvent) {
        eventIdToReturn = createdEvent.getId();
        message = "성공: 이벤트 생성됨";
        Logger.log(`캘린더 (행 ${rowIndexForLog}, 제목 ${title}): 새 이벤트 생성 완료 (ID: ${eventIdToReturn}).`);

        if (targetColorId && typeof createdEvent.setColor === 'function') {
          try {
            createdEvent.setColor(targetColorId);
            Logger.log(` -> 생성된 이벤트에 색상 ID '${targetColorId}' 적용 완료.`);
          } catch (colorError) {
            Logger.log(` -> 생성된 이벤트 색상 적용 실패 (ID: ${targetColorId}): ${colorError.toString()}`);
            message += ` (색상 적용 실패: ${colorError.message.substring(0,30)})`;
          }
        }
      } else {
        // API 호출은 성공했으나 이벤트 객체가 반환되지 않은 경우
        Logger.log(`캘린더 (행 ${rowIndexForLog}, 제목 ${title}): 이벤트 생성 API 호출 후 유효한 이벤트 객체를 받지 못했습니다.`);
        message = `오류: 이벤트 생성 실패 (API가 유효한 객체를 반환하지 않음)`;
        eventIdToReturn = null; 
      }

    } catch (createError) {
      Logger.log(`캘린더 (행 ${rowIndexForLog}, 제목 ${title}): 이벤트 생성 중 예외 발생 - ${createError.toString()}\nStack: ${createError.stack}`);
      message = `오류: 이벤트 생성 실패 (${createError.message.substring(0, 70)}...). ` + message;
      eventIdToReturn = null;
    }
  }
  
  return { id: eventIdToReturn, message: message.trim() };
}

// ================================================================================
// Google Tasks 처리 함수
// ================================================================================

/**
 * 단일 Google Task를 처리하고, 결과 ID와 메시지를 반환합니다.
 * @param {string} taskListId 대상 Task List ID
 * @param {Array} rowData 시트의 행 데이터 배열 (0-indexed)
 * @param {string|null} syncId 현재 저장된 Task ID
 * @param {string} timeZone 스프레드시트의 시간대 (Task 마감일 파싱 시 참고용, 실제 API는 UTC)
 * @param {number} rowIndexForLog 로깅을 위한 실제 시트 행 번호 (1-indexed)
 * @return {{id: string|null, message: string}} 처리 후 Task ID와 결과 메시지
 */
function processTask(taskListId, rowData, syncId, timeZone, rowIndexForLog) {
  let taskIdToReturn = syncId;
  let message = "";

  // 1. 데이터 추출
  const title = rowData[COL_TITLE - 1] ? rowData[COL_TITLE - 1].toString().trim() : '';
  if (!title) {
    return { id: syncId, message: "오류: 제목 없음. 처리 불가." };
  }

  const dueDateStr = COL_DUE_DATE > 0 && rowData[COL_DUE_DATE - 1] instanceof Date ? Utilities.formatDate(rowData[COL_DUE_DATE - 1], timeZone, "yyyy-MM-dd") : (COL_DUE_DATE > 0 && rowData[COL_DUE_DATE - 1] ? rowData[COL_DUE_DATE - 1].toString().trim() : '');
  const grade = COL_GRADE > 0 && rowData[COL_GRADE - 1] ? rowData[COL_GRADE - 1].toString().trim() : '';
  const description = COL_DESC > 0 && rowData[COL_DESC - 1] ? rowData[COL_DESC - 1].toString().trim() : '';
  const statusSheet = COL_STATUS > 0 && rowData[COL_STATUS - 1] ? rowData[COL_STATUS - 1].toString().trim().toLowerCase() : '';

  const taskNotes = formatDescription(grade, description, rowData[COL_STATUS -1] ? rowData[COL_STATUS -1].toString().trim() : '');
  let task = null;

  // 2. 기존 Task 조회
  if (syncId) {
    try {
      task = Tasks.Tasks.get(taskListId, syncId);
    } catch (e) {
      if (e.message && e.message.toLowerCase().includes('not found')) {
         Logger.log(`Task (행 ${rowIndexForLog}, ID ${syncId}): Task 못 찾음 (삭제된 것으로 간주). ID 초기화.`);
         message = "정보: 이전 Task 못 찾음. ID 초기화됨.";
      } else {
         Logger.log(`Task (행 ${rowIndexForLog}, ID ${syncId}): 기존 Task 조회 오류 - ${e.toString()}. ID 초기화.`);
         message = `오류: 기존 Task(${syncId}) 조회 실패. ID 초기화됨. (${e.message.substring(0,30)})`;
      }
      task = null; // 조회 실패 시 task 객체 null
      syncId = null; // ID도 null로 하여 새로 생성 유도
    }
  }

  // 3. Task 데이터 준비 (API용)
  let newTaskData = { title: title, notes: taskNotes };
  
  // 마감일 설정 (Tasks API는 RFC3339 UTC 형식 요구)
  if (dueDateStr) {
    try {
      const dateRegex = /^\d{4}-\d{2}-\d{2}$/;
      if (!dateRegex.test(dueDateStr)) throw new Error(`마감일(${dueDateStr}) 형식 오류. 'yyyy-MM-dd' 필요.`);
      const parts = dueDateStr.split('-');
      // 날짜만 있는 경우, 해당 날짜의 자정 UTC로 설정.
      const dueDateObj = new Date(Date.UTC(parseInt(parts[0],10), parseInt(parts[1],10) - 1, parseInt(parts[2],10)));
      newTaskData.due = dueDateObj.toISOString(); // "YYYY-MM-DDTHH:00:00.000Z"
    } catch (e) {
      Logger.log(`Task (행 ${rowIndexForLog}, 제목 ${title}): 잘못된 마감일 형식(${dueDateStr}). 마감일 미설정. 오류: ${e.toString()}`);
      message += `경고: 마감일 형식 오류(${dueDateStr}). `;
      newTaskData.due = task ? task.due : null; // 기존 값 유지 시도 (업데이트 시)
    }
  } else {
    newTaskData.due = null; // 마감일 없으면 null로 명시 (제거 의도)
  }

  // 상태 설정
  if (statusSheet === '운영완료' || statusSheet === '취소됨') {
      newTaskData.status = 'completed';
  } else {
      newTaskData.status = 'needsAction';
  }

  // 4. Task 업데이트 또는 생성
  if (task) { // Task 업데이트
    let patchData = {};
    
    // 제목, 설명 변경 확인
    if (task.title !== newTaskData.title) patchData.title = newTaskData.title;
    if ((task.notes || '') !== (newTaskData.notes || '')) patchData.notes = newTaskData.notes || ""; 

    // 마감일 변경 확인 (YYYY-MM-DD 부분만 비교)
    const existingDue = task.due ? task.due.substring(0, 10) : null; 
    const newDue = newTaskData.due ? newTaskData.due.substring(0, 10) : null; 

    if (existingDue !== newDue) {
        patchData.due = newTaskData.due; // null일 경우 마감일 제거
    }
    
    // 상태 변경 확인
    if (task.status !== newTaskData.status) patchData.status = newTaskData.status;

    // 변경 사항이 있을 경우에만 Patch API 호출
    if (Object.keys(patchData).length > 0) {
      try {
        Tasks.Tasks.patch(patchData, taskListId, syncId); 
        message = "성공: Task 업데이트됨. " + message;
        Logger.log(`Task (행 ${rowIndexForLog}, ID ${syncId}): 업데이트 완료.`);
      } catch (patchError) {
        Logger.log(`Task (행 ${rowIndexForLog}, ID ${syncId}): 업데이트 실패 - ${patchError.toString()}`);
        message = `오류: Task 업데이트 실패 (${patchError.message.substring(0,50)}...). ` + message;
      }
    } else {
      message = "정보: 변경 사항 없음. " + message;
      Logger.log(`Task (행 ${rowIndexForLog}, ID ${syncId}): 변경 사항 없음.`);
    }
    taskIdToReturn = syncId; // 업데이트 시도 후 ID는 유지

  } else { // Task 생성
    try {
      const createdTask = Tasks.Tasks.insert(newTaskData, taskListId);
      taskIdToReturn = createdTask.id;
      message = "성공: Task 생성됨. " + message; // 이전 조회 실패 메시지가 있다면 여기에 합쳐짐
      Logger.log(`Task (행 ${rowIndexForLog}, 제목 ${title}): 새 Task 생성 완료 (ID: ${taskIdToReturn}).`);
    } catch (insertError) {
      Logger.log(`Task (행 ${rowIndexForLog}, 제목 ${title}): 생성 실패 - ${insertError.toString()}`);
      message = `오류: Task 생성 실패 (${insertError.message.substring(0,50)}...). ` + message;
      taskIdToReturn = null;
    }
  }
  
  return { id: taskIdToReturn, message: message.trim() };
}

// ================================================================================
// 유틸리티 및 헬퍼 함수들
// ================================================================================

/**
 * 날짜 문자열(YYYY-MM-DD)과 시간 문자열(HH:mm)을 합쳐 Date 객체를 생성합니다. (내부 유틸리티)
 * @param {string} dateStr 'YYYY-MM-DD'
 * @param {string} timeStr 'HH:mm'
 * @param {string} timeZone 스프레드시트의 시간대
 * @return {Date} 파싱된 Date 객체
 * @throws {Error} 날짜/시간 형식이 잘못되었거나 파싱에 실패한 경우
 */
function parseDateTimeInternal(dateStr, timeStr, timeZone) {
  if (!dateStr || !/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) {
    throw new Error(`잘못된 날짜 형식: '${dateStr}'. 'YYYY-MM-dd' 형식이 필요합니다.`);
  }
  if (!timeStr || !/^\d{2}:\d{2}(:\d{2})?$/.test(timeStr)) { // HH:mm 또는 HH:mm:ss 허용
    throw new Error(`잘못된 시간 형식: '${timeStr}'. 'HH:mm' 형식이 필요합니다.`);
  }

  const dateTimeString = `${dateStr} ${timeStr}`;
  try {
    // Utilities.parseDate를 사용하여 지정된 시간대의 날짜/시간 문자열을 Date 객체로 변환
    const parsedDate = Utilities.parseDate(dateTimeString, timeZone, "yyyy-MM-dd HH:mm");
    if (isNaN(parsedDate.getTime())) {
        throw new Error(`'${dateTimeString}'을(를) 유효한 날짜/시간으로 변환할 수 없습니다.`);
    }
    Logger.log(`parseDateTimeInternal: input='${dateTimeString}', timeZone='${timeZone}', output (ISO)='${parsedDate.toISOString()}'`);
    return parsedDate;
  } catch (e) {
     throw new Error(`날짜/시간 파싱 실패 ('${dateTimeString}', 시간대: ${timeZone}): ${e.message || e.toString()}`);
  }
}

/**
 * 캘린더 이벤트 또는 Task 노트에 사용될 설명 문자열을 포맷합니다.
 * @param {string} grade 학년 정보
 * @param {string} description 사용자 입력 설명
 * @param {string} status 시트의 원본 상태값
 * @return {string} 조합된 설명 문자열
 */
function formatDescription(grade, description, status) {
  let parts = [];
  if (grade) parts.push(`학년: ${grade}`);
  if (description) parts.push(`설명: ${description}`);
  if (status) parts.push(`상태 (시트): ${status}`); // 명시적으로 시트 상태임을 표시
  
  return parts.join('\n');
}

// ================================================================================
// 동기화 이력 관리 함수들 (삭제 감지용)
// ================================================================================

/**
 * 삭제된 항목을 확인합니다.
 * @param {Set} currentSyncIds 현재 시트에 있는 Sync ID들의 Set
 * @return {number} 삭제 예정인 항목의 개수
 */
function checkDeletedItems(currentSyncIds) {
  const syncHistory = getSyncHistory();
  let deletedCount = 0;
  
  // 이력에는 있지만 현재 시트에는 없는 ID를 찾습니다.
  for (const [syncId, itemInfo] of Object.entries(syncHistory)) {
    if (syncId && !currentSyncIds.has(syncId)) {
      deletedCount++;
    }
  }
  
  return deletedCount;
}

/**
 * '동기화이력' 시트에서 동기화 이력을 가져옵니다.
 * @return {Object} syncId를 키로 하는 이력 객체 { syncId: { type, title, lastUpdated } }
 */
function getSyncHistory() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let historySheet = ss.getSheetByName(SYNC_HISTORY_SHEET_NAME);
  
  if (!historySheet) {
    // 이력 시트가 없으면 새로 생성하고 숨김 처리
    Logger.log(`'${SYNC_HISTORY_SHEET_NAME}' 시트가 없어 새로 생성합니다.`);
    historySheet = ss.insertSheet(SYNC_HISTORY_SHEET_NAME);
    historySheet.getRange(1, 1, 1, 4).setValues([['Sync ID', 'Type', 'Title', 'Last Updated']]);
    historySheet.hideSheet(); // 사용자에게 보이지 않도록 숨김
    return {};
  }
  
  const data = historySheet.getDataRange().getValues();
  const history = {};
  
  for (let i = 1; i < data.length; i++) { // 1행(헤더) 제외
    const syncId = data[i][0] ? data[i][0].toString().trim() : null;
    if (syncId) {
      history[syncId] = {
        type: data[i][1],
        title: data[i][2],
        lastUpdated: data[i][3]
      };
    }
  }
  
  return history;
}

/**
 * '동기화이력' 시트를 현재 동기화된 항목들로 업데이트합니다.
 * @param {Array} currentItems 현재 시트의 항목들 [{syncId, type, title}, ...]
 */
function updateSyncHistory(currentItems) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let historySheet = ss.getSheetByName(SYNC_HISTORY_SHEET_NAME);
  
  if (!historySheet) {
    // 시트가 없는 경우 (사용자가 실수로 삭제했을 경우 대비)
    historySheet = ss.insertSheet(SYNC_HISTORY_SHEET_NAME);
    historySheet.hideSheet();
  }
  
  // 기존 내용을 지우고 헤더 설정
  historySheet.clear();
  historySheet.getRange(1, 1, 1, 4).setValues([['Sync ID', 'Type', 'Title', 'Last Updated']]);
  
  // 현재 동기화된 데이터로 채우기
  if (currentItems.length > 0) {
    const historyData = currentItems.map(item => [
      item.syncId,
      item.type,
      item.title,
      new Date() // 마지막 업데이트 시간 기록
    ]);
    historySheet.getRange(2, 1, historyData.length, 4).setValues(historyData);
  }
}

// ================================================================================
// 시간 기반 트리거 관련 함수 (자동 동기화용)
// ================================================================================

/**
 * 시간 기반 트리거에서 호출될 함수입니다. (UI 알림 없음)
 * 사용자가 직접 실행하지 않고, 설정된 주기에 따라 자동으로 실행됩니다.
 */
function triggerSync() {
  Logger.log("시간 기반 트리거: 자동 동기화 프로세스 시작");
  try {
    // 1. 설정 로드
    if (!loadSettings()) {
      Logger.log("triggerSync: 설정 로드 실패. 동기화를 중단합니다.");
      // 트리거 실행 시 오류가 발생하면 관리자(스크립트 소유자)에게 이메일 알림
      MailApp.sendEmail(Session.getEffectiveUser().getEmail(), 
                        "[자동화 스크립트] 동기화 오류 알림: 설정 로드 실패", 
                        "자동 동기화 중 설정을 불러오는 데 실패했습니다. Apps Script 로그 및 '설정' 시트를 확인해주세요.");
      return;
    }
    Logger.log("triggerSync: 설정 로드 완료.");

    // 2. 동기화 실행 (UI 없는 버전)
    syncSheetToApis_forTrigger(); 

    Logger.log("시간 기반 트리거: 자동 동기화 프로세스 완료");

  } catch (e) {
    // 예측하지 못한 심각한 오류 발생 시
    Logger.log(`triggerSync: 동기화 중 심각한 오류 발생 - ${e.toString()}\nStack: ${e.stack}`);
    MailApp.sendEmail(Session.getEffectiveUser().getEmail(),
                      "[자동화 스크립트] 심각한 오류 발생",
                      `자동 동기화 중 예측하지 못한 오류가 발생했습니다.\n\n오류 내용: ${e.toString()}\n\nApps Script 실행 기록 로그를 확인해주세요.`);
  }
}

/**
 * (트리거용) 구글 시트 데이터를 구글 캘린더 및 구글 Tasks와 동기화합니다.
 * syncSheetToApis 함수에서 UI 관련 코드(alert, toast 등)가 제거된 버전입니다.
 */
function syncSheetToApis_forTrigger() {
  // 설정은 triggerSync 함수에서 이미 로드되었으므로 여기서는 loadSettings() 호출 불필요.
  // CALENDAR_ID, SHEET_NAME 등의 전역 변수는 이미 채워져 있음.

  Logger.log("syncSheetToApis_forTrigger: 동기화 작업 시작");

  // Tasks API 고급 서비스 활성화 확인 (UI 알림 없음)
  if (typeof Tasks === 'undefined' || !Tasks.Tasks || !Tasks.Tasklists) {
    const msg = "Google Tasks API 고급 서비스가 활성화되지 않았습니다. 관리자 확인 필요.";
    Logger.log(msg);
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    Logger.log(`데이터 시트 '${SHEET_NAME}'를 찾을 수 없습니다.`);
    return;
  }

  const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
  if (!calendar) {
    const msg = `캘린더 ID '${CALENDAR_ID}'를 찾을 수 없거나 접근 권한이 없습니다.`;
    Logger.log(msg);
    return;
  }

  try {
    Tasks.Tasklists.get(TASK_LIST_ID);
  } catch (e) {
    const msg = `Task List ID '${TASK_LIST_ID}'에 접근할 수 없습니다. 오류: ${e.toString()}`;
    Logger.log(msg);
    return;
  }

  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  const numDataRows = values.length - HEADER_ROW;

  if (numDataRows <= 0) {
      Logger.log("syncSheetToApis_forTrigger: 동기화할 데이터가 없습니다.");
      return;
  }

  // Sync ID 및 결과 메시지 기록을 위한 열 번호 유효성 확인 (UI 알림 없음)
  if (isNaN(COL_SYNC_ID) || COL_SYNC_ID < 1 || isNaN(COL_SYNC_RESULT) || COL_SYNC_RESULT < 1) {
      Logger.log("설정 오류: 'Sync ID 열 번호' 또는 '동기화 결과 열 번호'가 올바르게 설정되지 않았습니다.");
      return;
  }
  
  // 동기화 데이터 범위 설정 및 읽기 (syncSheetToApis와 동일)
  const firstCol = Math.min(COL_SYNC_ID, COL_SYNC_RESULT);
  const lastCol = Math.max(COL_SYNC_ID, COL_SYNC_RESULT);
  const numColsForSyncData = lastCol - firstCol + 1;
  
  const syncDataRange = sheet.getRange(DATA_START_ROW, firstCol, numDataRows, numColsForSyncData);
  const syncDataValues = syncDataRange.getValues();

  const syncIdColIndexInArray = COL_SYNC_ID - firstCol;
  const syncResultColIndexInArray = COL_SYNC_RESULT - firstCol;

  // 동기화 이력 준비
  const syncHistory = getSyncHistory();
  const processedSyncIds = new Set();
  const currentItems = [];

  Logger.log(`syncSheetToApis_forTrigger: 총 ${numDataRows}개 행 처리 예정.`);

  // 각 행 데이터 처리 (syncSheetToApis와 동일 로직)
  for (let i = HEADER_ROW; i < values.length; i++) {
    const rowData = values[i];
    const rowIndexInSheet = i + 1;
    const dataRowIndex = i - HEADER_ROW;

    if (!syncDataValues[dataRowIndex]) {
        Logger.log(`오류: 행 ${rowIndexInSheet}의 syncDataValues 인덱스 접근 불가.`);
         syncDataValues[dataRowIndex] = new Array(numColsForSyncData).fill('');
    }

    let currentSyncId = syncDataValues[dataRowIndex][syncIdColIndexInArray] ? syncDataValues[dataRowIndex][syncIdColIndexInArray].toString().trim() : null;
    let resultMessage = "";

    const title = rowData[COL_TITLE - 1] ? rowData[COL_TITLE - 1].toString().trim() : '';
    if (!title) {
      resultMessage = "오류: 제목 없음";
      Logger.log(`행 ${rowIndexInSheet}: ${resultMessage}. 건너뜁니다.`);
      syncDataValues[dataRowIndex][syncResultColIndexInArray] = resultMessage;
      continue;
    }

    const type = rowData[COL_TYPE - 1] ? rowData[COL_TYPE - 1].toString().trim().toLowerCase() : '';
    Logger.log(`syncSheetToApis_forTrigger: 행 ${rowIndexInSheet} 처리 중: 구분='${type}', 제목='${title}', SyncID='${currentSyncId || '없음'}'`);

    if (currentSyncId) {
      processedSyncIds.add(currentSyncId);
    }

    try {
      let syncResult = { id: currentSyncId, message: "정보: 처리되지 않음" };

      if (type === 'calendar') {
        syncResult = processCalendarEvent(calendar, rowData, currentSyncId, ss.getSpreadsheetTimeZone(), rowIndexInSheet);
      } else if (type === 'tasks') {
        syncResult = processTask(TASK_LIST_ID, rowData, currentSyncId, ss.getSpreadsheetTimeZone(), rowIndexInSheet);
      } else {
        resultMessage = `경고: 알 수 없는 구분 값 '${rowData[COL_TYPE - 1]}'`;
        Logger.log(`행 ${rowIndexInSheet}: ${resultMessage}. 건너뜁니다.`);
        syncResult = { id: null, message: resultMessage };
      }

      syncDataValues[dataRowIndex][syncIdColIndexInArray] = syncResult.id;
      syncDataValues[dataRowIndex][syncResultColIndexInArray] = syncResult.message;

      if (syncResult.id) {
        currentItems.push({
          syncId: syncResult.id,
          type: type,
          title: title
        });
        if (syncResult.id !== currentSyncId) {
          processedSyncIds.add(syncResult.id);
        }
      }

    } catch (e) {
      resultMessage = `심각한 오류: ${e.message.substring(0, 100)}... (로그 확인)`;
      Logger.log(`행 ${rowIndexInSheet} 처리 중 심각한 오류: ${e.toString()}\nStack: ${e.stack}`);
      syncDataValues[dataRowIndex][syncResultColIndexInArray] = resultMessage;
    }
  }

  // 시트에서 삭제된 항목들 처리 (syncSheetToApis와 동일 로직)
  Logger.log("syncSheetToApis_forTrigger: 삭제된 항목 확인 중...");
  for (const [syncId, itemInfo] of Object.entries(syncHistory)) {
    if (!processedSyncIds.has(syncId)) {
      Logger.log(`syncSheetToApis_forTrigger: 삭제 감지: ${itemInfo.type} - ${itemInfo.title} (ID: ${syncId})`);
      
      try {
        if (itemInfo.type === 'calendar') {
          const event = calendar.getEventById(syncId);
          if (event) {
            event.deleteEvent();
            Logger.log(`syncSheetToApis_forTrigger: 캘린더 이벤트 삭제 완료: ${itemInfo.title} (ID: ${syncId})`);
          }
        } else if (itemInfo.type === 'tasks') {
          try {
            const task = Tasks.Tasks.get(TASK_LIST_ID, syncId);
            if (task) {
              Tasks.Tasks.remove(TASK_LIST_ID, syncId);
              Logger.log(`syncSheetToApis_forTrigger: Task 삭제 완료: ${itemInfo.title} (ID: ${syncId})`);
            }
          } catch (e) {
            if (!e.message.toLowerCase().includes('not found')) {
              throw e;
            }
          }
        }
      } catch (e) {
        Logger.log(`syncSheetToApis_forTrigger: 삭제 중 오류 발생: ${itemInfo.type} - ${itemInfo.title} (ID: ${syncId}): ${e.toString()}`);
      }
    }
  }

  // 동기화 이력 업데이트 및 시트 결과 반영
  updateSyncHistory(currentItems);
  syncDataRange.setValues(syncDataValues);

  Logger.log("syncSheetToApis_forTrigger: 동기화 작업 완료.");
}