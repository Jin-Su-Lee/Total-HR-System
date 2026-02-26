// ============================================================
// 지수회계법인 — 환급사업부 스텝관리 시스템
// Code.gs (백엔드)
// ============================================================

// ============ 전역 상수 ============

var MASTER_SPREADSHEET_ID = '1iBSDGrjV_SMX4Uftt3PsQoCUwZP66WlZ6fTndEUHlmY';
var LINKED_ATTENDANCE_SS_ID = '1MegbwXBYrYn4gM1u7SC6Jj8Go0zCUYFCHIakjDo1TVs';

var SHEET_NAMES = {
  STAFF: '스텝',
  EVALUATION: '성과평가',
  CONTRACT: '계약이력',
  ADMIN: '관리자',
  SETTINGS: '설정',
  LOG: '로그'
};

var STATUS_LIST = [
  '후보등록',
  '서류심사',
  '면접',
  '합격',
  '교육',
  '근무중',
  '계약종료',
  '재계약'
];

var STATUS_TRANSITIONS = {
  '후보등록': ['서류심사'],
  '서류심사': ['면접', '후보등록'],
  '면접': ['합격', '서류심사'],
  '합격': ['교육'],
  '교육': ['근무중'],
  '근무중': ['계약종료', '재계약'],
  '계약종료': ['재계약'],
  '재계약': ['교육', '근무중']
};

var ROLE_LEVELS = {
  '최고관리자': 4,
  '총괄관리자': 3,
  '조직관리자': 2,
  '직원': 1
};

var EVAL_CRITERIA = ['정확성', '속도', '협업', '출석', '태도']; // 기본 평가항목 (전 팀 공통)

var DDAY_ALERTS = [30, 14, 7, 3]; // 계약만료 알림 일수

// ============ 시트 컬럼 정의 ============

var STAFF_COLUMNS = [
  '스텝ID', '이름', '연락처', '이메일', '생년월일',
  '성별', '주소', '지원경로', '조직', '직무',
  '상태', '계약시작일', '계약종료일', '급여형태', '급여액',
  '은행명', '계좌번호', '비고', '이력서URL', '등록일', '수정일'
];
// 21 컬럼

var RESUME_FOLDER_NAME = '스텝관리_이력서';

var EVALUATION_COLUMNS = [
  '평가ID', '스텝ID', '스텝이름', '조직', '평가기간',
  '정확성', '속도', '협업', '출석', '태도',
  '총점', '평균', '등급', '코멘트', '평가자', '평가일', '추가항목'
];
// 17 컬럼 — 추가항목: 팀별 추가 평가기준 JSON (예: {"일 평균 처리개수": 4})

var CONTRACT_COLUMNS = [
  '계약ID', '스텝ID', '스텝이름', '계약구분', '계약시작일',
  '계약종료일', '급여형태', '급여액', '계약상태', '비고',
  '등록일', '등록자'
];

var ADMIN_COLUMNS = ['이메일', '이름', '역할', '조직', '상태', '등록일'];

var SETTINGS_COLUMNS = ['키', '값', '설명', '수정일'];

var LOG_COLUMNS = ['로그ID', '일시', '사용자', '작업유형', '대상', '상세내용'];

// ============ 실행단위 캐싱 ============

var _ssCache = {};
var _sheetCache = {};
var _masterSS = null;
var _adminCache = null;

// ============ 헬퍼 함수 ============

function getMasterSS() {
  if (!_masterSS) {
    _masterSS = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID);
  }
  return _masterSS;
}

function getSheet(sheetName) {
  if (!_sheetCache[sheetName]) {
    _sheetCache[sheetName] = getMasterSS().getSheetByName(sheetName);
  }
  return _sheetCache[sheetName];
}

function generateId_(sheet, prefix) {
  var data = sheet.getDataRange().getValues();
  var maxNum = 0;
  for (var i = 1; i < data.length; i++) {
    var id = data[i][0].toString();
    if (id.startsWith(prefix + '-')) {
      var num = parseInt(id.split('-')[1]);
      if (num > maxNum) maxNum = num;
    }
  }
  return prefix + '-' + String(maxNum + 1).padStart(3, '0');
}

function formatDate_(date) {
  if (!date) return '';
  if (typeof date === 'string') return date;
  return Utilities.formatDate(date, 'Asia/Seoul', 'yyyy-MM-dd');
}

function parseDate_(dateStr) {
  if (!dateStr) return null;
  if (dateStr instanceof Date) return dateStr;
  var parts = dateStr.split('-');
  if (parts.length === 3) {
    return new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
  }
  return null;
}

function getNow_() {
  return Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
}

function getToday_() {
  return Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
}

function writeLog_(user, action, target, detail) {
  try {
    var sheet = getSheet(SHEET_NAMES.LOG);
    var logId = generateId_(sheet, 'LOG');
    var row = [logId, getNow_(), user, action, target, detail];
    sheet.appendRow(row);
  } catch (e) {
    Logger.log('로그 기록 실패: ' + e.message);
  }
}

// ============ 웹앱 진입점 ============

function doGet(e) {
  var user = Session.getActiveUser().getEmail();
  if (!isAuthorizedUser(user)) {
    return HtmlService.createHtmlOutput(
      '<div style="text-align:center;padding:60px;font-family:sans-serif;">' +
      '<h2>접근 권한이 없습니다</h2>' +
      '<p>관리자에게 문의하세요.</p>' +
      '<p style="color:#888;">' + user + '</p></div>'
    );
  }

  var template = HtmlService.createTemplateFromFile('Index');
  template.userEmail = user;
  template.userRole = getUserRole(user);
  template.userName = getUserName(user);
  template.userOrg = getUserOrg(user);

  return template.evaluate()
    .setTitle('지수회계법인 - 스텝관리')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ============ 인증/권한 ============

function getAdminList_() {
  if (_adminCache) return _adminCache;
  var sheet = getSheet(SHEET_NAMES.ADMIN);
  var data = sheet.getDataRange().getValues();
  var admins = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) {
      admins.push({
        email: data[i][0].toString().trim().toLowerCase(),
        name: data[i][1] || '',
        role: data[i][2] || '직원',
        org: data[i][3] || '',
        status: data[i][4] || '활성'
      });
    }
  }
  _adminCache = admins;
  return admins;
}

function isAuthorizedUser(email) {
  if (!email) return false;
  // @zsoo.kr 도메인 확인
  if (!email.toLowerCase().endsWith('@zsoo.kr')) return false;

  var admins = getAdminList_();
  for (var i = 0; i < admins.length; i++) {
    if (admins[i].email === email.toLowerCase() && admins[i].status === '활성') {
      return true;
    }
  }
  return false;
}

function getUserRole(email) {
  if (!email) return '직원';
  var admins = getAdminList_();
  for (var i = 0; i < admins.length; i++) {
    if (admins[i].email === email.toLowerCase()) {
      return admins[i].role || '직원';
    }
  }
  return '직원';
}

function getUserName(email) {
  if (!email) return '';
  var admins = getAdminList_();
  for (var i = 0; i < admins.length; i++) {
    if (admins[i].email === email.toLowerCase()) {
      return admins[i].name || email.split('@')[0];
    }
  }
  return email.split('@')[0];
}

function getUserOrg(email) {
  if (!email) return '';
  var admins = getAdminList_();
  for (var i = 0; i < admins.length; i++) {
    if (admins[i].email === email.toLowerCase()) {
      return admins[i].org || '';
    }
  }
  return '';
}

function hasPermission_(userRole, requiredLevel) {
  var userLevel = ROLE_LEVELS[userRole] || 1;
  return userLevel >= requiredLevel;
}

// 멀티 조직 지원 — 쉼표 구분 문자열 → 배열
function getUserOrgs_(email) {
  var orgStr = getUserOrg(email);
  if (!orgStr) return [];
  return orgStr.split(',').map(function(o) { return o.trim(); });
}

// 조직 접근 권한 확인 (최고관리자/총괄관리자는 전체 접근)
function canAccessOrg_(userRole, userOrgs, targetOrg) {
  if (userRole === '최고관리자' || userRole === '총괄관리자') return true;
  if (!targetOrg) return true;
  for (var i = 0; i < userOrgs.length; i++) {
    if (userOrgs[i] === targetOrg) return true;
  }
  return false;
}

// ============ 초기화 ============

function initializeSheets() {
  var ss = getMasterSS();

  // 스텝 시트
  var staffSheet = ss.getSheetByName(SHEET_NAMES.STAFF);
  if (!staffSheet) {
    staffSheet = ss.insertSheet(SHEET_NAMES.STAFF);
    staffSheet.appendRow(STAFF_COLUMNS);
    staffSheet.setFrozenRows(1);
    staffSheet.getRange(1, 1, 1, STAFF_COLUMNS.length)
      .setFontWeight('bold')
      .setBackground('#4472C4')
      .setFontColor('#FFFFFF');
    staffSheet.setColumnWidth(1, 100);  // 스텝ID
    staffSheet.setColumnWidth(2, 80);   // 이름
    staffSheet.setColumnWidth(7, 200);  // 주소
    staffSheet.setColumnWidth(11, 80);  // 상태
  }

  // 성과평가 시트
  var evalSheet = ss.getSheetByName(SHEET_NAMES.EVALUATION);
  if (!evalSheet) {
    evalSheet = ss.insertSheet(SHEET_NAMES.EVALUATION);
    evalSheet.appendRow(EVALUATION_COLUMNS);
    evalSheet.setFrozenRows(1);
    evalSheet.getRange(1, 1, 1, EVALUATION_COLUMNS.length)
      .setFontWeight('bold')
      .setBackground('#548235')
      .setFontColor('#FFFFFF');
  }

  // 계약이력 시트
  var contractSheet = ss.getSheetByName(SHEET_NAMES.CONTRACT);
  if (!contractSheet) {
    contractSheet = ss.insertSheet(SHEET_NAMES.CONTRACT);
    contractSheet.appendRow(CONTRACT_COLUMNS);
    contractSheet.setFrozenRows(1);
    contractSheet.getRange(1, 1, 1, CONTRACT_COLUMNS.length)
      .setFontWeight('bold')
      .setBackground('#BF8F00')
      .setFontColor('#FFFFFF');
  }

  // 관리자 시트
  var adminSheet = ss.getSheetByName(SHEET_NAMES.ADMIN);
  if (!adminSheet) {
    adminSheet = ss.insertSheet(SHEET_NAMES.ADMIN);
    adminSheet.appendRow(ADMIN_COLUMNS);
    adminSheet.setFrozenRows(1);
    adminSheet.getRange(1, 1, 1, ADMIN_COLUMNS.length)
      .setFontWeight('bold')
      .setBackground('#C00000')
      .setFontColor('#FFFFFF');
  }

  // 설정 시트
  var settingsSheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet(SHEET_NAMES.SETTINGS);
    settingsSheet.appendRow(SETTINGS_COLUMNS);
    settingsSheet.setFrozenRows(1);
    settingsSheet.getRange(1, 1, 1, SETTINGS_COLUMNS.length)
      .setFontWeight('bold')
      .setBackground('#7030A0')
      .setFontColor('#FFFFFF');
    // 기본 설정값 삽입
    var defaults = [
      ['DDAY_ALERT_DAYS', '30,14,7,3', '계약만료 알림 일수', getToday_()],
      ['DEFAULT_ORG', '환급사업부', '기본 조직', getToday_()],
      ['EVAL_SCALE', '5', '평가 최대 점수', getToday_()]
    ];
    for (var i = 0; i < defaults.length; i++) {
      settingsSheet.appendRow(defaults[i]);
    }
  }

  // 로그 시트
  var logSheet = ss.getSheetByName(SHEET_NAMES.LOG);
  if (!logSheet) {
    logSheet = ss.insertSheet(SHEET_NAMES.LOG);
    logSheet.appendRow(LOG_COLUMNS);
    logSheet.setFrozenRows(1);
    logSheet.getRange(1, 1, 1, LOG_COLUMNS.length)
      .setFontWeight('bold')
      .setBackground('#404040')
      .setFontColor('#FFFFFF');
  }

  Logger.log('시트 초기화 완료: 6개 시트 생성/확인');
  return { success: true, message: '시트 초기화가 완료되었습니다.' };
}

function seedAdminData() {
  var user = Session.getActiveUser().getEmail();
  var sheet = getSheet(SHEET_NAMES.ADMIN);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim().toLowerCase() === user.toLowerCase()) {
      Logger.log('이미 관리자로 등록됨: ' + user);
      return { success: true, message: '이미 관리자로 등록되어 있습니다.' };
    }
  }
  var row = [user, user.split('@')[0], '최고관리자', '', '활성', getToday_()];
  sheet.getRange(sheet.getLastRow() + 1, 1, 1, ADMIN_COLUMNS.length).setValues([row]);
  Logger.log('관리자 시드 완료: ' + user);
  return { success: true, message: user + ' 최고관리자로 등록되었습니다.' };
}

// ============ 초기 데이터 로드 ============

function getInitialData() {
  try {
    var user = Session.getActiveUser().getEmail();
    var role = getUserRole(user);
    var org = getUserOrg(user);

    var staffList = getStaffList({});
    var statusCounts = {};
    for (var s = 0; s < STATUS_LIST.length; s++) {
      statusCounts[STATUS_LIST[s]] = 0;
    }
    if (staffList.success) {
      for (var i = 0; i < staffList.data.length; i++) {
        var st = staffList.data[i].status;
        if (statusCounts.hasOwnProperty(st)) {
          statusCounts[st]++;
        }
      }
    }

    // D-day 알림 대상
    var alerts = getDdayAlerts_();

    return {
      success: true,
      data: {
        staffList: staffList.success ? staffList.data : [],
        statusCounts: statusCounts,
        statusList: STATUS_LIST,
        statusTransitions: STATUS_TRANSITIONS,
        alerts: alerts,
        userRole: role,
        userOrg: org,
        totalCount: staffList.success ? staffList.data.length : 0
      }
    };
  } catch (e) {
    Logger.log('getInitialData 오류: ' + e.message);
    return { success: false, message: '초기 데이터 로드 실패: ' + e.message };
  }
}

// ============ 스텝 CRUD ============

function getStaffList(filters) {
  try {
    var user = Session.getActiveUser().getEmail();
    var role = getUserRole(user);
    var userOrgs = getUserOrgs_(user);

    var sheet = getSheet(SHEET_NAMES.STAFF);
    var data = sheet.getDataRange().getValues();
    var staffList = [];

    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;

      var staff = rowToStaffObject_(data[i]);

      // 조직관리자 강제 필터링
      if (!canAccessOrg_(role, userOrgs, staff.org)) continue;

      // 필터 적용
      if (filters) {
        if (filters.status && staff.status !== filters.status) continue;
        if (filters.org && staff.org !== filters.org) continue;
        if (filters.keyword) {
          var kw = filters.keyword.toLowerCase();
          var searchable = (staff.name + staff.phone + staff.email + staff.org).toLowerCase();
          if (searchable.indexOf(kw) === -1) continue;
        }
      }

      staff.dday = calculateDday_(staff.contractEnd);
      staffList.push(staff);
    }

    return { success: true, data: staffList };
  } catch (e) {
    Logger.log('getStaffList 오류: ' + e.message);
    return { success: false, message: '목록 조회 실패: ' + e.message };
  }
}

function addStaff(data) {
  try {
    var user = Session.getActiveUser().getEmail();
    var sheet = getSheet(SHEET_NAMES.STAFF);
    var staffId = generateId_(sheet, 'STAFF');
    var now = getNow_();

    var row = [
      staffId,
      data.name || '',
      data.phone || '',
      data.email || '',
      data.birthDate || '',
      data.gender || '',
      data.address || '',
      data.source || '',
      data.org || '',
      data.duty || '',
      data.status || '후보등록',
      data.contractStart || '',
      data.contractEnd || '',
      data.payType || '',
      data.payAmount || '',
      data.bankName || '',
      data.bankAccount || '',
      data.memo || '',
      data.resumeUrl || '',
      now,
      now
    ];

    sheet.getRange(sheet.getLastRow() + 1, 1, 1, STAFF_COLUMNS.length).setValues([row]);

    var savedStaff = rowToStaffObject_(row);
    savedStaff.dday = calculateDday_(savedStaff.contractEnd);

    writeLog_(user, '등록', staffId, data.name + ' 스텝 등록');

    return { success: true, message: '스텝이 등록되었습니다.', staff: savedStaff };
  } catch (e) {
    Logger.log('addStaff 오류: ' + e.message);
    return { success: false, message: '등록 실패: ' + e.message };
  }
}

function updateStaff(data) {
  try {
    var user = Session.getActiveUser().getEmail();
    var sheet = getSheet(SHEET_NAMES.STAFF);
    var allData = sheet.getDataRange().getValues();
    var targetRow = -1;

    for (var i = 1; i < allData.length; i++) {
      if (allData[i][0].toString() === data.id) {
        targetRow = i + 1; // 시트는 1-based
        break;
      }
    }

    if (targetRow === -1) {
      return { success: false, message: '해당 스텝을 찾을 수 없습니다.' };
    }

    var now = getNow_();
    var row = [
      data.id,
      data.name || '',
      data.phone || '',
      data.email || '',
      data.birthDate || '',
      data.gender || '',
      data.address || '',
      data.source || '',
      data.org || '',
      data.duty || '',
      data.status || allData[targetRow - 1][10],
      data.contractStart || '',
      data.contractEnd || '',
      data.payType || '',
      data.payAmount || '',
      data.bankName || '',
      data.bankAccount || '',
      data.memo || '',
      data.resumeUrl || allData[targetRow - 1][18] || '',
      formatDate_(allData[targetRow - 1][19]), // 등록일 유지
      now
    ];

    sheet.getRange(targetRow, 1, 1, STAFF_COLUMNS.length).setValues([row]);

    var updatedStaff = rowToStaffObject_(row);
    updatedStaff.dday = calculateDday_(updatedStaff.contractEnd);

    writeLog_(user, '수정', data.id, data.name + ' 정보 수정');

    return { success: true, message: '수정되었습니다.', staff: updatedStaff };
  } catch (e) {
    Logger.log('updateStaff 오류: ' + e.message);
    return { success: false, message: '수정 실패: ' + e.message };
  }
}

function deleteStaff(id) {
  try {
    var user = Session.getActiveUser().getEmail();
    var role = getUserRole(user);
    if (!hasPermission_(role, ROLE_LEVELS['총괄관리자'])) {
      return { success: false, message: '삭제 권한이 없습니다.' };
    }

    var sheet = getSheet(SHEET_NAMES.STAFF);
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString() === id) {
        var staffName = data[i][1];
        sheet.deleteRow(i + 1);
        writeLog_(user, '삭제', id, staffName + ' 스텝 삭제');
        return { success: true, message: '삭제되었습니다.' };
      }
    }

    return { success: false, message: '해당 스텝을 찾을 수 없습니다.' };
  } catch (e) {
    Logger.log('deleteStaff 오류: ' + e.message);
    return { success: false, message: '삭제 실패: ' + e.message };
  }
}

function updateStaffStatus(id, newStatus) {
  try {
    var user = Session.getActiveUser().getEmail();
    var sheet = getSheet(SHEET_NAMES.STAFF);
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString() === id) {
        var currentStatus = data[i][10].toString();
        var staffName = data[i][1];

        // 상태 전이 규칙 검증
        var allowed = STATUS_TRANSITIONS[currentStatus] || [];
        if (allowed.indexOf(newStatus) === -1) {
          return {
            success: false,
            message: '\'' + currentStatus + '\' 에서 \'' + newStatus + '\' 로 변경할 수 없습니다. 허용: ' + allowed.join(', ')
          };
        }

        var targetRow = i + 1;
        var existingRow = sheet.getRange(targetRow, 1, 1, STAFF_COLUMNS.length).getValues()[0];
        existingRow[10] = newStatus;  // 상태 (11번째 컬럼, 0-based: 10)
        existingRow[20] = getNow_();  // 수정일 (21번째 컬럼, 0-based: 20)
        sheet.getRange(targetRow, 1, 1, STAFF_COLUMNS.length).setValues([existingRow]);

        // 계약이력 자동 기록 (근무중 → 계약종료 또는 재계약)
        if (currentStatus === '근무중' && (newStatus === '계약종료' || newStatus === '재계약')) {
          addContractHistory_(data[i], newStatus, user);
        }

        writeLog_(user, '상태변경', id, staffName + ': ' + currentStatus + ' → ' + newStatus);

        // 변경된 스텝 객체 반환
        data[i][10] = newStatus;
        var updatedStaff = rowToStaffObject_(data[i]);
        updatedStaff.dday = calculateDday_(updatedStaff.contractEnd);

        return { success: true, message: '상태가 변경되었습니다.', staff: updatedStaff };
      }
    }

    return { success: false, message: '해당 스텝을 찾을 수 없습니다.' };
  } catch (e) {
    Logger.log('updateStaffStatus 오류: ' + e.message);
    return { success: false, message: '상태 변경 실패: ' + e.message };
  }
}

function getStaffById(id) {
  try {
    var sheet = getSheet(SHEET_NAMES.STAFF);
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString() === id) {
        var staff = rowToStaffObject_(data[i]);
        staff.dday = calculateDday_(staff.contractEnd);

        // 성과평가 이력
        staff.evaluations = getEvaluationsByStaffId_(id);
        // 계약이력
        staff.contracts = getContractsByStaffId_(id);

        return { success: true, data: staff };
      }
    }

    return { success: false, message: '해당 스텝을 찾을 수 없습니다.' };
  } catch (e) {
    Logger.log('getStaffById 오류: ' + e.message);
    return { success: false, message: '조회 실패: ' + e.message };
  }
}

// ============ 성과평가 ============

function addEvaluation(data) {
  try {
    var user = Session.getActiveUser().getEmail();
    var role = getUserRole(user);
    if (!hasPermission_(role, ROLE_LEVELS['조직관리자'])) {
      return { success: false, message: '평가 등록 권한이 없습니다.' };
    }
    // 조직관리자는 자기 팀 스텝만 평가 가능
    if (role === '조직관리자') {
      var userOrgs = getUserOrgs_(user);
      if (!canAccessOrg_(role, userOrgs, data.org)) {
        return { success: false, message: '담당 조직의 스텝만 평가할 수 있습니다.' };
      }
    }
    var sheet = getSheet(SHEET_NAMES.EVALUATION);
    var evalId = generateId_(sheet, 'EVAL');

    // 기본 5항목 점수
    var baseScores = [
      Number(data.accuracy) || 0,
      Number(data.speed) || 0,
      Number(data.teamwork) || 0,
      Number(data.attendance) || 0,
      Number(data.attitude) || 0
    ];
    var baseTotal = baseScores.reduce(function(a, b) { return a + b; }, 0);
    var criteriaCount = 5;

    // 추가 항목 점수 (팀별 템플릿)
    var extraScores = data.extraScores || {};
    var extraTotal = 0;
    var extraKeys = Object.keys(extraScores);
    for (var e = 0; e < extraKeys.length; e++) {
      var score = Number(extraScores[extraKeys[e]]) || 0;
      extraScores[extraKeys[e]] = score;
      extraTotal += score;
      criteriaCount++;
    }

    var total = baseTotal + extraTotal;
    var avg = criteriaCount > 0 ? Math.round((total / criteriaCount) * 10) / 10 : 0;
    var grade = calculateGrade_(avg);
    var extraJson = extraKeys.length > 0 ? JSON.stringify(extraScores) : '';

    var row = [
      evalId,
      data.staffId || '',
      data.staffName || '',
      data.org || '',
      data.period || '',
      baseScores[0], baseScores[1], baseScores[2], baseScores[3], baseScores[4],
      total,
      avg,
      grade,
      data.comment || '',
      getUserName(user),
      getToday_(),
      extraJson
    ];

    sheet.getRange(sheet.getLastRow() + 1, 1, 1, EVALUATION_COLUMNS.length).setValues([row]);

    var savedEval = {
      id: evalId,
      staffId: data.staffId,
      staffName: data.staffName,
      org: data.org,
      period: data.period,
      accuracy: baseScores[0],
      speed: baseScores[1],
      teamwork: baseScores[2],
      attendance: baseScores[3],
      attitude: baseScores[4],
      extraScores: extraKeys.length > 0 ? extraScores : null,
      total: total,
      average: avg,
      grade: grade,
      comment: data.comment || '',
      evaluator: getUserName(user),
      evalDate: getToday_()
    };

    writeLog_(user, '평가등록', evalId, data.staffName + ' 성과평가 등록 (등급: ' + grade + ')');

    return { success: true, message: '평가가 등록되었습니다.', evaluation: savedEval };
  } catch (e) {
    Logger.log('addEvaluation 오류: ' + e.message);
    return { success: false, message: '평가 등록 실패: ' + e.message };
  }
}

function getEvaluationsByStaffId_(staffId) {
  var sheet = getSheet(SHEET_NAMES.EVALUATION);
  var data = sheet.getDataRange().getValues();
  var evals = [];

  for (var i = 1; i < data.length; i++) {
    if (data[i][1].toString() === staffId) {
      evals.push({
        id: data[i][0],
        staffId: data[i][1],
        staffName: data[i][2],
        org: data[i][3],
        period: data[i][4],
        accuracy: data[i][5],
        speed: data[i][6],
        teamwork: data[i][7],
        attendance: data[i][8],
        attitude: data[i][9],
        total: data[i][10],
        average: data[i][11],
        grade: data[i][12],
        comment: data[i][13],
        evaluator: data[i][14],
        evalDate: formatDate_(data[i][15]),
        extraScores: parseExtraScores_(data[i][16])
      });
    }
  }
  return evals;
}

function calculateGrade_(avg) {
  if (avg >= 4.5) return 'S';
  if (avg >= 3.5) return 'A';
  if (avg >= 2.5) return 'B';
  if (avg >= 1.5) return 'C';
  return 'D';
}

function parseExtraScores_(value) {
  if (!value) return null;
  try {
    var parsed = JSON.parse(value.toString());
    if (typeof parsed === 'object' && parsed !== null && !Array.isArray(parsed)) return parsed;
    return null;
  } catch (e) {
    return null;
  }
}

// ============ 계약이력 ============

function addContractHistory_(staffRow, contractType, user) {
  try {
    var sheet = getSheet(SHEET_NAMES.CONTRACT);
    var contractId = generateId_(sheet, 'CONT');

    var row = [
      contractId,
      staffRow[0],           // 스텝ID
      staffRow[1],           // 스텝이름
      contractType === '재계약' ? '재계약' : '만료',
      formatDate_(staffRow[11]), // 계약시작일
      formatDate_(staffRow[12]), // 계약종료일
      staffRow[13],          // 급여형태
      staffRow[14],          // 급여액
      contractType === '재계약' ? '재계약진행' : '종료',
      '',
      getNow_(),
      getUserName(user)
    ];

    sheet.getRange(sheet.getLastRow() + 1, 1, 1, CONTRACT_COLUMNS.length).setValues([row]);
  } catch (e) {
    Logger.log('계약이력 기록 실패: ' + e.message);
  }
}

function getContractsByStaffId_(staffId) {
  var sheet = getSheet(SHEET_NAMES.CONTRACT);
  var data = sheet.getDataRange().getValues();
  var contracts = [];

  for (var i = 1; i < data.length; i++) {
    if (data[i][1].toString() === staffId) {
      contracts.push({
        id: data[i][0],
        staffId: data[i][1],
        staffName: data[i][2],
        type: data[i][3],
        startDate: formatDate_(data[i][4]),
        endDate: formatDate_(data[i][5]),
        payType: data[i][6],
        payAmount: data[i][7],
        contractStatus: data[i][8],
        memo: data[i][9],
        createdAt: formatDate_(data[i][10]),
        createdBy: data[i][11]
      });
    }
  }
  return contracts;
}

// ============ 이력서 관리 ============

function getOrCreateResumeFolder_() {
  var settingsSheet = getSheet(SHEET_NAMES.SETTINGS);
  var settingsData = settingsSheet.getDataRange().getValues();
  for (var i = 1; i < settingsData.length; i++) {
    if (settingsData[i][0] === 'RESUME_FOLDER_ID' && settingsData[i][1]) {
      try {
        return DriveApp.getFolderById(settingsData[i][1].toString());
      } catch (e) { /* 폴더 삭제됨 — 재생성 */ }
    }
  }
  var folder = DriveApp.createFolder(RESUME_FOLDER_NAME);
  var found = false;
  for (var j = 1; j < settingsData.length; j++) {
    if (settingsData[j][0] === 'RESUME_FOLDER_ID') {
      settingsSheet.getRange(j + 1, 2).setValue(folder.getId());
      found = true;
      break;
    }
  }
  if (!found) {
    settingsSheet.appendRow(['RESUME_FOLDER_ID', folder.getId(), '이력서 저장 폴더', getNow_()]);
  }
  return folder;
}

function uploadResumeAndExtractText(base64Data, fileName, staffId, staffName) {
  try {
    var user = Session.getActiveUser().getEmail();

    // 1. base64 → Blob
    var decoded = Utilities.base64Decode(base64Data);
    var blob = Utilities.newBlob(decoded, 'application/pdf', fileName);

    // 2. Drive에 원본 PDF 저장
    var folder = getOrCreateResumeFolder_();
    var safeName = (staffId ? staffId + '_' : '') + (staffName || 'unknown') + '_' + fileName;
    blob.setName(safeName);
    var pdfFile = folder.createFile(blob);
    var pdfUrl = pdfFile.getUrl();

    // 3. OCR: PDF → Google Docs 변환 → 텍스트 추출
    var extractedText = '';
    try {
      var resource = {
        title: 'TEMP_OCR_' + safeName,
        mimeType: 'application/vnd.google-apps.document'
      };
      var docFile = Drive.Files.insert(resource, blob, { ocr: true, ocrLanguage: 'ko' });
      var doc = DocumentApp.openById(docFile.id);
      extractedText = doc.getBody().getText();
      DriveApp.getFileById(docFile.id).setTrashed(true);
    } catch (ocrErr) {
      Logger.log('OCR 변환 실패 (PDF 저장은 완료): ' + ocrErr.message);
    }

    // 4. 텍스트 파싱
    var parsed = parseResumeText_(extractedText);

    writeLog_(user, '이력서업로드', staffId || 'NEW', fileName + ' 업로드 및 텍스트 추출');

    return {
      success: true,
      message: '이력서가 업로드되었습니다.',
      resumeUrl: pdfUrl,
      extractedText: extractedText.substring(0, 2000),
      parsed: parsed
    };
  } catch (e) {
    Logger.log('uploadResumeAndExtractText 오류: ' + e.message);
    return { success: false, message: '이력서 업로드 실패: ' + e.message };
  }
}

function parseResumeText_(text) {
  var result = { name: '', phone: '', email: '', birthDate: '', gender: '', address: '' };
  if (!text) return result;

  // 전처리: 탭→스페이스, 라인 정보 보존
  var cleanText = text.replace(/\t/g, ' ');
  var lines = cleanText.split('\n').map(function(l) { return l.trim(); }).filter(function(l) { return l.length > 0; });
  var fullText = lines.join(' ');
  // 연속 공백 축소 (OCR 표 깨짐 대응)
  var compact = fullText.replace(/\s{2,}/g, ' ');

  // ===== 이메일 (가장 확실한 패턴) =====
  var emailMatch = compact.match(/[\w.\-+]+@[\w.\-]+\.\w{2,}/);
  if (emailMatch) result.email = emailMatch[0];

  // ===== 전화번호 =====
  var phonePatterns = [
    // 라벨 + 번호 (OCR 공백 분산 대응: 연 락 처)
    /(?:연\s*락\s*처|전\s*화\s*번?\s*호?|핸\s*드\s*폰|휴\s*대\s*폰|HP|H\.P|Mobile|Tel|Phone)\s*[：:\.]?\s*\(?(01[0-9])\)?[\s\-.]?(\d{3,4})[\s\-.]?(\d{4})/i,
    // 순수 번호 (괄호 포함)
    /\(?(01[0-9])\)?[\s\-.]?(\d{3,4})[\s\-.]?(\d{4})/
  ];
  for (var pi = 0; pi < phonePatterns.length; pi++) {
    var pm = compact.match(phonePatterns[pi]);
    if (pm && pm[3]) {
      result.phone = pm[1] + '-' + pm[2] + '-' + pm[3];
      break;
    }
    if (pm && !pm[3]) {
      var rawPhone = pm[0].replace(/[\s.\-\(\)]/g, '');
      if (rawPhone.length >= 10 && rawPhone.length <= 11) {
        result.phone = rawPhone.replace(/(\d{3})(\d{3,4})(\d{4})/, '$1-$2-$3');
        break;
      }
    }
  }

  // ===== 생년월일 =====
  var birthFound = false;

  // 패턴1: 라벨 + YYYY.MM.DD (OCR 공백 분산 대응)
  var birthLabeled = compact.match(/(?:생\s*년\s*월\s*일|생\s*년|생\s*일|출\s*생\s*일?|Birth\s*(?:date)?)\s*[：:\.]?\s*(\d{4})\s*[\.\-\/년\s]\s*(\d{1,2})\s*[\.\-\/월\s]\s*(\d{1,2})\s*일?/i);
  if (birthLabeled) {
    var by = birthLabeled[1], bm2 = birthLabeled[2], bd2 = birthLabeled[3];
    if (parseInt(by) >= 1950 && parseInt(by) <= 2010 && parseInt(bm2) >= 1 && parseInt(bm2) <= 12 && parseInt(bd2) >= 1 && parseInt(bd2) <= 31) {
      result.birthDate = by + '-' + ('0' + bm2).slice(-2) + '-' + ('0' + bd2).slice(-2);
      birthFound = true;
    }
  }

  // 패턴2: 주민등록번호 (YYMMDD-N, 마스킹 포함)
  if (!birthFound) {
    var juminMatch = compact.match(/(\d{2})(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])\s*[\-–—]\s*([1-4])/);
    if (juminMatch) {
      var jPrefix = (juminMatch[4] === '1' || juminMatch[4] === '2') ? '19' : '20';
      result.birthDate = jPrefix + juminMatch[1] + '-' + juminMatch[2] + '-' + juminMatch[3];
      birthFound = true;
      if (!result.gender) {
        result.gender = (juminMatch[4] === '1' || juminMatch[4] === '3') ? '남' : '여';
      }
    }
  }

  // 패턴3: YYYY.MM.DD (라벨 없음, 1950~2010 범위 첫 매칭)
  if (!birthFound) {
    var dateRegex = /((?:19|20)\d{2})\s*[\.\-\/년]\s*(\d{1,2})\s*[\.\-\/월]\s*(\d{1,2})\s*일?/g;
    var dm;
    while ((dm = dateRegex.exec(compact)) !== null) {
      var dy = parseInt(dm[1]), dmm = parseInt(dm[2]), ddd = parseInt(dm[3]);
      if (dy >= 1950 && dy <= 2010 && dmm >= 1 && dmm <= 12 && ddd >= 1 && ddd <= 31) {
        result.birthDate = dm[1] + '-' + ('0' + dm[2]).slice(-2) + '-' + ('0' + dm[3]).slice(-2);
        birthFound = true;
        break;
      }
    }
  }

  // 패턴4: YYYYMMDD 8자리 연속 (라벨 근처)
  if (!birthFound) {
    var bd8 = compact.match(/(?:생년월일|생년|출생|생일)[^\d]{0,10}((?:19|20)\d{2})(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])/);
    if (bd8) {
      result.birthDate = bd8[1] + '-' + bd8[2] + '-' + bd8[3];
    }
  }

  // ===== 성별 =====
  if (!result.gender) {
    var genderPatterns = [
      /(?:성\s*별|gender)\s*[：:\.]?\s*(남|여|남성|여성|남자|여자|Male|Female)/i,
      /성\s*별[^가-힣]{0,10}(남|여)/
    ];
    for (var gi = 0; gi < genderPatterns.length; gi++) {
      var gm = compact.match(genderPatterns[gi]);
      if (gm) {
        var gv = gm[1] || gm[0];
        result.gender = (gv.indexOf('남') > -1 || (typeof gv === 'string' && gv.toLowerCase() === 'male')) ? '남' : '여';
        break;
      }
    }
  }

  // ===== 주소 =====
  var provinces = '서울|부산|대구|인천|광주|대전|울산|세종|경기|강원|충북|충남|전북|전남|경북|경남|제주';
  var addrPatterns = [
    // 라벨 + 주소
    new RegExp('(?:주\\s*소|현\\s*주\\s*소|거\\s*주\\s*지|자\\s*택|Address)\\s*[：:\\.]?\\s*((?:' + provinces + ')[^\\n]{3,80})', 'i'),
    // 시/도 + 시/군/구 (충분히 긴 것)
    new RegExp('((?:' + provinces + ')(?:특별시|광역시|특별자치시|도|특별자치도)?\\s*[가-힣]+(?:시|군|구)[^\\n]{3,60})')
  ];
  for (var ai = 0; ai < addrPatterns.length; ai++) {
    var am = compact.match(addrPatterns[ai]);
    if (am) {
      var addr = (am[1] || am[0]).trim();
      // 다른 필드 라벨이 섞이면 잘라냄
      addr = addr.split(/(?:연락처|전화|이메일|성별|생년월일|학력|경력)/)[0].trim();
      result.address = addr.substring(0, 80);
      break;
    }
  }

  // ===== 이름 =====
  var excluded = ['이력서', '경력서', '자기소개', '인적사항', '경력사항', '학력사항', '자격사항', '지원서', '입사지원'];

  // 패턴1: 명시적 라벨
  var namePatterns = [
    /(?:이\s*름|성\s*명|Name|지\s*원\s*자\s*명?)\s*[：:\.]?\s*([가-힣]{2,4})/i,
    /(?:성\s*명|이\s*름)\s{2,}([가-힣]{2,4})/,
    /^([가-힣]{2,4})\s*(?:이력서|RESUME)/im,
    /(?:이력서|RESUME)\s*[-–]?\s*([가-힣]{2,4})/im,
    /인\s*적\s*사\s*항[^가-힣]{0,15}([가-힣]{2,4})/
  ];
  for (var ni = 0; ni < namePatterns.length; ni++) {
    var nm = compact.match(namePatterns[ni]);
    if (nm && nm[1] && excluded.indexOf(nm[1]) === -1) {
      result.name = nm[1];
      break;
    }
  }

  // 폴백1: 각 줄에서 라벨+이름 패턴
  if (!result.name) {
    for (var li = 0; li < Math.min(25, lines.length); li++) {
      var tblNm = lines[li].match(/(?:성\s*명|이\s*름|성명|이름)\s*[：:\.]?\s*([가-힣]{2,4})/);
      if (tblNm && excluded.indexOf(tblNm[1]) === -1) {
        result.name = tblNm[1];
        break;
      }
    }
  }

  // 폴백2: 첫 10줄에서 단독 한글 이름 (2~4글자)
  if (!result.name) {
    for (var li2 = 0; li2 < Math.min(10, lines.length); li2++) {
      var lineNm = lines[li2].match(/^([가-힣]{2,4})$/);
      if (lineNm && excluded.indexOf(lineNm[1]) === -1) {
        result.name = lineNm[1];
        break;
      }
      var startNm = lines[li2].match(/^([가-힣]{2,4})\s/);
      if (startNm && lines[li2].length < 20 && excluded.indexOf(startNm[1]) === -1) {
        result.name = startNm[1];
        break;
      }
    }
  }

  Logger.log('parseResumeText_ 결과: ' + JSON.stringify(result));
  return result;
}

function migrateAddResumeColumn() {
  var sheet = getSheet(SHEET_NAMES.STAFF);
  var header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  // 이미 이력서URL 컬럼이 있으면 스킵
  for (var i = 0; i < header.length; i++) {
    if (header[i] === '이력서URL') {
      return { success: true, message: '이력서URL 컬럼이 이미 존재합니다.' };
    }
  }
  // 18번째 컬럼(비고) 뒤에 삽입
  sheet.insertColumnAfter(18);
  sheet.getRange(1, 19).setValue('이력서URL')
    .setFontWeight('bold')
    .setBackground('#4472C4')
    .setFontColor('#FFFFFF');
  return { success: true, message: '이력서URL 컬럼이 추가되었습니다. (19번째 위치)' };
}

// ============ 비즈니스 로직 ============

function rowToStaffObject_(row) {
  return {
    id: row[0] ? row[0].toString() : '',
    name: row[1] || '',
    phone: row[2] || '',
    email: row[3] || '',
    birthDate: formatDate_(row[4]),
    gender: row[5] || '',
    address: row[6] || '',
    source: row[7] || '',
    org: row[8] || '',
    duty: row[9] || '',
    status: row[10] || '',
    contractStart: formatDate_(row[11]),
    contractEnd: formatDate_(row[12]),
    payType: row[13] || '',
    payAmount: row[14] || '',
    bankName: row[15] || '',
    bankAccount: row[16] || '',
    memo: row[17] || '',
    resumeUrl: row[18] || '',
    createdAt: formatDate_(row[19]),
    updatedAt: formatDate_(row[20])
  };
}

function calculateDday_(endDateStr) {
  if (!endDateStr) return null;
  var endDate = parseDate_(endDateStr);
  if (!endDate) return null;

  var today = new Date();
  today.setHours(0, 0, 0, 0);
  endDate.setHours(0, 0, 0, 0);

  var diff = Math.ceil((endDate.getTime() - today.getTime()) / (1000 * 60 * 60 * 24));
  return diff;
}

function getDdayAlerts_() {
  var sheet = getSheet(SHEET_NAMES.STAFF);
  var data = sheet.getDataRange().getValues();
  var alerts = [];

  for (var i = 1; i < data.length; i++) {
    if (!data[i][0] || data[i][10] !== '근무중') continue;

    var dday = calculateDday_(formatDate_(data[i][12]));
    if (dday === null) continue;

    for (var d = 0; d < DDAY_ALERTS.length; d++) {
      if (dday <= DDAY_ALERTS[d] && dday >= 0) {
        alerts.push({
          staffId: data[i][0].toString(),
          staffName: data[i][1],
          org: data[i][8],
          contractEnd: formatDate_(data[i][12]),
          dday: dday,
          level: dday <= 3 ? 'danger' : dday <= 7 ? 'warning' : 'info'
        });
        break;
      }
    }
  }

  // D-day 가까운 순 정렬
  alerts.sort(function(a, b) { return a.dday - b.dday; });
  return alerts;
}

// ============ 대시보드 통계 ============

function getDashboardStats() {
  try {
    var sheet = getSheet(SHEET_NAMES.STAFF);
    var data = sheet.getDataRange().getValues();

    var stats = {
      total: 0,
      byStatus: {},
      byOrg: {},
      recentAdded: [],
      expiringContracts: []
    };

    for (var s = 0; s < STATUS_LIST.length; s++) {
      stats.byStatus[STATUS_LIST[s]] = 0;
    }

    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      stats.total++;

      var status = data[i][10] || '';
      if (stats.byStatus.hasOwnProperty(status)) {
        stats.byStatus[status]++;
      }

      var org = data[i][8] || '미지정';
      stats.byOrg[org] = (stats.byOrg[org] || 0) + 1;
    }

    stats.alerts = getDdayAlerts_();

    return { success: true, data: stats };
  } catch (e) {
    Logger.log('getDashboardStats 오류: ' + e.message);
    return { success: false, message: '통계 조회 실패: ' + e.message };
  }
}

// ============ 대시보드 상세 데이터 ============

function getDashboardData() {
  try {
    var user = Session.getActiveUser().getEmail();
    var role = getUserRole(user);
    var userOrgs = getUserOrgs_(user);

    // 스텝 데이터 (권한 필터링 적용)
    var sheet = getSheet(SHEET_NAMES.STAFF);
    var data = sheet.getDataRange().getValues();

    var total = 0;
    var working = 0;
    var byOrg = {};
    var byOrgStatus = {};
    var filteredStaff = [];

    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      var org = data[i][8] || '미지정';
      var status = data[i][10] || '';

      if (!canAccessOrg_(role, userOrgs, org)) continue;

      total++;
      if (status === '근무중') working++;

      byOrg[org] = (byOrg[org] || 0) + 1;
      if (!byOrgStatus[org]) byOrgStatus[org] = {};
      byOrgStatus[org][status] = (byOrgStatus[org][status] || 0) + 1;

      filteredStaff.push({
        id: data[i][0].toString(),
        name: data[i][1] || '',
        email: data[i][3] || '',
        org: org,
        status: status,
        contractEnd: formatDate_(data[i][12])
      });
    }

    // D-day 알림 (권한 필터)
    var allAlerts = getDdayAlerts_();
    var alerts = [];
    for (var a = 0; a < allAlerts.length; a++) {
      if (canAccessOrg_(role, userOrgs, allAlerts[a].org)) {
        alerts.push(allAlerts[a]);
      }
    }

    // 성과평가 필요 스텝
    var evalNeeded = getEvalNeededStaff_(filteredStaff);

    // 팀별 출근율
    var attendanceRate = getAttendanceRateForDashboard_(filteredStaff);

    return {
      success: true,
      data: {
        total: total,
        working: working,
        alertCount: alerts.length,
        byOrg: byOrg,
        byOrgStatus: byOrgStatus,
        alerts: alerts,
        evalNeeded: evalNeeded,
        attendanceRate: attendanceRate
      }
    };
  } catch (e) {
    Logger.log('getDashboardData 오류: ' + e.message);
    return { success: false, message: '대시보드 데이터 조회 실패: ' + e.message };
  }
}

function getEvalNeededStaff_(staffList) {
  try {
    var evalSheet = getSheet(SHEET_NAMES.EVALUATION);
    var evalData = evalSheet.getDataRange().getValues();

    // staffId → 최신 평가일 매핑
    var latestEvalMap = {};
    for (var i = 1; i < evalData.length; i++) {
      if (!evalData[i][0]) continue;
      var sid = evalData[i][1] ? evalData[i][1].toString() : '';
      var evalDate = formatDate_(evalData[i][15]);
      if (!latestEvalMap[sid] || evalDate > latestEvalMap[sid]) {
        latestEvalMap[sid] = evalDate;
      }
    }

    var today = new Date();
    today.setHours(0, 0, 0, 0);
    var result = [];

    for (var j = 0; j < staffList.length; j++) {
      var st = staffList[j];
      if (st.status !== '근무중') continue;

      var lastEval = latestEvalMap[st.id] || null;
      var daysSince = -1;

      if (lastEval) {
        var evalD = parseDate_(lastEval);
        if (evalD) {
          evalD.setHours(0, 0, 0, 0);
          daysSince = Math.floor((today.getTime() - evalD.getTime()) / (1000 * 60 * 60 * 24));
        }
      }

      // 평가 없음 또는 60일 이상 경과
      if (!lastEval || daysSince >= 60) {
        result.push({
          staffId: st.id,
          staffName: st.name,
          org: st.org,
          lastEvalDate: lastEval || '없음',
          daysSince: lastEval ? daysSince : -1
        });
      }
    }

    // 경과일 큰 순 정렬 (평가 없는 것 우선)
    result.sort(function(a, b) {
      if (a.daysSince === -1 && b.daysSince === -1) return 0;
      if (a.daysSince === -1) return -1;
      if (b.daysSince === -1) return 1;
      return b.daysSince - a.daysSince;
    });

    return result;
  } catch (e) {
    Logger.log('getEvalNeededStaff_ 오류: ' + e.message);
    return [];
  }
}

function getAttendanceRateForDashboard_(staffList) {
  try {
    var ss = getLinkedAttSS_();
    if (!ss) return { overall: 0 };

    // 직원 시트: 이메일 → { empId, org }
    var empSheet = ss.getSheetByName('직원');
    if (!empSheet) return { overall: 0 };
    var empData = empSheet.getDataRange().getValues();
    var emailToEmp = {};
    for (var i = 1; i < empData.length; i++) {
      var email = empData[i][2] ? empData[i][2].toString().trim().toLowerCase() : '';
      if (email) {
        emailToEmp[email] = {
          empId: empData[i][0] ? empData[i][0].toString() : '',
          org: empData[i][4] || ''
        };
      }
    }

    // 스텝 이메일 → 근태 직원 매핑
    var empIdToOrg = {};
    for (var j = 0; j < staffList.length; j++) {
      var semail = (staffList[j].email || '').trim().toLowerCase();
      if (semail && emailToEmp[semail]) {
        empIdToOrg[emailToEmp[semail].empId] = staffList[j].org;
      }
    }

    // 출퇴근기록 시트: 이번 달 출근일수 집계
    var attSheet = ss.getSheetByName('출퇴근기록');
    if (!attSheet) return { overall: 0 };
    var attData = attSheet.getDataRange().getValues();
    var today = getToday_();
    var yearMonth = today.substring(0, 7);

    var orgAttDays = {};  // 팀별 총 출근일
    var orgEmpCount = {}; // 팀별 직원 수

    // 직원별 출근일 카운트
    var empDays = {};
    for (var a = 1; a < attData.length; a++) {
      var empId = attData[a][1] ? attData[a][1].toString() : '';
      if (!empId || !empIdToOrg[empId]) continue;
      var attDate = formatDate_(attData[a][2]);
      if (!attDate || attDate.substring(0, 7) !== yearMonth) continue;
      empDays[empId] = (empDays[empId] || 0) + 1;
    }

    // 팀별 집계
    for (var eid in empIdToOrg) {
      var eorg = empIdToOrg[eid];
      orgEmpCount[eorg] = (orgEmpCount[eorg] || 0) + 1;
      orgAttDays[eorg] = (orgAttDays[eorg] || 0) + (empDays[eid] || 0);
    }

    // 이번 달 영업일 수 계산 (주말 제외)
    var year = parseInt(yearMonth.substring(0, 4));
    var month = parseInt(yearMonth.substring(5, 7)) - 1;
    var daysInMonth = new Date(year, month + 1, 0).getDate();
    var todayDate = today.substring(8, 10);
    var maxDay = Math.min(parseInt(todayDate), daysInMonth);
    var businessDays = 0;
    for (var d = 1; d <= maxDay; d++) {
      var dow = new Date(year, month, d).getDay();
      if (dow !== 0 && dow !== 6) businessDays++;
    }
    if (businessDays === 0) businessDays = 1;

    var result = {};
    var totalAttDays = 0;
    var totalExpected = 0;

    for (var oname in orgEmpCount) {
      var expected = orgEmpCount[oname] * businessDays;
      var actual = orgAttDays[oname] || 0;
      result[oname] = expected > 0 ? Math.round((actual / expected) * 1000) / 10 : 0;
      totalAttDays += actual;
      totalExpected += expected;
    }

    result.overall = totalExpected > 0 ? Math.round((totalAttDays / totalExpected) * 1000) / 10 : 0;

    return result;
  } catch (e) {
    Logger.log('getAttendanceRateForDashboard_ 오류: ' + e.message);
    return { overall: 0 };
  }
}

// ============ 트리거 관리 ============

function setupTriggers() {
  removeAllTriggers_();

  // D-day 알림: 매일 09:00 실행
  ScriptApp.newTrigger('sendDdayAlertEmails')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();

  Logger.log('트리거 설정 완료');
  return { success: true, message: '트리거가 설정되었습니다. (D-day 알림 매일 09:00)' };
}

function removeAllTriggers_() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
}

// ============ 트리거: D-day 알림 메일 ============

function sendDdayAlertEmails() {
  try {
    var alerts = getDdayAlerts_();
    if (alerts.length === 0) return;

    // 관리자 목록 조회
    var admins = getAdminList_();
    var managerEmails = [];
    for (var i = 0; i < admins.length; i++) {
      if (ROLE_LEVELS[admins[i].role] >= ROLE_LEVELS['조직관리자']) {
        managerEmails.push(admins[i].email);
      }
    }

    if (managerEmails.length === 0) return;

    var body = '📋 스텝 계약만료 알림\n\n';
    for (var j = 0; j < alerts.length; j++) {
      var a = alerts[j];
      body += '• ' + a.staffName + ' (' + a.org + ') - D-' + a.dday + ' (' + a.contractEnd + ')\n';
    }
    body += '\n스텝관리 시스템에서 확인해주세요.';

    for (var k = 0; k < managerEmails.length; k++) {
      MailApp.sendEmail({
        to: managerEmails[k],
        subject: '[스텝관리] 계약만료 알림 (' + alerts.length + '건)',
        body: body
      });
    }

    Logger.log('D-day 알림 메일 발송 완료: ' + alerts.length + '건, ' + managerEmails.length + '명');
  } catch (e) {
    Logger.log('알림 메일 발송 실패: ' + e.message);
  }
}

// ============ 데이터 연계 (근태관리 시스템) ============

var _linkedAttSS = null;

function getLinkedAttSS_() {
  if (_linkedAttSS) return _linkedAttSS;
  try {
    _linkedAttSS = SpreadsheetApp.openById(LINKED_ATTENDANCE_SS_ID);
    return _linkedAttSS;
  } catch (e) {
    Logger.log('근태관리 스프레드시트 연결 실패: ' + e.message);
    return null;
  }
}

function getStaffAttendanceSummary(staffEmail) {
  try {
    var ss = getLinkedAttSS_();
    if (!ss) return { success: false, message: '근태관리 시스템에 연결할 수 없습니다.' };

    var result = { attendance: null, leave: null, payroll: null };

    // 1. 직원 정보 조회
    var empSheet = ss.getSheetByName('직원');
    if (!empSheet) return { success: false, message: '근태관리 직원 시트를 찾을 수 없습니다.' };
    var empData = empSheet.getDataRange().getValues();
    var empId = '';
    var empName = '';
    var empOrg = '';
    var leaveBalance = '';
    for (var i = 1; i < empData.length; i++) {
      if (empData[i][2] && empData[i][2].toString().trim().toLowerCase() === staffEmail.toLowerCase()) {
        empId = empData[i][0].toString();
        empName = empData[i][1] || '';
        empOrg = empData[i][4] || '';
        leaveBalance = empData[i][11] || 0;
        break;
      }
    }
    if (!empId) return { success: true, data: null, message: '근태관리 시스템에 등록되지 않은 스텝입니다.' };

    // 2. 이번 달 출퇴근 요약
    var attSheet = ss.getSheetByName('출퇴근기록');
    if (attSheet) {
      var attData = attSheet.getDataRange().getValues();
      var today = getToday_();
      var yearMonth = today.substring(0, 7);
      var totalDays = 0;
      var lateDays = 0;
      var totalHours = 0;
      var lastClockIn = '';

      for (var a = 1; a < attData.length; a++) {
        if (attData[a][1] && attData[a][1].toString() === empId) {
          var attDate = formatDate_(attData[a][2]);
          if (attDate && attDate.substring(0, 7) === yearMonth) {
            totalDays++;
            if (attData[a][6] === '지각') lateDays++;
            var hours = attData[a][8] ? parseFloat(attData[a][8]) : 0;
            totalHours += hours;
          }
          var ci = attData[a][3] ? attData[a][3].toString() : '';
          if (ci && ci > lastClockIn) lastClockIn = ci;
        }
      }
      result.attendance = {
        monthDays: totalDays,
        lateDays: lateDays,
        totalHours: Math.round(totalHours * 100) / 100,
        lastClockIn: lastClockIn
      };
    }

    // 3. 휴가 요약
    var leaveSheet = ss.getSheetByName('휴가');
    if (leaveSheet) {
      var leaveData = leaveSheet.getDataRange().getValues();
      var pendingCount = 0;
      var approvedCount = 0;
      for (var l = 1; l < leaveData.length; l++) {
        if (leaveData[l][1] && leaveData[l][1].toString() === empId) {
          var status = leaveData[l][10] || '';
          if (status === '대기') pendingCount++;
          if (status === '승인') approvedCount++;
        }
      }
      result.leave = {
        balance: leaveBalance,
        pending: pendingCount,
        approved: approvedCount
      };
    }

    // 4. 최근 급여
    var paySheet = ss.getSheetByName('급여');
    if (paySheet) {
      var payData = paySheet.getDataRange().getValues();
      var latestPay = null;
      for (var p = 1; p < payData.length; p++) {
        if (payData[p][1] && payData[p][1].toString() === empId) {
          var period = payData[p][3] ? payData[p][3].toString() : '';
          if (!latestPay || period > latestPay.period) {
            latestPay = {
              period: period,
              totalPay: payData[p][17] || 0,
              workDays: payData[p][5] || 0,
              workHours: payData[p][6] || 0
            };
          }
        }
      }
      result.payroll = latestPay;
    }

    return { success: true, data: result };
  } catch (e) {
    Logger.log('근태 연동 조회 오류: ' + e.message);
    return { success: false, message: '근태 데이터 조회 실패: ' + e.message };
  }
}

// ============ 전체 성과평가 조회 ============

function getAllEvaluations() {
  try {
    var user = Session.getActiveUser().getEmail();
    var role = getUserRole(user);
    var userOrgs = getUserOrgs_(user);

    var sheet = getSheet(SHEET_NAMES.EVALUATION);
    var data = sheet.getDataRange().getValues();
    var evals = [];

    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      var org = data[i][3] ? data[i][3].toString() : '';
      if (!canAccessOrg_(role, userOrgs, org)) continue;

      evals.push({
        id: data[i][0],
        staffId: data[i][1],
        staffName: data[i][2],
        org: org,
        period: data[i][4],
        accuracy: data[i][5],
        speed: data[i][6],
        teamwork: data[i][7],
        attendance: data[i][8],
        attitude: data[i][9],
        total: data[i][10],
        average: data[i][11],
        grade: data[i][12],
        comment: data[i][13],
        evaluator: data[i][14],
        evalDate: formatDate_(data[i][15]),
        extraScores: parseExtraScores_(data[i][16])
      });
    }

    return { success: true, data: evals };
  } catch (e) {
    Logger.log('getAllEvaluations 오류: ' + e.message);
    return { success: false, message: '평가 목록 조회 실패: ' + e.message };
  }
}

// ============ 전체 계약이력 조회 ============

function getAllContracts() {
  try {
    var user = Session.getActiveUser().getEmail();
    var role = getUserRole(user);
    var userOrgs = getUserOrgs_(user);

    var sheet = getSheet(SHEET_NAMES.CONTRACT);
    var data = sheet.getDataRange().getValues();
    var contracts = [];

    // 스텝ID → 조직 매핑 (권한 필터용)
    var staffSheet = getSheet(SHEET_NAMES.STAFF);
    var staffData = staffSheet.getDataRange().getValues();
    var staffOrgMap = {};
    for (var s = 1; s < staffData.length; s++) {
      if (staffData[s][0]) {
        staffOrgMap[staffData[s][0].toString()] = staffData[s][8] || '';
      }
    }

    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      var staffId = data[i][1] ? data[i][1].toString() : '';
      var org = staffOrgMap[staffId] || '';
      if (!canAccessOrg_(role, userOrgs, org)) continue;

      contracts.push({
        id: data[i][0],
        staffId: staffId,
        staffName: data[i][2],
        org: org,
        type: data[i][3],
        startDate: formatDate_(data[i][4]),
        endDate: formatDate_(data[i][5]),
        payType: data[i][6],
        payAmount: data[i][7],
        status: data[i][8],
        memo: data[i][9],
        createdAt: formatDate_(data[i][10]),
        creator: data[i][11]
      });
    }

    return { success: true, data: contracts };
  } catch (e) {
    Logger.log('getAllContracts 오류: ' + e.message);
    return { success: false, message: '계약이력 조회 실패: ' + e.message };
  }
}

// ============ 설정 CRUD ============

function getSettings() {
  try {
    var sheet = getSheet(SHEET_NAMES.SETTINGS);
    var data = sheet.getDataRange().getValues();
    var settings = [];

    for (var i = 1; i < data.length; i++) {
      if (data[i][0]) {
        settings.push({
          key: data[i][0].toString(),
          value: data[i][1] ? data[i][1].toString() : '',
          description: data[i][2] || '',
          updatedAt: formatDate_(data[i][3])
        });
      }
    }

    return { success: true, data: settings };
  } catch (e) {
    Logger.log('getSettings 오류: ' + e.message);
    return { success: false, message: '설정 조회 실패: ' + e.message };
  }
}

function updateSetting(key, value) {
  try {
    var user = Session.getActiveUser().getEmail();
    var role = getUserRole(user);
    if (!hasPermission_(role, ROLE_LEVELS['최고관리자'])) {
      return { success: false, message: '설정 변경 권한이 없습니다.' };
    }

    var sheet = getSheet(SHEET_NAMES.SETTINGS);
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString() === key) {
        sheet.getRange(i + 1, 2).setValue(value);
        sheet.getRange(i + 1, 4).setValue(getNow_());
        writeLog_(user, '설정변경', key, key + ' = ' + value);
        return { success: true, message: '설정이 저장되었습니다.' };
      }
    }

    // 키가 없으면 새로 생성
    sheet.appendRow([key, value, '', getNow_()]);
    writeLog_(user, '설정추가', key, key + ' = ' + value);
    return { success: true, message: '설정이 추가되었습니다.' };
  } catch (e) {
    Logger.log('updateSetting 오류: ' + e.message);
    return { success: false, message: '설정 저장 실패: ' + e.message };
  }
}

// ============ 관리자 CRUD ============

function getAdminListForSettings() {
  try {
    var user = Session.getActiveUser().getEmail();
    var role = getUserRole(user);
    if (!hasPermission_(role, ROLE_LEVELS['최고관리자'])) {
      return { success: false, message: '관리자 목록 조회 권한이 없습니다.' };
    }
    var admins = getAdminList_();
    return { success: true, data: admins };
  } catch (e) {
    Logger.log('getAdminListForSettings 오류: ' + e.message);
    return { success: false, message: '관리자 목록 조회 실패: ' + e.message };
  }
}

function addAdmin(data) {
  try {
    var user = Session.getActiveUser().getEmail();
    var role = getUserRole(user);
    if (!hasPermission_(role, ROLE_LEVELS['최고관리자'])) {
      return { success: false, message: '관리자 추가 권한이 없습니다.' };
    }

    var sheet = getSheet(SHEET_NAMES.ADMIN);
    var existing = sheet.getDataRange().getValues();
    var email = (data.email || '').trim().toLowerCase();
    if (!email) return { success: false, message: '이메일을 입력해주세요.' };

    for (var i = 1; i < existing.length; i++) {
      if (existing[i][0] && existing[i][0].toString().trim().toLowerCase() === email) {
        return { success: false, message: '이미 등록된 관리자입니다.' };
      }
    }

    var row = [email, data.name || '', data.role || '직원', data.org || '', '활성', getToday_()];
    sheet.getRange(sheet.getLastRow() + 1, 1, 1, ADMIN_COLUMNS.length).setValues([row]);
    _adminCache = null;

    writeLog_(user, '관리자추가', email, (data.name || email) + ' 관리자 추가 (' + (data.role || '직원') + ')');
    return { success: true, message: '관리자가 추가되었습니다.' };
  } catch (e) {
    Logger.log('addAdmin 오류: ' + e.message);
    return { success: false, message: '관리자 추가 실패: ' + e.message };
  }
}

function updateAdmin(data) {
  try {
    var user = Session.getActiveUser().getEmail();
    var role = getUserRole(user);
    if (!hasPermission_(role, ROLE_LEVELS['최고관리자'])) {
      return { success: false, message: '관리자 수정 권한이 없습니다.' };
    }

    var sheet = getSheet(SHEET_NAMES.ADMIN);
    var existing = sheet.getDataRange().getValues();
    var email = (data.email || '').trim().toLowerCase();

    for (var i = 1; i < existing.length; i++) {
      if (existing[i][0] && existing[i][0].toString().trim().toLowerCase() === email) {
        sheet.getRange(i + 1, 2).setValue(data.name || existing[i][1]);
        sheet.getRange(i + 1, 3).setValue(data.role || existing[i][2]);
        sheet.getRange(i + 1, 4).setValue(data.org || '');
        sheet.getRange(i + 1, 5).setValue(data.status || existing[i][4]);
        _adminCache = null;

        writeLog_(user, '관리자수정', email, (data.name || email) + ' 정보 수정');
        return { success: true, message: '관리자 정보가 수정되었습니다.' };
      }
    }

    return { success: false, message: '해당 관리자를 찾을 수 없습니다.' };
  } catch (e) {
    Logger.log('updateAdmin 오류: ' + e.message);
    return { success: false, message: '관리자 수정 실패: ' + e.message };
  }
}

function deleteAdmin(email) {
  try {
    var user = Session.getActiveUser().getEmail();
    var role = getUserRole(user);
    if (!hasPermission_(role, ROLE_LEVELS['최고관리자'])) {
      return { success: false, message: '관리자 삭제 권한이 없습니다.' };
    }

    if (email.toLowerCase() === user.toLowerCase()) {
      return { success: false, message: '자기 자신은 삭제할 수 없습니다.' };
    }

    var sheet = getSheet(SHEET_NAMES.ADMIN);
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim().toLowerCase() === email.toLowerCase()) {
        sheet.deleteRow(i + 1);
        _adminCache = null;
        writeLog_(user, '관리자삭제', email, email + ' 관리자 삭제');
        return { success: true, message: '관리자가 삭제되었습니다.' };
      }
    }

    return { success: false, message: '해당 관리자를 찾을 수 없습니다.' };
  } catch (e) {
    Logger.log('deleteAdmin 오류: ' + e.message);
    return { success: false, message: '관리자 삭제 실패: ' + e.message };
  }
}

// ============ 평가 템플릿 CRUD ============

function getEvalTemplates() {
  try {
    var sheet = getSheet(SHEET_NAMES.SETTINGS);
    var data = sheet.getDataRange().getValues();
    var templates = {};

    for (var i = 1; i < data.length; i++) {
      var key = data[i][0] ? data[i][0].toString() : '';
      if (key.indexOf('EVAL_TEMPLATE_') === 0) {
        var orgName = key.replace('EVAL_TEMPLATE_', '');
        var val = data[i][1] ? data[i][1].toString().trim() : '';
        if (val) {
          templates[orgName] = val.split(',').map(function(c) { return c.trim(); });
        }
      }
    }

    if (!templates['DEFAULT']) {
      templates['DEFAULT'] = EVAL_CRITERIA.slice();
    }

    return { success: true, data: templates };
  } catch (e) {
    Logger.log('getEvalTemplates 오류: ' + e.message);
    return { success: false, message: '평가 템플릿 조회 실패: ' + e.message };
  }
}

function saveEvalTemplate(org, criteriaStr) {
  try {
    var user = Session.getActiveUser().getEmail();
    var role = getUserRole(user);
    if (!hasPermission_(role, ROLE_LEVELS['최고관리자'])) {
      return { success: false, message: '평가 템플릿 변경 권한이 없습니다.' };
    }

    var criteria = criteriaStr.split(',').map(function(c) { return c.trim(); }).filter(function(c) { return c.length > 0; });

    // 기본 5항목 필수 검증
    for (var b = 0; b < EVAL_CRITERIA.length; b++) {
      if (criteria.indexOf(EVAL_CRITERIA[b]) === -1) {
        return { success: false, message: '기본 항목 "' + EVAL_CRITERIA[b] + '"은(는) 필수입니다.' };
      }
    }

    var key = 'EVAL_TEMPLATE_' + org;
    var result = updateSetting(key, criteria.join(','));

    if (result.success) {
      writeLog_(user, '템플릿변경', key, org + ' 평가 템플릿: ' + criteria.join(','));
    }
    return { success: true, message: org + ' 평가 템플릿이 저장되었습니다.', data: criteria };
  } catch (e) {
    Logger.log('saveEvalTemplate 오류: ' + e.message);
    return { success: false, message: '템플릿 저장 실패: ' + e.message };
  }
}

// ============ 마이그레이션: 평가 템플릿 ============

function migrateEvalTemplates() {
  try {
    // 1. 평가 시트에 추가항목 컬럼 헤더 추가
    var evalSheet = getSheet(SHEET_NAMES.EVALUATION);
    var header = evalSheet.getRange(1, 1, 1, evalSheet.getLastColumn()).getValues()[0];
    var hasExtra = false;
    for (var i = 0; i < header.length; i++) {
      if (header[i] === '추가항목') { hasExtra = true; break; }
    }
    if (!hasExtra) {
      var nextCol = header.length + 1;
      evalSheet.getRange(1, nextCol).setValue('추가항목');
    }

    // 2. 설정 시트에 기본 템플릿 시드
    var settingsSheet = getSheet(SHEET_NAMES.SETTINGS);
    var settingsData = settingsSheet.getDataRange().getValues();
    var existingKeys = {};
    for (var s = 1; s < settingsData.length; s++) {
      if (settingsData[s][0]) existingKeys[settingsData[s][0].toString()] = true;
    }

    if (!existingKeys['EVAL_TEMPLATE_DEFAULT']) {
      settingsSheet.appendRow(['EVAL_TEMPLATE_DEFAULT', '정확성,속도,협업,출석,태도', '기본 평가 템플릿', getToday_()]);
    }
    if (!existingKeys['EVAL_TEMPLATE_검토팀']) {
      settingsSheet.appendRow(['EVAL_TEMPLATE_검토팀', '정확성,속도,협업,출석,태도,일 평균 처리개수', '검토팀 평가 템플릿', getToday_()]);
    }

    Logger.log('migrateEvalTemplates 완료');
    return { success: true, message: '평가 템플릿 마이그레이션 완료' };
  } catch (e) {
    Logger.log('migrateEvalTemplates 오류: ' + e.message);
    return { success: false, message: '마이그레이션 실패: ' + e.message };
  }
}

// ============ 마이그레이션: 관리자 권한 설정 ============

function setupAdminPermissions() {
  var permMap = {
    '서기훈': { role: '최고관리자', org: '' },
    '김현아': { role: '최고관리자', org: '' },
    '김종섭': { role: '최고관리자', org: '' },
    '문지영': { role: '총괄관리자', org: '고객지원팀,인용확인팀' },
    '이진수': { role: '총괄관리자', org: '' },
    '이천보': { role: '조직관리자', org: '검토팀' },
    '이은별': { role: '조직관리자', org: '분류팀,뉴터칭콜팀' },
    '김다니엘': { role: '조직관리자', org: '신고팀,작성팀' },
    '정서림': { role: '조직관리자', org: '세무서대응팀' }
  };

  var sheet = getSheet(SHEET_NAMES.ADMIN);
  var data = sheet.getDataRange().getValues();
  var updated = 0;
  var notFound = [];

  for (var i = 1; i < data.length; i++) {
    var name = (data[i][1] || '').toString().trim();
    if (permMap[name]) {
      sheet.getRange(i + 1, 3).setValue(permMap[name].role);
      sheet.getRange(i + 1, 4).setValue(permMap[name].org);
      sheet.getRange(i + 1, 5).setValue('활성');
      updated++;
      delete permMap[name];
    }
  }

  // 매칭 안 된 사람 리스트
  for (var nm in permMap) {
    notFound.push(nm);
  }

  _adminCache = null;
  var msg = updated + '명 권한 업데이트 완료.';
  if (notFound.length > 0) msg += ' 미발견: ' + notFound.join(', ');
  Logger.log('setupAdminPermissions: ' + msg);
  return { success: true, message: msg };
}
