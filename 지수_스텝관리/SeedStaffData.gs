// ============================================================
// 지수회계법인 — 환급사업부 스텝관리 시스템
// SeedStaffData.gs (시드 데이터)
// ============================================================

// ============ 스텝 시드 데이터 ============

/**
 * 55명의 스텝 데이터를 '스텝' 시트에 일괄 등록한다.
 * 근태관리 시스템의 직원 목록과 동일한 인원을 사용한다.
 * 재직 → 근무중, 퇴사 → 계약종료 로 매핑한다.
 */
function seedStaffData() {
  var sheet = getSheet(SHEET_NAMES.STAFF);
  var existing = sheet.getDataRange().getValues();
  if (existing.length > 1) {
    Logger.log('스텝 시트에 이미 데이터가 존재합니다. 시드 중단.');
    return { success: false, message: '스텝 시트에 이미 데이터가 있습니다. 초기화 후 재시도하세요.' };
  }

  var now = getNow_();

  // 3개월 후 날짜를 계산하는 내부 헬퍼
  function addMonths_(dateStr, months) {
    var parts = dateStr.split('-');
    var d = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1 + months, parseInt(parts[2]));
    var yyyy = d.getFullYear();
    var mm = ('0' + (d.getMonth() + 1)).slice(-2);
    var dd = ('0' + d.getDate()).slice(-2);
    return yyyy + '-' + mm + '-' + dd;
  }

  // 퇴사자의 퇴사일을 비고에서 추출하는 헬퍼
  // 퇴사일이 있으면 그것을 계약종료일로 사용
  function getResignDate_(memo) {
    if (!memo) return null;
    var match = memo.match(/퇴사:(\d{4}-\d{2}-\d{2})/);
    return match ? match[1] : null;
  }

  // 사원번호를 비고에서 추출하는 헬퍼
  function getShifteeId_(memo) {
    if (!memo) return '';
    var match = memo.match(/사원번호:(\d+)/);
    return match ? 'Shiftee#' + match[1] : '';
  }

  // 원본 직원 데이터: [empId, name, phone, email, org, job, hireDate, hourlyPay, dailyPay, bank, account, annualLeave, status, memo]
  var employees = [
    // ── 검토팀 (14명) ──
    ['EMP-0001', '고서현', '', '', '검토팀', '경정청구팀', '2025-09-22', 11000, 88000, '', '', 7, '재직', '사원번호:710'],
    ['EMP-0002', '김도연3', '', '', '검토팀', '경정청구팀', '2025-08-20', 11000, 88000, '', '', 10, '재직', '사원번호:209'],
    ['EMP-0003', '김승국', '', '', '검토팀', '경정청구팀', '2025-04-01', 11000, 88000, '', '', 8, '재직', '사원번호:254'],
    ['EMP-0004', '김지연', '', '', '검토팀', '경정청구팀', '2025-07-16', 11000, 88000, '', '', 7, '재직', '사원번호:285'],
    ['EMP-0005', '김지현', '', '', '검토팀', '경정청구팀', '2025-01-13', 11000, 88000, '', '', 2, '재직', '사원번호:243'],
    ['EMP-0006', '김진웅', '', '', '검토팀', '경정청구팀', '2025-11-10', 11000, 88000, '', '', 3, '재직', '사원번호:895'],
    ['EMP-0007', '맹현진', '', '', '검토팀', '경정청구팀', '2025-12-15', 11000, 88000, '', '', 2, '재직', '사원번호:926'],
    ['EMP-0008', '박도겸', '', '', '검토팀', '경정청구팀', '2024-08-14', 11000, 88000, '', '', 14, '재직', '사원번호:226'],
    ['EMP-0009', '배병수', '', '', '검토팀', '경정청구팀', '2025-03-31', 11000, 88000, '', '', 0, '재직', '사원번호:253'],
    ['EMP-0010', '유지훈', '', '', '검토팀', '경정청구팀', '2025-03-04', 11000, 88000, '', '', 10, '재직', '사원번호:250'],
    ['EMP-0011', '이주현', '', '', '검토팀', '경정청구팀', '2025-05-07', 11000, 88000, '', '', 0, '재직', '사원번호:260'],
    ['EMP-0012', '임현지', '', '', '검토팀', '경정청구팀', '2025-11-24', 11000, 88000, '', '', 2, '재직', '사원번호:952'],
    ['EMP-0013', '정재현', '', '', '검토팀', '경정청구팀', '2025-08-20', 11000, 88000, '', '', 5, '재직', '사원번호:303'],
    ['EMP-0014', '한승헌', '', '', '검토팀', '경정청구팀', '2024-12-30', 11000, 88000, '', '', 3, '재직', '사원번호:623'],
    // ── 고객지원팀 (3명) ──
    ['EMP-0015', '양태정', '', '', '고객지원팀', '경정청구팀', '2025-07-15', 11000, 88000, '', '', 1, '재직', '사원번호:911'],
    ['EMP-0016', '정솔잎', '', '', '고객지원팀', '경정청구팀', '2025-11-13', 11000, 88000, '', '', 2, '재직', '사원번호:913'],
    ['EMP-0017', '정은화', '', '', '고객지원팀', '경정청구팀', '2026-02-05', 11000, 88000, '', '', 0, '재직', '사원번호:914'],
    // ── 뉴터칭콜팀 (4명) ──
    ['EMP-0018', '김하얀', '', '', '뉴터칭콜팀', '경정청구팀', '2025-08-01', 11000, 88000, '', '', 5, '재직', '사원번호:530'],
    ['EMP-0019', '석빛날', '', '', '뉴터칭콜팀', '경정청구팀', '2026-02-23', 11000, 88000, '', '', 0, '재직', '사원번호:549'],
    ['EMP-0020', '안선희', '', '', '뉴터칭콜팀', '경정청구팀', '2025-12-16', 11000, 88000, '', '', 1, '재직', '사원번호:541'],
    ['EMP-0021', '유은지', '', '', '뉴터칭콜팀', '경정청구팀', '2025-07-02', 11000, 88000, '', '', 6, '재직', '사원번호:529'],
    // ── 분류팀 (11명) ──
    ['EMP-0022', '구서현', '', '', '분류팀', '경정청구팀', '2026-02-09', 12000, 96000, '', '', 0, '재직', '사원번호:642'],
    ['EMP-0023', '김영현', '', '', '분류팀', '경정청구팀', '2025-12-17', 12000, 96000, '', '', 0, '재직', '사원번호:639'],
    ['EMP-0024', '박한별', '', '', '분류팀', '경정청구팀', '2026-02-23', 12000, 96000, '', '', 0, '재직', '사원번호:643'],
    ['EMP-0025', '서은총', '', '', '분류팀', '경정청구팀', '2025-09-01', 12000, 96000, '', '', 5, '재직', '사원번호:631'],
    ['EMP-0026', '소지원', '', '', '분류팀', '경정청구팀', '2026-02-24', 12000, 96000, '', '', 0, '재직', '사원번호:644'],
    ['EMP-0027', '송빈나', '', '', '분류팀', '경정청구팀', '2025-12-08', 12000, 96000, '', '', 0, '재직', '사원번호:635'],
    ['EMP-0028', '원희랑', '', '', '분류팀', '경정청구팀', '2025-02-10', 12000, 96000, '', '', 2, '재직', '사원번호:264'],
    ['EMP-0029', '유승희', '', '', '분류팀', '경정청구팀', '2025-09-23', 12000, 96000, '', '', 2, '재직', '사원번호:634'],
    ['EMP-0030', '이연아', '', '', '분류팀', '경정청구팀', '2025-08-20', 12000, 96000, '', '', 4, '재직', '사원번호:630'],
    ['EMP-0031', '이향림', '', '', '분류팀', '경정청구팀', '2025-12-29', 12000, 96000, '', '', 0, '재직', '사원번호:641'],
    ['EMP-0032', '황은선', '', '', '분류팀', '경정청구팀', '2025-09-08', 12000, 96000, '', '', 5, '재직', '사원번호:632'],
    // ── 신고팀 (5명) ──
    ['EMP-0033', '신주용', '', '', '신고팀', '경정청구팀', '2025-11-11', 11000, 88000, '', '', 3, '재직', '사원번호:417'],
    ['EMP-0034', '윤은지', '', '', '신고팀', '경정청구팀', '2025-10-20', 11000, 88000, '', '', 0, '재직', '사원번호:416'],
    ['EMP-0035', '임은일', '', '', '신고팀', '경정청구팀', '2025-12-09', 11000, 88000, '', '', 1, '재직', '사원번호:418'],
    ['EMP-0036', '조윤서', '', '', '신고팀', '경정청구팀', '2026-01-19', 11000, 88000, '', '', 0, '재직', '사원번호:421'],
    ['EMP-0037', '최서영', '', '', '신고팀', '경정청구팀', '2026-01-05', 11000, 88000, '', '', 0, '재직', '사원번호:420'],
    // ── 인용확인팀 (5명) ──
    ['EMP-0038', '김세율', '', '', '인용확인팀', '경정청구팀', '2026-01-16', 11000, 88000, '', '', 0, '재직', '사원번호:871'],
    ['EMP-0039', '김예진', '', '', '인용확인팀', '경정청구팀', '2026-02-05', 11000, 88000, '', '', 0, '재직', '사원번호:874'],
    ['EMP-0040', '김지석', '', '', '인용확인팀', '경정청구팀', '2025-09-10', 11000, 88000, '', '', 4, '재직', '사원번호:863'],
    ['EMP-0041', '민경빈', '', '', '인용확인팀', '경정청구팀', '2026-01-07', 11000, 88000, '', '', 1, '재직', '사원번호:870'],
    ['EMP-0042', '서건규', '', '', '인용확인팀', '경정청구팀', '2025-07-14', 11000, 88000, '', '', 2, '재직', '사원번호:860'],
    // ── 작성팀 (4명) ──
    ['EMP-0043', '김민기', '', '', '작성팀', '경정청구팀', '2025-08-18', 11000, 88000, '', '', 4, '재직', '사원번호:738'],
    ['EMP-0044', '김지은', '', '', '작성팀', '경정청구팀', '2025-12-29', 11000, 88000, '', '', 0, '재직', '사원번호:746'],
    ['EMP-0045', '박유림', '', '', '작성팀', '경정청구팀', '2025-06-23', 11000, 88000, '', '', 7, '재직', '사원번호:736'],
    ['EMP-0046', '주유빈', '', '', '작성팀', '경정청구팀', '2025-12-08', 11000, 88000, '', '', 1, '재직', '사원번호:744'],
    // ── 퇴사자 (9명) ──
    ['EMP-0047', '양지현', '', '', '뉴터칭콜팀', '경정청구팀', '2025-04-14', 11000, 88000, '', '', 4, '퇴사', '퇴사:2025-11-11,사원번호:524'],
    ['EMP-0048', '주하림', '', '', '작성팀', '경정청구팀', '2025-05-14', 11000, 88000, '', '', 5, '퇴사', '퇴사:2025-11-13,사원번호:735'],
    ['EMP-0049', '최하은', '', '', '분류팀', '경정청구팀', '2024-05-27', 12000, 96000, '', '', 22, '퇴사', '퇴사:2025-12-10,사원번호:612'],
    ['EMP-0050', '김희연', '', '', '분류팀', '경정청구팀', '2025-12-29', 12000, 96000, '', '', 0, '퇴사', '퇴사:2026-02-25,사원번호:628'],
    ['EMP-0051', '권희율', '', '', '뉴터칭콜팀', '경정청구팀', '2025-12-17', 11000, 88000, '', '', 2, '퇴사', '퇴사:2026-02-27,사원번호:543'],
    ['EMP-0052', '김예원', '', '', '분류팀', '경정청구팀', '2026-01-05', 12000, 96000, '', '', 1, '퇴사', '퇴사:2026-02-27,사원번호:611'],
    ['EMP-0053', '정현채', '', '', '분류팀', '경정청구팀', '2025-12-23', 12000, 96000, '', '', 1, '퇴사', '퇴사:2026-02-27,사원번호:622'],
    ['EMP-0054', '이소희', '', '', '뉴터칭콜팀', '경정청구팀', '2026-01-07', 11000, 88000, '', '', 1, '퇴사', '퇴사:2026-03-31,사원번호:546'],
    ['EMP-0055', '손혜미', '', '', '분류팀', '경정청구팀', '2025-03-04', 12000, 96000, '', '', 9, '퇴사', '퇴사:2026-04-03,사원번호:265']
  ];

  // STAFF_COLUMNS 매핑 (20컬럼):
  // [스텝ID, 이름, 연락처, 이메일, 생년월일, 성별, 주소, 지원경로, 조직, 직무,
  //  상태, 계약시작일, 계약종료일, 급여형태, 급여액, 은행명, 계좌번호, 비고, 등록일, 수정일]
  var staffRows = [];
  for (var i = 0; i < employees.length; i++) {
    var emp = employees[i];
    var empId = emp[0];          // EMP-XXXX
    var name = emp[1];           // 이름
    var phone = emp[2];          // 연락처
    var email = emp[3];          // 이메일
    var org = emp[4];            // 조직
    var job = emp[5];            // 직무
    var hireDate = emp[6];       // 입사일 → 계약시작일
    var hourlyPay = emp[7];      // 시급
    var dailyPay = emp[8];       // 일급
    var bank = emp[9];           // 은행명
    var account = emp[10];       // 계좌번호
    var empStatus = emp[12];     // 재직/퇴사
    var memo = emp[13];          // 비고 (사원번호, 퇴사일 포함)

    // staffId: STF-0001 ~ STF-0055
    var seq = i + 1;
    var staffId = 'STF-' + ('0000' + seq).slice(-4);

    // 상태 매핑: 재직 → 근무중, 퇴사 → 계약종료
    var status = (empStatus === '퇴사') ? '계약종료' : '근무중';

    // 계약시작일 = 입사일
    var contractStart = hireDate;

    // 계약종료일: 퇴사자는 퇴사일 사용, 재직자는 입사일 + 3개월
    var resignDate = getResignDate_(memo);
    var contractEnd = resignDate ? resignDate : addMonths_(hireDate, 3);

    // 비고: Shiftee 사원번호 + 근태시스템 직원ID
    var shifteeTag = getShifteeId_(memo);
    var noteItems = [];
    if (shifteeTag) noteItems.push(shifteeTag);
    noteItems.push('근태ID:' + empId);
    var note = noteItems.join(', ');

    // 급여형태: 시급제
    var payType = '시급제';
    // 급여액: 시급 사용
    var payAmount = hourlyPay;

    var row = [
      staffId,        // 스텝ID
      name,           // 이름
      phone,          // 연락처
      email,          // 이메일
      '',             // 생년월일 (미보유)
      '',             // 성별 (미보유)
      '',             // 주소 (미보유)
      '내부배치',      // 지원경로 (환급사업부 내부 배치)
      org,            // 조직
      job,            // 직무
      status,         // 상태 (근무중/계약종료)
      contractStart,  // 계약시작일
      contractEnd,    // 계약종료일
      payType,        // 급여형태
      payAmount,      // 급여액
      bank,           // 은행명
      account,        // 계좌번호
      note,           // 비고
      now,            // 등록일
      now             // 수정일
    ];

    staffRows.push(row);
  }

  // 배치 쓰기 (20컬럼)
  sheet.getRange(2, 1, staffRows.length, STAFF_COLUMNS.length).setValues(staffRows);

  Logger.log('스텝 시드 데이터 등록 완료: ' + staffRows.length + '명');
  return { success: true, message: staffRows.length + '명의 스텝 데이터가 등록되었습니다.', data: { count: staffRows.length } };
}


// ============ 성과평가 시드 데이터 ============

/**
 * 근무중(재직) 스텝에 대해 1인 1건의 샘플 성과평가를 등록한다.
 * 점수 범위: 3~5 (무작위)
 */
function seedEvalData() {
  var evalSheet = getSheet(SHEET_NAMES.EVALUATION);
  var existingEval = evalSheet.getDataRange().getValues();
  if (existingEval.length > 1) {
    Logger.log('성과평가 시트에 이미 데이터가 존재합니다. 시드 중단.');
    return { success: false, message: '성과평가 시트에 이미 데이터가 있습니다. 초기화 후 재시도하세요.' };
  }

  // 스텝 시트에서 근무중인 스텝 목록 조회
  var staffSheet = getSheet(SHEET_NAMES.STAFF);
  var staffData = staffSheet.getDataRange().getValues();
  if (staffData.length <= 1) {
    Logger.log('스텝 시트에 데이터가 없습니다. seedStaffData()를 먼저 실행하세요.');
    return { success: false, message: '스텝 데이터가 없습니다. seedStaffData()를 먼저 실행하세요.' };
  }

  var now = getNow_();
  var today = getToday_();

  // 헤더에서 컬럼 인덱스 확인
  var header = staffData[0];
  var idxStaffId = 0;   // 스텝ID
  var idxName = 1;      // 이름
  var idxOrg = 8;       // 조직
  var idxStatus = 10;   // 상태

  // 등급 판정 함수
  function getGrade_(avg) {
    if (avg >= 4.5) return 'S';
    if (avg >= 4.0) return 'A';
    if (avg >= 3.5) return 'B';
    if (avg >= 3.0) return 'C';
    return 'D';
  }

  // 고정 시드 랜덤 (재현 가능) — 간단한 선형 합동 생성기
  var seed = 42;
  function nextRand_() {
    seed = (seed * 1103515245 + 12345) & 0x7fffffff;
    return seed;
  }

  // 3~5 범위 랜덤 점수
  function randScore_() {
    return (nextRand_() % 3) + 3; // 3, 4, 5
  }

  var evalRows = [];
  var evalCount = 0;

  for (var i = 1; i < staffData.length; i++) {
    var staffId = staffData[i][idxStaffId];
    var staffName = staffData[i][idxName];
    var staffOrg = staffData[i][idxOrg];
    var staffStatus = staffData[i][idxStatus];

    // 근무중인 스텝만 평가
    if (staffStatus !== '근무중') continue;

    evalCount++;
    var evalId = 'EVL-' + ('0000' + evalCount).slice(-4);

    // 무작위 점수 생성 (3~5)
    var accuracy = randScore_();   // 정확성
    var speed = randScore_();      // 속도
    var teamwork = randScore_();   // 협업
    var attendance = randScore_(); // 출석
    var attitude = randScore_();   // 태도

    var total = accuracy + speed + teamwork + attendance + attitude;
    var average = Math.round((total / 5) * 10) / 10; // 소수점 1자리
    var grade = getGrade_(average);

    // EVALUATION_COLUMNS 매핑 (16컬럼):
    // [평가ID, 스텝ID, 스텝이름, 조직, 평가기간, 정확성, 속도, 협업, 출석, 태도,
    //  총점, 평균, 등급, 코멘트, 평가자, 평가일]
    var evalRow = [
      evalId,                           // 평가ID
      staffId,                          // 스텝ID
      staffName,                        // 스텝이름
      staffOrg,                         // 조직
      '2026-01-01 ~ 2026-02-28',       // 평가기간
      accuracy,                         // 정확성
      speed,                            // 속도
      teamwork,                         // 협업
      attendance,                       // 출석
      attitude,                         // 태도
      total,                            // 총점
      average,                          // 평균
      grade,                            // 등급
      '시드 데이터 샘플 평가입니다.',    // 코멘트
      '시스템',                         // 평가자
      today                             // 평가일
    ];

    evalRows.push(evalRow);
  }

  if (evalRows.length === 0) {
    Logger.log('근무중 스텝이 없어 평가 데이터를 생성할 수 없습니다.');
    return { success: false, message: '근무중 스텝이 없습니다.' };
  }

  // 배치 쓰기 (16컬럼)
  evalSheet.getRange(2, 1, evalRows.length, EVALUATION_COLUMNS.length).setValues(evalRows);

  Logger.log('성과평가 시드 데이터 등록 완료: ' + evalRows.length + '건');
  return { success: true, message: evalRows.length + '건의 성과평가 데이터가 등록되었습니다.', data: { count: evalRows.length } };
}


// ============ 전체 시드 실행 ============

/**
 * 스텝 데이터와 성과평가 데이터를 순차적으로 등록한다.
 */
function seedAllStaffData() {
  var results = [];

  // 1) 스텝 데이터 등록
  var staffResult = seedStaffData();
  results.push('스텝: ' + staffResult.message);
  Logger.log('seedStaffData 결과: ' + staffResult.message);

  // 스텝 데이터 등록 실패 시 평가 데이터도 건너뜀
  if (!staffResult.success) {
    return {
      success: false,
      message: results.join(' / '),
      data: results
    };
  }

  // 2) 성과평가 데이터 등록
  var evalResult = seedEvalData();
  results.push('평가: ' + evalResult.message);
  Logger.log('seedEvalData 결과: ' + evalResult.message);

  var allSuccess = staffResult.success && evalResult.success;

  return {
    success: allSuccess,
    message: results.join(' / '),
    data: results
  };
}
