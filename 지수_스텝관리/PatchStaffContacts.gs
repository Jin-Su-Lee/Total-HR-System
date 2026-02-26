/**
 * PatchStaffContacts.gs
 * =====================
 * One-time migration script: patches contact info (연락처, 이메일, 생년월일)
 * into the staff-management spreadsheet from Shiftee employee data.
 *
 * Staff sheet columns:
 *   1:스텝ID, 2:이름, 3:연락처, 4:이메일, 5:생년월일, ...
 *
 * Usage: Run patchStaffContactInfo() from the Apps Script editor.
 * Generated: 2026-02-24 18:06:21
 */

function patchStaffContactInfo() {
  var SPREADSHEET_ID = '1iBSDGrjV_SMX4Uftt3PsQoCUwZP66WlZ6fTndEUHlmY';
  var SHEET_NAME = '스텝';

  // [stfId, phone, email, birthday]
  var contactData = [
    ['STF-0001', '', 'seohyun001212@naver.com', ''],
    ['STF-0002', '', 'kdy5008@zsoo.kr', ''],
    ['STF-0003', '', 'ksg2064@zsoo.kr', ''],
    ['STF-0004', '', 'kjy1098@zsoo.kr', ''],
    ['STF-0005', '', 'kjh9759@zsoo.kr', ''],
    ['STF-0006', '010-4993-2506', 'kjw2506@zsoo.kr', ''],
    ['STF-0007', '', 'mhj3644@zsoo.kr', ''],
    ['STF-0008', '', 'pdg6637@zsoo.kr', ''],
    ['STF-0009', '', 'bbs1014@zsoo.kr', ''],
    ['STF-0010', '010-3272-6231', 'yjh6231@zsoo.kr', ''],
    ['STF-0011', '', 'ljh8254@zsoo.kr', ''],
    ['STF-0012', '', 'lhj6115@zsoo.kr', ''],
    ['STF-0013', '', 'jjh3912@zsoo.kr', ''],
    ['STF-0014', '', 'hsh6716@zsoo.kr', ''],
    ['STF-0015', '', 'ytj8057@zsoo.kr', ''],
    ['STF-0016', '', 'jsy5826@zsoo.kr', ''],
    ['STF-0017', '', 'jeh5329@zsoo.kr', ''],
    ['STF-0018', '', 'khy4717@zsoo.kr', ''],
    ['STF-0019', '', 'sbn5423@zsoo.kr', ''],
    ['STF-0020', '', 'ash5569@zsoo.kr', ''],
    ['STF-0021', '', 'yej8561@zsoo.kr', ''],
    ['STF-0022', '', 'ksh5951@zsoo.kr', ''],
    ['STF-0023', '', 'kyh6437@zsoo.kr', ''],
    ['STF-0024', '', 'phb7818@zsoo.kr', ''],
    ['STF-0025', '', 'sec5765@zsoo.kr', ''],
    ['STF-0026', '', 'sjw5025@zsoo.kr', ''],
    ['STF-0027', '', 'sbn6503@zsoo.kr', ''],
    ['STF-0028', '', 'whr8222@zsoo.kr', ''],
    ['STF-0029', '', 'ysh5778@zsoo.kr', ''],
    ['STF-0030', '', 'lya1561@zsoo.kr', ''],
    ['STF-0031', '', 'lhr0902@zsoo.kr', ''],
    ['STF-0032', '', 'hes1142@zsoo.kr', ''],
    ['STF-0033', '010-4932-6558', 'sjy6558@zsoo.kr', ''],
    ['STF-0034', '010-9922-3395', 'yej3395@zsoo.kr', ''],
    ['STF-0035', '010-7515-3926', 'lei3926@zsoo.kr', ''],
    ['STF-0036', '010-4592-1817', 'jys1817@zsoo.kr', ''],
    ['STF-0038', '', 'ksy5687@zsoo.kr', ''],
    ['STF-0039', '', 'kyj4981@zsoo.kr', ''],
    ['STF-0040', '', 'js1448@zsoo.kr', ''],
    ['STF-0041', '', 'mgb8874@zsoo.kr', ''],
    ['STF-0042', '', 'sgg5969@zsoo.kr', ''],
    ['STF-0043', '', 'kmg6332@zsoo.kr', ''],
    ['STF-0044', '010-8671-2363', 'kje2363@zsoo.kr', ''],
    ['STF-0046', '010-7551-4001', 'jyb4001@zsoo.kr', ''],
    ['STF-0047', '', 'yjh0804@zsoo.kr', ''],
    ['STF-0048', '', 'jhr6910@zsoo.kr', ''],
    ['STF-0049', '', 'che4491@zsoo.kr', ''],
    ['STF-0050', '', 'khy1408@zsoo.kr', ''],
    ['STF-0051', '', 'khy5569@zsoo.kr', ''],
    ['STF-0052', '', 'kyw4288@zsoo.kr', ''],
    ['STF-0053', '', 'jhc2731@zsoo.kr', ''],
    ['STF-0054', '', 'lsh4340@zsoo.kr', ''],
    ['STF-0055', '', 'shm1533@zsoo.kr', '']
  ];

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    Logger.log('ERROR: Sheet "' + SHEET_NAME + '" not found');
    return;
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('ERROR: No data rows found');
    return;
  }

  // Read all STF IDs from column 1 (rows 2..lastRow)
  var stfIds = sheet.getRange(2, 1, lastRow - 1, 1).getValues(); // [[stfId], ...]
  var idToRow = {};
  for (var i = 0; i < stfIds.length; i++) {
    var id = String(stfIds[i][0]).trim();
    if (id) {
      idToRow[id] = i + 2; // sheet row (1-indexed, skip header)
    }
  }

  var updated = 0;
  var notFound = [];

  // Read existing contact columns (3:연락처, 4:이메일, 5:생년월일) for batch update
  var existingData = sheet.getRange(2, 3, lastRow - 1, 3).getValues(); // [[phone, email, birthday], ...]

  for (var d = 0; d < contactData.length; d++) {
    var stfId = contactData[d][0];
    var phone = contactData[d][1];
    var email = contactData[d][2];
    var birthday = contactData[d][3];

    var row = idToRow[stfId];
    if (!row) {
      notFound.push(stfId);
      continue;
    }

    var arrIdx = row - 2; // index into existingData array

    // Only overwrite if we have new data (don't blank out existing data)
    if (phone) {
      existingData[arrIdx][0] = phone;
    }
    if (email) {
      existingData[arrIdx][1] = email;
    }
    if (birthday) {
      existingData[arrIdx][2] = birthday;
    }

    updated++;
  }

  // Batch write all contact columns at once
  sheet.getRange(2, 3, lastRow - 1, 3).setValues(existingData);

  Logger.log('Patch complete: ' + updated + ' staff updated');
  if (notFound.length > 0) {
    Logger.log('Not found in sheet: ' + notFound.join(', '));
  }

  return {
    success: true,
    message: updated + '명 연락처 업데이트 완료',
    updated: updated,
    notFound: notFound
  };
}
