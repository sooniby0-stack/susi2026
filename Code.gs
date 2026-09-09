/**
 * 수시지원 관리 - Google Apps Script 백엔드
 *
 * 사용법
 * 1) 구글시트를 새로 만들고 이 스크립트를 [확장 프로그램 > Apps Script]에 붙여넣는다.
 * 2) 상단 실행 드롭다운에서 setup 함수를 한 번 실행해 헤더/기본 설정을 만든다.
 *    (처음 실행 시 권한 승인 창이 뜨면 허용한다)
 * 3) [배포 > 새 배포] > 유형: 웹앱
 *      - 실행 계정: 나
 *      - 액세스 권한이 있는 사용자: 전체
 *    배포 후 나오는 웹앱 URL을 web/config.js 의 APPS_SCRIPT_URL 에 붙여넣는다.
 * 4) 교사 비밀번호를 바꾸고 싶으면 setTeacherPassword("새비번") 함수를 실행한다.
 */

const SHEET_NAME = 'data';
const HEADERS = [
  '학번', '비밀번호', '이름', '반', '순번',
  '학교명', '학과', '전형유형', '전형명', '모집인원',
  '수능최저', '면접', '내신점수', '전년도평균',
  '원서마감일', '합격발표일', '수험번호', '판단', '기타의견', '추천전형',
  '우선순위'
];

// ---------- 초기 설정 ----------

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  }
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('TEACHER_PASSWORD')) {
    props.setProperty('TEACHER_PASSWORD', '1234');
  }
}

function setTeacherPassword(pw) {
  PropertiesService.getScriptProperties().setProperty('TEACHER_PASSWORD', String(pw));
}

// 기존에 이미 쓰던 시트에 '우선순위' 컬럼(맨 뒤)을 반영하고 싶을 때 한 번 실행.
// 실제 데이터 위치는 건드리지 않고, 1행(헤더 표시용 글자)만 최신 HEADERS로 다시 씁니다.
function migrateAddPriority() {
  const sheet = getSheet_();
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
}

// ---------- 시트 읽기/쓰기 ----------

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('시트가 없습니다. setup()을 먼저 실행하세요.');
  return sheet;
}

function readAll_() {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  return values
    .filter(r => String(r[0]).trim() !== '') // 학번 없는 빈 행 무시
    .map(r => {
      const obj = {};
      HEADERS.forEach((h, i) => { obj[h] = r[i]; });
      return obj;
    });
}

function writeAll_(rows) {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, HEADERS.length).clearContent();
  }
  if (rows.length === 0) return;
  const values = rows.map(r => HEADERS.map(h => (r[h] !== undefined && r[h] !== null) ? r[h] : ''));
  sheet.getRange(2, 1, values.length, HEADERS.length).setValues(values);
}

function isTeacherPasswordValid_(pw) {
  const stored = PropertiesService.getScriptProperties().getProperty('TEACHER_PASSWORD') || '1234';
  return String(pw) === String(stored);
}

// ---------- HTTP 엔드포인트 ----------

function doGet(e) {
  const action = e.parameter.action;
  let result;
  try {
    if (action === 'all') {
      result = { ok: true, rows: readAll_() };
    } else if (action === 'student') {
      const id = e.parameter.id;
      const rows = readAll_().filter(r => String(r['학번']) === String(id));
      result = { ok: true, rows: rows };
    } else {
      result = { ok: false, error: '알 수 없는 요청입니다' };
    }
  } catch (err) {
    result = { ok: false, error: String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOut_({ ok: false, error: '잘못된 요청 형식입니다' });
  }

  const action = body.action;
  let result = { ok: false, error: '알 수 없는 요청입니다' };

  try {
    if (action === 'login') {
      result = handleLogin_(body);
    } else if (action === 'saveApplications') {
      result = handleSaveApplications_(body);
    } else if (action === 'changePassword') {
      result = handleChangePassword_(body);
    } else if (action === 'teacherLogin') {
      result = { ok: isTeacherPasswordValid_(body.password) };
      if (!result.ok) result.error = '비밀번호가 일치하지 않습니다';
    } else if (action === 'addStudent') {
      result = handleAddStudent_(body);
    } else if (action === 'bulkAddStudents') {
      result = handleBulkAddStudents_(body);
    } else if (action === 'deleteStudent') {
      result = handleDeleteStudent_(body);
    } else if (action === 'resetPassword') {
      result = handleResetPassword_(body);
    }
  } catch (err) {
    result = { ok: false, error: String(err) };
  }

  return jsonOut_(result);
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------- 액션 구현 ----------

function handleLogin_(body) {
  const rows = readAll_();
  const mine = rows.filter(r => String(r['학번']) === String(body.id));
  if (mine.length === 0) return { ok: false, error: '존재하지 않는 학번입니다' };
  const master = mine.find(r => Number(r['순번']) === 0) || mine[0];
  if (String(master['비밀번호']) !== String(body.pw)) {
    return { ok: false, error: '비밀번호가 일치하지 않습니다' };
  }
  const applications = mine
    .filter(r => Number(r['순번']) > 0)
    .sort((a, b) => Number(a['순번']) - Number(b['순번']));
  return {
    ok: true,
    name: master['이름'],
    cls: master['반'],
    applications: applications
  };
}

function handleSaveApplications_(body) {
  // body: { id, pw, applications: [ {학교명, 학과, ...}, ... ] }
  const rows = readAll_();
  const mine = rows.filter(r => String(r['학번']) === String(body.id));
  if (mine.length === 0) return { ok: false, error: '존재하지 않는 학번입니다' };
  const master = mine.find(r => Number(r['순번']) === 0) || mine[0];
  if (String(master['비밀번호']) !== String(body.pw)) {
    return { ok: false, error: '비밀번호가 일치하지 않습니다' };
  }
  const others = rows.filter(r => String(r['학번']) !== String(body.id));

  const newMaster = Object.assign({}, master, {
    순번: 0, 학교명: '', 학과: '', 전형유형: '', 전형명: '', 모집인원: '',
    수능최저: '', 면접: '', 내신점수: '', 전년도평균: '', 원서마감일: '',
    합격발표일: '', 수험번호: '', 판단: '', 기타의견: '', 추천전형: '', 우선순위: ''
  });

  const newRows = (body.applications || []).map((app, i) => Object.assign(
    { 학번: master['학번'], 비밀번호: master['비밀번호'], 이름: master['이름'], 반: master['반'] },
    { 순번: i + 1 },
    app
  ));

  writeAll_(others.concat([newMaster], newRows));
  return { ok: true };
}

function handleChangePassword_(body) {
  const rows = readAll_();
  const mine = rows.filter(r => String(r['학번']) === String(body.id));
  if (mine.length === 0) return { ok: false, error: '존재하지 않는 학번입니다' };
  const master = mine.find(r => Number(r['순번']) === 0) || mine[0];
  if (String(master['비밀번호']) !== String(body.oldPw)) {
    return { ok: false, error: '기존 비밀번호가 일치하지 않습니다' };
  }
  rows.filter(r => String(r['학번']) === String(body.id)).forEach(r => { r['비밀번호'] = body.newPw; });
  writeAll_(rows);
  return { ok: true };
}

function handleAddStudent_(body) {
  // body: { teacherPassword, id, name, cls }
  if (!isTeacherPasswordValid_(body.teacherPassword)) return { ok: false, error: '교사 인증 실패' };
  const rows = readAll_();
  if (rows.some(r => String(r['학번']) === String(body.id))) {
    return { ok: false, error: '이미 존재하는 학번입니다' };
  }
  rows.push({
    학번: body.id, 비밀번호: body.id, 이름: body.name, 반: body.cls, 순번: 0
  });
  writeAll_(rows);
  return { ok: true };
}

function handleBulkAddStudents_(body) {
  // body: { teacherPassword, students: [ {id, name, cls}, ... ] }
  if (!isTeacherPasswordValid_(body.teacherPassword)) return { ok: false, error: '교사 인증 실패' };
  const rows = readAll_();
  const existingIds = new Set(rows.map(r => String(r['학번'])));
  const students = body.students || [];

  const added = [];
  const skipped = [];

  students.forEach(s => {
    const id = String(s.id || '').trim();
    const name = String(s.name || '').trim();
    const cls = String(s.cls || '').trim();
    if (!id || !name) { skipped.push({ id: id || '(빈 학번)', reason: '학번 또는 이름 누락' }); return; }
    if (existingIds.has(id)) { skipped.push({ id, reason: '이미 존재하는 학번' }); return; }
    rows.push({ 학번: id, 비밀번호: id, 이름: name, 반: cls, 순번: 0 });
    existingIds.add(id);
    added.push(id);
  });

  if (added.length > 0) writeAll_(rows);
  return { ok: true, addedCount: added.length, added: added, skipped: skipped };
}

function handleDeleteStudent_(body) {
  if (!isTeacherPasswordValid_(body.teacherPassword)) return { ok: false, error: '교사 인증 실패' };
  const rows = readAll_().filter(r => String(r['학번']) !== String(body.id));
  writeAll_(rows);
  return { ok: true };
}

function handleResetPassword_(body) {
  if (!isTeacherPasswordValid_(body.teacherPassword)) return { ok: false, error: '교사 인증 실패' };
  const rows = readAll_();
  const mine = rows.filter(r => String(r['학번']) === String(body.id));
  if (mine.length === 0) return { ok: false, error: '존재하지 않는 학번입니다' };
  mine.forEach(r => { r['비밀번호'] = String(body.id); });
  writeAll_(rows);
  return { ok: true };
}
