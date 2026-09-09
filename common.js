// 공통 API 헬퍼
// GAS 웹앱은 application/json Content-Type을 보내면 브라우저가 preflight(OPTIONS)를
// 요청하는데 Apps Script가 이를 처리하지 못해 실패한다. 그래서 Content-Type을
// 지정하지 않고(text/plain 기본값) 문자열 그대로 보낸다. GAS 쪽에서는
// e.postData.contents 를 JSON.parse 하므로 문제 없다.

async function apiPost(payload) {
  const res = await fetch(APPS_SCRIPT_URL, {
    method: 'POST',
    body: JSON.stringify(payload)
  });
  if (!res.ok) throw new Error('네트워크 오류 (' + res.status + ')');
  return res.json();
}

async function apiGet(params) {
  const url = new URL(APPS_SCRIPT_URL);
  Object.entries(params).forEach(([k, v]) => url.searchParams.set(k, v));
  const res = await fetch(url.toString());
  if (!res.ok) throw new Error('네트워크 오류 (' + res.status + ')');
  return res.json();
}

const JUDGE_OPTIONS = ['상향', '적정', '안정', '미정'];
const PRIORITY_OPTIONS = ['1순위', '2순위', '3순위', '4순위', '5순위', '6순위'];
const MAX_APPLICATIONS = 6;

const FIELD_LABELS = [
  ['우선순위', 'select', PRIORITY_OPTIONS],
  ['학교명', 'text'], ['학과', 'text'], ['전형유형', 'text'], ['전형명', 'text'],
  ['모집인원', 'text'], ['수능최저', 'text'], ['면접', 'text'], ['내신점수', 'text'],
  ['전년도평균', 'text'], ['원서마감일', 'date'], ['합격발표일', 'date'],
  ['수험번호', 'text'], ['판단', 'select', JUDGE_OPTIONS], ['기타의견', 'text'], ['추천전형', 'checkbox']
];

function judgeClass(judge) {
  const map = { '상향': 'judge-reach', '적정': 'judge-match', '안정': 'judge-safe', '미정': 'judge-unset' };
  return map[judge] || 'judge-unset';
}

function escapeHtml(str) {
  return String(str ?? '').replace(/[&<>"']/g, c => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
  })[c]);
}

function toCSV(rows, columns) {
  const esc = v => {
    const s = String(v ?? '');
    return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  };
  const lines = [columns.join(',')];
  rows.forEach(r => lines.push(columns.map(c => esc(r[c])).join(',')));
  return '\uFEFF' + lines.join('\n'); // BOM 포함: 엑셀에서 한글 깨짐 방지
}

function downloadCSV(filename, rows, columns) {
  const csv = toCSV(rows, columns);
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}
