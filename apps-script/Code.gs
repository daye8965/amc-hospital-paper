/**
 * 병원보 편집계획안 → 대시보드 데이터 API (Google Apps Script)
 *
 * 원본: "★병원보 편집계획안(NEW)" (asanpr.members@gmail.com 소유)
 * 탭 하나가 한 호(號)이고, 탭 이름이 호수입니다(예: "786").
 *
 * 배포: 배포 > 새 배포 > 웹 앱 (실행: 나 / 액세스: 링크가 있는 모든 사용자)
 * 나온 /exec URL 을 병원보 대시보드의 GOOGLE_APPS_SCRIPT_URL 에 넣습니다.
 * 이후 수정은 배포 > 배포 관리 > 연필 > 새 버전 으로 해야 URL 이 유지됩니다.
 */

const SPREADSHEET_ID = '1A_pBtxztIUlwBl2an8-hWiVHE38lt3_0nS8HsKJ2doQ';

// 헤더 이름 → 필드. 탭마다 표기가 다를 수 있어 별칭을 둡니다.
const FIELD_ALIASES = {
  page:    ['지면', '면'],
  section: ['코너명', '코너', '구분', '카테고리'],   // 옛 호는 '카테고리'를 씁니다
  title:   ['콘텐츠', '제목', '내용'],
  note:    ['비고'],
  staff:   ['담당', '담당자']
};

// 데이터 탭으로 인정하는 최소 조건
const REQUIRED_FIELDS = ['section', 'title', 'staff'];

function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify(collectRows_()))
    .setMimeType(ContentService.MimeType.JSON);
}

function collectRows_() {
  const out = [];
  SpreadsheetApp.openById(SPREADSHEET_ID).getSheets().forEach(function (sheet) {
    const name = sheet.getName().trim();
    const issue = Number((name.match(/\d+/) || [])[0]);
    if (!issue) return;                       // 호수 탭이 아니면 건너뜀

    const values = sheet.getDataRange().getDisplayValues();
    const headerRow = findHeaderRow_(values);
    if (headerRow < 0) return;

    const map = buildColumnMap_(values[headerRow]);
    const meta = readIssueMeta_(values, headerRow);

    for (let i = headerRow + 1; i < values.length; i++) {
      const row = values[i];
      const title   = cell_(row, map.title);
      const section = cell_(row, map.section);
      const staff   = cell_(row, map.staff);
      if (!title && !section && !staff) continue;      // 빈 줄
      if (!staff) continue;                            // 담당 없는 줄은 집계에서 제외
      // 헤더가 두 줄인 탭이 있어 머리글이 데이터로 새어 나옵니다
      if (staff === '담당' || title === '콘텐츠' || section === '코너명' || section === '카테고리') continue;

      const page = cell_(row, map.page);
      out.push({
        issue: issue,
        year: meta.year,          // 탭에 연도가 없으면 null
        month: meta.month,        // 발행 월
        day: meta.day,
        issueLabel: meta.label,   // 예: "제786호(12월 1일 자)"
        page: page,
        // 지면이 '연기'이거나 비어 있으면 아직 지면이 안 잡힌 건
        status: /연기/.test(page) ? '연기' : (page ? '게재' : '미정'),
        section: section,
        title: title,
        note: cell_(row, map.note),
        staff: staff
      });
    }
  });
  return out;
}

function cell_(row, idx) {
  return idx === undefined || idx >= row.length ? '' : String(row[idx] || '').trim();
}

/** 헤더 위쪽 줄에서 "제786호(12월 1일 자)" 같은 제목을 찾아 호수·발행일을 꺼냅니다. */
function readIssueMeta_(values, headerRow) {
  const meta = { label: '', year: null, month: null, day: null };
  for (let i = 0; i < headerRow; i++) {
    const line = values[i].join(' ');
    if (!/제\s*\d+\s*호/.test(line)) continue;
    meta.label = line.replace(/\s+/g, ' ').trim();
    const y = line.match(/(20\d{2})\s*년/);
    if (y) meta.year = Number(y[1]);
    const md = line.match(/(\d{1,2})\s*월\s*(\d{1,2})\s*일/);
    if (md) { meta.month = Number(md[1]); meta.day = Number(md[2]); }
    break;
  }
  return meta;
}

function findHeaderRow_(values) {
  const limit = Math.min(values.length, 12);
  for (let i = 0; i < limit; i++) {
    const map = buildColumnMap_(values[i]);
    if (REQUIRED_FIELDS.every(function (f) { return map[f] !== undefined; })) return i;
  }
  return -1;
}

function buildColumnMap_(headerCells) {
  const map = {};
  headerCells.forEach(function (c, idx) {
    const name = String(c || '').replace(/\s+/g, '');
    if (!name) return;
    Object.keys(FIELD_ALIASES).forEach(function (field) {
      if (map[field] !== undefined) return;
      if (FIELD_ALIASES[field].indexOf(name) >= 0) map[field] = idx;
    });
  });
  return map;
}

/**
 * 배포 전 점검용. 편집기에서 실행하고 로그를 보세요.
 * 탭이 몇 개인지, 연도가 탭에 적혀 있는지 여기서 확인합니다.
 */
function testCollect() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const names = ss.getSheets().map(function (s) { return s.getName(); });
  Logger.log('탭 %s개: %s', names.length, names.join(', '));

  const rows = collectRows_();
  const byIssue = {}, byStaff = {}, noYear = [];
  rows.forEach(function (r) {
    byIssue[r.issue] = (byIssue[r.issue] || 0) + 1;
    byStaff[r.staff] = (byStaff[r.staff] || 0) + 1;
    if (!r.year && noYear.indexOf(r.issue) < 0) noYear.push(r.issue);
  });
  Logger.log('총 %s건', rows.length);
  Logger.log('호별 건수: %s', JSON.stringify(byIssue));
  Logger.log('담당자별 건수: %s', JSON.stringify(byStaff));
  Logger.log('연도가 안 적힌 호: %s', JSON.stringify(noYear));
  Logger.log('첫 행: %s', JSON.stringify(rows[0]));
}
