/**
 * 병원보 편집계획안 → 대시보드 데이터 API (Google Apps Script)
 *
 * 원본: "★병원보 편집계획안(NEW)" (asanpr.members@gmail.com 소유)
 * 탭 하나가 한 호(號)이고, 탭 이름이 호수입니다(예: "786").
 *
 * ── 속도 ──
 * 탭이 122개라 전부 읽으면 40초가 걸리고, 웹 앱은 60초에서 끊깁니다.
 * 바뀌는 건 최근 몇 호뿐이고 지나간 호는 그대로이므로,
 *   · 최근 LIVE_ISSUES 호는 매번 새로 읽고
 *   · 그보다 오래된 호는 Drive 에 쌓아둔 캐시 파일에서 가져옵니다.
 * 캐시는 요청이 올 때 빠진 호만 채웁니다. 시간 트리거는 쓰지 않습니다.
 *
 * ── 지나간 호를 고쳤을 때 ──
 * 캐시에 이미 들어간 호를 수정하면 대시보드에 반영되지 않습니다.
 * 그때는 편집기에서 rebuildCache() 를 한 번 실행하거나, /exec?rebuild=1 을 열면 됩니다.
 *
 * 배포: 배포 > 새 배포 > 웹 앱 (실행: 나 / 액세스: 링크가 있는 모든 사용자)
 * 이후 수정은 배포 > 배포 관리 > 연필 > 새 버전 으로 해야 URL 이 유지됩니다.
 */

const SPREADSHEET_ID = '1A_pBtxztIUlwBl2an8-hWiVHE38lt3_0nS8HsKJ2doQ';
const CACHE_FILENAME = 'amc-hospital-paper-cache.json';
const LIVE_ISSUES = 4;          // 최근 몇 호를 매번 새로 읽을지

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

function doGet(e) {
  const rebuild = !!(e && e.parameter && e.parameter.rebuild);
  return ContentService
    .createTextOutput(JSON.stringify(collectRows_(rebuild)))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ────────────────────────────── 수집 ────────────────────────────── */

function collectRows_(forceRebuild) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // 호수 탭만, 최신 호가 앞으로
  const entries = [];
  ss.getSheets().forEach(function (sheet) {
    const m = sheet.getName().trim().match(/\d+/);
    if (m) entries.push({ issue: Number(m[0]), sheet: sheet });
  });
  entries.sort(function (a, b) { return b.issue - a.issue; });

  const live = entries.slice(0, LIVE_ISSUES);
  const past = entries.slice(LIVE_ISSUES);

  const cache = forceRebuild ? { rows: [], issues: [] } : readCache_();
  const cached = {};
  cache.issues.forEach(function (i) { cached[i] = true; });

  // 캐시에 없는 지난 호만 새로 읽어 채웁니다 (평소엔 0~1개)
  const missing = past.filter(function (en) { return !cached[en.issue]; });
  if (missing.length) {
    missing.forEach(function (en) {
      cache.rows = cache.rows.concat(readSheet_(en.sheet, en.issue));
      cache.issues.push(en.issue);
    });
    writeCache_(cache);
  }

  // 최근 호는 캐시를 믿지 않고 항상 새로 읽습니다
  const isLive = {};
  live.forEach(function (en) { isLive[en.issue] = true; });
  const out = cache.rows.filter(function (r) { return !isLive[r.issue]; });
  live.forEach(function (en) {
    readSheet_(en.sheet, en.issue, true).forEach(function (r) { out.push(r); });
  });
  return out;
}

function readSheet_(sheet, issue, withLegend) {
  const out = [];
  const range = sheet.getDataRange();
  const values = range.getDisplayValues();
  const headerRow = findHeaderRow_(values);
  if (headerRow < 0) return out;

  const map = buildColumnMap_(values[headerRow]);
  const meta = readIssueMeta_(values, headerRow);
  // 지면 칸이 세로 병합된 곳(18면에 두 꼭지 등)은 첫 줄에만 값이 들어옵니다.
  // 그대로 두면 아래 줄이 '지면 미정'으로 잘못 잡혀서, 병합 범위를 받아 채웁니다.
  const pageFill = mergedPageValues_(range, values, map.page);
  const bgs = range.getBackgrounds();                 // 셀 배경색 (시트의 색 표기)
  const zone = findZone_(values, map.section);        // 표지 ~ 이달의 후원자 구간
  const legend = legend_(values, bgs);                // 2행 담당자 이름 칸 색

  for (let i = headerRow + 1; i < values.length; i++) {
    const row = values[i];
    const title   = cell_(row, map.title);
    const section = cell_(row, map.section);
    const staff   = cell_(row, map.staff);
    if (!title && !section && !staff) continue;      // 빈 줄
    if (!staff) continue;                            // 담당 없는 줄은 집계에서 제외
    // 헤더가 두 줄인 탭이 있어 머리글이 데이터로 새어 나옵니다
    if (staff === '담당' || title === '콘텐츠' || section === '코너명' || section === '카테고리') continue;

    const page = (pageFill[i] !== undefined) ? pageFill[i] : cell_(row, map.page);
    out.push({
      issue: issue,
      gid: sheet.getSheetId(),   // 대시보드에서 그 호 탭으로 바로 열기 위해
      year: meta.year,          // 탭에 연도가 없으면 null (대시보드가 호수로 역산합니다)
      month: meta.month,
      day: meta.day,
      issueLabel: meta.label,
      page: page,
      status: /연기/.test(page) ? '연기' : (page ? '게재' : '미정'),
      section: section,
      title: title,
      note: cell_(row, map.note),
      staff: staff,
      zone: !!(zone && i >= zone.start && i <= zone.end),
      staffBg: colorOf_(bgs, i, map.staff),
      legend: withLegend ? legend : undefined
    });
  }
  return out;
}

/** 표지 ~ 이달의 후원자 구간의 행 범위 (없으면 null) */
function findZone_(values, sectionCol) {
  if (sectionCol === undefined) return null;
  let start = -1, end = -1;
  for (let i = 0; i < values.length; i++) {
    const v = cell_(values[i], sectionCol);
    if (start < 0 && v === '표지') start = i;
    if (start >= 0 && v === '이달의 후원자') { end = i; break; }
  }
  return (start >= 0 && end >= start) ? { start: start, end: end } : null;
}

/** 담당자 이름 → 배경색 (2행의 이름 칸) */
function legend_(values, bgs) {
  const out = {};
  const names = values[1] || [];
  names.forEach(function (v, j) {
    const n = String(v || '').trim();
    const bg = (bgs[1] || [])[j];
    if (n && n.length <= 4 && bg && !isWhite_(bg)) out[n] = bg;
  });
  return out;
}

/** 해당 칸의 배경색 (흰색이면 빈 값) */
function colorOf_(bgs, i, col) {
  if (col === undefined) return '';
  const bg = (bgs[i] || [])[col];
  return (bg && !isWhite_(bg)) ? bg : '';
}

function isWhite_(bg) {
  const c = String(bg).toLowerCase();
  return c === '#ffffff' || c === '#fff' || c === '';
}

/** 병합된 지면 칸을 아래 줄까지 펼칩니다. 행 번호(0부터) → 지면 값 */
function mergedPageValues_(range, values, pageCol) {
  const fill = {};
  if (pageCol === undefined) return fill;
  const col1 = pageCol + 1;                      // getMergedRanges 는 1부터 셉니다
  // getMergedRanges 는 Sheet 가 아니라 Range 의 메서드입니다
  range.getMergedRanges().forEach(function (rng) {
    if (rng.getColumn() > col1 || rng.getLastColumn() < col1) return;
    const first = rng.getRow() - 1, last = rng.getLastRow() - 1;
    const v = String((values[first] || [])[pageCol] || '').trim();
    if (!v) return;
    for (let r = first; r <= last; r++) fill[r] = v;
  });
  return fill;
}

/* ────────────────────────────── 캐시 ────────────────────────────── */

function cacheFile_() {
  const it = DriveApp.getFilesByName(CACHE_FILENAME);
  return it.hasNext() ? it.next() : null;
}

function readCache_() {
  try {
    const f = cacheFile_();
    if (!f) return { rows: [], issues: [] };
    const c = JSON.parse(f.getBlob().getDataAsString());
    if (!c || !Array.isArray(c.rows) || !Array.isArray(c.issues)) return { rows: [], issues: [] };
    return c;
  } catch (err) {
    return { rows: [], issues: [] };   // 깨졌으면 없는 셈 치고 다시 만듭니다
  }
}

function writeCache_(cache) {
  const payload = JSON.stringify({ t: new Date().toISOString(), rows: cache.rows, issues: cache.issues });
  const f = cacheFile_();
  if (f) f.setContent(payload);
  else DriveApp.createFile(CACHE_FILENAME, payload, MimeType.PLAIN_TEXT);
}

/** 지나간 호를 고쳤을 때 한 번 실행하세요. 40초쯤 걸립니다. */
function rebuildCache() {
  const rows = collectRows_(true);
  Logger.log('캐시 새로 만듦: %s꼭지', rows.length);
}

/* ────────────────────────────── 파싱 ────────────────────────────── */

function cell_(row, idx) {
  return idx === undefined || idx >= row.length ? '' : String(row[idx] || '').trim();
}

/** 헤더 위쪽에서 "제786호(12월 1일 자)" 같은 제목을 찾아 발행일을 꺼냅니다. */
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

/* ────────────────────────────── 점검 ────────────────────────────── */

/** 편집기에서 실행하고 로그를 보세요. 두 번째 실행부터 빨라지는지 확인용입니다. */
function testCollect() {
  const t0 = Date.now();
  const rows = collectRows_(false);
  const sec = ((Date.now() - t0) / 1000).toFixed(1);

  const cache = readCache_();
  const byStaff = {};
  rows.forEach(function (r) { byStaff[r.staff] = (byStaff[r.staff] || 0) + 1; });

  Logger.log('%s초 걸림 · 총 %s꼭지', sec, rows.length);
  Logger.log('캐시에 담긴 호: %s개 (최근 %s호는 매번 새로 읽음)', cache.issues.length, LIVE_ISSUES);
  Logger.log('담당자별: %s', JSON.stringify(byStaff));
  Logger.log('첫 행: %s', JSON.stringify(rows[0]));
}
