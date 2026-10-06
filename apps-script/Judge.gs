/**
 * 과학송 경연대회 심사 서버 (Google Apps Script)
 *
 * 채점 페이지(judge.html)가 보내는 점수를 구글 시트에 쌓고 학년별 순위를 냅니다.
 * 엔트리 폼의 Code.gs 와는 별개의 프로젝트로 만드세요.
 * (같은 프로젝트에 넣으면 doGet/doPost 가 겹쳐 엔트리 폼이 멈춥니다.)
 *
 * 시트 구성
 *   설정      : 대회명, 심사 코드, 관리자 코드
 *   순서표    : 팀ID | 학년 | 순서 | 팀원 | 곡 제목   ← 웹앱의 "팀 편집"에서 저장하면 채워짐
 *   심사기록  : 채점관 한 명이 한 팀을 채점할 때마다 한 줄 (다시 제출하면 그 줄을 고침)
 *   심사 집계 N학년 : 자동으로 다시 쓰이는 순위표 (직접 고치지 않음)
 */

// 다른 스프레드시트에 기록하려면 그 시트의 ID 를 넣으세요. 비워 두면 이 스크립트가 붙어 있는 시트를 씁니다.
var SPREADSHEET_ID = '';

var SHEETS = { config: '설정', roster: '순서표', records: '심사기록', summary: '심사 집계 ' };
var ROSTER_HEAD = ['팀ID', '학년', '순서', '팀원', '곡 제목'];

/**
 * 과학송 경연대회 채점 기준(루브릭). 채점 화면, 인쇄용 표, 서버 점수 검증이 모두 이 표를 씁니다.
 * bands: 수준별 점수 구간 [최저, 최고] — 매우 우수, 우수, 보통, 노력 필요 순.
 * pick:  수준을 눌렀을 때 먼저 들어가는 점수(채점관이 구간 안에서 조정).
 */
var RUBRIC = {
  levels: ['매우 우수', '우수', '보통', '노력 필요'],
  criteria: [
    {
      id: 'science', name: '과학 개념 정확성', short: '과학', max: 30,
      ask: '가사에 담긴 과학 개념과 용어가 정확한가',
      bands: [[27, 30], [21, 26], [15, 20], [0, 14]], pick: [28, 23, 17, 10],
      desc: [
        '핵심 개념과 용어가 모두 정확하고, 개념 사이의 관계(원인과 결과, 조건)까지 가사에 담겨 있다. 가사만 들어도 단원의 핵심을 설명할 수 있다.',
        '핵심 개념은 정확하지만 용어 사용이 일부 어색하거나, 개념을 하나씩 따로 설명하는 데 그친다.',
        '개념을 나열하는 수준이거나 사소한 오류가 한두 곳 있다. 고른 단원과의 관련은 분명하다.',
        '뚜렷한 오개념이 있거나, 과학 내용이 거의 드러나지 않는다.'
      ],
      student: {
        goal: '가사만 들어도 배운 내용을 설명할 수 있게',
        checks: ['핵심 용어를 교과서와 같은 뜻으로 썼다', '개념을 늘어놓지 않고 "왜", "어떻게"를 담았다', '틀린 내용이 없는지 교과서로 다시 확인했다']
      }
    },
    {
      id: 'creative', name: '창의성과 독창성', short: '창의', max: 25,
      ask: '개념을 자기만의 방식으로 풀어냈는가',
      bands: [[23, 25], [18, 22], [13, 17], [0, 12]], pick: [24, 20, 15, 8],
      desc: [
        '비유, 이야기, 말놀이 등 자기만의 발상으로 개념을 풀어내 기억에 남는다. 원곡의 특징을 주제에 맞게 재치 있게 살렸다.',
        '눈에 띄는 표현이나 구성이 몇 군데 있으나, 곡 전체로 이어지지는 않는다.',
        '교과서 문장을 멜로디에 맞춰 옮긴 부분이 대부분이다.',
        '원곡 가사를 거의 그대로 두었거나, 다른 작품과 구별되는 점을 찾기 어렵다.'
      ],
      student: {
        goal: '우리 팀만의 비유와 이야기로',
        checks: ['교과서 문장을 그대로 옮기지 않고 우리 말로 바꿨다', '비유, 이야기, 말놀이 가운데 하나 이상을 넣었다', '원곡의 재미있는 부분을 주제에 맞게 살렸다']
      }
    },
    {
      id: 'music', name: '음악 완성도', short: '음악', max: 20,
      ask: '가사와 멜로디가 잘 맞고 노래가 안정적인가',
      bands: [[18, 20], [14, 17], [10, 13], [0, 9]], pick: [19, 15, 11, 6],
      desc: [
        '가사의 글자 수와 강세가 멜로디에 자연스럽게 맞고, 음정과 박자가 안정적이다. 가사가 또렷하게 들린다.',
        '전체적으로 안정적이나, 가사가 멜로디에 어색하게 얹힌 곳이나 음정·박자가 흔들리는 곳이 일부 있다.',
        '음정·박자가 자주 흔들리거나, 가사가 잘 들리지 않는 구간이 여러 곳 있다.',
        '노래가 끝까지 이어지지 않거나, 가사를 알아듣기 어렵다.'
      ],
      student: {
        goal: '가사가 또렷이 들리게',
        checks: ['가사 글자 수가 멜로디에 맞는다', '음정과 박자를 맞춰 끝까지 불렀다', '반주에 목소리가 묻히지 않는다']
      }
    },
    {
      id: 'video', name: '영상 완성도', short: '영상', max: 15,
      ask: '영상이 가사의 과학 내용을 잘 보여 주는가',
      bands: [[14, 15], [11, 13], [8, 10], [0, 7]], pick: [14, 12, 9, 5],
      desc: [
        '장면이 가사의 과학 내용을 그림, 실험, 연기 등으로 보여 주어 이해를 돕는다. 자막, 화질, 음량, 편집이 고르다.',
        '영상이 가사와 대체로 어울리나, 과학 내용을 보여 주는 장면이 일부에 그치거나 자막·음량 등 편집에 아쉬운 곳이 있다.',
        '가사와 관계가 약한 장면이 많거나, 자막이 없고 화면·음량이 고르지 않아 보기에 불편하다.',
        '영상이 완성되지 않았거나, 화면과 소리를 알아보기 어렵다.'
      ],
      student: {
        goal: '화면이 가사를 설명해 주게',
        checks: ['장면이 가사의 과학 내용을 보여 준다', '가사 자막을 넣었다', '소리 크기와 화면이 처음부터 끝까지 고르다']
      }
    },
    {
      id: 'delivery', name: '전달력과 협동', short: '전달', max: 10,
      ask: '두 사람이 함께 만들어 자신 있게 전달하는가',
      bands: [[9, 10], [7, 8], [5, 6], [0, 4]], pick: [9, 7, 5, 3],
      desc: [
        '두 사람의 역할이 고르게 드러나고, 자신 있는 목소리와 표현으로 내용을 끝까지 전달한다.',
        '두 사람이 모두 참여했으나 한 사람의 비중이 크거나, 전달이 약해지는 구간이 있다.',
        '한 사람이 대부분을 맡았거나, 소극적이어서 내용이 잘 전달되지 않는다.',
        '한 사람의 참여를 확인하기 어렵다.'
      ],
      student: {
        goal: '둘이 함께, 자신 있게',
        checks: ['두 사람이 모두 노래나 영상에 나온다', '역할을 고르게 나눴다', '자신 있는 목소리로 끝까지 불렀다']
      }
    }
  ],
  notes: [
    '수준을 먼저 고른 뒤, 그 구간 안에서 점수를 정합니다. 점수는 정수로 줍니다.',
    '뚜렷한 오개념이 하나라도 있으면 과학 개념 정확성은 "보통" 이하로 채점합니다.',
    '원곡과 멜로디 출처를 밝히지 않은 작품은 창의성과 독창성을 "보통" 이하로 채점합니다.',
    '최종 점수는 채점한 심사위원들의 총점 평균입니다. 동점이면 과학 개념 정확성 평균, 그다음 창의성과 독창성 평균이 높은 팀이 앞서고, 그래도 같으면 공동 순위입니다.'
  ]
};

/**
 * 심사 집계. 채점 페이지(브라우저)와 Apps Script 서버가 같은 코드를 씁니다.
 * roster:  [{id, grade, order, members, title}]
 * records: [{team(팀 ID), judge, scores:[정수 5개], comment, at}]
 * 팀 점수 = 그 팀을 채점한 사람들의 총점 평균.
 * 동점 = 과학 개념 정확성(첫 항목) 평균 -> 창의성과 독창성(둘째 항목) 평균 -> 공동 순위.
 */
function aggregate(roster, records, criteria) {
  function norm(s) { return String(s == null ? '' : s).replace(/\s+/g, ' ').trim(); }
  // 팀은 고유 ID로 찾고, ID가 없는 옛 자료만 학년-순서로 찾는다
  function has(v) { return v != null && String(v) !== ''; }
  function teamKey(t) { return has(t.id) ? String(t.id) : Number(t.grade) + '-' + Number(t.order); }
  function recKey(r) { return has(r.team) ? String(r.team) : Number(r.grade) + '-' + Number(r.order); }
  function sum(a) { var s = 0; for (var i = 0; i < a.length; i++) s += a[i]; return s; }
  // 분수 a.s/a.n 과 b.s/b.n 을 정수 곱으로 비교 (반올림 오차 없음)
  function frac(as, an, bs, bn) { return bs * an - as * bn; }

  var teams = {}, order = [];
  roster.forEach(function (t) {
    var k = teamKey(t);
    if (teams[k]) return;
    teams[k] = { key: k, id: k, grade: Number(t.grade), order: Number(t.order), members: norm(t.members),
      title: norm(t.title), judgeMap: {} };
    order.push(k);
  });

  var orphans = 0, judgeStat = {};
  records.forEach(function (r, idx) {
    var t = teams[recKey(r)], name = norm(r.judge);
    if (!t || !name) { orphans++; return; }
    var scores = criteria.map(function (c, i) { return Number(r.scores[i]) || 0; });
    var cur = { judge: name, scores: scores, total: sum(scores), comment: norm(r.comment),
      at: String(r.at || ''), idx: idx };
    var prev = t.judgeMap[name];
    if (!prev || cur.at > prev.at || (cur.at === prev.at && cur.idx > prev.idx)) t.judgeMap[name] = cur;
  });

  var grades = {};
  order.forEach(function (k) {
    var t = teams[k];
    var js = Object.keys(t.judgeMap).sort().map(function (n) { return t.judgeMap[n]; });
    js.forEach(function (j) {
      delete j.idx;
      var st = judgeStat[j.judge] || (judgeStat[j.judge] = { judge: j.judge, counts: {}, total: 0 });
      st.counts[t.grade] = (st.counts[t.grade] || 0) + 1; st.total++;
    });
    delete t.judgeMap;
    t.judges = js; t.n = js.length;
    t.critSum = criteria.map(function (c, i) { return sum(js.map(function (j) { return j.scores[i]; })); });
    t.totalSum = sum(t.critSum);
    t.critAvg = t.n ? t.critSum.map(function (s) { return s / t.n; }) : criteria.map(function () { return null; });
    t.totalAvg = t.n ? t.totalSum / t.n : null;
    t.rank = null; t.award = ''; t.fewJudges = false;
    var g = grades[t.grade] || (grades[t.grade] = { grade: t.grade, teams: [], maxJudges: 0, hasTie: false, scored: 0 });
    g.teams.push(t);
  });

  function cmp(a, b) {
    return frac(a.totalSum, a.n, b.totalSum, b.n) ||
      frac(a.critSum[0], a.n, b.critSum[0], b.n) ||
      frac(a.critSum[1], a.n, b.critSum[1], b.n);
  }

  Object.keys(grades).forEach(function (gk) {
    var g = grades[gk];
    var ranked = g.teams.filter(function (t) { return t.n > 0; });
    var rest = g.teams.filter(function (t) { return t.n === 0; });
    ranked.sort(function (a, b) { return cmp(a, b) || a.order - b.order; });
    rest.sort(function (a, b) { return a.order - b.order; });
    g.scored = ranked.length;
    ranked.forEach(function (t, i) {
      t.rank = (i > 0 && cmp(ranked[i - 1], t) === 0) ? ranked[i - 1].rank : i + 1;
      if (t.n > g.maxJudges) g.maxJudges = t.n;
    });
    var count = {};
    ranked.forEach(function (t) { count[t.rank] = (count[t.rank] || 0) + 1; });
    ranked.forEach(function (t) {
      t.fewJudges = t.n < g.maxJudges;
      t.tied = count[t.rank] > 1;
      if (t.rank <= 2) t.award = (t.tied ? '공동 ' : '') + t.rank + '등';
      if (t.tied && t.rank <= 2) g.hasTie = true;
    });
    g.teams = ranked.concat(rest);
  });

  var judges = Object.keys(judgeStat).sort().map(function (n) { return judgeStat[n]; });
  return { grades: grades, judges: judges, orphans: orphans };
}

var CRITERIA = RUBRIC.criteria;
var RECORD_HEAD = ['제출일시', '팀ID', '학년', '순서', '팀원', '곡 제목', '채점관']
  .concat(CRITERIA.map(function (c) { return c.name; })).concat(['총점', '한줄평']);

/* ---------- 처음 한 번 실행 ---------- */

function setup() {
  var ss = ss_();
  var cfg = ss.getSheetByName(SHEETS.config);
  if (!cfg) {
    cfg = ss.insertSheet(SHEETS.config);
    cfg.getRange(1, 1, 4, 3).setValues([
      ['항목', '값', '설명'],
      ['대회명', '2026 과학송 경연대회', '채점 화면과 발표 화면 위에 나오는 이름'],
      ['심사 코드', randomDigits_(6), '채점관에게 알려 주는 코드'],
      ['관리자 코드', randomDigits_(8), '결과 화면과 팀 편집을 여는 코드 (채점관에게 알리지 않음)']
    ]);
    cfg.getRange('B2:B4').setNumberFormat('@');
    styleHead_(cfg, 3);
    cfg.setColumnWidth(1, 110); cfg.setColumnWidth(2, 220); cfg.setColumnWidth(3, 420);
  }
  rosterSheet_();
  recordSheet_();
  writeSummary_();
  var c = readConfig_();
  Logger.log('준비 완료. 심사 코드: ' + c.judgeCode + ' / 관리자 코드: ' + c.adminCode + ' (설정 시트에서 바꿀 수 있습니다)');
}

function rebuildSummary() { writeSummary_(); }

function onOpen() {
  try {
    SpreadsheetApp.getUi().createMenu('과학송 심사')
      .addItem('처음 설정 (시트 만들기)', 'setup')
      .addItem('집계 다시 만들기', 'rebuildSummary')
      .addToUi();
  } catch (err) { /* 독립 스크립트에서는 메뉴가 없습니다 */ }
}

/* ---------- 웹앱 입구 ---------- */

function doPost(e) {
  var req;
  try { req = JSON.parse(e.postData.contents); }
  catch (err) { return json_({ ok: false, code: 'BAD_REQUEST', error: '요청을 읽지 못했습니다.' }); }
  return json_(handle_(req));
}

function doGet(e) {
  var p = (e && e.parameter) || {};
  if (!p.action) return ContentService.createTextOutput('과학송 심사 서버가 작동 중입니다.');
  var req;
  try { req = p.payload ? JSON.parse(p.payload) : {}; }
  catch (err) { req = {}; }
  req.action = p.action;
  var out = JSON.stringify(handle_(req));
  if (p.callback && /^[A-Za-z_$][\w$]{0,60}$/.test(p.callback)) {
    return ContentService.createTextOutput(p.callback + '(' + out + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(out).setMimeType(ContentService.MimeType.JSON);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- 요청 처리 ---------- */

function handle_(req) {
  try {
    req = req || {};
    var cfg = readConfig_();
    switch (String(req.action || '')) {
      case 'ping':
        return { ok: true, title: cfg.title };
      case 'roster':
        requireCode_(req.code, cfg.judgeCode, '심사 코드');
        var judge = cleanJudge_(req.judge);
        return { ok: true, title: cfg.title, roster: readRoster_(),
          mine: readRecords_().filter(function (r) { return r.judge === judge; }) };
      case 'submit':
        requireCode_(req.code, cfg.judgeCode, '심사 코드');
        return submit_(req);
      case 'results':
        requireCode_(req.code, cfg.adminCode, '관리자 코드');
        return { ok: true, title: cfg.title, judgeCode: cfg.judgeCode, roster: readRoster_(),
          records: readRecords_(), at: new Date().toISOString() };
      case 'saveRoster':
        requireCode_(req.code, cfg.adminCode, '관리자 코드');
        return saveRoster_(req, cfg);
      default:
        fail_('BAD_REQUEST', '알 수 없는 요청입니다.');
    }
  } catch (err) {
    return { ok: false, code: err.code || 'ERROR', error: err.userMessage || ('서버 오류: ' + err.message) };
  }
}

function submit_(req) {
  var judge = cleanJudge_(req.judge);
  var scores = req.scores;
  if (!Array.isArray(scores) || scores.length !== CRITERIA.length) fail_('BAD_SCORE', '다섯 항목을 모두 채점해 주세요.');
  scores = scores.map(function (v, i) {
    var n = Number(v), c = CRITERIA[i];
    if (v === null || v === '' || !isFinite(n) || Math.floor(n) !== n || n < 0 || n > c.max) {
      fail_('BAD_SCORE', c.name + ' 점수는 0~' + c.max + ' 사이의 정수여야 합니다.');
    }
    return n;
  });
  var comment = String(req.comment == null ? '' : req.comment).replace(/\s+/g, ' ').trim().slice(0, 300);

  // 팀은 고유 ID로 찾는다. ID 없이 온 요청(옛 화면)은 학년과 순서로 찾는다.
  var id = cleanText_(req.team), team = null;
  readRoster_().forEach(function (t) {
    if (id ? t.id === id : (t.grade === Number(req.grade) && t.order === Number(req.order))) team = team || t;
  });
  if (!team) fail_('NO_TEAM', '순서표에 없는 팀입니다. 화면을 새로 고친 뒤 다시 해 주세요.');

  var total = 0; scores.forEach(function (s) { total += s; });
  var now = new Date();
  var row = [now, team.id, team.grade, team.order, team.members, team.title, judge].concat(scores).concat([total, comment]);

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) fail_('BUSY', '다른 요청을 처리하는 중입니다. 잠시 뒤 다시 눌러 주세요.');
  try {
    var sh = recordSheet_();
    var last = sh.getLastRow(), hit = 0;
    if (last > 1) {
      var keys = sh.getRange(2, 1, last - 1, 7).getValues();
      for (var i = 0; i < keys.length; i++) {
        if (cleanText_(keys[i][1]) === team.id && cleanText_(keys[i][6]) === judge) { hit = i + 2; break; }
      }
    }
    if (hit) sh.getRange(hit, 1, 1, row.length).setValues([row]);
    else sh.getRange(last + 1, 1, 1, row.length).setValues([row]);
    SpreadsheetApp.flush();
    try { writeSummary_(); } catch (err) { /* 집계 시트 갱신 실패가 제출을 막지 않게 */ }
  } finally {
    lock.releaseLock();
  }
  return { ok: true, updated: !!hit,
    record: { team: team.id, grade: team.grade, order: team.order, judge: judge, scores: scores, comment: comment, at: now.toISOString() } };
}

/** 웹앱의 "팀 편집"에서 보낸 팀 목록으로 순서표를 통째로 다시 씁니다. */
function saveRoster_(req, cfg) {
  var list = req.roster;
  if (!Array.isArray(list) || list.length > 300) fail_('BAD_ROSTER', '팀 목록을 읽지 못했습니다.');
  var seen = {}, count = {}, rows = [];
  list.forEach(function (t) {
    t = t || {};
    var grade = Number(t.grade);
    if (Math.floor(grade) !== grade || grade < 1 || grade > 9) fail_('BAD_ROSTER', '학년은 1~9 사이의 숫자여야 합니다.');
    var members = cleanText_(t.members).slice(0, 100), title = cleanText_(t.title).slice(0, 100);
    if (!members && !title) return;                       // 빈 줄은 버린다
    var id = cleanText_(t.id);
    if (!/^[A-Za-z0-9_-]{1,24}$/.test(id) || seen[id]) id = newTeamId_(seen);
    seen[id] = true;
    count[grade] = (count[grade] || 0) + 1;               // 보낸 차례가 곧 발표 순서
    rows.push([id, grade, count[grade], members, title]);
  });
  rows.sort(function (a, b) { return a[1] - b[1] || a[2] - b[2]; });

  var title = req.title == null ? null : cleanText_(req.title).slice(0, 60);
  var judgeCode = req.judgeCode == null ? null : cleanText_(req.judgeCode);
  if (title !== null && !title) fail_('BAD_CONFIG', '대회 이름을 입력해 주세요.');
  if (judgeCode !== null) {
    if (!/^\S{4,20}$/.test(judgeCode)) fail_('BAD_CONFIG', '심사 코드는 띄어쓰기 없이 4~20자로 정해 주세요.');
    if (judgeCode === cfg.adminCode) fail_('BAD_CONFIG', '심사 코드는 관리자 코드와 달라야 합니다.');
  }

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) fail_('BUSY', '다른 요청을 처리하는 중입니다. 잠시 뒤 다시 눌러 주세요.');
  try {
    var sh = rosterSheet_();
    sh.clear();
    sh.getRange(1, 1, 1, ROSTER_HEAD.length).setValues([ROSTER_HEAD]);
    styleHead_(sh, ROSTER_HEAD.length);
    if (rows.length) sh.getRange(2, 1, rows.length, ROSTER_HEAD.length).setValues(rows);
    if (title !== null) setConfig_('대회명', title);
    if (judgeCode !== null) setConfig_('심사 코드', judgeCode);
    SpreadsheetApp.flush();
    try { writeSummary_(); } catch (err) { /* 집계 시트 갱신 실패가 저장을 막지 않게 */ }
  } finally {
    lock.releaseLock();
  }
  var now = readConfig_();
  return { ok: true, title: now.title, judgeCode: now.judgeCode, roster: readRoster_() };
}

/* ---------- 시트 읽기 ---------- */

function ss_() {
  return SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
}

function readConfig_() {
  var sh = ss_().getSheetByName(SHEETS.config);
  if (!sh) fail_('NOT_SETUP', '서버 준비가 끝나지 않았습니다. Apps Script 에서 setup 을 한 번 실행해 주세요.');
  var map = {};
  sh.getDataRange().getValues().forEach(function (r) { map[cleanText_(r[0])] = cleanText_(r[1]); });
  return { title: map['대회명'] || '과학송 경연대회', judgeCode: map['심사 코드'] || '', adminCode: map['관리자 코드'] || '' };
}

function setConfig_(key, value) {
  var sh = ss_().getSheetByName(SHEETS.config);
  var rows = sh.getDataRange().getValues();
  for (var i = 0; i < rows.length; i++) {
    if (cleanText_(rows[i][0]) === key) { sh.getRange(i + 1, 2, 1, 1).setValues([[value]]); return; }
  }
  sh.getRange(rows.length + 1, 1, 1, 2).setValues([[key, value]]);
}

function rosterSheet_() {
  var ss = ss_(), sh = ss.getSheetByName(SHEETS.roster);
  if (!sh) {
    sh = ss.insertSheet(SHEETS.roster);
    sh.getRange(1, 1, 1, ROSTER_HEAD.length).setValues([ROSTER_HEAD]);
    styleHead_(sh, ROSTER_HEAD.length);
    sh.setColumnWidth(4, 200); sh.setColumnWidth(5, 320);
  }
  return sh;
}

function readRoster_() {
  var sh = ss_().getSheetByName(SHEETS.roster);
  if (!sh || sh.getLastRow() < 2) return [];
  var out = [], seen = {};
  sh.getRange(2, 1, sh.getLastRow() - 1, ROSTER_HEAD.length).getValues().forEach(function (r) {
    var g = String(r[1]).match(/\d+/), o = String(r[2]).match(/\d+/);
    var members = cleanText_(r[3]), title = cleanText_(r[4]);
    if (!g || !o || (!members && !title)) return;
    var grade = Number(g[0]), order = Number(o[0]);
    var id = cleanText_(r[0]) || (grade + '-' + order);    // 시트에 직접 적어 ID가 빈 줄
    while (seen[id]) id += 'x';
    seen[id] = true;
    out.push({ id: id, grade: grade, order: order, members: members, title: title });
  });
  out.sort(function (a, b) { return a.grade - b.grade || a.order - b.order; });
  return out;
}

function recordSheet_() {
  var ss = ss_(), sh = ss.getSheetByName(SHEETS.records);
  if (!sh) {
    sh = ss.insertSheet(SHEETS.records);
    sh.getRange(1, 1, 1, RECORD_HEAD.length).setValues([RECORD_HEAD]);
    styleHead_(sh, RECORD_HEAD.length);
    sh.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm:ss');
  }
  return sh;
}

function readRecords_() {
  var sh = ss_().getSheetByName(SHEETS.records);
  if (!sh || sh.getLastRow() < 2) return [];
  var n = CRITERIA.length;
  return sh.getRange(2, 1, sh.getLastRow() - 1, RECORD_HEAD.length).getValues().map(function (r) {
    var at = r[0] instanceof Date ? r[0].toISOString() : String(r[0]);
    return { at: at, team: cleanText_(r[1]), grade: Number(r[2]), order: Number(r[3]), judge: cleanText_(r[6]),
      scores: r.slice(7, 7 + n).map(Number), comment: cleanText_(r[8 + n]) };
  }).filter(function (r) { return r.judge && r.team; });
}

/* ---------- 집계 시트 쓰기 ---------- */

function writeSummary_() {
  var ss = ss_(), roster = readRoster_();
  var result = aggregate(roster, readRecords_(), CRITERIA);
  var head = ['순위', '수상', '순서', '팀원', '곡 제목']
    .concat(CRITERIA.map(function (c) { return c.short + ' (' + c.max + ')'; }))
    .concat(['총점 평균', '채점 인원', '확인']);
  var live = {};
  Object.keys(result.grades).forEach(function (gk) {
    var g = result.grades[gk];
    var name = SHEETS.summary + g.grade + '학년';
    live[name] = true;
    var sh = ss.getSheetByName(name) || ss.insertSheet(name);
    sh.clear();
    var rows = g.teams.map(function (t) {
      var note = t.n === 0 ? '채점 전' : (t.tied && t.rank <= 2 ? '공동 순위' : (t.fewJudges ? '채점 인원이 적음' : ''));
      return [t.rank === null ? '' : t.rank, t.award, t.order, t.members, t.title]
        .concat(t.critAvg.map(function (v) { return v === null ? '' : round2_(v); }))
        .concat([t.totalAvg === null ? '' : round2_(t.totalAvg), t.n, note]);
    });
    sh.getRange(1, 1, 1, head.length).setValues([head]);
    styleHead_(sh, head.length);
    if (rows.length) {
      sh.getRange(2, 1, rows.length, head.length).setValues(rows);
      g.teams.forEach(function (t, i) {
        if (t.rank === 1) sh.getRange(i + 2, 1, 1, head.length).setBackground('#FCE9A9').setFontWeight('bold');
        else if (t.rank === 2) sh.getRange(i + 2, 1, 1, head.length).setBackground('#E3ECF7');
      });
    }
    sh.setColumnWidth(4, 180); sh.setColumnWidth(5, 280);
  });
  // 팀이 모두 빠진 학년의 옛 집계는 비운다
  ss.getSheets().forEach(function (sh) {
    var name = sh.getName();
    if (name.indexOf(SHEETS.summary) === 0 && !live[name]) sh.clear();
  });
}

/* ---------- 작은 도구 ---------- */

function fail_(code, message) {
  var err = new Error(message);
  err.code = code; err.userMessage = message;
  throw err;
}

function requireCode_(given, expected, label) {
  if (!expected) fail_('NOT_SETUP', '설정 시트에 ' + label + '가 비어 있습니다.');
  if (cleanText_(given) !== expected) fail_('BAD_CODE', label + '가 맞지 않습니다.');
}

function cleanText_(v) { return String(v == null ? '' : v).replace(/\s+/g, ' ').trim(); }

function cleanJudge_(v) {
  var name = cleanText_(v).slice(0, 20);
  if (!name) fail_('NO_NAME', '채점관 이름을 입력해 주세요.');
  return name;
}

function round2_(v) { return Math.round(v * 100) / 100; }

function randomDigits_(n) {
  var s = '';
  for (var i = 0; i < n; i++) s += Math.floor(Math.random() * 10);
  return s;
}

function newTeamId_(seen) {
  var abc = 'abcdefghijkmnpqrstuvwxyz23456789', id;
  do {
    id = 't';
    for (var i = 0; i < 6; i++) id += abc.charAt(Math.floor(Math.random() * abc.length));
  } while (seen[id]);
  return id;
}

function styleHead_(sheet, cols) {
  sheet.getRange(1, 1, 1, cols).setFontWeight('bold').setBackground('#1F2F57').setFontColor('#FFFFFF');
  sheet.setFrozenRows(1);
}
