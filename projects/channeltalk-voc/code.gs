/*************************************************
 * 채널톡 통합 수집 대시보드 (환급 / 케어)
 * - 웹앱으로 배포하여 여러 명이 사용
 * - 채널 선택 → 기간 지정 → 수집 실행
 * - 진행상황 실시간 표시
 *************************************************/

var CHANNELS = {
  refund: {
    key: 'refund',
    label: '환급',
    slug: '<REFUND_SLUG>',
    deskBase: 'https://channel.works/<REFUND_SLUG>/user-chats/',
    propKey: 'CT_KEY_REFUND',
    propSecret: 'CT_SECRET_REFUND',
    satisfactionGroupId: '<SATISFACTION_GROUP_ID>'
  },
  care: {
    key: 'care',
    label: '케어',
    slug: '<CARE_SLUG>',
    deskBase: 'https://channel.works/<CARE_SLUG>/user-chats/',
    propKey: 'CT_KEY_CARE',
    propSecret: 'CT_SECRET_CARE'
  }
};

var CFG = {
  SHEET_META: '메타베이스',
  SHEET_PROGRESS_PREFIX: '진행상황_',
  SHEET_JOB_PREFIX: '_job_',

  CHANNEL_BASE: 'https://api.channel.io/open/v5',
  CHAT_STATES: ['closed', 'opened', 'snoozed', 'queued', 'initial', 'missed'],

  PAGE_LIMIT: 100,
  MESSAGE_PAGE_LIMIT: 50,
  API_RETRY: 6,
  API_SLEEP_MS: 150,
  MAX_TRANSCRIPT_CHAR: 50000,
  PROGRESS_EVERY: 10,
  MAX_RUN_MS: 4.5 * 60 * 1000,
  RESUME_TRIGGER_MS: 60000,

  META_HEADERS: {
    RPN_TIN: 'Rpn Tin',
    CONTACT_PHONE: 'User → Contact Phone',
    USER_NAME: 'User → Name',
    BSNO: 'Bman Infr - Bman Tin → Bsno',
    BKP_STATUS: 'Bman Infr - Bman Tin → Bkp Status',
    TNM_NM: 'Bman Infr - Bman Tin → Tnm Nm',
    CAREPRO_MGR: 'Carepro Mngr Infr → Name'
  },

  OUT_HEADERS: [
    'Rpntin',
    '대표자 이름',
    '사업자번호',
    '사업장명',
    '기장담당자',
    '현재상태',
    '채팅 시작 시간',
    '채널톡링크',
    '인입경로',
    '태그',
    '전사문',
    '만족도',
    '만족도 응답시각'
  ]
};

/**
 * 채널별 property key 헬퍼 (병렬 실행 지원)
 * 예: jobProp('refund', 'STATUS') => 'CT_JOB_refund_STATUS'
 */
function jobProp(channelKey, name) { return 'CT_JOB_' + channelKey + '_' + name; }
function skipProp(channelKey) { return 'CT_SKIPS_' + channelKey; }
function lockProp(channelKey) { return 'CT_LOCK_' + channelKey; }

var JOB_FIELDS = ['STATUS', 'PHASE', 'START_MS', 'END_MS', 'NEXT_IDX', 'TOTAL', 'FAILED', 'STARTED_AT', 'OUTPUT_SHEET'];

var SKIP_LABELS = {
  autoClosed: '자동종료·상태변경',
  outOfRange: '기간 외 메시지',
  emptyText: '파일·이미지만',
  fetchError: 'API 오류'
};

/* =========================
 * 웹앱 진입점
 * ========================= */
var WEBHOOK_SHEET = '만족도_웹훅';
var WEBHOOK_HEADERS = ['수신시각', 'chatId', '점수', '영향 요소', '기타 사유', '이름', '전체 payload'];

function doGet(e) {
  return HtmlService.createTemplateFromFile('dashboard')
    .evaluate()
    .setTitle('채널톡 수집 대시보드')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 채널톡 워크플로우 웹훅 수신
 * URL: <배포URL>/exec?secret=<시크릿토큰>
 * Body: JSON (chatId, score, factor, etc, name)
 */
function doPost(e) {
  try {
    var raw = (e && e.postData && e.postData.contents) ? e.postData.contents : '';

    // Secret token 검증 (WEBHOOK_SECRET script property 세팅되어 있으면)
    var expected = PropertiesService.getScriptProperties().getProperty('WEBHOOK_SECRET');
    if (expected) {
      var provided = (e && e.parameter && e.parameter.secret) || '';
      if (provided !== expected) {
        return jsonResponse({ ok: false, error: 'unauthorized' }, 401);
      }
    }

    var payload = {};
    try { payload = JSON.parse(raw); } catch (parseErr) {}

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ensureSheet(ss, WEBHOOK_SHEET);
    if (sh.getLastRow() === 0) {
      writeSheetValues(sh, 1, 1, [WEBHOOK_HEADERS]);
    }

    var chatId = String(payload.chatId || payload.userChatId ||
      (payload.userChat && payload.userChat.id) ||
      (payload.entity && payload.entity.id) || '');
    var score = extractWebhookField(payload, ['score', '점수', 'rating']);
    var factor = extractWebhookField(payload, ['factor', '영향요소', '영향 요소', '요소', 'reason']);
    var etc = extractWebhookField(payload, ['etc', '기타', '기타사유', '기타 사유', 'comment', 'other']);
    var name = extractWebhookField(payload, ['name', 'userName']);

    sh.appendRow([
      new Date(),
      chatId,
      score !== null && score !== undefined ? String(score) : '',
      factor || '',
      etc || '',
      name || '',
      raw.slice(0, 8000)
    ]);

    return jsonResponse({ ok: true, chatId: chatId, saved: { score: score, factor: factor, etc: etc } });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err.message || err) }, 500);
  }
}

function jsonResponse(obj, code) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function extractWebhookField(obj, keys) {
  if (!obj || typeof obj !== 'object') return null;
  // 직접 매치
  for (var i = 0; i < keys.length; i++) {
    if (obj[keys[i]] !== undefined && obj[keys[i]] !== null && obj[keys[i]] !== '') {
      return obj[keys[i]];
    }
  }
  // 깊이 검색
  for (var k in obj) {
    if (!obj.hasOwnProperty(k)) continue;
    var v = obj[k];
    if (v && typeof v === 'object' && !Array.isArray(v)) {
      var found = extractWebhookField(v, keys);
      if (found !== null) return found;
    }
  }
  return null;
}

/**
 * 대시보드가 웹훅 저장 시트 데이터를 chatId 기준으로 매칭 조회
 */
function api_getStoredSatisfactionData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(WEBHOOK_SHEET);
  if (!sh) return { sheetExists: false, byChatId: {}, count: 0 };
  var last = sh.getLastRow();
  if (last < 2) return { sheetExists: true, byChatId: {}, count: 0 };

  var vals = sh.getRange(2, 1, last - 1, WEBHOOK_HEADERS.length).getValues();
  var byChatId = {};
  var tz = Session.getScriptTimeZone();
  vals.forEach(function(row) {
    var chatId = String(row[1] || '').trim();
    if (!chatId) return;
    byChatId[chatId] = {
      receivedAt: row[0] instanceof Date ? Utilities.formatDate(row[0], tz, 'yyyy-MM-dd HH:mm') : String(row[0] || ''),
      score: row[2],
      factor: String(row[3] || ''),
      etc: String(row[4] || ''),
      name: String(row[5] || '')
    };
  });
  return { sheetExists: true, byChatId: byChatId, count: Object.keys(byChatId).length };
}

/**
 * 캡처용 샘플 데이터 삽입 - 이미 매칭된 만족도 응답자에게 임시 점수/영향요소/기타사유 주입
 * 배포 후 실제 웹훅 데이터가 들어오면 자연스럽게 대체됨
 */
function api_seedFakeSatisfactionData(channelKey) {
  var ch = CHANNELS[channelKey];
  if (!ch || !ch.satisfactionGroupId) return { ok: false, error: '만족도 없는 채널' };

  var res = api_getSatisfactionRespondents(channelKey, '');
  if (!res.ok) return res;
  var respondents = res.respondents || [];
  if (!respondents.length) return { ok: false, error: '응답자 없음' };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ensureSheet(ss, WEBHOOK_SHEET);
  if (sh.getLastRow() === 0) {
    writeSheetValues(sh, 1, 1, [WEBHOOK_HEADERS]);
  }

  // 이미 저장된 chatId skip
  var existing = {};
  if (sh.getLastRow() > 1) {
    var vals = sh.getRange(2, 2, sh.getLastRow() - 1, 1).getValues();
    vals.forEach(function(r) { existing[String(r[0]).trim()] = true; });
  }

  var factors = [
    { name: '🙎‍♂️ 상담의 친절도', weight: 50 },
    { name: '⚡ 상담의 신속성', weight: 30 },
    { name: '✅ 안내의 정확성', weight: 20 }
  ];
  var etcSamples = [
    '친절하게 답변해주셔서 감사합니다',
    '빠른 응대에 만족합니다',
    '기다림 없이 바로 해결되어 좋았어요',
    '문의사항이 명확하게 해결됐어요',
    '정확한 안내 감사드립니다'
  ];

  var rows = [];
  var seedCount = 0;
  var skipped = 0;
  respondents.forEach(function(r) {
    if (existing[r.chatId]) { skipped++; return; }

    // 점수 분포: 5점 30%, 4점 55%, 3점 12%, 2점 3%
    var sr = Math.random();
    var score;
    if (sr < 0.30) score = 5;
    else if (sr < 0.85) score = 4;
    else if (sr < 0.97) score = 3;
    else score = 2;

    // 영향요소
    var fr = Math.random() * 100;
    var factor = factors[0].name;
    var cum = 0;
    for (var i = 0; i < factors.length; i++) {
      cum += factors[i].weight;
      if (fr < cum) { factor = factors[i].name; break; }
    }

    // 기타 사유: 20% 확률
    var etc = Math.random() < 0.20 ? etcSamples[Math.floor(Math.random() * etcSamples.length)] : '';

    rows.push([
      new Date(),
      r.chatId,
      String(score),
      factor,
      etc,
      r.name || '',
      '{"seed":true}'
    ]);
    seedCount++;
  });

  if (rows.length) {
    writeSheetValues(sh, sh.getLastRow() + 1, 1, rows);
  }

  return { ok: true, seeded: seedCount, alreadyExisted: skipped, totalRespondents: respondents.length };
}

/**
 * 캡처용 샘플 데이터 삭제 - seed 표시 있는 행만 정리
 */
function api_clearFakeSatisfactionData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(WEBHOOK_SHEET);
  if (!sh) return { ok: true, deleted: 0 };
  var last = sh.getLastRow();
  if (last < 2) return { ok: true, deleted: 0 };
  var vals = sh.getRange(2, 1, last - 1, WEBHOOK_HEADERS.length).getValues();
  var deletedRows = 0;
  for (var i = vals.length - 1; i >= 0; i--) {
    var raw = String(vals[i][6] || '');
    if (raw.indexOf('"seed":true') >= 0) {
      sh.deleteRow(i + 2);
      deletedRows++;
    }
  }
  return { ok: true, deleted: deletedRows };
}

/**
 * 이미 수집된 전사문에서 만족도 form 텍스트 추출
 * 채널톡 UI 에는 form이 보이지만 API 응답에는 없을 수도 있고
 * 반대로 전사문에는 캡처된 경우가 있어서 확인해봄
 */
function api_scanTranscriptsForSurvey(channelKey) {
  var ch = CHANNELS[channelKey];
  if (!ch) return { ok: false, error: '알 수 없는 채널' };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var suffix = '_' + ch.label;
  var sheets = ss.getSheets().filter(function(s) {
    var name = s.getName();
    return /^\d{8}~\d{8}_/.test(name) && name.substr(name.length - suffix.length) === suffix;
  });

  var results = [];
  var stats = { totalScanned: 0, matched: 0, sheetsScanned: sheets.length };
  var samples = [];

  sheets.forEach(function(sh) {
    var last = sh.getLastRow();
    if (last < 2) return;
    var numCols = Math.max(sh.getLastColumn(), CFG.OUT_HEADERS.length);
    var vals = sh.getRange(2, 1, last - 1, numCols).getValues();
    for (var i = 0; i < vals.length; i++) {
      var row = vals[i];
      stats.totalScanned++;
      var transcript = String(row[10] || '');
      if (!/상담\s*만족도\s*점수/.test(transcript)) continue;

      stats.matched++;
      var parsed = parseSurveyForm(transcript);
      var link = String(row[7] || '');
      var chatIdMatch = link.match(/user-chats\/([a-f0-9]{20,40})/);
      var chatId = chatIdMatch ? chatIdMatch[1] : '';

      var record = {
        chatId: chatId,
        name: String(row[1] || ''),
        bizName: String(row[3] || ''),
        respondedAt: String(row[6] || ''),
        link: link,
        score: parsed.score,
        factor: parsed.factor,
        etc: parsed.etc
      };
      results.push(record);
      if (samples.length < 3) {
        // 원본 컨텍스트 발췌 (성공/실패 판단용)
        var pos = transcript.search(/상담\s*만족도\s*점수/);
        var snippet = transcript.slice(Math.max(0, pos - 50), pos + 500);
        samples.push({ chatId: chatId, name: record.name, snippet: snippet, parsed: parsed });
      }
    }
  });

  return {
    ok: true,
    channel: channelKey,
    stats: stats,
    respondents: results,
    samples: samples,
    note: stats.matched > 0
      ? '✅ 전사문에서 ' + stats.matched + '건 발견'
      : '❌ 전사문에도 form 내용 없음'
  };
}

/**
 * 전사문 스캔 결과를 만족도_웹훅 시트에 반영 (payload에 "source":"transcript" 표시)
 */
function api_importSurveyFromTranscripts(channelKey) {
  var scan = api_scanTranscriptsForSurvey(channelKey);
  if (!scan.ok) return scan;
  if (!scan.respondents.length) return { ok: false, error: '전사문에 form 내용 없음' };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ensureSheet(ss, WEBHOOK_SHEET);
  if (sh.getLastRow() === 0) {
    writeSheetValues(sh, 1, 1, [WEBHOOK_HEADERS]);
  }

  // 이미 저장된 chatId 스킵
  var existing = {};
  if (sh.getLastRow() > 1) {
    var evals = sh.getRange(2, 2, sh.getLastRow() - 1, 1).getValues();
    evals.forEach(function(r) { existing[String(r[0]).trim()] = true; });
  }

  var rows = [];
  var imported = 0;
  var skipped = 0;
  scan.respondents.forEach(function(r) {
    if (!r.chatId || existing[r.chatId]) { skipped++; return; }
    if (r.score === null && !r.factor && !r.etc) return;
    rows.push([
      new Date(),
      r.chatId,
      r.score !== null && r.score !== undefined ? String(r.score) : '',
      r.factor || '',
      r.etc || '',
      r.name || '',
      JSON.stringify({ source: 'transcript' })
    ]);
    imported++;
  });

  if (rows.length) writeSheetValues(sh, sh.getLastRow() + 1, 1, rows);

  return { ok: true, imported: imported, skipped: skipped, scanned: scan.stats.totalScanned };
}

/**
 * 특정 chatId 의 모든 메시지 raw dump - form 메시지 실제 저장 위치 찾기용
 */
function api_dumpAllMessages(channelKey, chatId) {
  var ch = CHANNELS[channelKey];
  if (!ch) return { ok: false, error: '알 수 없는 채널' };
  var keys = getApiKeys(channelKey);
  if (!keys) return { ok: false, error: 'API 키 없음' };
  if (!chatId) return { ok: false, error: 'chatId 필요' };

  var headers = buildHeaders(keys.key, keys.secret);
  var allMessages = [];
  var since = '';
  for (var p = 0; p < 5; p++) {
    var params = { limit: 100, sortOrder: 'desc' };
    if (since) params.since = since;
    try {
      var data = ctGet('/user-chats/' + chatId + '/messages', params, headers);
      var msgs = data.messages || [];
      msgs.forEach(function(m) {
        allMessages.push({
          id: m.id,
          personType: m.personType,
          personId: m.personId,
          type: m.type,
          plainText: m.plainText || '',
          blocks: m.blocks || [],
          log: m.log || null,
          options: m.options || [],
          createdAt: m.createdAt,
          allKeys: Object.keys(m)
        });
      });
      if (!data.next) break;
      since = data.next;
    } catch (e) { break; }
  }

  // form 관련 후보 필터
  var candidates = allMessages.filter(function(m) {
    var t = (m.plainText || '') + JSON.stringify(m.blocks || []);
    return /만족도|점수|영향|기타|form|저장됨/i.test(t) ||
           (m.log && /form|answer|submit/i.test(String(m.log.action || '')));
  });

  return {
    ok: true,
    chatId: chatId,
    totalMessages: allMessages.length,
    formCandidates: candidates.length,
    candidates: candidates.slice(0, 10),
    allMessagesBrief: allMessages.map(function(m) {
      return {
        id: m.id,
        personType: m.personType,
        type: m.type,
        hasPlainText: !!m.plainText,
        plainTextPreview: (m.plainText || '').slice(0, 80),
        blockCount: (m.blocks || []).length,
        blockTypes: (m.blocks || []).map(function(b) { return b.type; }),
        hasLog: !!m.log,
        logAction: m.log ? m.log.action : null,
        options: m.options
      };
    })
  };
}

/**
 * 특정 chatId 진단 실행 (Apps Script 편집기에서 직접 실행)
 * 결과는 로그 + '_debug_dump' 시트에 저장
 */
function runDumpKangMinKyung() {
  var chatId = '6a8d7e42f2f275548a4a'; // 강민경 chat
  var res = api_dumpAllMessages('refund', chatId);

  Logger.log('==============================');
  Logger.log('chatId: ' + res.chatId);
  Logger.log('총 메시지: ' + res.totalMessages);
  Logger.log('form 후보 (만족도 관련 키워드 포함): ' + res.formCandidates);
  Logger.log('==============================');

  (res.candidates || []).forEach(function(c, i) {
    Logger.log('----- 후보 #' + (i + 1) + ' -----');
    Logger.log('id=' + c.id + ', personType=' + c.personType + ', type=' + c.type);
    Logger.log('plainText: ' + String(c.plainText || '(empty)').slice(0, 300));
    Logger.log('blocks (' + (c.blocks || []).length + '개): ' + JSON.stringify(c.blocks).slice(0, 800));
    Logger.log('log: ' + JSON.stringify(c.log));
    Logger.log('options: ' + JSON.stringify(c.options));
    Logger.log('allKeys: ' + (c.allKeys || []).join(','));
  });

  Logger.log('==============================');
  Logger.log('전체 메시지 요약:');
  (res.allMessagesBrief || []).forEach(function(m, i) {
    Logger.log('[' + (i + 1) + '] type=' + m.type + ' pType=' + m.personType +
      ' text="' + (m.plainTextPreview || '') + '"' +
      ' blocks=' + (m.blockCount || 0) +
      ' blockTypes=' + JSON.stringify(m.blockTypes) +
      ' logAction=' + (m.logAction || '-') +
      ' options=' + JSON.stringify(m.options));
  });

  // 시트에도 저장 (시각적으로 확인 가능)
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('_debug_dump') || ss.insertSheet('_debug_dump');
  sh.clearContents();

  var candidatesById = {};
  (res.candidates || []).forEach(function(c) { candidatesById[c.id] = c; });

  var rows = [
    ['chatId', res.chatId, '', '', '', '', '', '', '', '', ''],
    ['총 메시지', res.totalMessages, '', '', '', '', '', '', '', '', ''],
    ['form 후보', res.formCandidates, '', '', '', '', '', '', '', '', ''],
    ['', '', '', '', '', '', '', '', '', '', ''],
    ['idx', 'personType', 'type', 'plainText', 'blockCount', 'blockTypes', 'blocksJSON', 'logAction', 'logJSON', 'options', 'allKeys']
  ];
  (res.allMessagesBrief || []).forEach(function(m, i) {
    var full = candidatesById[m.id] || {};
    rows.push([
      i + 1,
      m.personType || '',
      m.type || '',
      m.plainTextPreview || '',
      m.blockCount || 0,
      JSON.stringify(m.blockTypes || []),
      full.blocks ? JSON.stringify(full.blocks).slice(0, 5000) : '',
      m.logAction || '',
      full.log ? JSON.stringify(full.log) : '',
      JSON.stringify(m.options || []),
      full.allKeys ? full.allKeys.join(',') : ''
    ]);
  });
  sh.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  Logger.log('시트에도 저장됨: _debug_dump 탭에서 확인');
  return res;
}

/**
 * 웹훅 URL 확인용 (배포 URL 반환)
 */
function api_getWebhookInfo() {
  var url = ScriptApp.getService().getUrl();
  var secret = PropertiesService.getScriptProperties().getProperty('WEBHOOK_SECRET');
  return {
    url: url || '',
    hasSecret: !!secret,
    secretHint: secret ? (secret.slice(0, 4) + '****') : null
  };
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/* =========================
 * 시트 메뉴 (수동 관리용)
 * ========================= */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('채널톡 대시보드')
    .addItem('환급 API 키 설정', 'setChannelKeysRefund')
    .addItem('케어 API 키 설정', 'setChannelKeysCare')
    .addItem('현재 작업 상태 확인', 'showJobStatus')
    .addItem('수집 중지', 'cancelChannelTalkJob')
    .addToUi();
}

function setChannelKeysRefund() { setChannelKeys('refund'); }
function setChannelKeysCare() { setChannelKeys('care'); }

function setChannelKeys(channelKey) {
  var ui = SpreadsheetApp.getUi();
  var ch = CHANNELS[channelKey];
  if (!ch) return;

  var keyRes = ui.prompt(ch.label + ' - Access Key', ui.ButtonSet.OK_CANCEL);
  if (keyRes.getSelectedButton() !== ui.Button.OK) return;
  var secRes = ui.prompt(ch.label + ' - Access Secret', ui.ButtonSet.OK_CANCEL);
  if (secRes.getSelectedButton() !== ui.Button.OK) return;

  var key = String(keyRes.getResponseText() || '').trim();
  var sec = String(secRes.getResponseText() || '').trim();
  if (!key || !sec) {
    ui.alert('키가 비어 있습니다.');
    return;
  }

  PropertiesService.getScriptProperties()
    .setProperty(ch.propKey, key)
    .setProperty(ch.propSecret, sec);
  ui.alert(ch.label + ' 키 저장 완료');
}

function showJobStatus() {
  var lines = [];
  Object.keys(CHANNELS).forEach(function(k) {
    var job = getJobState(k);
    var chLabel = CHANNELS[k].label;
    if (job) {
      lines.push(chLabel + ' - ' + job.status + ' (' + job.nextIdx + '/' + job.total + ')');
    } else {
      lines.push(chLabel + ' - 진행 없음');
    }
  });
  SpreadsheetApp.getUi().alert(lines.join('\n'));
}

/* =========================
 * 웹앱 API (client가 호출)
 * ========================= */
function api_getChannels() {
  return Object.keys(CHANNELS).map(function(k) {
    var ch = CHANNELS[k];
    return {
      key: ch.key,
      label: ch.label,
      hasApiKey: !!getApiKeys(k),
      hasSatisfaction: !!ch.satisfactionGroupId
    };
  });
}

function api_getJobState(channelKey) {
  if (channelKey && CHANNELS[channelKey]) {
    return getJobStateEnriched(channelKey);
  }
  // 인자 없을 때: 진행 중인 첫 채널 (backward compat)
  var keys = Object.keys(CHANNELS);
  for (var i = 0; i < keys.length; i++) {
    var s = getJobStateEnriched(keys[i]);
    if (s.running) return s;
  }
  return { running: false };
}

function getJobStateEnriched(channelKey) {
  var job = getJobState(channelKey);
  if (!job) return { running: false, channel: channelKey };
  var ch = CHANNELS[channelKey];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var collected = 0;
  if (job.outputSheet) {
    var shOut = ss.getSheetByName(job.outputSheet);
    if (shOut) collected = Math.max(0, shOut.getLastRow() - 1);
  }
  return {
    running: true,
    channel: channelKey,
    channelLabel: ch ? ch.label : channelKey,
    status: job.status,
    phase: job.phase || '',
    nextIdx: job.nextIdx,
    total: job.total,
    failed: job.failed,
    collected: collected,
    outputSheet: job.outputSheet,
    progressPct: job.total ? Math.round((job.nextIdx / job.total) * 100) : 0,
    skips: loadSkipCounts(channelKey)
  };
}

function api_getLastProgress(channelKey) {
  if (!CHANNELS[channelKey]) return null;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(CFG.SHEET_PROGRESS_PREFIX + channelKey);
  if (!sh) return null;
  var last = sh.getLastRow();
  if (last < 1) return null;
  var vals = readSheetValues(sh, 1, 1, Math.min(last, 25), 2);
  var map = {};
  for (var i = 0; i < vals.length; i++) {
    var k = String(vals[i][0] || '').trim();
    if (!k) continue;
    if (k.indexOf('[') === 0) break;
    map[k] = vals[i][1];
  }
  return {
    channel: channelKey,
    channelLabel: CHANNELS[channelKey].label,
    status: String(map['상태'] || ''),
    period: String(map['기간'] || ''),
    outputSheet: String(map['출력 탭'] || ''),
    candidates: Number(map['후보 대화방'] || 0),
    scanned: Number(map['처리한 대화방'] || 0),
    collected: Number(map['확정 수집'] || 0),
    failed: Number(map['전사 실패(건너뜀)'] || 0),
    progressPct: String(map['진행률'] || ''),
    note: String(map['비고'] || ''),
    updatedAt: String(map['갱신시각'] || ''),
    skips: {
      autoClosed: Number(map['스킵 - ' + SKIP_LABELS.autoClosed] || 0),
      outOfRange: Number(map['스킵 - ' + SKIP_LABELS.outOfRange] || 0),
      emptyText: Number(map['스킵 - ' + SKIP_LABELS.emptyText] || 0),
      fetchError: Number(map['스킵 - ' + SKIP_LABELS.fetchError] || 0)
    }
  };
}

function api_getAllChannelStatus() {
  var out = [];
  Object.keys(CHANNELS).forEach(function(k) {
    var running = getJobState(k);
    if (running) {
      out.push({
        channel: k,
        channelLabel: CHANNELS[k].label,
        running: true,
        state: getJobStateEnriched(k)
      });
    } else {
      var last = api_getLastProgress(k);
      out.push({
        channel: k,
        channelLabel: CHANNELS[k].label,
        running: false,
        state: last
      });
    }
  });
  return out;
}

function api_getSpreadsheetUrl() {
  return SpreadsheetApp.getActiveSpreadsheet().getUrl();
}

/* =========================
 * Phase 2 - 대시보드 데이터 집계
 * ========================= */
function api_getDashboardData(channelKey, sheetName) {
  if (!CHANNELS[channelKey]) return null;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var chLabel = CHANNELS[channelKey].label;
  var suffix = '_' + chLabel;

  var outputSheets;
  if (sheetName) {
    var only = ss.getSheetByName(sheetName);
    outputSheets = only ? [only] : [];
  } else {
    outputSheets = ss.getSheets().filter(function(s) {
      var name = s.getName();
      return /^\d{8}~\d{8}_/.test(name) && name.substr(name.length - suffix.length) === suffix;
    });
  }

  var totalChats = 0;
  var totalTags = 0;
  var tagCounts = {};
  var inflowCounts = {};
  var managerCounts = {};
  var ratingSum = 0;
  var ratingCount = 0;
  var ratingDist = {};
  var respondents = [];

  outputSheets.forEach(function(sh) {
    var last = sh.getLastRow();
    if (last < 2) return;
    var numCols = Math.max(sh.getLastColumn(), CFG.OUT_HEADERS.length);
    var vals = sh.getRange(2, 1, last - 1, numCols).getValues();
    for (var i = 0; i < vals.length; i++) {
      var row = vals[i];
      totalChats++;

      var repName = String(row[1] || '').trim();
      var bizName = String(row[3] || '').trim();
      var manager = String(row[4] || '').trim();
      var startTime = String(row[6] || '').trim();
      var deskLink = String(row[7] || '').trim();
      var inflow = String(row[8] || '').trim() || '미확인';
      var tagsRaw = String(row[9] || '').trim();
      var rating = row[11];
      var reviewedAt = String(row[12] || '').trim();

      if (manager) managerCounts[manager] = (managerCounts[manager] || 0) + 1;
      inflowCounts[inflow] = (inflowCounts[inflow] || 0) + 1;

      if (tagsRaw) {
        var tags = tagsRaw.split(',').map(function(t) { return t.trim(); }).filter(Boolean);
        totalTags += tags.length;
        for (var t = 0; t < tags.length; t++) {
          tagCounts[tags[t]] = (tagCounts[tags[t]] || 0) + 1;
        }
      }

      var rNum = rating === '' || rating === null || rating === undefined ? null : Number(rating);
      if (rNum !== null && !isNaN(rNum)) {
        ratingSum += rNum;
        ratingCount++;
        var key = String(rNum);
        ratingDist[key] = (ratingDist[key] || 0) + 1;
        respondents.push({
          name: repName || '(이름없음)',
          bizName: bizName,
          manager: manager,
          rating: rNum,
          inflow: inflow,
          tags: tagsRaw,
          startTime: startTime,
          reviewedAt: reviewedAt,
          link: deskLink
        });
      }
    }
  });

  var topTags = Object.keys(tagCounts).map(function(k) {
    return { name: k, count: tagCounts[k] };
  }).sort(function(a, b) { return b.count - a.count; }).slice(0, 15);

  var inflowList = Object.keys(inflowCounts).map(function(k) {
    return { name: k, count: inflowCounts[k] };
  }).sort(function(a, b) { return b.count - a.count; });

  var managerList = Object.keys(managerCounts).map(function(k) {
    return { name: k, count: managerCounts[k] };
  }).sort(function(a, b) { return b.count - a.count; }).slice(0, 20);

  respondents.sort(function(a, b) { return b.rating - a.rating; });

  var history = getExecutionHistory(ss, channelKey);
  var executedCount = history.length;
  var executedSum = history.reduce(function(acc, h) { return acc + (h.collected || 0); }, 0);

  return {
    channel: channelKey,
    channelLabel: chLabel,
    sheetName: sheetName || '',
    generatedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'),
    kpi: {
      totalChats: totalChats,
      totalTags: totalTags,
      uniqueTags: Object.keys(tagCounts).length,
      inflowKinds: inflowList.length,
      managerCount: Object.keys(managerCounts).length,
      executedCount: executedCount,
      cumulativeCollected: executedSum,
      outputSheetCount: outputSheets.length,
      avgRating: ratingCount > 0 ? Math.round(ratingSum / ratingCount * 100) / 100 : null,
      ratingResponseCount: ratingCount,
      ratingResponseRate: totalChats > 0 ? Math.round(ratingCount / totalChats * 1000) / 10 : 0
    },
    topTags: topTags,
    inflowDistribution: inflowList,
    managerList: managerList,
    ratingDistribution: ratingDist,
    respondents: respondents.slice(0, 100),
    history: history.slice(0, 10)
  };
}

/**
 * 만족도 응답자 조회 - 채널의 satisfactionGroupId 그룹에서 메시지를 스캔해 chatId 추출 → 수집된 데이터와 매칭
 */
function api_getSatisfactionRespondents(channelKey, sheetName) {
  var ch = CHANNELS[channelKey];
  if (!ch) return { ok: false, error: '알 수 없는 채널' };
  if (!ch.satisfactionGroupId) return { ok: false, error: '이 채널은 만족도 조사가 없습니다.' };
  var keys = getApiKeys(channelKey);
  if (!keys) return { ok: false, error: 'API 키 없음' };

  var headers = buildHeaders(keys.key, keys.secret);
  var respondentChatIds = {}; // chatId -> 응답시각(ms)
  var since = '';
  var pagesScanned = 0;
  var earliestMs = null;
  var latestMs = null;
  var last7dCount = 0;
  var last30dCount = 0;
  var now = new Date().getTime();
  var sevenDaysAgo = now - 7 * 24 * 60 * 60 * 1000;
  var thirtyDaysAgo = now - 30 * 24 * 60 * 60 * 1000;

  while (pagesScanned < 30) {
    var params = { limit: 100, sortOrder: 'desc' };
    if (since) params.since = since;
    var data;
    try {
      data = ctGet('/groups/' + ch.satisfactionGroupId + '/messages', params, headers);
    } catch (e) {
      return { ok: false, error: 'groups API 오류: ' + String(e.message || e) };
    }
    pagesScanned++;
    var msgs = data.messages || [];
    if (!msgs.length) break;

    msgs.forEach(function(m) {
      var texts = [m.plainText || ''];
      (m.blocks || []).forEach(function(b) {
        if (b.value) texts.push(String(b.value));
      });
      var combined = texts.join(' ');
      var re = /user-chats\/([a-f0-9]{20,40})/g;
      var match;
      while ((match = re.exec(combined)) !== null) {
        if (!respondentChatIds[match[1]]) {
          var ts = Number(m.createdAt || 0);
          respondentChatIds[match[1]] = ts;
          if (ts) {
            if (earliestMs === null || ts < earliestMs) earliestMs = ts;
            if (latestMs === null || ts > latestMs) latestMs = ts;
            if (ts >= sevenDaysAgo) last7dCount++;
            if (ts >= thirtyDaysAgo) last30dCount++;
          }
        }
      }
    });

    if (!data.next) break;
    since = data.next;
  }

  // 수집된 시트에서 chatId → row 매핑
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets;
  if (sheetName) {
    var s = ss.getSheetByName(sheetName);
    sheets = s ? [s] : [];
  } else {
    var suffix = '_' + ch.label;
    sheets = ss.getSheets().filter(function(s) {
      var name = s.getName();
      return /^\d{8}~\d{8}_/.test(name) && name.substr(name.length - suffix.length) === suffix;
    });
  }

  var chatDataMap = {};
  var totalChatsInScope = 0;
  sheets.forEach(function(sh) {
    var last = sh.getLastRow();
    if (last < 2) return;
    var numCols = Math.max(sh.getLastColumn(), CFG.OUT_HEADERS.length);
    var vals = sh.getRange(2, 1, last - 1, numCols).getValues();
    vals.forEach(function(row) {
      totalChatsInScope++;
      var link = String(row[7] || '');
      var m = link.match(/user-chats\/([a-f0-9]{20,40})/);
      if (m) chatDataMap[m[1]] = row;
    });
  });

  var matched = [];
  Object.keys(respondentChatIds).forEach(function(chatId) {
    if (chatDataMap[chatId]) {
      var row = chatDataMap[chatId];
      matched.push({
        chatId: chatId,
        name: String(row[1] || ''),
        bizName: String(row[3] || ''),
        manager: String(row[4] || ''),
        startTime: String(row[6] || ''),
        link: String(row[7] || ''),
        inflow: String(row[8] || ''),
        tags: String(row[9] || ''),
        respondedAt: respondentChatIds[chatId] ? Utilities.formatDate(new Date(respondentChatIds[chatId]), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') : ''
      });
    }
  });

  matched.sort(function(a, b) {
    return (b.respondedAt || '').localeCompare(a.respondedAt || '');
  });

  var tz = Session.getScriptTimeZone();
  return {
    ok: true,
    channel: channelKey,
    groupId: ch.satisfactionGroupId,
    pagesScanned: pagesScanned,
    totalRespondentsScanned: Object.keys(respondentChatIds).length,
    matchedInSheets: matched.length,
    totalChatsInScope: totalChatsInScope,
    responseRate: totalChatsInScope > 0 ? Math.round(matched.length / totalChatsInScope * 1000) / 10 : 0,
    last7dCount: last7dCount,
    last30dCount: last30dCount,
    earliestAt: earliestMs ? Utilities.formatDate(new Date(earliestMs), tz, 'yyyy-MM-dd') : '',
    latestAt: latestMs ? Utilities.formatDate(new Date(latestMs), tz, 'yyyy-MM-dd HH:mm') : '',
    respondents: matched
  };
}

/**
 * 만족도 응답자 각각의 채팅에서 숫자 점수를 추출 (서술형 제외).
 * chatIds 를 받아서 { chatId: [scores...] } 반환
 */
function api_getSatisfactionScores(channelKey, chatIds) {
  var ch = CHANNELS[channelKey];
  if (!ch) return { ok: false, error: '알 수 없는 채널' };
  var keys = getApiKeys(channelKey);
  if (!keys) return { ok: false, error: 'API 키 없음' };
  if (!chatIds || !chatIds.length) return { ok: false, error: 'chatId 없음' };

  var headers = buildHeaders(keys.key, keys.secret);
  var scoresPerChat = {};
  var diagnostics = {};
  var startedAt = new Date().getTime();
  var deadline = startedAt + 4.5 * 60 * 1000;

  for (var i = 0; i < chatIds.length; i++) {
    if (new Date().getTime() > deadline) break;
    var chatId = String(chatIds[i]);
    try {
      var data = ctGet('/user-chats/' + chatId + '/messages', { limit: 100, sortOrder: 'desc' }, headers);
      var msgs = data.messages || [];
      var extracted = extractNumericScoresFromMessages(msgs);
      scoresPerChat[chatId] = extracted.scores;
      if (i < 3) diagnostics[chatId] = extracted.diag;
    } catch (e) {
      scoresPerChat[chatId] = [];
      diagnostics[chatId] = { error: String(e.message || e) };
    }
  }

  return {
    ok: true,
    processedCount: Object.keys(scoresPerChat).length,
    totalRequested: chatIds.length,
    scores: scoresPerChat,
    diagnostics: diagnostics,
    elapsedSec: Math.round((new Date().getTime() - startedAt) / 1000)
  };
}

/**
 * 메시지 배열에서 사용자가 남긴 숫자형 답변을 추출 (서술형 제외)
 * 대응 패턴:
 *  - "5", "5점", "5/5", "5 / 5"
 *  - "★★★★★" (별 개수)
 *  - "매우 만족" / "만족" / "보통" / "불만족" / "매우 불만족" → 5~1점
 */
/**
 * 채널톡 만족도 조사 form 메시지 파싱
 * 구조:
 *   상담 만족도 점수(1~5점)를 선택해 주세요. *
 *   5
 *   해당 점수에 가장 큰 영향을 미친 요소가 무엇인가요? *
 *   🙎‍♂️ 상담의 친절도
 *   '기타'로 선택하신 사유를 작성해 주세요. (선택)
 *   답변 없음
 */
function extractNumericScoresFromMessages(msgs) {
  var answers = []; // { score, factor, etc, createdAt, raw }
  var diag = { userTexts: [], botTexts: [], typesCount: {}, blockSamples: [], surveyFound: 0 };

  for (var i = 0; i < msgs.length; i++) {
    var m = msgs[i];
    var mType = String(m.type || 'text');
    diag.typesCount[mType] = (diag.typesCount[mType] || 0) + 1;

    var text = String(m.plainText || '');
    var blockValues = [];
    (m.blocks || []).forEach(function(b) {
      if (b && b.value) blockValues.push(String(b.value));
      if (b && b.type && /form|choice|button|review|rating/i.test(String(b.type)) && diag.blockSamples.length < 5) {
        diag.blockSamples.push({ type: b.type, block: b });
      }
    });
    var combined = (text + '\n' + blockValues.join('\n')).trim();
    if (!combined) continue;

    var pType = String(m.personType || '').toLowerCase();
    if (pType === 'user' && diag.userTexts.length < 10) diag.userTexts.push(combined.slice(0, 200));
    if (pType !== 'user' && diag.botTexts.length < 10) diag.botTexts.push(combined.slice(0, 200));

    // 만족도 조사 form 메시지 검색
    if (/상담\s*만족도\s*점수/.test(combined)) {
      diag.surveyFound++;
      var parsed = parseSurveyForm(combined);
      if (parsed && (parsed.score !== null || parsed.factor || parsed.etc)) {
        parsed.createdAt = Number(m.createdAt || 0);
        parsed.raw = combined.slice(0, 300);
        answers.push(parsed);
      }
    }
  }

  return { scores: answers, diag: diag };
}

function parseSurveyForm(text) {
  var score = null;
  var factor = null;
  var etc = null;

  // 점수: "상담 만족도 점수(1~5점)를 선택해 주세요. *\n5"
  var scoreMatch = text.match(/상담\s*만족도\s*점수[^\n]*\n\s*([1-5])\s*(?:\n|$)/);
  if (scoreMatch) score = Number(scoreMatch[1]);
  else {
    // fallback: 어디든 첫 1~5 숫자
    var alt = text.match(/상담\s*만족도\s*점수[^]{0,200}?([1-5])(?:\s*점)?/);
    if (alt) score = Number(alt[1]);
  }

  // 영향 요소: "해당 점수에 가장 큰 영향을 미친 요소가 무엇인가요? *\n🙎‍♂️ 상담의 친절도"
  var factorMatch = text.match(/해당\s*점수에\s*가장\s*큰\s*영향[^\n]*\n\s*([^\n]+)/);
  if (factorMatch) {
    factor = factorMatch[1].trim();
    if (factor === '답변 없음' || factor === '') factor = null;
  }

  // 기타 사유
  var etcMatch = text.match(/['`'’‘]?기타['`'’‘]?로\s*선택하신\s*사유[^\n]*\n\s*([^\n]+)/);
  if (etcMatch) {
    etc = etcMatch[1].trim();
    if (etc === '답변 없음' || etc === '') etc = null;
  }

  return { score: score, factor: factor, etc: etc };
}

function api_getChannelSheets(channelKey) {
  if (!CHANNELS[channelKey]) return [];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var suffix = '_' + CHANNELS[channelKey].label;
  var sheets = ss.getSheets().filter(function(s) {
    var name = s.getName();
    return /^\d{8}~\d{8}_/.test(name) && name.substr(name.length - suffix.length) === suffix;
  });
  return sheets.map(function(s) {
    var rows = Math.max(0, s.getLastRow() - 1);
    var m = s.getName().match(/^(\d{8})~(\d{8})_/);
    return {
      name: s.getName(),
      period: m ? (m[1] + ' ~ ' + m[2]) : s.getName(),
      startYmd: m ? m[1] : '',
      rows: rows
    };
  }).sort(function(a, b) { return b.startYmd.localeCompare(a.startYmd); });
}

function getExecutionHistory(ss, channelKey) {
  var sh = ss.getSheetByName(CFG.SHEET_PROGRESS_PREFIX + channelKey);
  if (!sh) return [];
  var last = sh.getLastRow();
  if (last < 1) return [];
  var vals = sh.getRange(1, 1, last, 6).getValues();
  var histIdx = -1;
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0] || '') === '[실행 완료 이력]') {
      histIdx = i + 2;
      break;
    }
  }
  if (histIdx < 0 || histIdx >= vals.length) return [];
  var chLabel = CHANNELS[channelKey] ? CHANNELS[channelKey].label : channelKey;
  var out = [];
  for (var r = histIdx; r < vals.length; r++) {
    var row = vals[r];
    if (!row[0]) continue;
    var period = String(row[1] || '');
    var sheetName = '';
    var m = period.match(/(\d{8})\s*~\s*(\d{8})/);
    if (m) sheetName = m[1] + '~' + m[2] + '_' + chLabel;
    out.push({
      completedAt: formatMaybeDate(row[0]),
      period: period,
      sheetName: sheetName,
      candidates: Number(row[2] || 0),
      collected: Number(row[3] || 0),
      failed: Number(row[4] || 0),
      status: String(row[5] || ''),
      successRate: (row[2] && Number(row[2]) > 0)
        ? Math.round(Number(row[3]) / Number(row[2]) * 1000) / 10
        : 0
    });
  }
  out.reverse();
  return out;
}

function api_getChannelHistory(channelKey) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return getExecutionHistory(ss, channelKey);
}

function api_deleteHistoryEntry(payload) {
  var channelKey = String((payload && payload.channel) || '');
  var completedAt = String((payload && payload.completedAt) || '');
  var period = String((payload && payload.period) || '');
  var deleteSheet = !!(payload && payload.deleteSheet);

  if (!CHANNELS[channelKey]) return { ok: false, error: '알 수 없는 채널' };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(CFG.SHEET_PROGRESS_PREFIX + channelKey);
  if (!sh) return { ok: false, error: '진행상황 시트를 찾을 수 없습니다.' };

  var last = sh.getLastRow();
  if (last < 1) return { ok: false, error: '이력 없음' };

  var vals = sh.getRange(1, 1, last, 6).getValues();
  var histIdx = -1;
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0] || '') === '[실행 완료 이력]') {
      histIdx = i + 2;
      break;
    }
  }
  if (histIdx < 0) return { ok: false, error: '실행 완료 이력 섹션을 찾을 수 없습니다.' };

  var deletedRow = -1;
  for (var r = histIdx; r < vals.length; r++) {
    var row = vals[r];
    if (!row[0]) continue;
    var rowCompletedAt = formatMaybeDate(row[0]);
    var rowPeriod = String(row[1] || '');
    if (rowCompletedAt === completedAt && rowPeriod === period) {
      sh.deleteRow(r + 1);
      deletedRow = r + 1;
      break;
    }
  }

  if (deletedRow < 0) return { ok: false, error: '매칭되는 이력을 찾을 수 없습니다.' };

  var sheetDeleted = false;
  if (deleteSheet) {
    var chLabel = CHANNELS[channelKey].label;
    var m = period.match(/(\d{8})\s*~\s*(\d{8})/);
    if (m) {
      var sheetName = m[1] + '~' + m[2] + '_' + chLabel;
      var outputSh = ss.getSheetByName(sheetName);
      if (outputSh) {
        ss.deleteSheet(outputSh);
        sheetDeleted = true;
      }
    }
  }

  return { ok: true, sheetDeleted: sheetDeleted };
}

function formatMaybeDate(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  }
  return String(v || '');
}

function api_startCollection(payload) {
  var channelKey = String((payload && payload.channel) || '').trim();
  var startYmd = String((payload && payload.startYmd) || '').trim();
  var endYmd = String((payload && payload.endYmd) || '').trim();

  if (!CHANNELS[channelKey]) {
    return { ok: false, error: '채널을 선택하세요.' };
  }
  if (!getApiKeys(channelKey)) {
    return { ok: false, error: CHANNELS[channelKey].label + ' 채널의 API 키가 설정되지 않았습니다.' };
  }

  var startDate = parseYmd(startYmd);
  var endDate = parseYmd(endYmd);
  if (!startDate || !endDate) {
    return { ok: false, error: '날짜 형식 오류: YYYYMMDD' };
  }
  if (startDate > endDate) {
    return { ok: false, error: '시작일이 종료일보다 늦습니다.' };
  }

  // 같은 채널 중복 방지 (다른 채널은 병렬 실행 허용)
  var existing = getJobState(channelKey);
  if (existing && existing.status === 'running') {
    return { ok: false, error: CHANNELS[channelKey].label + ' 채널은 이미 진행 중입니다.' };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    mustSheet(ss, CFG.SHEET_META);
  } catch (e) {
    return { ok: false, error: '"메타베이스" 시트가 없습니다. 먼저 시트를 만들어 주세요.' };
  }

  var startMs = startDate.getTime();
  var endMs = endOfDay(endDate).getTime();
  var periodLabel = formatYmd(startDate) + ' ~ ' + formatYmd(endDate);

  // 프론트 폴링이 바로 상태를 볼 수 있도록 미리 running 세팅 (per-channel)
  var props = PropertiesService.getScriptProperties();
  props.setProperty(jobProp(channelKey, 'STATUS'), 'running');
  props.setProperty(jobProp(channelKey, 'PHASE'), 'preparing');
  props.setProperty(jobProp(channelKey, 'START_MS'), String(startMs));
  props.setProperty(jobProp(channelKey, 'END_MS'), String(endMs));
  props.setProperty(jobProp(channelKey, 'NEXT_IDX'), '0');
  props.setProperty(jobProp(channelKey, 'TOTAL'), '0');
  props.setProperty(jobProp(channelKey, 'FAILED'), '0');
  props.setProperty(jobProp(channelKey, 'STARTED_AT'), String(new Date().getTime()));
  props.setProperty(jobProp(channelKey, 'OUTPUT_SHEET'), '');
  props.setProperty(skipProp(channelKey), JSON.stringify({ autoClosed: 0, outOfRange: 0, emptyText: 0, fetchError: 0 }));

  writeProgress(ss, channelKey, {
    status: '준비중',
    period: periodLabel,
    outputSheet: '(준비중)',
    candidates: '-',
    scanned: 0,
    collected: 0,
    failed: 0,
    nextIdx: 0,
    progressPct: '0%',
    note: '대화방 목록 조회 중 (10~30초)'
  });

  try {
    var keys = getApiKeys(channelKey);
    var headers = buildHeaders(keys.key, keys.secret);
    var candidates = fetchChatCandidates(headers, channelKey, startMs, endMs);
    var outputSheetName = buildPeriodSheetName(startMs, endMs, channelKey);

    initJob(ss, channelKey, candidates, startMs, endMs, outputSheetName);
    initOutputSheet(ss, outputSheetName);
    props.setProperty(jobProp(channelKey, 'PHASE'), 'collecting');

    writeProgress(ss, channelKey, {
      status: '진행중',
      period: periodLabel,
      outputSheet: outputSheetName,
      candidates: candidates.length,
      scanned: 0,
      collected: 0,
      failed: 0,
      nextIdx: 0,
      progressPct: '0%',
      note: '수집 시작'
    });

    var result = runChannelTalkExportResume(channelKey, false);
    return {
      ok: true,
      channel: channelKey,
      total: candidates.length,
      outputSheet: outputSheetName,
      done: !!(result && result.done),
      collected: result ? result.collected : 0
    };
  } catch (e) {
    clearJobState(channelKey);
    writeProgress(ss, channelKey, {
      status: '오류',
      period: periodLabel,
      note: String(e.message || e)
    });
    return { ok: false, error: String(e.message || e) };
  }
}

function api_cancelJob(channelKey) {
  if (channelKey && CHANNELS[channelKey]) {
    cancelChannelJobSilent(channelKey);
  } else {
    // 인자 없을 때: 진행 중인 모든 채널 중지 (backward compat)
    Object.keys(CHANNELS).forEach(function(k) { cancelChannelJobSilent(k); });
  }
  return { ok: true };
}

/* =========================
 * 이어서 실행 - 채널별 트리거 진입점
 * ========================= */
function resumeTrigger_refund() { return runChannelTalkExportResume('refund', true); }
function resumeTrigger_care() { return runChannelTalkExportResume('care', true); }

function runChannelTalkExportResume(channelKey, fromTrigger) {
  if (!channelKey || !CHANNELS[channelKey]) return null;

  // 짧은 스크립트 락으로 flag 확인/설정 → 두 채널 병렬 처리 허용
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return null;
  var token;
  try {
    var props = PropertiesService.getScriptProperties();
    var existing = props.getProperty(lockProp(channelKey));
    if (existing) {
      var t = Number(existing);
      if (new Date().getTime() - t < 6 * 60 * 1000) return null; // 활성 락
    }
    token = String(new Date().getTime());
    props.setProperty(lockProp(channelKey), token);
  } finally {
    lock.releaseLock();
  }

  try {
    return runChannelTalkExportResumeInternal(channelKey, fromTrigger);
  } finally {
    var lock2 = LockService.getScriptLock();
    if (lock2.tryLock(3000)) {
      try {
        var props2 = PropertiesService.getScriptProperties();
        if (props2.getProperty(lockProp(channelKey)) === token) {
          props2.deleteProperty(lockProp(channelKey));
        }
      } finally {
        lock2.releaseLock();
      }
    }
  }
}

function runChannelTalkExportResumeInternal(channelKey, fromTrigger) {
  var started = new Date().getTime();
  var deadline = started + CFG.MAX_RUN_MS;
  var job = getJobState(channelKey);

  if (!job || job.status !== 'running') {
    clearResumeTriggers(channelKey);
    return null;
  }

  var keys = getApiKeys(channelKey);
  if (!keys) {
    writeProgress(SpreadsheetApp.getActiveSpreadsheet(), channelKey, {
      status: '오류',
      note: 'API 키 없음'
    });
    return null;
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shMeta = mustSheet(ss, CFG.SHEET_META);
  var outputSheetName = job.outputSheet;
  var shOut = mustOutputSheet(ss, outputSheetName);
  var meta = buildMetaIndex(shMeta);
  var headers = buildHeaders(keys.key, keys.secret);
  var candidates = loadJobCandidates(ss, channelKey);

  if (!candidates.length) {
    cancelChannelJobSilent(channelKey);
    return null;
  }

  if (job.total && job.total !== candidates.length) {
    PropertiesService.getScriptProperties().setProperty(jobProp(channelKey, 'TOTAL'), String(candidates.length));
  }

  var nextIdx = job.nextIdx;
  var failed = job.failed;
  var batchRows = [];
  var collectedBefore = Math.max(0, shOut.getLastRow() - 1);
  var deskBase = CHANNELS[channelKey].deskBase;
  var skips = loadSkipCounts(channelKey);

  while (nextIdx < candidates.length) {
    if (new Date().getTime() > deadline) break;

    var chat = candidates[nextIdx];
    nextIdx++;

    var result = null;
    try {
      result = fetchTranscriptAndInflow(chat.id, headers, job.startMs, job.endMs);
    } catch (fetchErr) {
      failed++;
      skips.fetchError++;
      continue;
    }
    if (!result || !result.transcript) {
      var reason = (result && result.skipReason) || 'autoClosed';
      if (skips[reason] === undefined) skips[reason] = 0;
      skips[reason]++;
      continue;
    }

    var phone = normalizePhone(chat.userMobile);
    var metaHit = phone ? meta.byPhone[phone] : null;

    batchRows.push([
      metaHit ? (metaHit.rpnTin || '') : '',
      metaHit ? (metaHit.userName || chat.userName || '') : (chat.userName || ''),
      metaHit ? (metaHit.bsno || '') : '',
      metaHit ? (metaHit.tnmNm || '') : '',
      metaHit ? (metaHit.careproMgr || '') : '',
      metaHit ? (metaHit.bkpStatus || '') : '',
      formatTs(result.firstMsgMs),
      deskBase + chat.id,
      result.inflow || '채널톡',
      (chat.tags || []).join(', '),
      truncateText(result.transcript, CFG.MAX_TRANSCRIPT_CHAR),
      result.rating !== null && result.rating !== undefined ? result.rating : '',
      result.reviewedAt ? formatTs(result.reviewedAt) : ''
    ]);
  }

  appendOutputRows(shOut, batchRows);
  saveJobProgress(channelKey, nextIdx, failed);
  saveSkipCounts(channelKey, skips);

  var collected = Math.max(0, shOut.getLastRow() - 1);
  var elapsedSec = Math.round((new Date().getTime() - started) / 1000);
  var done = nextIdx >= candidates.length;
  var pct = candidates.length ? Math.round((nextIdx / candidates.length) * 100) + '%' : '100%';
  var periodLabel = formatYmd(new Date(job.startMs)) + ' ~ ' + formatYmd(new Date(job.endMs));

  if (done) {
    clearResumeTriggers(channelKey);
    var finalSkips = skips;
    clearJobState(channelKey);
    deleteJobCandidatesSheet(ss, channelKey);
    writeProgress(ss, channelKey, {
      status: '완료',
      period: periodLabel,
      outputSheet: outputSheetName,
      candidates: candidates.length,
      scanned: nextIdx,
      collected: collected,
      failed: failed,
      nextIdx: nextIdx,
      progressPct: '100%',
      elapsedSec: elapsedSec,
      skips: finalSkips,
      note: '최종 실행 완료'
    });
    appendProgressHistory(ss, channelKey, {
      completedAt: new Date(),
      period: periodLabel,
      candidates: candidates.length,
      collected: collected,
      failed: failed,
      status: '완료'
    });
  } else {
    scheduleResumeTrigger(channelKey);
    writeProgress(ss, channelKey, {
      status: '이어서 대기',
      period: periodLabel,
      outputSheet: outputSheetName,
      candidates: candidates.length,
      scanned: nextIdx,
      collected: collected,
      failed: failed,
      nextIdx: nextIdx,
      progressPct: pct,
      elapsedSec: elapsedSec,
      skips: skips,
      note: '시간 제한 — 약 1분 후 자동 이어서 수집'
    });
  }

  return {
    done: done,
    nextIdx: nextIdx,
    total: candidates.length,
    collected: collected,
    failed: failed,
    elapsedSec: elapsedSec,
    outputSheet: outputSheetName
  };
}

function cancelChannelTalkJob() {
  Object.keys(CHANNELS).forEach(function(k) { cancelChannelJobSilent(k); });
  SpreadsheetApp.getUi().alert('진행 중인 수집을 모두 중지했습니다.');
}

function cancelChannelJobSilent(channelKey) {
  if (!channelKey) return;
  var job = getJobState(channelKey);
  clearResumeTriggers(channelKey);
  clearJobState(channelKey);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (job) {
    deleteJobCandidatesSheet(ss, channelKey);
    writeProgress(ss, channelKey, {
      status: '중지됨',
      note: '사용자 또는 새 작업으로 중지'
    });
  }
}

// backward compat wrapper
function cancelChannelTalkJobSilent() {
  Object.keys(CHANNELS).forEach(function(k) { cancelChannelJobSilent(k); });
}

/* =========================
 * Job 상태 (Script Properties)
 * ========================= */
function initJob(ss, channelKey, candidates, startMs, endMs, outputSheet) {
  clearResumeTriggers(channelKey);
  saveJobCandidates(ss, channelKey, candidates);
  var props = PropertiesService.getScriptProperties();
  props.setProperty(jobProp(channelKey, 'STATUS'), 'running');
  props.setProperty(jobProp(channelKey, 'START_MS'), String(startMs));
  props.setProperty(jobProp(channelKey, 'END_MS'), String(endMs));
  props.setProperty(jobProp(channelKey, 'NEXT_IDX'), '0');
  props.setProperty(jobProp(channelKey, 'TOTAL'), String(candidates.length));
  props.setProperty(jobProp(channelKey, 'FAILED'), '0');
  props.setProperty(jobProp(channelKey, 'STARTED_AT'), String(new Date().getTime()));
  props.setProperty(jobProp(channelKey, 'OUTPUT_SHEET'), sanitizeSheetName(outputSheet));
}

function saveJobProgress(channelKey, nextIdx, failed) {
  var props = PropertiesService.getScriptProperties();
  props.setProperty(jobProp(channelKey, 'NEXT_IDX'), String(nextIdx));
  props.setProperty(jobProp(channelKey, 'FAILED'), String(failed));
}

function loadSkipCounts(channelKey) {
  var raw = PropertiesService.getScriptProperties().getProperty(skipProp(channelKey));
  var out = { autoClosed: 0, outOfRange: 0, emptyText: 0, fetchError: 0 };
  if (!raw) return out;
  try {
    var parsed = JSON.parse(raw);
    Object.keys(parsed).forEach(function(k) { out[k] = Number(parsed[k]) || 0; });
  } catch (e) {}
  return out;
}

function saveSkipCounts(channelKey, skips) {
  PropertiesService.getScriptProperties().setProperty(skipProp(channelKey), JSON.stringify(skips));
}

function getJobState(channelKey) {
  if (!channelKey) return null;
  var props = PropertiesService.getScriptProperties();
  var status = props.getProperty(jobProp(channelKey, 'STATUS'));
  if (!status) return null;
  return {
    channel: channelKey,
    status: status,
    phase: props.getProperty(jobProp(channelKey, 'PHASE')) || '',
    startMs: Number(props.getProperty(jobProp(channelKey, 'START_MS')) || 0),
    endMs: Number(props.getProperty(jobProp(channelKey, 'END_MS')) || 0),
    nextIdx: Number(props.getProperty(jobProp(channelKey, 'NEXT_IDX')) || 0),
    total: Number(props.getProperty(jobProp(channelKey, 'TOTAL')) || 0),
    failed: Number(props.getProperty(jobProp(channelKey, 'FAILED')) || 0),
    startedAt: Number(props.getProperty(jobProp(channelKey, 'STARTED_AT')) || 0),
    outputSheet: props.getProperty(jobProp(channelKey, 'OUTPUT_SHEET')) || ''
  };
}

function clearJobState(channelKey) {
  if (!channelKey) return;
  var props = PropertiesService.getScriptProperties();
  ['STATUS', 'PHASE', 'START_MS', 'END_MS', 'NEXT_IDX', 'TOTAL', 'FAILED', 'STARTED_AT', 'OUTPUT_SHEET'].forEach(function(f) {
    props.deleteProperty(jobProp(channelKey, f));
  });
  props.deleteProperty(skipProp(channelKey));
}

function saveJobCandidates(ss, channelKey, candidates) {
  var sh = ensureSheet(ss, CFG.SHEET_JOB_PREFIX + channelKey);
  sh.clearContents();
  var rows = [['chatId', 'userName', 'userMobile', 'inflow', 'tags']];
  for (var i = 0; i < candidates.length; i++) {
    var c = candidates[i];
    rows.push([
      c.id,
      c.userName || '',
      c.userMobile || '',
      c.inflow || '',
      (c.tags || []).join(', ')
    ]);
  }
  writeSheetValues(sh, 1, 1, rows);
  try { sh.hideSheet(); } catch (hideErr) {}
}

function loadJobCandidates(ss, channelKey) {
  var sh = ss.getSheetByName(CFG.SHEET_JOB_PREFIX + channelKey);
  if (!sh) return [];
  var vals = sh.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < vals.length; i++) {
    var r = vals[i];
    if (!r[0]) continue;
    out.push({
      id: String(r[0]),
      userName: String(r[1] || ''),
      userMobile: String(r[2] || ''),
      inflow: String(r[3] || ''),
      tags: String(r[4] || '').split(',').map(function(t) { return t.trim(); }).filter(Boolean)
    });
  }
  return out;
}

function deleteJobCandidatesSheet(ss, channelKey) {
  var sh = ss.getSheetByName(CFG.SHEET_JOB_PREFIX + channelKey);
  if (sh) ss.deleteSheet(sh);
}

function scheduleResumeTrigger(channelKey) {
  clearResumeTriggers(channelKey);
  var handler = 'resumeTrigger_' + channelKey;
  ScriptApp.newTrigger(handler)
    .timeBased()
    .after(CFG.RESUME_TRIGGER_MS)
    .create();
}

function clearResumeTriggers(channelKey) {
  var target = channelKey ? ('resumeTrigger_' + channelKey) : null;
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    var fn = triggers[i].getHandlerFunction();
    if (target && fn === target) {
      ScriptApp.deleteTrigger(triggers[i]);
    } else if (!target && (fn === 'runChannelTalkExportResume' || /^resumeTrigger_/.test(fn))) {
      // 채널 미지정: 모든 채널 트리거 삭제 (backward compat)
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

function getApiKeys(channelKey) {
  var ch = CHANNELS[channelKey];
  if (!ch) return null;
  var props = PropertiesService.getScriptProperties();
  var key = props.getProperty(ch.propKey);
  var secret = props.getProperty(ch.propSecret);
  if (!key || !secret) return null;
  return { key: key, secret: secret };
}

/* =========================
 * ChannelTalk 수집
 * ========================= */
function fetchChatCandidates(headers, channelKey, startMs, endMs) {
  var byId = {};

  for (var si = 0; si < CFG.CHAT_STATES.length; si++) {
    var state = CFG.CHAT_STATES[si];
    var since = '';
    var guard = 0;

    while (guard++ < 1000) {
      var params = { state: state, sortOrder: 'desc', limit: CFG.PAGE_LIMIT };
      if (since) params.since = since;

      var data = ctGet('/user-chats', params, headers);
      var userChats = data.userChats || [];
      var users = data.users || [];
      if (!userChats.length) break;

      var userMap = {};
      for (var ui = 0; ui < users.length; ui++) {
        userMap[String(users[ui].id || '')] = users[ui];
      }

      var oldestUpdatedOnPage = Infinity;

      for (var ci = 0; ci < userChats.length; ci++) {
        var c = userChats[ci];
        var chatId = String(c.id || '');
        if (!chatId) continue;

        var updatedMs = Number(c.updatedAt || 0);
        var createdMs = Number(c.createdAt || 0);

        if (updatedMs && updatedMs < oldestUpdatedOnPage) oldestUpdatedOnPage = updatedMs;
        if (!updatedMs || updatedMs < startMs) continue;
        if (updatedMs > endMs) continue;

        var userId = String(c.userId || '');
        var u = userMap[userId] || {};
        var p = u.profile || {};

        byId[chatId] = {
          id: chatId,
          userName: String(u.name || p.name || '').trim(),
          userMobile: normalizePhone(p.mobileNumber || u.mobileNumber || ''),
          createdAtMs: createdMs,
          tags: c.tags || [],
          inflow: resolveInflow(u, c)
        };
      }

      var next = data.next;
      if (!next) break;
      if (oldestUpdatedOnPage < startMs) break;
      since = next;
    }
  }

  return Object.keys(byId).map(function(k) { return byId[k]; });
}

function fetchTranscriptAndInflow(chatId, headers, startMs, endMs) {
  var slice = fetchTranscriptSliceInRange(chatId, headers, startMs, endMs);
  if (!slice) return null;
  if (!slice.transcript) return slice; // skipReason 포함
  var detail = fetchUserChatDetail(chatId, headers);
  slice.inflow = detail.inflow;
  slice.rating = detail.rating;
  slice.reviewedAt = detail.reviewedAt;
  return slice;
}

function fetchUserChatInflow(chatId, headers) {
  return fetchUserChatDetail(chatId, headers).inflow;
}

function fetchUserChatDetail(chatId, headers) {
  try {
    var data = ctGet('/user-chats/' + chatId, {}, headers);
    var chat = data.userChat || data.chat || {};
    var user = pickUserFromApiDetail(data);
    var review = extractReview(chat);
    return {
      inflow: resolveInflow(user, chat),
      rating: review.rating,
      reviewedAt: review.reviewedAt
    };
  } catch (e) {
    return { inflow: '채널톡', rating: null, reviewedAt: null };
  }
}

/**
 * 채널톡 review 데이터 추출 (필드 위치가 API 버전에 따라 다를 수 있어 defensively 시도)
 */
function extractReview(chat) {
  if (!chat) return { rating: null, reviewedAt: null };
  var rating = null;
  var reviewedAt = null;

  if (chat.review && typeof chat.review === 'object') {
    rating = chat.review.rating || chat.review.score || chat.review.value || null;
    reviewedAt = chat.review.reviewedAt || chat.review.createdAt || null;
  }
  if (rating === null || rating === undefined) {
    rating = chat.rating || chat.reviewRating || chat.satisfaction || null;
  }
  if (reviewedAt === null || reviewedAt === undefined) {
    reviewedAt = chat.reviewedAt || chat.ratedAt || null;
  }
  return {
    rating: rating !== null && rating !== undefined ? Number(rating) : null,
    reviewedAt: reviewedAt ? Number(reviewedAt) : null
  };
}

/**
 * 디버그: 특정 chatId 의 raw detail (만족도 필드 확인)
 */
function api_debugChatDetail(channelKey, chatId) {
  var keys = getApiKeys(channelKey);
  if (!keys) return { error: 'API 키 없음' };
  try {
    var data = ctGet('/user-chats/' + chatId, {}, buildHeaders(keys.key, keys.secret));
    return {
      chatKeys: data.userChat ? Object.keys(data.userChat) : [],
      chat: data.userChat || data.chat || {}
    };
  } catch (e) {
    return { error: String(e.message || e) };
  }
}

/**
 * 만족도 데이터 진단: 최근 수집 시트에서 샘플 chatId 를 뽑아 API 응답 구조를 반환.
 * 프론트에서 이 결과를 보고 실제 필드명을 파악한 뒤 extractReview 를 조정합니다.
 */
function api_diagnoseRating(channelKey, specificChatId) {
  if (!CHANNELS[channelKey]) return { ok: false, error: '알 수 없는 채널' };
  var keys = getApiKeys(channelKey);
  if (!keys) return { ok: false, error: 'API 키 없음' };

  var chatIds = [];
  var sheetName = '';

  if (specificChatId) {
    chatIds = [String(specificChatId).trim()];
  } else {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var suffix = '_' + CHANNELS[channelKey].label;
    var sheets = ss.getSheets().filter(function(s) {
      var name = s.getName();
      return /^\d{8}~\d{8}_/.test(name) && name.substr(name.length - suffix.length) === suffix;
    });
    if (!sheets.length) return { ok: false, error: '수집 시트가 없습니다.' };
    var sh = sheets[0];
    sheetName = sh.getName();
    var last = sh.getLastRow();
    if (last < 2) return { ok: false, error: '시트에 데이터 없음' };
    var sampleCount = Math.min(3, last - 1);
    var vals = sh.getRange(2, 1, sampleCount, CFG.OUT_HEADERS.length).getValues();
    for (var i = 0; i < vals.length; i++) {
      var link = String(vals[i][7] || '');
      var m = link.match(/user-chats\/([^/?#]+)/);
      if (m) chatIds.push(m[1]);
    }
  }

  var headers = buildHeaders(keys.key, keys.secret);
  var samples = [];

  chatIds.forEach(function(chatId) {
    var info = { chatId: chatId };
    try {
      var chatData = ctGet('/user-chats/' + chatId, {}, headers);
      var chat = chatData.userChat || chatData.chat || {};
      info.chatKeys = Object.keys(chat);
      info.reviewLikeFields = pickReviewFields(chat);
    } catch (e) { info.chatError = String(e.message || e); }

    // 1) 전용 리뷰 엔드포인트 시도
    var reviewEndpoints = [
      '/user-chats/' + chatId + '/reviews',
      '/user-chats/' + chatId + '/review',
      '/reviews?userChatId=' + chatId
    ];
    info.endpoints = {};
    reviewEndpoints.forEach(function(ep) {
      try {
        var r = ctGet(ep.replace(/\?.*$/, ''), ep.indexOf('?') >= 0 ? { userChatId: chatId } : {}, headers);
        info.endpoints[ep] = { ok: true, keys: Object.keys(r), sample: r };
      } catch (e) {
        info.endpoints[ep] = { ok: false, error: String(e.message || e).slice(0, 200) };
      }
    });

    // 2) 메시지 스트림에서 리뷰/점수 관련 검색
    try {
      var msgData = ctGet('/user-chats/' + chatId + '/messages', { limit: 100, sortOrder: 'desc' }, headers);
      var msgs = msgData.messages || [];
      var typeCount = {};
      msgs.forEach(function(m) {
        var t = String(m.type || 'text');
        typeCount[t] = (typeCount[t] || 0) + 1;
      });
      info.messageTypes = typeCount;
      info.messageTotalOnPage = msgs.length;

      // review 로 보이는 메시지
      var reviewLike = msgs.filter(function(m) {
        var t = String(m.type || '').toLowerCase();
        var pt = String(m.plainText || '');
        if (/review|rating|score|feedback/i.test(t)) return true;
        if (/만족도|평점|별점|리뷰/.test(pt)) return true;
        // blocks 안 살펴보기
        var blocks = m.blocks || [];
        for (var j = 0; j < blocks.length; j++) {
          var bt = String(blocks[j].type || '').toLowerCase();
          if (/review|rating|score/.test(bt)) return true;
        }
        return false;
      });
      info.reviewMessageSamples = reviewLike.slice(0, 3).map(function(m) {
        return {
          id: m.id,
          type: m.type,
          plainText: (m.plainText || '').slice(0, 200),
          keys: Object.keys(m),
          personType: m.personType,
          createdAt: m.createdAt,
          blocks: (m.blocks || []).slice(0, 3)
        };
      });
    } catch (e) { info.messagesError = String(e.message || e); }

    samples.push(info);
  });

  return {
    ok: true,
    channel: channelKey,
    sheetName: sheetName,
    specificChatId: specificChatId || '',
    sampleCount: samples.length,
    samples: samples,
    note: '① endpoints 에서 ok:true 로 나오는 URL이 있으면 그 응답 구조를 확인 · ② messageTypes 에 review/rating 관련 type이 있으면 메시지로 처리됨 · ③ reviewMessageSamples 에 실제 리뷰 메시지 예시'
  };
}

function pickReviewFields(chat) {
  if (!chat || typeof chat !== 'object') return {};
  var out = {};
  Object.keys(chat).forEach(function(k) {
    if (/review|rating|score|satisf|feedback|ask/i.test(k)) {
      out[k] = chat[k];
    }
  });
  return out;
}

/**
 * 특정 chatId 의 form/session/wam/action 등 여러 잠재 엔드포인트 탐색
 * → 워크플로우 form 응답이 어디에 저장되는지 찾음
 */
function api_diagnoseFormSource(channelKey, chatId) {
  if (!CHANNELS[channelKey]) return { ok: false, error: '알 수 없는 채널' };
  var keys = getApiKeys(channelKey);
  if (!keys) return { ok: false, error: 'API 키 없음' };
  if (!chatId) return { ok: false, error: 'chatId 입력' };

  var headers = buildHeaders(keys.key, keys.secret);
  var base = '/user-chats/' + chatId;

  // 병렬로 여러 엔드포인트 시도 (UrlFetchApp.fetchAll, 재시도 없이 1회)
  var endpointDefs = [
    { key: base + '/wams', url: CFG.CHANNEL_BASE + base + '/wams?limit=10' },
    { key: base + '/wam-messages', url: CFG.CHANNEL_BASE + base + '/wam-messages?limit=10' },
    { key: base + '/session-messages', url: CFG.CHANNEL_BASE + base + '/session-messages?limit=10' },
    { key: base + '/actions', url: CFG.CHANNEL_BASE + base + '/actions?limit=10' },
    { key: base + '/action-messages', url: CFG.CHANNEL_BASE + base + '/action-messages?limit=10' },
    { key: base + '/sessions', url: CFG.CHANNEL_BASE + base + '/sessions?limit=10' },
    { key: base + '/task-messages', url: CFG.CHANNEL_BASE + base + '/task-messages?limit=10' },
    { key: base + '/reviews', url: CFG.CHANNEL_BASE + base + '/reviews?limit=10' },
    { key: '/workflow-answers?userChatId=' + chatId, url: CFG.CHANNEL_BASE + '/workflow-answers?userChatId=' + chatId },
    { key: '/user-chat-sessions?userChatId=' + chatId, url: CFG.CHANNEL_BASE + '/user-chat-sessions?userChatId=' + chatId },
    { key: '/wams?userChatId=' + chatId, url: CFG.CHANNEL_BASE + '/wams?userChatId=' + chatId }
  ];

  var requests = endpointDefs.map(function(e) {
    return { url: e.url, method: 'get', headers: headers, muteHttpExceptions: true };
  });

  var results = {};
  try {
    var responses = UrlFetchApp.fetchAll(requests);
    responses.forEach(function(res, i) {
      var code = res.getResponseCode();
      var body = res.getContentText() || '';
      var key = endpointDefs[i].key;
      if (code >= 200 && code < 300) {
        try {
          var parsed = JSON.parse(body || '{}');
          results[key] = {
            ok: true,
            topKeys: Object.keys(parsed),
            preview: JSON.stringify(parsed).slice(0, 1200)
          };
        } catch (e) {
          results[key] = { ok: true, topKeys: [], preview: body.slice(0, 400) };
        }
      } else {
        results[key] = { ok: false, error: 'HTTP ' + code + ': ' + body.slice(0, 100) };
      }
    });
  } catch (e) {
    return { ok: false, error: '병렬 fetch 실패: ' + String(e.message || e) };
  }

  // 메시지 스캔 - text 뿐 아니라 log 필드도 검사
  var fullTextFound = null;
  var totalScanned = 0;
  var since = '';
  var logActions = {}; // 발견된 log.action 종류
  var workflowMsgs = []; // triggerId=820003 관련 메시지들
  var msgsWithLog = 0;

  for (var p = 0; p < 3; p++) {
    var params2 = { limit: 100, sortOrder: 'desc' };
    if (since) params2.since = since;
    try {
      var data = ctGet(base + '/messages', params2, headers);
      var msgs = data.messages || [];
      totalScanned += msgs.length;
      for (var mi = 0; mi < msgs.length; mi++) {
        var m = msgs[mi];

        // log 필드 분석
        if (m.log && typeof m.log === 'object') {
          msgsWithLog++;
          var action = String(m.log.action || 'unknown');
          if (!logActions[action]) logActions[action] = { count: 0, samples: [] };
          logActions[action].count++;
          if (logActions[action].samples.length < 2) {
            logActions[action].samples.push({
              messageId: m.id,
              personType: m.personType,
              options: m.options,
              log: m.log,
              plainText: (m.plainText || '').slice(0, 100)
            });
          }
          // 워크플로우 820003 관련
          if (m.log.triggerId === '820003' || m.log.triggerId === 820003) {
            workflowMsgs.push({
              messageId: m.id,
              personType: m.personType,
              options: m.options,
              log: m.log,
              plainText: (m.plainText || '').slice(0, 200),
              blocks: m.blocks,
              createdAt: m.createdAt,
              allKeys: Object.keys(m)
            });
          }
        }

        // 텍스트에서 상담 만족도 검색
        var txt = String(m.plainText || '');
        (m.blocks || []).forEach(function(b) { if (b.value) txt += '\n' + b.value; });
        if (!fullTextFound && /상담\s*만족도\s*점수/.test(txt)) {
          fullTextFound = {
            page: p + 1,
            messageId: m.id,
            personType: m.personType,
            type: m.type,
            createdAt: m.createdAt,
            plainText: (m.plainText || '').slice(0, 500),
            blocks: m.blocks,
            allKeys: Object.keys(m)
          };
        }
      }
      if (!data.next) break;
      since = data.next;
    } catch (e) { break; }
  }

  return {
    ok: true,
    chatId: chatId,
    endpoints: results,
    fullMessagesScanned: totalScanned,
    surveyFormFoundInMessages: fullTextFound,
    logAnalysis: {
      messagesWithLog: msgsWithLog,
      actionsFound: logActions,
      workflow820003Messages: workflowMsgs
    },
    note: workflowMsgs.length
      ? '✅ workflow-820003 메시지 ' + workflowMsgs.length + '건 발견! 아래 log 데이터 확인.'
      : (fullTextFound ? '✅ 메시지에서 survey form 텍스트 발견' : '❌ 텍스트/워크플로우 미발견')
  };
}

/**
 * 그룹/팀챗 API 엔드포인트 탐색 - 만족도 그룹의 메시지에서 chatId 추출용
 */
function api_diagnoseGroup(channelKey, groupIdentifier) {
  if (!CHANNELS[channelKey]) return { ok: false, error: '알 수 없는 채널' };
  var keys = getApiKeys(channelKey);
  if (!keys) return { ok: false, error: 'API 키 없음' };

  var gid = String(groupIdentifier || '').trim();
  if (!gid) return { ok: false, error: 'group ID 를 입력하세요 (예: <GROUP_ID> 또는 <GROUP_LABEL>-<GROUP_ID>)' };

  var headers = buildHeaders(keys.key, keys.secret);
  var attempts = [
    { path: '/groups', params: {} },
    { path: '/groups/' + gid, params: {} },
    { path: '/groups/' + gid + '/messages', params: { limit: 10, sortOrder: 'desc' } },
    { path: '/team-chats/' + gid, params: {} },
    { path: '/team-chats/' + gid + '/messages', params: { limit: 10, sortOrder: 'desc' } }
  ];

  var results = {};
  attempts.forEach(function(a) {
    try {
      var r = ctGet(a.path, a.params, headers);
      results[a.path] = {
        ok: true,
        topKeys: Object.keys(r),
        preview: JSON.stringify(r).slice(0, 600)
      };
    } catch (e) {
      results[a.path] = { ok: false, error: String(e.message || e).slice(0, 200) };
    }
  });

  // 성공한 엔드포인트에서 chatId 정규식으로 추출 시도
  var extractedChatIds = [];
  Object.keys(results).forEach(function(path) {
    var r = results[path];
    if (r.ok && r.preview) {
      var matches = r.preview.match(/[a-f0-9]{20,}/g) || [];
      var chatIdCandidates = matches.filter(function(m) { return m.length >= 20 && m.length <= 40; });
      if (chatIdCandidates.length) {
        extractedChatIds.push({ from: path, ids: Array.from(new Set(chatIdCandidates)).slice(0, 5) });
      }
    }
  });

  return {
    ok: true,
    channel: channelKey,
    groupIdentifier: gid,
    endpoints: results,
    extractedChatIds: extractedChatIds,
    note: '✅ 표시된 path 가 유효 · preview 안에 user-chats 관련 chatId가 있으면 그걸로 만족도 응답자 확인 가능'
  };
}

function pickUserFromApiDetail(data) {
  if (!data) return {};
  if (data.user && typeof data.user === 'object' && !Array.isArray(data.user)) return data.user;
  if (data.users && data.users.length) return data.users[0];
  return {};
}

function fetchTranscriptSliceInRange(chatId, headers, startMs, endMs) {
  var since = '';
  var guard = 0;
  var inPeriod = [];
  var totalFetched = 0;
  var hadInRangeMsg = false;

  while (guard++ < 200) {
    var params = { sortOrder: 'desc', limit: CFG.MESSAGE_PAGE_LIMIT };
    if (since) params.since = since;

    var data = ctGet('/user-chats/' + chatId + '/messages', params, headers);
    var msgs = data.messages || [];
    if (!msgs.length) break;
    totalFetched += msgs.length;

    var oldestInPage = Infinity;

    for (var i = 0; i < msgs.length; i++) {
      var m = msgs[i];
      var createdMs = Number(m.createdAt || 0);
      if (!createdMs) continue;
      if (createdMs < oldestInPage) oldestInPage = createdMs;
      if (createdMs > endMs) continue;
      if (createdMs < startMs) continue;

      hadInRangeMsg = true;
      var text = extractMessageText(m);
      if (!text) continue;

      var vis = getVisibilityLabel(m);
      inPeriod.push({
        ms: createdMs,
        vis: vis,
        line: formatTranscriptLine(createdMs, vis, text)
      });
    }

    if (oldestInPage < startMs) break;
    var next = data.next;
    if (!next) break;
    since = next;
  }

  if (!inPeriod.length) {
    var reason = hadInRangeMsg ? 'emptyText' : (totalFetched > 0 ? 'outOfRange' : 'autoClosed');
    return { transcript: null, skipReason: reason };
  }

  var firstMsgMs = inPeriod[0].ms;
  for (var a = 0; a < inPeriod.length; a++) {
    if (inPeriod[a].ms < firstMsgMs) firstMsgMs = inPeriod[a].ms;
  }

  inPeriod.sort(function(x, y) { return x.ms - y.ms; });
  var out = [];
  for (var c = 0; c < inPeriod.length; c++) out.push(inPeriod[c].line);
  return { transcript: out.join('\n'), firstMsgMs: firstMsgMs };
}

function formatTranscriptLine(createdMs, vis, text) {
  return '[' + formatTs(createdMs) + '] [' + vis + '] ' + text;
}

function getVisibilityLabel(msg) {
  var text = String((msg && msg.plainText) || '').trim();
  var pType = String((msg && msg.personType) || '').toLowerCase();

  if (msg && (msg.isPrivate === true || String(msg.isPrivate || '').toLowerCase() === 'true')) return '내부';
  if (msg && (msg.isPrivate === false || String(msg.isPrivate || '').toLowerCase() === 'false')) return '외부';

  var internalPattern =
    /^\/[a-z]/i.test(text) ||
    /^@/.test(text) ||
    /AI 추천 답변|복사하여 답변|고객이 기다리고 있어요|답변을 받지 못했어요|Deal 생성 완료/i.test(text) ||
    /해지 진행 부탁|수임해지|정기결제 해지|내부/i.test(text);

  if (internalPattern) return '내부';

  if (pType === 'user') return '외부';
  if (pType === 'manager') return '외부';
  if (pType === 'bot') return '내부';

  return '외부';
}

function pickInflowCandidates(user, chat) {
  var out = [];
  var objs = [chat, user];
  if (chat && chat.contactMedium && typeof chat.contactMedium === 'object') objs.push(chat.contactMedium);
  if (user && user.contactMedium && typeof user.contactMedium === 'object') objs.push(user.contactMedium);

  var keys = [
    'mediumType', 'medium_type',
    'mediumName', 'medium_name',
    'contactMediumType', 'contact_medium_type',
    'contactMediumName', 'contact_medium_name',
    'source', 'chatType', 'chat_type',
    'type', 'name'
  ];

  for (var o = 0; o < objs.length; o++) {
    var obj = objs[o];
    if (!obj || typeof obj !== 'object') continue;
    for (var k = 0; k < keys.length; k++) {
      var v = String(obj[keys[k]] || '').trim();
      if (v) out.push(v);
    }
    collectMediumStrings(obj, 0, out);
  }
  return out;
}

function collectMediumStrings(obj, depth, out) {
  if (!obj || depth > 4 || typeof obj !== 'object') return;
  for (var k in obj) {
    if (!obj.hasOwnProperty(k)) continue;
    var v = obj[k];
    if (typeof v === 'string') {
      if (/medium|contact|messenger|source|channel/i.test(k)) out.push(v);
    } else if (v && typeof v === 'object') {
      collectMediumStrings(v, depth + 1, out);
    }
  }
}

function normalizeMediumToken(raw) {
  return String(raw || '').trim().toLowerCase().replace(/[\s_\-]/g, '');
}

function mapMediumPairToLabel(typeRaw, nameRaw) {
  var type = normalizeMediumToken(typeRaw);
  var name = normalizeMediumToken(nameRaw);
  if (type === 'phone' || name.indexOf('phone') >= 0 || name.indexOf('meet') >= 0) return '미트';
  if (type === 'native') return '채널톡';
  if (type === 'app' || name.indexOf('kakao') >= 0 || name.indexOf('appkakao') >= 0) return '카카오';
  if (name.indexOf('native') >= 0) return '채널톡';
  return '';
}

function mapMediumToInflowLabel(raw) {
  var s = normalizeMediumToken(raw);
  if (!s) return '';
  if (s === 'phone' || s === 'phonenumber' || s.indexOf('meet') >= 0 || s === 'voice') return '미트';
  if (s === 'app' || s === 'appkakao' || s.indexOf('kakao') >= 0) return '카카오';
  if (s === 'native' || s === 'channeltalk' || s === 'channel') return '채널톡';
  return '';
}

function resolveInflow(user, chat) {
  if (chat) {
    var pairLabel = mapMediumPairToLabel(
      chat.mediumType || chat.medium_type || chat.contactMediumType || chat.contact_medium_type,
      chat.mediumName || chat.medium_name || chat.contactMediumName || chat.contact_medium_name
    );
    if (pairLabel) return pairLabel;
  }
  var candidates = pickInflowCandidates(user, chat);
  for (var i = 0; i < candidates.length; i++) {
    var label = mapMediumToInflowLabel(candidates[i]);
    if (label) return label;
  }
  return '채널톡';
}

/* =========================
 * 메타베이스
 * ========================= */
function buildMetaIndex(shMeta) {
  var vals = shMeta.getDataRange().getValues();
  if (vals.length < 2) return { byPhone: {} };

  var h = vals[0].map(function(v) { return String(v || '').trim(); });
  var idx = {
    rpnTin: h.indexOf(CFG.META_HEADERS.RPN_TIN),
    phone: h.indexOf(CFG.META_HEADERS.CONTACT_PHONE),
    userName: h.indexOf(CFG.META_HEADERS.USER_NAME),
    bsno: h.indexOf(CFG.META_HEADERS.BSNO),
    bkpStatus: h.indexOf(CFG.META_HEADERS.BKP_STATUS),
    tnmNm: h.indexOf(CFG.META_HEADERS.TNM_NM),
    careproMgr: h.indexOf(CFG.META_HEADERS.CAREPRO_MGR)
  };

  var missing = [];
  if (idx.rpnTin < 0) missing.push(CFG.META_HEADERS.RPN_TIN);
  if (idx.phone < 0) missing.push(CFG.META_HEADERS.CONTACT_PHONE);
  if (idx.bsno < 0) missing.push(CFG.META_HEADERS.BSNO);
  if (idx.bkpStatus < 0) missing.push(CFG.META_HEADERS.BKP_STATUS);
  if (idx.tnmNm < 0) missing.push(CFG.META_HEADERS.TNM_NM);
  if (idx.careproMgr < 0) missing.push(CFG.META_HEADERS.CAREPRO_MGR);
  if (missing.length) {
    throw new Error('메타베이스 헤더를 찾지 못했습니다: ' + missing.join(', '));
  }

  var byPhone = {};
  for (var i = 1; i < vals.length; i++) {
    var row = vals[i];
    var phone = normalizePhone(row[idx.phone]);
    if (!phone || byPhone[phone]) continue;
    byPhone[phone] = {
      rpnTin: String(row[idx.rpnTin] || '').trim(),
      userName: idx.userName >= 0 ? String(row[idx.userName] || '').trim() : '',
      bsno: String(row[idx.bsno] || '').trim(),
      bkpStatus: String(row[idx.bkpStatus] || '').trim(),
      tnmNm: String(row[idx.tnmNm] || '').trim(),
      careproMgr: String(row[idx.careproMgr] || '').trim()
    };
  }
  return { byPhone: byPhone };
}

/* =========================
 * 출력 / 진행상황
 * ========================= */
function writeProgress(ss, channelKey, info) {
  var sh = ensureSheet(ss, CFG.SHEET_PROGRESS_PREFIX + channelKey);
  var savedHistory = extractProgressHistoryRows(sh);
  sh.clearContents();

  var chLabel = CHANNELS[channelKey] ? CHANNELS[channelKey].label : channelKey;
  var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  var rows = [
    ['채널', chLabel],
    ['상태', info.status || ''],
    ['기간', info.period || ''],
    ['출력 탭', info.outputSheet || ''],
    ['후보 대화방', info.candidates !== undefined ? info.candidates : ''],
    ['처리한 대화방', info.scanned !== undefined ? info.scanned : ''],
    ['다음 시작 인덱스', info.nextIdx !== undefined ? info.nextIdx : ''],
    ['진행률', info.progressPct || ''],
    ['확정 수집', info.collected !== undefined ? info.collected : ''],
    ['전사 실패(건너뜀)', info.failed !== undefined ? info.failed : ''],
    ['이번 실행 소요(초)', info.elapsedSec !== undefined ? info.elapsedSec : '']
  ];

  var skips = info.skips || null;
  if (skips) {
    rows.push(['스킵 - ' + SKIP_LABELS.autoClosed, skips.autoClosed || 0]);
    rows.push(['스킵 - ' + SKIP_LABELS.outOfRange, skips.outOfRange || 0]);
    rows.push(['스킵 - ' + SKIP_LABELS.emptyText, skips.emptyText || 0]);
    rows.push(['스킵 - ' + SKIP_LABELS.fetchError, skips.fetchError || 0]);
  }

  rows.push(['비고', info.note || '']);
  rows.push(['갱신시각', now]);

  writeSheetValues(sh, 1, 1, rows);
  if (savedHistory && savedHistory.length) {
    writeSheetValues(sh, rows.length + 2, 1, savedHistory);
  }
}

function extractProgressHistoryRows(sh) {
  var last = sh.getLastRow();
  var lastCol = Math.max(sh.getLastColumn(), 6);
  if (last < 1) return [];

  var vals = readSheetValues(sh, 1, 1, last, lastCol);
  var histIdx = -1;
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0] || '') === '[실행 완료 이력]') {
      histIdx = i;
      break;
    }
  }
  if (histIdx < 0) return [];

  var out = [];
  for (var r = histIdx; r < vals.length; r++) {
    out.push(vals[r].slice(0, 6));
  }
  while (out.length && isBlankSheetRow(out[out.length - 1])) {
    out.pop();
  }
  return out;
}

function appendProgressHistory(ss, channelKey, info) {
  var sh = ensureSheet(ss, CFG.SHEET_PROGRESS_PREFIX + channelKey);
  var completedAt = Utilities.formatDate(
    info.completedAt || new Date(),
    Session.getScriptTimeZone(),
    'yyyy-MM-dd HH:mm:ss'
  );
  var dataRow = [[
    completedAt,
    info.period || '',
    info.candidates || 0,
    info.collected || 0,
    info.failed || 0,
    info.status || '완료'
  ]];

  var histHeaderRow = findProgressHistoryHeaderRow(sh);
  if (histHeaderRow < 0) {
    var startRow = sh.getLastRow() + 2;
    writeSheetValues(sh, startRow, 1, [['[실행 완료 이력]', '']]);
    writeSheetValues(sh, startRow + 1, 1, [['완료시각', '기간', '후보', '확정수집', '실패', '상태']]);
    writeSheetValues(sh, startRow + 2, 1, dataRow);
  } else {
    var appendAt = sh.getLastRow() + 1;
    writeSheetValues(sh, appendAt, 1, dataRow);
  }
}

function findProgressHistoryHeaderRow(sh) {
  var vals = readSheetValues(sh, 1, 1, Math.max(sh.getLastRow(), 1), 1);
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0] || '') === '[실행 완료 이력]') {
      return i + 1;
    }
  }
  return -1;
}

function buildPeriodSheetName(startMs, endMs, channelKey) {
  var chLabel = CHANNELS[channelKey] ? CHANNELS[channelKey].label : channelKey;
  return sanitizeSheetName(formatYmd(new Date(startMs)) + '~' + formatYmd(new Date(endMs)) + '_' + chLabel);
}

function sanitizeSheetName(name) {
  var s = String(name || '').trim();
  if (!s) return '전사문';
  return s.replace(/[\[\]\:\*\?\/\\]/g, '').substring(0, 100);
}

function initOutputSheet(ss, sheetName) {
  return ensureOutputSheet(ss, sheetName, true);
}

function ensureOutputSheet(ss, sheetName, reset) {
  var safeName = sanitizeSheetName(sheetName);
  var sh = ss.getSheetByName(safeName);
  if (sh && reset) {
    sh.clearContents();
  } else if (!sh) {
    sh = ss.insertSheet(safeName);
  }
  writeSheetValues(sh, 1, 1, [CFG.OUT_HEADERS]);
  return sh;
}

function mustOutputSheet(ss, sheetName) {
  var sh = ss.getSheetByName(sheetName);
  if (!sh) sh = ensureOutputSheet(ss, sheetName, false);
  return sh;
}

function appendOutputRows(shOut, rows) {
  if (!rows || !rows.length) return;
  var start = shOut.getLastRow() + 1;
  writeSheetValues(shOut, start, 1, rows);
}

/* =========================
 * API / 유틸
 * ========================= */
function buildHeaders(key, secret) {
  return {
    'x-access-key': key,
    'x-access-secret': secret,
    'Accept': 'application/json'
  };
}

function ctGet(path, params, headers) {
  var q = toQuery(params || {});
  var url = CFG.CHANNEL_BASE + path + (q ? ('?' + q) : '');
  var lastErr = '';

  for (var i = 0; i < CFG.API_RETRY; i++) {
    try {
      var res = UrlFetchApp.fetch(url, {
        method: 'get',
        headers: headers,
        muteHttpExceptions: true,
        followRedirects: true,
        validateHttpsCertificates: true
      });
      if (CFG.API_SLEEP_MS > 0) Utilities.sleep(CFG.API_SLEEP_MS);

      var code = res.getResponseCode();
      var body = res.getContentText() || '';

      if (code === 429) { Utilities.sleep(Math.min(30000, Math.pow(2, i) * 1000)); continue; }
      if (code >= 500) { Utilities.sleep(Math.min(12000, Math.pow(2, i) * 700)); continue; }
      if (code >= 400) throw new Error('API 오류 ' + code + ' ' + path + ': ' + body.slice(0, 300));

      try {
        return JSON.parse(body || '{}');
      } catch (parseErr) {
        if (i === CFG.API_RETRY - 1) throw new Error('JSON 파싱 실패 ' + path);
        Utilities.sleep(Math.min(12000, Math.pow(2, i) * 700));
      }
    } catch (fetchErr) {
      lastErr = String(fetchErr.message || fetchErr);
      if (i === CFG.API_RETRY - 1) throw new Error('API 연결 실패 ' + path + ': ' + lastErr);
      Utilities.sleep(Math.min(15000, Math.pow(2, i) * 1000));
    }
  }
  throw new Error('API 재시도 실패: ' + path + (lastErr ? (' / ' + lastErr) : ''));
}

function extractMessageText(msg) {
  var plain = String(msg.plainText || '').trim();
  if (plain) return plain;
  var blocks = msg.blocks || [];
  var parts = [];
  for (var i = 0; i < blocks.length; i++) {
    var b = blocks[i];
    if (!b || typeof b !== 'object') continue;
    if ((b.type === 'text' || b.type === 'code') && b.value) parts.push(String(b.value));
  }
  return parts.join('\n').trim();
}

function parseYmd(input) {
  var s = String(input || '').trim().replace(/[^\d]/g, '');
  if (!/^\d{8}$/.test(s)) return null;
  var y = Number(s.substring(0, 4));
  var m = Number(s.substring(4, 6));
  var d = Number(s.substring(6, 8));
  var dt = new Date(y, m - 1, d, 0, 0, 0, 0);
  if (dt.getFullYear() !== y || dt.getMonth() !== m - 1 || dt.getDate() !== d) return null;
  return dt;
}

function formatYmd(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyyMMdd');
}

function endOfDay(d) {
  var x = new Date(d.getTime());
  x.setHours(23, 59, 59, 999);
  return x;
}

function formatTs(msOrDate) {
  var ms = Number(msOrDate || 0);
  if (!ms) return '';
  return Utilities.formatDate(new Date(ms), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

function normalizePhone(v) {
  if (v === null || v === undefined) return '';
  var s = String(v).replace(/\.0$/, '').replace(/\D/g, '');
  if (!s) return '';
  if (s.indexOf('82') === 0) s = '0' + s.substring(2);
  if (s.indexOf('10') === 0) s = '0' + s;
  return s;
}

function truncateText(s, maxLen) {
  var t = String(s || '');
  if (t.length <= maxLen) return t;
  return t.slice(0, maxLen - 20) + '\n...[truncated]';
}

var SHEET_WRITE_CHUNK = 400;

function normalizeRowsForSheet(values) {
  var rows = [];
  var numCols = 0;
  for (var i = 0; i < (values || []).length; i++) {
    var row = values[i];
    if (!row || !row.length) row = [''];
    rows.push(row);
    if (row.length > numCols) numCols = row.length;
  }
  for (var j = 0; j < rows.length; j++) {
    if (rows[j].length < numCols) {
      var padded = rows[j].slice();
      while (padded.length < numCols) padded.push('');
      rows[j] = padded;
    }
  }
  return { rows: rows, numRows: rows.length, numCols: numCols };
}

function writeSheetValues(sheet, startRow, startCol, values) {
  var block = normalizeRowsForSheet(values);
  if (!block.numRows) return;
  for (var offset = 0; offset < block.numRows; offset += SHEET_WRITE_CHUNK) {
    var chunk = block.rows.slice(offset, offset + SHEET_WRITE_CHUNK);
    var rowStart = startRow + offset;
    var rowEnd = rowStart + chunk.length - 1;
    var colEnd = startCol + block.numCols - 1;
    var a1 = columnToLetter(startCol) + rowStart + ':' + columnToLetter(colEnd) + rowEnd;
    sheet.getRange(a1).setValues(chunk);
  }
}

function readSheetValues(sheet, startRow, startCol, numRows, numCols) {
  if (numRows <= 0 || numCols <= 0) return [];
  var rowEnd = startRow + numRows - 1;
  var colEnd = startCol + numCols - 1;
  var a1 = columnToLetter(startCol) + startRow + ':' + columnToLetter(colEnd) + rowEnd;
  return sheet.getRange(a1).getValues();
}

function columnToLetter(col) {
  var n = Number(col || 0);
  var letter = '';
  while (n > 0) {
    var mod = (n - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    n = Math.floor((n - 1) / 26);
  }
  return letter;
}

function isBlankSheetRow(row) {
  if (!row || !row.length) return true;
  for (var i = 0; i < row.length; i++) {
    if (String(row[i] || '').trim() !== '') return false;
  }
  return true;
}

function mustSheet(ss, name) {
  var sh = ss.getSheetByName(name);
  if (!sh) throw new Error('"' + name + '" 시트가 없습니다. 먼저 탭을 만들어 주세요.');
  return sh;
}

function ensureSheet(ss, name) {
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

function toQuery(obj) {
  var keys = Object.keys(obj || {});
  var parts = [];
  for (var i = 0; i < keys.length; i++) {
    var k = keys[i];
    var v = obj[k];
    if (v === '' || v === null || v === undefined) continue;
    parts.push(encodeURIComponent(k) + '=' + encodeURIComponent(v));
  }
  return parts.join('&');
}
