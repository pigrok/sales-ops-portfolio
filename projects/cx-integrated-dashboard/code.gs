/*************************************************
 * 채널톡 통합 수집 대시보드 (환급 / 케어)
 * - 웹앱으로 배포하여 여러 명이 사용
 * - 채널 선택 → 기간 지정 → 수집 실행
 * - 진행상황 실시간 표시
 *************************************************/

// 외부 해지 워크플로 시트 (케어 채널 실제 해지 완료 기록)
// 컬럼: 타임스탬프,고객명,RPN TIN,사업자번호,복수사업자번호,성사여부,채널톡링크,해지요청일자,해지처리일자,첫결제일,해지요청사유,사유상세,기장담당자,요청자
var CANCEL_WF_SHEET_ID = '<CANCEL_WF_SHEET_ID>';
var CANCEL_WF_TAB = ''; // 빈 문자열이면 첫 시트 사용

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
    RPN_TIN: 'Bman Sbcbe History - Bman Tin → Rpn Tin',
    CONTACT_PHONE: 'User → Contact Phone',
    USER_NAME: 'User → Name',
    BSNO: 'Bsno',
    BKP_STATUS: 'Bkp Status',
    TNM_NM: 'Tnm Nm',
    CAREPRO_MGR: 'Carepro Mngr Infr → Name',
    CAREPRO_MGR_EMAIL: 'Carepro Mngr Infr → Email'
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
    '만족도 응답시각',
    '첫응대자',
    '최종응대자'
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

// 스냅샷: 수집 완료 시점의 대시보드 + 만족도 응답자 데이터를 JSON으로 저장 → raw 시트 삭제해도 뷰 완전 유지
var SNAPSHOT_SHEET = '_스냅샷_대시보드';
var SNAPSHOT_HEADERS = [
  'snapshotId', 'channelKey', 'sheetName', 'period', 'collectedAt', 'rawSheetState',
  'dashboardJson1', 'dashboardJson2', 'dashboardJson3',
  'satisfactionJson1', 'satisfactionJson2', 'satisfactionJson3'
];
var SNAPSHOT_JSON_COL_LIMIT = 45000; // Google Sheets 셀당 문자 한도(50000)에 여유
var SNAPSHOT_DASHBOARD_COL_START = 7; // dashboardJson1 위치
var SNAPSHOT_SATISFACTION_COL_START = 10; // satisfactionJson1 위치

function doGet(e) {
  return HtmlService.createTemplateFromFile('dashboard')
    .evaluate()
    .setTitle('CX 케어·환급 통합 대시보드')
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
  // 이벤트 단위 재설계: chatId → 웹훅 응답 배열 (receivedAtMs 포함) · 프런트에서 시간 근접 매칭
  var byChatId = {};
  var tz = Session.getScriptTimeZone();
  var total = 0;
  vals.forEach(function(row) {
    var chatId = String(row[1] || '').trim();
    if (!chatId) return;
    var receivedAtRaw = row[0];
    var receivedAtMs = 0;
    var receivedAtStr = '';
    if (receivedAtRaw instanceof Date) {
      receivedAtMs = receivedAtRaw.getTime();
      receivedAtStr = Utilities.formatDate(receivedAtRaw, tz, 'yyyy-MM-dd HH:mm');
    } else if (receivedAtRaw) {
      var parsed = new Date(String(receivedAtRaw));
      if (!isNaN(parsed.getTime())) {
        receivedAtMs = parsed.getTime();
        receivedAtStr = Utilities.formatDate(parsed, tz, 'yyyy-MM-dd HH:mm');
      } else {
        receivedAtStr = String(receivedAtRaw);
      }
    }
    // payload에서 factorLabel 추출 (폼 버전 태깅용)
    var factorLabel = '';
    var factorOptions = [];
    var submittedAt = 0;
    var payloadRaw = row[6];
    if (payloadRaw) {
      try {
        var pl = typeof payloadRaw === 'string' ? JSON.parse(payloadRaw) : (payloadRaw || {});
        factorLabel = String(pl.factorLabel || '');
        factorOptions = Array.isArray(pl.factorOptions) ? pl.factorOptions : [];
        submittedAt = Number(pl.submittedAt || 0);
      } catch (e) {}
    }
    // 레거시 데이터 fallback: factorLabel 없으면 factor 값으로 폼 종류 추정
    if (!factorLabel) {
      var factor = String(row[3] || '');
      if (/⚡|해결됨|해결되지\s*않음|늦은\s*답변|🔁/.test(factor)) {
        factorLabel = '해당 점수를 주신 배경이 무엇인가요?';
      } else if (/💁|🤖|답변\s*속도|문제\s*해결\s*여부/.test(factor)) {
        factorLabel = '해당 점수에 가장 큰 영향을 미친 요소가 무엇인가요?';
      }
    }
    var rec = {
      receivedAt: receivedAtStr,
      receivedAtMs: receivedAtMs,
      score: row[2],
      factor: String(row[3] || ''),
      etc: String(row[4] || ''),
      name: String(row[5] || ''),
      factorLabel: factorLabel,
      factorOptions: factorOptions,
      submittedAt: submittedAt
    };
    if (!byChatId[chatId]) byChatId[chatId] = [];
    byChatId[chatId].push(rec);
    total++;
  });
  Object.keys(byChatId).forEach(function(cid) {
    byChatId[cid].sort(function(a, b) { return (a.receivedAtMs || 0) - (b.receivedAtMs || 0); });
  });

  // 전체 데이터 기준 폼 라벨 세대 순위 (lastSeen desc) — 클라이언트가 "현재/이전" 라벨링에 사용
  var formStats = {}; // label -> { firstSeen, lastSeen, respondents }
  Object.keys(byChatId).forEach(function(cid) {
    byChatId[cid].forEach(function(rec) {
      var l = rec.factorLabel;
      if (!l) return;
      var ts = rec.submittedAt || rec.receivedAtMs || 0;
      if (!formStats[l]) formStats[l] = { firstSeen: Infinity, lastSeen: 0, respondents: 0 };
      var s = formStats[l];
      if (ts && ts < s.firstSeen) s.firstSeen = ts;
      if (ts && ts > s.lastSeen) s.lastSeen = ts;
      s.respondents++;
    });
  });
  var formRanking = Object.keys(formStats).map(function(l) {
    return { label: l, firstSeen: formStats[l].firstSeen, lastSeen: formStats[l].lastSeen, respondents: formStats[l].respondents };
  }).sort(function(a, b) { return (b.lastSeen || 0) - (a.lastSeen || 0); });

  return {
    sheetExists: true,
    byChatId: byChatId,
    count: Object.keys(byChatId).length,
    totalRecords: total,
    formRanking: formRanking
  };
}

// (제거) api_seedFakeSatisfactionData / api_clearFakeSatisfactionData: 초기 캡처용 시딩 함수
// (제거) api_scanTranscriptsForSurvey / api_importSurveyFromTranscripts: 이전 전사문 텍스트 기반 파싱
// 이제 fetchTranscriptSliceInRange가 수집 시점에 form.inputs로 직접 파싱해서 만족도_웹훅에 저장

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

// (제거) runDumpKangMinKyung: 특정 chatId 하드코딩된 dump 러너. 필요 시 api_dumpAllMessages 직접 호출

/**
 * 비즈봇 만족도 폼 응답 구조 프로브
 * - 여러 chatId를 받아 각 상담의 폼/응답 메시지 원본을 `_probe_bizbot` 시트에 덤프
 * - "상담 만족도 점수" 포함 메시지의 blocks/log/options를 전량 노출
 * - 한 상담에 여러 응답이 있는 경우 각각 행으로 저장
 */
function api_probeBizbotForm(channelKey, chatIds) {
  var ch = CHANNELS[channelKey];
  if (!ch) return { ok: false, error: '알 수 없는 채널' };
  var keys = getApiKeys(channelKey);
  if (!keys) return { ok: false, error: 'API 키 없음' };
  if (!chatIds || !chatIds.length) return { ok: false, error: 'chatIds 없음' };

  var headers = buildHeaders(keys.key, keys.secret);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('_probe_bizbot') || ss.insertSheet('_probe_bizbot');
  sh.clearContents();

  var out = [];
  out.push([
    'chatId', 'idx', 'createdAt', 'personType', 'personId', 'type',
    'plainText', 'blockCount', 'blockTypes', 'blocksJSON',
    'logAction', 'logJSON', 'options', 'form', 'allKeys', 'surveyKeywordHit'
  ]);

  var perChatSummary = [];

  chatIds.forEach(function(chatId) {
    chatId = String(chatId || '').trim();
    if (!chatId) return;

    var allMsgs = [];
    var since = '';
    for (var p = 0; p < 10; p++) {
      var params = { limit: 100, sortOrder: 'desc' };
      if (since) params.since = since;
      try {
        var data = ctGet('/user-chats/' + chatId + '/messages', params, headers);
        var msgs = data.messages || [];
        for (var i = 0; i < msgs.length; i++) allMsgs.push(msgs[i]);
        if (!data.next) break;
        since = data.next;
      } catch (e) { break; }
    }

    allMsgs.sort(function(a, b) { return Number(a.createdAt || 0) - Number(b.createdAt || 0); });

    var surveyCount = 0;
    var formLikeCount = 0;

    allMsgs.forEach(function(m, i) {
      var plain = String(m.plainText || '');
      var blocks = m.blocks || [];
      var blockTypes = blocks.map(function(b) { return b && b.type; });
      var blocksJson = JSON.stringify(blocks);
      var combined = plain + '\n' + blocksJson + '\n' + JSON.stringify(m.log || {}) + '\n' + JSON.stringify(m.options || []) + '\n' + JSON.stringify(m.form || {});
      var hit = /상담\s*만족도\s*점수|만족도\s*조사|survey/i.test(combined);
      var isFormLike = /form|choice|button|radio|input|select|answer|submit/i.test(blocksJson + JSON.stringify(m.log || {}));
      if (hit) surveyCount++;
      if (isFormLike) formLikeCount++;

      // 폼 관련되지 않은 일반 텍스트 메시지는 스킵 (덤프 크기 축소)
      if (!hit && !isFormLike && !m.form && !m.log) return;

      out.push([
        chatId,
        i + 1,
        m.createdAt ? Utilities.formatDate(new Date(Number(m.createdAt)), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss') : '',
        m.personType || '',
        m.personId || '',
        m.type || '',
        plain.slice(0, 400),
        blocks.length,
        JSON.stringify(blockTypes),
        blocksJson.slice(0, 8000),
        m.log ? String(m.log.action || '') : '',
        m.log ? JSON.stringify(m.log).slice(0, 2000) : '',
        JSON.stringify(m.options || []).slice(0, 1500),
        m.form ? JSON.stringify(m.form).slice(0, 3000) : '',
        Object.keys(m).join(','),
        hit ? 'YES' : ''
      ]);
    });

    perChatSummary.push({
      chatId: chatId,
      totalMessages: allMsgs.length,
      surveyKeywordHits: surveyCount,
      formLikeMessages: formLikeCount
    });
  });

  sh.getRange(1, 1, out.length, out[0].length).setValues(out);
  sh.setFrozenRows(1);

  return {
    ok: true,
    perChatSummary: perChatSummary,
    dumpRows: out.length - 1,
    note: '_probe_bizbot 시트에서 blocksJSON / logJSON / form 컬럼 확인'
  };
}

/**
 * 편집기 실행 진입점: Rok님이 제공한 2건에 프로브 실행
 */
function runProbeBizbot() {
  // 구 폼(2026 상반기) + 신 폼(2026-08) 각각 샘플
  var chatIds = ['6a826c7563a85ec227e4', '6a826bbd12aa12448966', '6a8d8553bc2cca004ceb'];
  var res = api_probeBizbotForm('refund', chatIds);
  Logger.log(JSON.stringify(res, null, 2));
  return res;
}

/**
 * 만족도_웹훅 시트 진단
 * - 총 행수, 컬럼별 결측 카운트
 * - factorLabel(payload) 유무 분포
 * - chatId별 레코드 병존 케이스 (score만 있는 legacy + score+factor 있는 신규 → 매칭 우선순위 이슈 원인)
 * - 각 카테고리별 샘플 5개 → `_만족도_웹훅_진단` 시트에 dump
 */
function api_diagnoseSatisfactionSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(WEBHOOK_SHEET);
  if (!sh) return { ok: false, error: WEBHOOK_SHEET + ' 시트 없음' };
  var last = sh.getLastRow();
  if (last < 2) return { ok: false, error: '데이터 없음' };

  var vals = sh.getRange(2, 1, last - 1, WEBHOOK_HEADERS.length).getValues();
  var tz = Session.getScriptTimeZone();

  var stats = {
    totalRows: vals.length,
    hasScore: 0,
    hasFactor: 0,
    hasEtc: 0,
    hasFactorLabelInPayload: 0,
    scoreOnlyNoFactor: 0,
    fullResponse: 0,        // score + factor 다 있음
    emptyAll: 0,            // 아무것도 없음
    apiScanRows: 0,
    webhookRows: 0,
    transcriptRows: 0,
    fakeRows: 0,
    unknownSourceRows: 0,
    payloadParseFailed: 0
  };

  // chatId별 병존 케이스 분석용
  var byChatId = {}; // chatId -> { total, scoreOnly, full, sources: {api-scan, webhook, transcript, ...} }
  var sampleScoreOnly = [];
  var sampleFull = [];
  var sampleLegacyPayload = [];

  vals.forEach(function(row, idx) {
    var chatId = String(row[1] || '').trim();
    var score = row[2];
    var factor = String(row[3] || '').trim();
    var etc = String(row[4] || '').trim();
    var payloadRaw = row[6];
    var receivedAt = row[0];

    var hasScore = (score !== '' && score !== null && score !== undefined && String(score) !== 'null');
    var hasFactor = !!factor;
    var hasEtc = !!etc;

    if (hasScore) stats.hasScore++;
    if (hasFactor) stats.hasFactor++;
    if (hasEtc) stats.hasEtc++;

    var source = 'unknown';
    var factorLabelInPayload = '';
    var payloadObj = null;
    if (payloadRaw) {
      try {
        payloadObj = typeof payloadRaw === 'string' ? JSON.parse(payloadRaw) : payloadRaw;
        if (payloadObj && payloadObj.source) source = String(payloadObj.source);
        if (payloadObj && payloadObj.factorLabel) {
          factorLabelInPayload = String(payloadObj.factorLabel);
          stats.hasFactorLabelInPayload++;
        }
      } catch (e) { stats.payloadParseFailed++; }
    }

    if (source === 'api-scan') stats.apiScanRows++;
    else if (source === 'webhook') stats.webhookRows++;
    else if (source === 'transcript') stats.transcriptRows++;
    else if (source === 'fake' || source === 'seed') stats.fakeRows++;
    else stats.unknownSourceRows++;

    if (hasScore && hasFactor) {
      stats.fullResponse++;
      if (sampleFull.length < 5) sampleFull.push({ row: idx + 2, chatId: chatId, score: score, factor: factor, source: source, factorLabelInPayload: factorLabelInPayload });
    } else if (hasScore && !hasFactor) {
      stats.scoreOnlyNoFactor++;
      if (sampleScoreOnly.length < 5) sampleScoreOnly.push({ row: idx + 2, chatId: chatId, score: score, factor: '(빈값)', source: source, factorLabelInPayload: factorLabelInPayload, payloadPreview: (payloadRaw ? String(payloadRaw).slice(0, 200) : '(없음)') });
    } else if (!hasScore && !hasFactor && !hasEtc) {
      stats.emptyAll++;
    }

    if (!factorLabelInPayload && hasScore && sampleLegacyPayload.length < 5) {
      sampleLegacyPayload.push({ row: idx + 2, chatId: chatId, score: score, factor: factor || '(빈값)', payloadPreview: (payloadRaw ? String(payloadRaw).slice(0, 300) : '(없음)') });
    }

    // 병존 분석
    if (chatId) {
      if (!byChatId[chatId]) byChatId[chatId] = { total: 0, scoreOnly: 0, full: 0, sources: {} };
      var g = byChatId[chatId];
      g.total++;
      if (hasScore && !hasFactor) g.scoreOnly++;
      if (hasScore && hasFactor) g.full++;
      g.sources[source] = (g.sources[source] || 0) + 1;
    }
  });

  // chatId별 병존 케이스 카운트
  var coexistCount = 0;
  var coexistSamples = [];
  Object.keys(byChatId).forEach(function(cid) {
    var g = byChatId[cid];
    if (g.scoreOnly > 0 && g.full > 0) {
      coexistCount++;
      if (coexistSamples.length < 10) coexistSamples.push({ chatId: cid, total: g.total, scoreOnly: g.scoreOnly, full: g.full, sources: g.sources });
    }
  });

  stats.uniqueChatIds = Object.keys(byChatId).length;
  stats.coexistChatIds = coexistCount;

  // 진단 시트에 dump
  var diagSh = ss.getSheetByName('_만족도_웹훅_진단') || ss.insertSheet('_만족도_웹훅_진단');
  diagSh.clearContents();
  var rows = [];
  rows.push(['=== 만족도_웹훅 시트 진단 ===', '']);
  rows.push(['생성 시각', Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss')]);
  rows.push(['', '']);
  rows.push(['--- 행 카운트 ---', '']);
  rows.push(['총 행수', stats.totalRows]);
  rows.push(['유니크 chatId', stats.uniqueChatIds]);
  rows.push(['score 있음', stats.hasScore]);
  rows.push(['factor 있음', stats.hasFactor]);
  rows.push(['etc 있음', stats.hasEtc]);
  rows.push(['score+factor 완전 응답', stats.fullResponse]);
  rows.push(['score만 있고 factor 없음', stats.scoreOnlyNoFactor]);
  rows.push(['아무것도 없음', stats.emptyAll]);
  rows.push(['', '']);
  rows.push(['--- payload 분석 ---', '']);
  rows.push(['factorLabel 태깅 있음', stats.hasFactorLabelInPayload]);
  rows.push(['payload 파싱 실패', stats.payloadParseFailed]);
  rows.push(['source=api-scan', stats.apiScanRows]);
  rows.push(['source=webhook', stats.webhookRows]);
  rows.push(['source=transcript', stats.transcriptRows]);
  rows.push(['source=fake/seed', stats.fakeRows]);
  rows.push(['source=unknown', stats.unknownSourceRows]);
  rows.push(['', '']);
  rows.push(['--- 병존 분석 ---', '']);
  rows.push(['chatId에 score-only + full 병존', stats.coexistChatIds]);
  rows.push(['(→ 매칭 시 어느 쪽 선택되냐에 따라 factor 유실 발생 가능)', '']);
  rows.push(['', '']);
  rows.push(['--- 병존 chatId 샘플 10 ---', '']);
  rows.push(['chatId', 'total', 'scoreOnly', 'full', 'sources']);
  coexistSamples.forEach(function(s) {
    rows.push([s.chatId, s.total, s.scoreOnly, s.full, JSON.stringify(s.sources)]);
  });
  rows.push(['', '']);
  rows.push(['--- score-only 샘플 5 ---', '']);
  rows.push(['시트행', 'chatId', 'score', 'factor', 'source', 'factorLabelInPayload', 'payloadPreview']);
  sampleScoreOnly.forEach(function(s) {
    rows.push([s.row, s.chatId, s.score, s.factor, s.source, s.factorLabelInPayload, s.payloadPreview]);
  });
  rows.push(['', '']);
  rows.push(['--- full 응답 샘플 5 ---', '']);
  rows.push(['시트행', 'chatId', 'score', 'factor', 'source', 'factorLabelInPayload']);
  sampleFull.forEach(function(s) {
    rows.push([s.row, s.chatId, s.score, s.factor, s.source, s.factorLabelInPayload]);
  });
  rows.push(['', '']);
  rows.push(['--- payload에 factorLabel 없는 행 샘플 (legacy 추정) ---', '']);
  rows.push(['시트행', 'chatId', 'score', 'factor', 'payloadPreview']);
  sampleLegacyPayload.forEach(function(s) {
    rows.push([s.row, s.chatId, s.score, s.factor, s.payloadPreview]);
  });

  // 6컬럼으로 통일해서 setValues
  var maxCols = 7;
  var padded = rows.map(function(r) {
    var row = r.slice();
    while (row.length < maxCols) row.push('');
    return row.slice(0, maxCols);
  });
  diagSh.getRange(1, 1, padded.length, maxCols).setValues(padded);

  return {
    ok: true,
    stats: stats,
    coexistSamples: coexistSamples,
    sampleScoreOnly: sampleScoreOnly,
    sampleFull: sampleFull,
    sampleLegacyPayload: sampleLegacyPayload,
    note: '_만족도_웹훅_진단 시트에 dump 완료'
  };
}

function runDiagnoseSatisfaction() {
  var res = api_diagnoseSatisfactionSheet();
  Logger.log(JSON.stringify(res, null, 2));
  return res;
}

/**
 * 만족도_웹훅 시트에서 factor 없는 행을 재스캔해 factor를 채워 넣음
 * - 각 factor-없음 행의 chatId + submittedAt로 API 재조회 → 새 파서로 재파싱
 * - factor 발견되면 해당 행의 factor/etc/payload 컬럼 업데이트
 * - 4.5분 타임아웃 · resume 인덱스 property로 이어서
 */
function api_backfillMissingFactors(channelKey) {
  var ch = CHANNELS[channelKey];
  if (!ch) return { ok: false, error: '알 수 없는 채널' };
  if (!ch.satisfactionGroupId) return { ok: false, error: '만족도 없는 채널' };
  var keys = getApiKeys(channelKey);
  if (!keys) return { ok: false, error: 'API 키 없음' };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(WEBHOOK_SHEET);
  if (!sh) return { ok: false, error: WEBHOOK_SHEET + ' 시트 없음' };
  var last = sh.getLastRow();
  if (last < 2) return { ok: false, error: '데이터 없음' };

  var vals = sh.getRange(2, 1, last - 1, WEBHOOK_HEADERS.length).getValues();
  var headers = buildHeaders(keys.key, keys.secret);
  var props = PropertiesService.getScriptProperties();
  var resumeKey = 'BACKFILL_FACTOR_RESUME__' + channelKey;
  var resumeAt = Number(props.getProperty(resumeKey) || 0);

  // 재스캔한 chatId 캐시 (한 chatId를 여러 행에서 반복 조회하지 않게)
  var scannedCache = {}; // chatId -> [{ score, factor, etc, factorLabel, factorOptions, submittedAt, messageId }]

  var stats = {
    totalScanned: 0,
    factorEmptyRows: 0,
    updated: 0,
    stillMissing: 0,
    apiCallsAvoided: 0,
    apiCalls: 0,
    resumedFrom: resumeAt,
    timedOut: false
  };
  var deadline = new Date().getTime() + 4.5 * 60 * 1000;

  var updatedRows = []; // { rowIdx: sheet row 1-based, factor, etc, payload }
  var i = resumeAt;
  for (; i < vals.length; i++) {
    if (new Date().getTime() > deadline) { stats.timedOut = true; break; }
    stats.totalScanned++;
    var row = vals[i];
    var chatId = String(row[1] || '').trim();
    var scoreRaw = row[2];
    var factorRaw = String(row[3] || '').trim();
    if (!chatId) continue;
    if (factorRaw) continue; // 이미 factor 있음 → 스킵
    stats.factorEmptyRows++;

    // 원래 payload 파싱해서 submittedAt 얻기 (매칭 키)
    var payloadRaw = row[6];
    var originalSubmittedAt = 0;
    if (payloadRaw) {
      try {
        var pl = typeof payloadRaw === 'string' ? JSON.parse(payloadRaw) : (payloadRaw || {});
        originalSubmittedAt = Number(pl.submittedAt || 0);
      } catch (e) {}
    }

    // 채팅 재스캔 (캐시 활용)
    var responses;
    if (scannedCache[chatId]) {
      responses = scannedCache[chatId];
      stats.apiCallsAvoided++;
    } else {
      try {
        var allMsgs = [];
        var since = '';
        for (var p = 0; p < 15; p++) {
          var params = { limit: 100, sortOrder: 'desc' };
          if (since) params.since = since;
          var data = ctGet('/user-chats/' + chatId + '/messages', params, headers);
          var mss = data.messages || [];
          for (var mi = 0; mi < mss.length; mi++) allMsgs.push(mss[mi]);
          if (!data.next) break;
          since = data.next;
        }
        responses = extractSatisfactionFromChatMessages(allMsgs);
        scannedCache[chatId] = responses;
        stats.apiCalls++;
      } catch (e) {
        continue;
      }
    }

    // 이 행의 submittedAt과 매칭되는 응답 찾기
    var matched = null;
    for (var r = 0; r < responses.length; r++) {
      if (originalSubmittedAt && Number(responses[r].submittedAt) === originalSubmittedAt) {
        matched = responses[r]; break;
      }
    }
    // submittedAt 없이 저장된 레거시 행이면 아무 응답 중 factor 있는 것 하나 택 (chatId 유니크 응답이 1개인 경우에만 안전)
    if (!matched && !originalSubmittedAt && responses.length === 1) {
      matched = responses[0];
    }

    if (matched && matched.factor) {
      var newPayload = JSON.stringify({
        source: 'backfill',
        submittedAt: matched.submittedAt,
        messageId: matched.messageId,
        factorLabel: matched.factorLabel || '',
        factorOptions: matched.factorOptions || [],
        originalSource: (function() {
          try { return (typeof payloadRaw === 'string' ? JSON.parse(payloadRaw) : (payloadRaw || {})).source || ''; } catch (e) { return ''; }
        })()
      });
      updatedRows.push({
        rowIdx: i + 2,
        score: matched.score !== null && matched.score !== undefined ? String(matched.score) : String(scoreRaw || ''),
        factor: matched.factor,
        etc: matched.etc || String(row[4] || ''),
        payload: newPayload
      });
      stats.updated++;
    } else {
      stats.stillMissing++;
    }
  }

  // 배치 업데이트: C(점수)/D(요소)/E(기타) 3컬럼 + G(payload) 1컬럼 별도 (F 이름은 유지)
  updatedRows.forEach(function(u) {
    sh.getRange(u.rowIdx, 3, 1, 3).setValues([[u.score, u.factor, u.etc]]);
    sh.getRange(u.rowIdx, 7).setValue(u.payload);
  });

  if (stats.timedOut) {
    props.setProperty(resumeKey, String(i));
    stats.resumeAt = i;
  } else {
    props.deleteProperty(resumeKey);
  }

  return {
    ok: true,
    stats: stats,
    note: stats.timedOut
      ? '시간 초과 · 다시 실행하면 ' + i + '행부터 이어서'
      : '완료 · factor 없음 ' + stats.factorEmptyRows + '건 중 ' + stats.updated + '건 채워짐 (' + stats.stillMissing + '건 여전히 미매칭)'
  };
}

function runBackfillFactors() {
  var res = api_backfillMissingFactors('refund');
  Logger.log(JSON.stringify(res, null, 2));
  return res;
}

/**
 * factor 여전히 없는 행들 진단 dump
 * - 각 행의 chatId·score·payload·submittedAt·receivedAt 원본
 * - API로 상담 재조회해서 실제 폼 응답이 몇 개 있는지, 각각 factor 유무 확인
 * - `_만족도_웹훅_미매칭_진단` 시트에 dump
 */
function api_diagnoseStillMissingFactors(channelKey) {
  var ch = CHANNELS[channelKey];
  if (!ch) return { ok: false, error: '알 수 없는 채널' };
  var keys = getApiKeys(channelKey);
  if (!keys) return { ok: false, error: 'API 키 없음' };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(WEBHOOK_SHEET);
  if (!sh) return { ok: false, error: WEBHOOK_SHEET + ' 시트 없음' };
  var last = sh.getLastRow();
  if (last < 2) return { ok: false, error: '데이터 없음' };

  var vals = sh.getRange(2, 1, last - 1, WEBHOOK_HEADERS.length).getValues();
  var tz = Session.getScriptTimeZone();
  var headers = buildHeaders(keys.key, keys.secret);

  // factor 여전히 없는 행 수집
  var missing = [];
  for (var i = 0; i < vals.length; i++) {
    var row = vals[i];
    var chatId = String(row[1] || '').trim();
    var factor = String(row[3] || '').trim();
    if (!chatId || factor) continue;
    var receivedAtRaw = row[0];
    var receivedAtStr = receivedAtRaw instanceof Date
      ? Utilities.formatDate(receivedAtRaw, tz, 'yyyy-MM-dd HH:mm:ss')
      : String(receivedAtRaw || '');
    var payloadRaw = row[6];
    var payloadObj = null;
    var payloadSubmittedAt = 0;
    var payloadSource = '';
    if (payloadRaw) {
      try {
        payloadObj = typeof payloadRaw === 'string' ? JSON.parse(payloadRaw) : payloadRaw;
        payloadSubmittedAt = Number(payloadObj.submittedAt || 0);
        payloadSource = String(payloadObj.source || '');
      } catch (e) {}
    }
    missing.push({
      sheetRow: i + 2,
      chatId: chatId,
      score: row[2],
      etc: String(row[4] || ''),
      name: String(row[5] || ''),
      receivedAt: receivedAtStr,
      payloadSubmittedAt: payloadSubmittedAt,
      payloadSubmittedAtHuman: payloadSubmittedAt ? Utilities.formatDate(new Date(payloadSubmittedAt), tz, 'yyyy-MM-dd HH:mm:ss') : '',
      payloadSource: payloadSource,
      payloadRaw: payloadRaw ? String(payloadRaw).slice(0, 300) : ''
    });
  }

  // 각 chatId에 대해 API 재조회 → 실제 폼 응답 얼마나 있는지 파악
  var chatScanCache = {};
  missing.forEach(function(m) {
    if (chatScanCache[m.chatId]) return;
    try {
      var allMsgs = [];
      var since = '';
      for (var p = 0; p < 15; p++) {
        var params = { limit: 100, sortOrder: 'desc' };
        if (since) params.since = since;
        var data = ctGet('/user-chats/' + m.chatId + '/messages', params, headers);
        var mss = data.messages || [];
        for (var mi = 0; mi < mss.length; mi++) allMsgs.push(mss[mi]);
        if (!data.next) break;
        since = data.next;
      }
      var responses = extractSatisfactionFromChatMessages(allMsgs);
      // 원시 form 메시지도 카운트 (submittedAt 없는 것 포함)
      var rawFormCount = 0;
      var rawFormSubmittedCount = 0;
      var rawFormSurveyCount = 0;
      allMsgs.forEach(function(mm) {
        if (mm.form && mm.form.type === 'custom') {
          rawFormCount++;
          if (mm.form.submittedAt) rawFormSubmittedCount++;
          if ((mm.form.inputs || []).some(function(inp) { return /만족도\s*점수/.test(String(inp.label || '')); })) {
            rawFormSurveyCount++;
          }
        }
      });
      chatScanCache[m.chatId] = {
        responses: responses,
        rawFormCount: rawFormCount,
        rawFormSubmittedCount: rawFormSubmittedCount,
        rawFormSurveyCount: rawFormSurveyCount,
        totalMsgs: allMsgs.length
      };
    } catch (e) {
      chatScanCache[m.chatId] = { error: String(e.message || e) };
    }
  });

  // 진단 시트에 dump
  var diagSh = ss.getSheetByName('_만족도_웹훅_미매칭_진단') || ss.insertSheet('_만족도_웹훅_미매칭_진단');
  diagSh.clearContents();
  var out = [[
    '시트행', 'chatId', 'score', 'etc', 'name', 'receivedAt',
    'payload.submittedAt', 'payloadSubmittedAtHuman', 'payload.source', 'payloadPreview',
    'API총메시지', 'API총폼', '만족도폼', '제출완료폼',
    'API응답0=score', 'API응답0=factor', 'API응답0=factorLabel', 'API응답0=submittedAt', '판정'
  ]];
  missing.forEach(function(m) {
    var s = chatScanCache[m.chatId] || {};
    var r0 = (s.responses && s.responses[0]) || {};
    var verdict = '';
    if (s.error) verdict = 'API 오류: ' + s.error;
    else if (!s.rawFormSurveyCount) verdict = '이 상담엔 만족도 폼 자체가 없음 (오배정 or 이전 삭제)';
    else if (!s.rawFormSubmittedCount) verdict = '폼은 있으나 사용자가 제출 안 함';
    else if (!s.responses.length) verdict = '제출됐지만 파서가 못 잡음 (라벨 확인 필요)';
    else if (s.responses.length && !s.responses.some(function(r) { return r.factor; })) verdict = '제출됐으나 factor 값 없음 (특이 케이스)';
    else if (m.payloadSubmittedAt && !s.responses.some(function(r) { return Number(r.submittedAt) === m.payloadSubmittedAt; })) verdict = 'submittedAt 매칭 실패 (기존 payload의 submittedAt이 이제 존재 안 함)';
    else verdict = '알 수 없음';

    out.push([
      m.sheetRow, m.chatId, m.score, m.etc, m.name, m.receivedAt,
      m.payloadSubmittedAt, m.payloadSubmittedAtHuman, m.payloadSource, m.payloadRaw,
      s.totalMsgs || 0, s.rawFormCount || 0, s.rawFormSurveyCount || 0, s.rawFormSubmittedCount || 0,
      r0.score !== undefined ? r0.score : '',
      r0.factor || '',
      r0.factorLabel || '',
      r0.submittedAt || '',
      verdict
    ]);
  });
  diagSh.getRange(1, 1, out.length, out[0].length).setValues(out);
  diagSh.setFrozenRows(1);

  // 판정별 카운트
  var verdictCounts = {};
  out.slice(1).forEach(function(r) {
    var v = String(r[r.length - 1]);
    verdictCounts[v] = (verdictCounts[v] || 0) + 1;
  });

  return {
    ok: true,
    totalMissing: missing.length,
    verdictCounts: verdictCounts,
    note: '_만족도_웹훅_미매칭_진단 시트에 dump 완료'
  };
}

function runDiagnoseStillMissing() {
  var res = api_diagnoseStillMissingFactors('refund');
  Logger.log(JSON.stringify(res, null, 2));
  return res;
}

/**
 * 채널톡 만족도 조사 폼 응답 파싱
 * form.type === 'custom' + submittedAt 존재 + inputs[].value 존재 시 유효 응답으로 인식
 * inputs 라벨 매칭:
 *   - "상담 만족도 점수" → score
 *   - "영향을 미친 요소" → factor
 *   - "기타.*사유" → etc
 */
function extractSatisfactionFromChatMessages(msgs) {
  var out = [];
  for (var i = 0; i < msgs.length; i++) {
    var m = msgs[i];
    var form = m.form;
    if (!form || form.type !== 'custom') continue;
    if (!form.submittedAt) continue;
    var inputs = form.inputs || [];
    var score = null, factor = null, etc = null;
    var isSatisfactionForm = false;
    var factorLabel = ''; // 폼 버전 지문 (신·구 폼 구분용)
    var factorOptions = []; // 요소 선택지 배열
    for (var j = 0; j < inputs.length; j++) {
      var inp = inputs[j];
      var label = String(inp.label || '');
      var val = inp.value;
      if (/만족도\s*점수/.test(label)) {
        isSatisfactionForm = true;
        if (val !== undefined && val !== null && val !== '') {
          var n = Number(val);
          if (!isNaN(n)) score = n;
        }
      } else if (/영향.*요소|점수.*배경/.test(label)) {
        // 신 폼: "해당 점수에 가장 큰 영향을 미친 요소가 무엇인가요?"
        // 구 폼: "해당 점수를 주신 배경이 무엇인가요?"
        factorLabel = label;
        factorOptions = Array.isArray(inp.options) ? inp.options.slice() : [];
        if (val !== undefined && val !== null && val !== '') factor = String(val);
      } else if (/기타.*사유|기타.*의견|기타.*이유/.test(label)) {
        if (val !== undefined && val !== null && val !== '') etc = String(val);
      }
    }
    if (!isSatisfactionForm) continue;
    if (score === null && !factor && !etc) continue;
    out.push({
      messageId: m.id,
      personId: m.personId || '',
      score: score,
      factor: factor,
      etc: etc,
      factorLabel: factorLabel,
      factorOptions: factorOptions,
      submittedAt: Number(form.submittedAt || 0),
      createdAt: Number(m.createdAt || 0)
    });
  }
  return out;
}

/**
 * 시트의 상담들을 순회하며 폼 응답을 스캔해 만족도_웹훅 시트에 append
 * - 중복 방지 키: chatId + '|' + submittedAt
 * - 시간 제한: 4.5분 (Apps Script 6분 한도)
 * - 스캔 완료된 chatId 는 script property 에 기록 → 다음 호출 시 스킵 (force=true면 무시)
 * - 자동 페이지네이션: 시간 초과되면 resume 인덱스 저장, 다시 호출하면 이어서
 */
function api_syncSatisfactionFromChats(channelKey, sheetName, force) {
  var ch = CHANNELS[channelKey];
  if (!ch) return { ok: false, error: '알 수 없는 채널' };
  if (!ch.satisfactionGroupId) return { ok: false, error: '만족도 조사 없는 채널', skipped: true };
  var keys = getApiKeys(channelKey);
  if (!keys) return { ok: false, error: 'API 키 없음' };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var srcSh = ss.getSheetByName(sheetName);
  if (!srcSh) return { ok: false, error: '시트 없음: ' + sheetName };
  var last = srcSh.getLastRow();
  if (last < 2) return { ok: false, error: '시트에 데이터 없음' };

  var numCols = Math.max(srcSh.getLastColumn(), CFG.OUT_HEADERS.length);
  var vals = srcSh.getRange(2, 1, last - 1, numCols).getValues();

  var webSh = ss.getSheetByName(WEBHOOK_SHEET) || ss.insertSheet(WEBHOOK_SHEET);
  if (webSh.getLastRow() === 0) writeSheetValues(webSh, 1, 1, [WEBHOOK_HEADERS]);

  // 기존 웹훅 시트에서 chatId|submittedAt 셋 로드 → 중복 방지
  var existingKeys = {};
  if (webSh.getLastRow() > 1) {
    var wrows = webSh.getRange(2, 1, webSh.getLastRow() - 1, WEBHOOK_HEADERS.length).getValues();
    for (var w = 0; w < wrows.length; w++) {
      var cid = String(wrows[w][1] || '').trim();
      if (!cid) continue;
      var payloadRaw = wrows[w][6];
      var sub = '';
      try {
        var pl = typeof payloadRaw === 'string' ? JSON.parse(payloadRaw) : (payloadRaw || {});
        sub = pl.submittedAt ? String(pl.submittedAt) : '';
      } catch (e) {}
      if (sub) existingKeys[cid + '|' + sub] = true;
    }
  }

  var props = PropertiesService.getScriptProperties();
  var scannedKey = 'SAT_SCANNED__' + channelKey + '__' + sheetName;
  var resumeKey = 'SAT_SYNC_RESUME__' + channelKey + '__' + sheetName;

  var scannedSet = {};
  if (!force) {
    var scannedRaw = props.getProperty(scannedKey);
    if (scannedRaw) {
      try {
        var arr = JSON.parse(scannedRaw);
        for (var s = 0; s < arr.length; s++) scannedSet[arr[s]] = true;
      } catch (e) {}
    }
  }

  var resumeAt = Number(props.getProperty(resumeKey) || 0);
  if (force) resumeAt = 0;

  var headers = buildHeaders(keys.key, keys.secret);
  var newRows = [];
  var stats = { scanned: 0, chatsWithResponse: 0, responses: 0, skippedExisting: 0, skippedCached: 0, errors: 0, resumedFrom: resumeAt, timedOut: false };
  var deadline = new Date().getTime() + 4.5 * 60 * 1000;
  var newlyScanned = [];

  var i = resumeAt;
  for (; i < vals.length; i++) {
    if (new Date().getTime() > deadline) { stats.timedOut = true; break; }
    var row = vals[i];
    var link = String(row[7] || '');
    var chatMatch = link.match(/user-chats\/([a-f0-9]{20,40})/);
    if (!chatMatch) continue;
    var chatId = chatMatch[1];
    if (scannedSet[chatId]) { stats.skippedCached++; continue; }

    var name = String(row[1] || '');
    stats.scanned++;

    try {
      var allMsgs = [];
      var since = '';
      for (var p = 0; p < 15; p++) {
        var params = { limit: 100, sortOrder: 'desc' };
        if (since) params.since = since;
        var data = ctGet('/user-chats/' + chatId + '/messages', params, headers);
        var mss = data.messages || [];
        for (var mi = 0; mi < mss.length; mi++) allMsgs.push(mss[mi]);
        if (!data.next) break;
        since = data.next;
      }
      var responses = extractSatisfactionFromChatMessages(allMsgs);
      if (responses.length) stats.chatsWithResponse++;
      for (var r = 0; r < responses.length; r++) {
        var rp = responses[r];
        var key = chatId + '|' + rp.submittedAt;
        if (existingKeys[key]) { stats.skippedExisting++; continue; }
        existingKeys[key] = true;
        newRows.push([
          new Date(rp.submittedAt || rp.createdAt),
          chatId,
          rp.score !== null && rp.score !== undefined ? String(rp.score) : '',
          rp.factor || '',
          rp.etc || '',
          name,
          JSON.stringify({
            source: 'api-scan',
            submittedAt: rp.submittedAt,
            messageId: rp.messageId,
            factorLabel: rp.factorLabel || '',
            factorOptions: rp.factorOptions || []
          })
        ]);
        stats.responses++;
      }
      scannedSet[chatId] = true;
      newlyScanned.push(chatId);
    } catch (e) {
      stats.errors++;
    }
  }

  if (newRows.length) writeSheetValues(webSh, webSh.getLastRow() + 1, 1, newRows);

  // 스캔 완료된 chatId 집합 저장 (PropertiesService 500KB 한도 안에서: 20자 × 5000개 ≈ 100KB, 여유 충분)
  var allScanned = Object.keys(scannedSet);
  try {
    props.setProperty(scannedKey, JSON.stringify(allScanned));
  } catch (e) {
    // 초과 시 최근 것만 유지
    props.setProperty(scannedKey, JSON.stringify(allScanned.slice(-4000)));
  }

  if (stats.timedOut) {
    props.setProperty(resumeKey, String(i));
    stats.resumeAt = i;
  } else {
    props.deleteProperty(resumeKey);
  }

  return {
    ok: true,
    stats: stats,
    totalChats: vals.length,
    completed: !stats.timedOut,
    note: stats.timedOut
      ? '시간 제한으로 중단. 다시 실행하면 ' + i + '번 상담부터 이어서 스캔합니다.'
      : '완료 · ' + newRows.length + '건 추가 · ' + stats.skippedCached + '건 캐시 스킵'
  };
}

/**
 * 매니저(담당자) → 팀 자동 매핑
 * /managers 로 매니저 목록 조회 + /groups 로 그룹 목록 + managerIds 로 역매핑
 */
function api_getManagerTeamMapping(channelKey) {
  var ch = CHANNELS[channelKey];
  if (!ch) return { ok: false, error: '알 수 없는 채널' };
  var keys = getApiKeys(channelKey);
  if (!keys) return { ok: false, error: 'API 키 없음' };

  var headers = buildHeaders(keys.key, keys.secret);
  var managers = [];
  var since = '';
  var guard = 0;

  // 매니저 전체 페이지네이션
  while (guard++ < 50) {
    var params = { limit: 500 };
    if (since) params.since = since;
    var data;
    try {
      data = ctGet('/managers', params, headers);
    } catch (e) {
      return { ok: false, error: 'managers API 오류: ' + String(e.message || e) };
    }
    var page = data.managers || [];
    managers = managers.concat(page);
    if (!data.next) break;
    since = data.next;
  }

  // 그룹 전체 페이지네이션
  var groups = [];
  since = '';
  guard = 0;
  while (guard++ < 50) {
    var params2 = { limit: 100 };
    if (since) params2.since = since;
    var data2;
    try {
      data2 = ctGet('/groups', params2, headers);
    } catch (e) {
      return { ok: false, error: 'groups API 오류: ' + String(e.message || e) };
    }
    var page2 = data2.groups || [];
    groups = groups.concat(page2);
    if (!data2.next) break;
    since = data2.next;
  }

  // 역매핑: managerId → [그룹 정보]
  var mgrTeams = {};
  groups.forEach(function(g) {
    (g.managerIds || []).forEach(function(mgrId) {
      if (!mgrTeams[mgrId]) mgrTeams[mgrId] = [];
      mgrTeams[mgrId].push({ id: g.id, title: g.title || g.name, scope: g.scope });
    });
  });

  var result = managers.map(function(m) {
    return {
      id: m.id,
      name: m.name,
      email: m.email,
      role: m.roleId,
      teams: mgrTeams[m.id] || []
    };
  });

  return {
    ok: true,
    channel: channelKey,
    totalManagers: managers.length,
    totalGroups: groups.length,
    managers: result,
    allGroups: groups.map(function(g) {
      return {
        id: g.id,
        title: g.title || g.name,
        scope: g.scope,
        memberCount: (g.managerIds || []).length
      };
    })
  };
}

/**
 * 상담 배정용 팀 관련 비공식 엔드포인트 탐색
 */
function api_probeRoutingTeamEndpoints(channelKey) {
  var ch = CHANNELS[channelKey];
  if (!ch) return { ok: false, error: '알 수 없는 채널' };
  var keys = getApiKeys(channelKey);
  if (!keys) return { ok: false, error: 'API 키 없음' };
  var headers = buildHeaders(keys.key, keys.secret);

  var attempts = [
    { key: '/teams', url: CFG.CHANNEL_BASE + '/teams?limit=10' },
    { key: '/routing-teams', url: CFG.CHANNEL_BASE + '/routing-teams?limit=10' },
    { key: '/assignment-teams', url: CFG.CHANNEL_BASE + '/assignment-teams?limit=10' },
    { key: '/manager-teams', url: CFG.CHANNEL_BASE + '/manager-teams?limit=10' },
    { key: '/routes', url: CFG.CHANNEL_BASE + '/routes?limit=10' },
    { key: '/user-chats-assignments', url: CFG.CHANNEL_BASE + '/user-chats-assignments?limit=10' }
  ];

  var requests = attempts.map(function(a) {
    return { url: a.url, method: 'get', headers: headers, muteHttpExceptions: true };
  });

  var results = {};
  try {
    var responses = UrlFetchApp.fetchAll(requests);
    responses.forEach(function(res, i) {
      var code = res.getResponseCode();
      var body = res.getContentText() || '';
      var key = attempts[i].key;
      if (code >= 200 && code < 300) {
        try {
          var parsed = JSON.parse(body || '{}');
          results[key] = {
            ok: true,
            topKeys: Object.keys(parsed),
            preview: JSON.stringify(parsed).slice(0, 800)
          };
        } catch (e) {
          results[key] = { ok: true, topKeys: [], preview: body.slice(0, 400) };
        }
      } else {
        results[key] = { ok: false, error: 'HTTP ' + code };
      }
    });
  } catch (e) {
    return { ok: false, error: '병렬 fetch 실패: ' + String(e.message || e) };
  }

  return {
    ok: true,
    channel: channelKey,
    endpoints: results,
    note: 'HTTP 200 응답 나오는 엔드포인트가 있으면 상담 배정용 팀 데이터 사용 가능'
  };
}

/**
 * 편집기 실행용 - 상담 배정용 팀 엔드포인트 탐색
 */
function runProbeRoutingTeams() {
  var res = api_probeRoutingTeamEndpoints('refund');
  if (!res.ok) { Logger.log('오류: ' + res.error); return; }
  Logger.log('==============================');
  Logger.log('상담 배정용 팀 엔드포인트 탐색 결과:');
  Object.keys(res.endpoints).forEach(function(ep) {
    var r = res.endpoints[ep];
    if (r.ok) {
      Logger.log('✅ ' + ep);
      Logger.log('   topKeys: ' + (r.topKeys || []).join(', '));
      Logger.log('   preview: ' + r.preview);
    } else {
      Logger.log('❌ ' + ep + '  → ' + r.error);
    }
  });
  return res;
}

/**
 * 편집기에서 직접 실행용 - 매니저-팀 매핑 결과를 로그 + 시트에 저장
 */
function runManagerTeamMapping() {
  var res = api_getManagerTeamMapping('refund');
  if (!res.ok) { Logger.log('오류: ' + res.error); return; }

  Logger.log('==============================');
  Logger.log('총 매니저: ' + res.totalManagers);
  Logger.log('총 그룹: ' + res.totalGroups);
  Logger.log('==============================');
  Logger.log('전체 그룹 목록 (title):');
  res.allGroups.forEach(function(g, i) {
    Logger.log('[' + (i + 1) + '] ' + g.title + ' (id=' + g.id + ', 멤버=' + g.memberCount + ', scope=' + g.scope + ')');
  });

  // 시트에 저장
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shGroups = ss.getSheetByName('_그룹_리스트') || ss.insertSheet('_그룹_리스트');
  shGroups.clearContents();
  var groupRows = [['id', 'title', 'scope', 'memberCount']];
  res.allGroups.forEach(function(g) {
    groupRows.push([g.id, g.title, g.scope || '', g.memberCount]);
  });
  shGroups.getRange(1, 1, groupRows.length, 4).setValues(groupRows);

  var shMgrs = ss.getSheetByName('_매니저_팀매핑') || ss.insertSheet('_매니저_팀매핑');
  shMgrs.clearContents();
  var mgrRows = [['id', 'name', 'email', '소속 그룹들 (title, 콤마구분)']];
  res.managers.forEach(function(m) {
    mgrRows.push([
      m.id,
      m.name,
      m.email || '',
      (m.teams || []).map(function(t) { return t.title; }).join(', ')
    ]);
  });
  shMgrs.getRange(1, 1, mgrRows.length, 4).setValues(mgrRows);

  Logger.log('시트에 저장: _그룹_리스트, _매니저_팀매핑');
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
    .createMenu('CX 통합 대시보드')
    .addItem('환급 API 키 설정', 'setChannelKeysRefund')
    .addItem('케어 API 키 설정', 'setChannelKeysCare')
    .addItem('현재 작업 상태 확인', 'showJobStatus')
    .addItem('수집 중지', 'cancelChannelTalkJob')
    .addSeparator()
    .addItem('_그룹_리스트/_매니저_팀매핑 수기 세팅 (케어)', 'setupManagerTeamMappingManual')
    .addItem('활성 탭 메타베이스 매칭 갱신', 'refreshMetaOnActiveOutputSheet')
    .addSeparator()
    .addItem('Groq API 키 설정', 'setGroqApiKey')
    .addItem('Groq 사용 가능 모델 조회', 'listGroqModels')
    .addItem('VOC 분류 파일럿 (활성 탭 10건)', 'runVocClassificationPilot')
    .addSeparator()
    .addItem('활성 탭 전체 VOC 분류 (Phase 2, LLM)', 'runVocClassificationFullBatch')
    .addItem('활성 탭 리포트 생성 (Phase 1+2 통합)', 'buildCareReport')
    .addSeparator()
    .addItem('자동 파이프라인 시작 (활성 탭)', 'startAutoAnalysisPipelineFromActive')
    .addItem('자동 파이프라인 취소', 'cancelAutoAnalysisPipeline')
    .addItem('자동 파이프라인 상태 확인', 'showAutoAnalysisPipelineStatus')
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
    skips: sanitizeSkips({
      autoClosed: map['스킵 - ' + SKIP_LABELS.autoClosed],
      outOfRange: map['스킵 - ' + SKIP_LABELS.outOfRange],
      emptyText: map['스킵 - ' + SKIP_LABELS.emptyText],
      fetchError: map['스킵 - ' + SKIP_LABELS.fetchError]
    })
  };
}

function api_getAllChannelStatus() {
  var out = [];
  Object.keys(CHANNELS).forEach(function(k) {
    var running = getJobState(k);
    var entry;
    if (running) {
      entry = {
        channel: k,
        channelLabel: CHANNELS[k].label,
        running: true,
        state: getJobStateEnriched(k)
      };
    } else {
      var last = api_getLastProgress(k);
      entry = {
        channel: k,
        channelLabel: CHANNELS[k].label,
        running: false,
        state: last
      };
    }
    // 자동 분석 파이프라인 상태 (케어 채널만)
    if (k === 'care') {
      try { entry.autoPipeline = getAutoPipelineStatusForStatus(); } catch (e) { entry.autoPipeline = null; }
    }
    out.push(entry);
  });
  return out;
}

/**
 * 채널 상태 카드에 표시할 자동 파이프라인 진행 상태
 * - 활성이면 stage/진행률
 * - 비활성이면 마지막 리포트 시트 존재 여부로 완료 판단
 */
function getAutoPipelineStatusForStatus() {
  var props = PropertiesService.getScriptProperties();
  var sheetName = props.getProperty(AUTO_PIPELINE_CFG.SHEET_PROP);
  var stage = props.getProperty(AUTO_PIPELINE_CFG.STAGE_PROP);
  var startedAt = props.getProperty(AUTO_PIPELINE_CFG.STARTED_AT_PROP);
  if (sheetName && stage) {
    var status = { active: true, targetSheet: sheetName, stage: stage, startedAt: startedAt };
    if (stage === 'LLM') {
      var progressProp = VOC_BATCH_CFG.RESUME_PROP_PREFIX + sheetName;
      var lastRow = Number(props.getProperty(progressProp) || 0);
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sh = ss.getSheetByName(sheetName);
      var totalRows = sh ? Math.max(0, sh.getLastRow() - 1) : 0;
      status.llmProgress = {
        lastRow: lastRow,
        total: totalRows,
        pct: totalRows > 0 ? Math.round(lastRow / totalRows * 100) : 0
      };
      // 예상 소요: batch당 (LLM응답 8초 + sleep 2.5초) 가정 → 배치=30건
      var remaining = Math.max(0, totalRows - lastRow);
      var remainingBatches = Math.ceil(remaining / VOC_BATCH_CFG.BATCH_SIZE);
      status.etaSec = remainingBatches * 10;
    }
    return status;
  }
  // 비활성: 최근 케어 시트에 대응하는 리포트 시트가 있으면 완료로 판단
  var last = api_getLastProgress('care');
  if (last && last.outputSheet) {
    var reportName = '_리포트_' + last.outputSheet;
    var ss2 = SpreadsheetApp.getActiveSpreadsheet();
    if (ss2.getSheetByName(reportName)) {
      return { active: false, done: true, reportSheet: reportName, targetSheet: last.outputSheet };
    }
    return { active: false, done: false, targetSheet: last.outputSheet };
  }
  return null;
}

function api_getSpreadsheetUrl() {
  return SpreadsheetApp.getActiveSpreadsheet().getUrl();
}

// ================================
// 스냅샷 시스템
// ================================

function ensureSnapshotSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SNAPSHOT_SHEET);
  if (!sh) {
    sh = ss.insertSheet(SNAPSHOT_SHEET);
    writeSheetValues(sh, 1, 1, [SNAPSHOT_HEADERS]);
    sh.hideSheet();
  }
  return sh;
}

function extractPeriodFromSheetName_(sheetName) {
  var m = String(sheetName || '').match(/^(\d{8})~(\d{8})/);
  return m ? m[1] + '~' + m[2] : '';
}

function makeSnapshotId_(channelKey, sheetName) {
  return channelKey + '__' + sheetName;
}

function splitJsonIntoChunks_(jsonStr, chunkLimit, maxChunks) {
  var chunks = [];
  var i = 0;
  while (i < jsonStr.length && chunks.length < maxChunks) {
    chunks.push(jsonStr.substr(i, chunkLimit));
    i += chunkLimit;
  }
  return { chunks: chunks, truncated: i < jsonStr.length, originalLen: jsonStr.length };
}

/**
 * 대시보드 + 만족도 응답자 스냅샷 저장 - raw 시트 삭제해도 뷰 완전 유지
 * - dashboardJson: api_getDashboardData 결과 (KPI/태그/해지 등)
 * - satisfactionJson: api_getSatisfactionRespondents 결과 (응답자 리스트)
 * - 각각 3개 컬럼으로 분할해 셀당 45KB 제한 회피
 */
function api_saveDashboardSnapshot(channelKey, sheetName) {
  var ch = CHANNELS[channelKey];
  if (!ch) return { ok: false, error: '알 수 없는 채널' };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rawSh = ss.getSheetByName(sheetName);
  if (!rawSh) return { ok: false, error: '원본 시트 없음: ' + sheetName };

  // ---- 대시보드 데이터 ----
  var dashboardData;
  try {
    dashboardData = api_getDashboardData(channelKey, sheetName);
  } catch (e) {
    return { ok: false, error: '대시보드 데이터 계산 실패: ' + String(e.message || e) };
  }
  var dashJson = JSON.stringify(dashboardData);
  var maxChunks = 3;
  var maxSize = SNAPSHOT_JSON_COL_LIMIT * maxChunks;
  if (dashJson.length > maxSize) {
    var lite = JSON.parse(dashJson);
    ['meetInboundChats', 'noMgrChats', 'unmatchedChats'].forEach(function(k) {
      if (Array.isArray(lite[k]) && lite[k].length > 30) {
        lite[k + '_truncated'] = lite[k].length;
        lite[k] = lite[k].slice(0, 30);
      }
    });
    dashJson = JSON.stringify(lite);
  }
  var dashSplit = splitJsonIntoChunks_(dashJson, SNAPSHOT_JSON_COL_LIMIT, maxChunks);

  // ---- 만족도 응답자 (만족도 채널만) ----
  var satJson = '';
  var satSplit = { chunks: ['', '', ''], originalLen: 0, truncated: false };
  var satRespondentsCount = 0;
  if (ch.satisfactionGroupId) {
    try {
      var satRes = api_getSatisfactionRespondents(channelKey, sheetName);
      if (satRes && satRes.ok) {
        satRespondentsCount = (satRes.respondents || []).length;
        satJson = JSON.stringify(satRes);
        if (satJson.length > maxSize) {
          // 응답자 리스트가 너무 크면 개별 필드 슬림화
          var satLite = JSON.parse(satJson);
          if (Array.isArray(satLite.respondents)) {
            satLite.respondents = satLite.respondents.map(function(r) {
              // 렌더에 필수인 필드만 유지 (transcript나 대용량 필드 제거)
              return {
                chatId: r.chatId, name: r.name, bizName: r.bizName,
                respondedAt: r.respondedAt, respondedAtMs: r.respondedAtMs,
                link: r.link, tags: r.tags, inflow: r.inflow,
                inPeriod: r.inPeriod
              };
            });
            satJson = JSON.stringify(satLite);
          }
        }
        satSplit = splitJsonIntoChunks_(satJson, SNAPSHOT_JSON_COL_LIMIT, maxChunks);
      }
    } catch (e) {
      Logger.log('[SNAPSHOT] 만족도 저장 실패: ' + e.message);
    }
  }

  // ---- 시트에 저장 ----
  var sh = ensureSnapshotSheet_();
  var period = extractPeriodFromSheetName_(sheetName);
  var snapshotId = makeSnapshotId_(channelKey, sheetName);

  var last = sh.getLastRow();
  var updateRow = -1;
  if (last > 1) {
    var idCol = sh.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < idCol.length; i++) {
      if (idCol[i][0] === snapshotId) { updateRow = i + 2; break; }
    }
  }

  var row = [
    snapshotId,
    channelKey,
    sheetName,
    period,
    new Date(),
    'alive',
    dashSplit.chunks[0] || '',
    dashSplit.chunks[1] || '',
    dashSplit.chunks[2] || '',
    satSplit.chunks[0] || '',
    satSplit.chunks[1] || '',
    satSplit.chunks[2] || ''
  ];

  if (updateRow > 0) {
    sh.getRange(updateRow, 1, 1, SNAPSHOT_HEADERS.length).setValues([row]);
  } else {
    sh.appendRow(row);
  }

  return {
    ok: true,
    snapshotId: snapshotId,
    dashboardBytes: dashSplit.originalLen,
    satisfactionBytes: satSplit.originalLen,
    satisfactionRespondents: satRespondentsCount,
    jsonBytes: dashSplit.originalLen + satSplit.originalLen,
    chunks: dashSplit.chunks.length,
    truncated: dashSplit.truncated || satSplit.truncated
  };
}

/**
 * 스냅샷 조회 - 대시보드 + 만족도 응답자 JSON 복원
 * @param {string} which - 'dashboard' | 'satisfaction' | 'both' (기본 both)
 */
function api_getDashboardSnapshot(channelKey, sheetName, which) {
  which = which || 'both';
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SNAPSHOT_SHEET);
  if (!sh) return { ok: false, error: '스냅샷 시트 없음' };
  var last = sh.getLastRow();
  if (last < 2) return { ok: false, error: '스냅샷 없음' };

  var snapshotId = makeSnapshotId_(channelKey, sheetName);
  var vals = sh.getRange(2, 1, last - 1, SNAPSHOT_HEADERS.length).getValues();
  for (var i = 0; i < vals.length; i++) {
    if (vals[i][0] === snapshotId) {
      var result = {
        ok: true,
        fromSnapshot: true,
        collectedAt: vals[i][4],
        rawSheetState: vals[i][5]
      };
      if (which === 'dashboard' || which === 'both') {
        var dashStr = String(vals[i][6] || '') + String(vals[i][7] || '') + String(vals[i][8] || '');
        if (dashStr) {
          try { result.dashboard = JSON.parse(dashStr); }
          catch (e) { return { ok: false, error: '대시보드 JSON 파싱 실패: ' + e.message }; }
        }
      }
      if (which === 'satisfaction' || which === 'both') {
        var satStr = String(vals[i][9] || '') + String(vals[i][10] || '') + String(vals[i][11] || '');
        if (satStr) {
          try { result.satisfaction = JSON.parse(satStr); }
          catch (e) { /* 만족도가 없거나 파싱 실패는 무시 */ }
        }
      }
      return result;
    }
  }
  return { ok: false, error: '해당 시트 스냅샷 없음: ' + sheetName };
}

/**
 * 대시보드 로드 우선순위: raw 있으면 raw / 없으면 스냅샷
 */
function api_getDashboardDataOrSnapshot(channelKey, sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rawSh = ss.getSheetByName(sheetName);
  if (rawSh && rawSh.getLastRow() > 1) {
    var data = api_getDashboardData(channelKey, sheetName);
    if (data) data.__fromSnapshot = false;
    return data;
  }
  var snap = api_getDashboardSnapshot(channelKey, sheetName, 'dashboard');
  if (snap.ok && snap.dashboard) {
    var data = snap.dashboard;
    data.__fromSnapshot = true;
    data.__snapshotCollectedAt = snap.collectedAt ? Utilities.formatDate(new Date(snap.collectedAt), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') : '';
    data.__rawSheetState = snap.rawSheetState;
    return data;
  }
  return { error: '시트도 스냅샷도 없음' };
}

/**
 * 만족도 응답자 로드 우선순위: raw 있으면 실시간 스캔 / 없으면 스냅샷
 * 클라이언트가 이 함수 하나로 호출해서 raw 유무 상관없이 결과 얻음
 */
function api_getSatisfactionRespondentsOrSnapshot(channelKey, sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rawSh = ss.getSheetByName(sheetName);
  if (rawSh && rawSh.getLastRow() > 1) {
    var res = api_getSatisfactionRespondents(channelKey, sheetName);
    if (res) res.__fromSnapshot = false;
    return res;
  }
  var snap = api_getDashboardSnapshot(channelKey, sheetName, 'satisfaction');
  if (snap.ok && snap.satisfaction) {
    var res = snap.satisfaction;
    res.__fromSnapshot = true;
    res.__snapshotCollectedAt = snap.collectedAt ? Utilities.formatDate(new Date(snap.collectedAt), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') : '';
    return res;
  }
  return { ok: false, error: '시트도 만족도 스냅샷도 없음' };
}

/**
 * 만족도 스캔(수집) 라우터 — raw 없으면 no-op 반환 (스냅샷 모드에서 스캔 스킵)
 */
function api_syncSatisfactionFromChatsOrSkip(channelKey, sheetName, force) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rawSh = ss.getSheetByName(sheetName);
  if (!rawSh || rawSh.getLastRow() < 2) {
    return { ok: true, skipped: true, reason: 'raw 없음(스냅샷 모드)', stats: { scanned: 0, responses: 0 } };
  }
  return api_syncSatisfactionFromChats(channelKey, sheetName, force);
}

/**
 * 시트 압축 - 전사문 컬럼(K, 11번째)만 비우기. 대시보드 재계산 여전히 가능하지만 시트 크기 90% 감소
 */
function api_compressSheet(channelKey, sheetName) {
  if (!CHANNELS[channelKey]) return { ok: false, error: '알 수 없는 채널' };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(sheetName);
  if (!sh) return { ok: false, error: '시트 없음: ' + sheetName };
  var last = sh.getLastRow();
  if (last < 2) return { ok: false, error: '데이터 없음' };

  // 압축 전에 스냅샷 저장 (안전 마진)
  var snap = api_saveDashboardSnapshot(channelKey, sheetName);
  if (!snap.ok) return { ok: false, error: '스냅샷 저장 실패, 압축 중단: ' + snap.error };

  // 전사문 컬럼 = K = 11번째 컬럼 (0-based로는 10)
  // OUT_HEADERS: ['rpnTin','userName','bsno','tnmNm','careproMgr','bkpStatus','firstMsgAt','link','inflow','tags','전사문','rating','reviewedAt','firstResponder','lastResponder']
  var transcriptCol = 11;
  var range = sh.getRange(2, transcriptCol, last - 1, 1);
  range.clearContent();

  // 스냅샷 상태 업데이트
  updateSnapshotState_(channelKey, sheetName, 'compressed');

  return {
    ok: true,
    sheetName: sheetName,
    rowsAffected: last - 1,
    snapshotJsonBytes: snap.jsonBytes,
    note: '전사문 컬럼 비움 · 대시보드 계속 정상 작동'
  };
}

function updateSnapshotState_(channelKey, sheetName, state) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SNAPSHOT_SHEET);
  if (!sh) return;
  var last = sh.getLastRow();
  if (last < 2) return;
  var snapshotId = makeSnapshotId_(channelKey, sheetName);
  var idCol = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < idCol.length; i++) {
    if (idCol[i][0] === snapshotId) {
      sh.getRange(i + 2, 6).setValue(state);
      return;
    }
  }
}

/**
 * 이력 항목: 시트 상태 (raw alive/compressed/스냅샷만) 조회
 */
function api_getSheetStatus(channelKey, sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rawSh = ss.getSheetByName(sheetName);
  var rawExists = !!rawSh;
  var rawRows = rawExists ? Math.max(0, rawSh.getLastRow() - 1) : 0;

  var snap = ss.getSheetByName(SNAPSHOT_SHEET);
  var hasSnapshot = false;
  var snapshotState = null;
  var snapshotAt = null;
  if (snap && snap.getLastRow() > 1) {
    var snapshotId = makeSnapshotId_(channelKey, sheetName);
    var vals = snap.getRange(2, 1, snap.getLastRow() - 1, 6).getValues();
    for (var i = 0; i < vals.length; i++) {
      if (vals[i][0] === snapshotId) {
        hasSnapshot = true;
        snapshotState = String(vals[i][5] || '');
        snapshotAt = vals[i][4];
        break;
      }
    }
  }

  // 전사문 상태 추정: 첫 5행 전사문 컬럼(K, 인덱스 10) 비어 있으면 compressed
  var compressed = false;
  if (rawExists && rawRows > 0) {
    var sample = rawSh.getRange(2, 11, Math.min(5, rawRows), 1).getValues();
    var nonEmpty = sample.filter(function(r) { return String(r[0] || '').trim(); }).length;
    if (nonEmpty === 0) compressed = true;
  }

  return {
    ok: true,
    rawExists: rawExists,
    rawRows: rawRows,
    compressed: compressed,
    hasSnapshot: hasSnapshot,
    snapshotState: snapshotState,
    snapshotAt: snapshotAt ? Utilities.formatDate(new Date(snapshotAt), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') : ''
  };
}

/**
 * 스냅샷 시스템 self-test — 대시보드 · 압축 · 스냅샷 복원까지 안전하게 QA
 * 사용자 액션 최소화 (한 번 실행하면 리포트 반환)
 * 원본 시트에 파괴적 변경 없이 검증
 */
function runSnapshotSelfTest() {
  var report = [];
  function log(msg) { report.push(msg); Logger.log(msg); }
  var pass = 0, fail = 0;
  function assert(cond, name) {
    if (cond) { pass++; log('✅ ' + name); }
    else { fail++; log('❌ ' + name); }
  }

  log('=== 스냅샷 시스템 Self-Test 시작 ===');

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // 최근 환급 수집 시트 하나 찾기
  var target = null;
  var suffix = '_' + CHANNELS.refund.label;
  var sheets = ss.getSheets().filter(function(s) {
    var name = s.getName();
    return /^\d{8}~\d{8}_/.test(name) && name.substr(name.length - suffix.length) === suffix;
  });
  if (!sheets.length) { log('❌ 환급 수집 시트 없음 — 테스트 중단'); return { ok: false, report: report }; }
  target = sheets[0].getName();
  log('🎯 대상 시트: ' + target);

  // Step 1: 대시보드 데이터 계산
  log('\n--- Step 1: api_getDashboardData 실행 ---');
  var t1 = new Date().getTime();
  var dashOriginal;
  try {
    dashOriginal = api_getDashboardData('refund', target);
    var t1e = new Date().getTime() - t1;
    log('실시간 계산 완료 · ' + t1e + 'ms');
    assert(!!dashOriginal, '대시보드 데이터 반환');
    assert(!!(dashOriginal.kpi && dashOriginal.kpi.totalChats), 'kpi.totalChats 존재 (' + (dashOriginal.kpi && dashOriginal.kpi.totalChats) + ')');
    assert(!!dashOriginal.topTags, 'topTags 존재');
    assert(!!dashOriginal.dowDistribution, 'dowDistribution 존재');
  } catch (e) {
    log('❌ 대시보드 계산 실패: ' + e.message); fail++;
    return { ok: false, report: report };
  }

  // Step 2: 스냅샷 저장
  log('\n--- Step 2: 스냅샷 저장 ---');
  var saveRes = api_saveDashboardSnapshot('refund', target);
  assert(saveRes.ok, '스냅샷 저장 성공');
  assert(saveRes.jsonBytes > 0, 'JSON 크기 > 0 (' + saveRes.jsonBytes + ' bytes)');
  assert(saveRes.chunks >= 1, '청크 개수 >= 1 (' + saveRes.chunks + ')');
  assert(!saveRes.truncated, '잘림 없음 (135KB 이내)');

  // Step 3: 스냅샷 조회 (대시보드)
  log('\n--- Step 3: 스냅샷 조회 (대시보드) ---');
  var loadRes = api_getDashboardSnapshot('refund', target, 'dashboard');
  assert(loadRes.ok, '스냅샷 조회 성공');
  assert(!!loadRes.dashboard, '대시보드 데이터 존재');
  assert((loadRes.dashboard.kpi && loadRes.dashboard.kpi.totalChats) === (dashOriginal.kpi && dashOriginal.kpi.totalChats), 'kpi.totalChats 일치');
  assert(JSON.stringify(loadRes.dashboard.topTags) === JSON.stringify(dashOriginal.topTags), 'topTags 일치');
  assert(JSON.stringify(loadRes.dashboard.dowDistribution) === JSON.stringify(dashOriginal.dowDistribution), 'dowDistribution 일치');

  // Step 3b: 만족도 스냅샷 조회
  log('\n--- Step 3b: 만족도 스냅샷 조회 ---');
  var satLoad = api_getDashboardSnapshot('refund', target, 'satisfaction');
  assert(satLoad.ok, '만족도 스냅샷 조회 성공');
  if (satLoad.satisfaction) {
    var count = (satLoad.satisfaction.respondents || []).length;
    assert(count > 0, '응답자 리스트 존재 (' + count + '명)');
    assert(!!satLoad.satisfaction.respondents[0].chatId, '응답자 chatId 있음');
  } else {
    log('⚠️ 만족도 데이터 없음 (환급 채널이지만 스냅샷 저장 시 만족도 API 실패했을 가능성)');
  }

  // Step 4: 라우터 함수 (raw 있음)
  log('\n--- Step 4: 라우터 (raw 있음 → 실시간) ---');
  var routerRes = api_getDashboardDataOrSnapshot('refund', target);
  assert(routerRes.__fromSnapshot === false, 'raw 있음 → 실시간 계산 (스냅샷 아님)');
  assert((routerRes.kpi && routerRes.kpi.totalChats) === (dashOriginal.kpi && dashOriginal.kpi.totalChats), 'kpi.totalChats 일치');

  // Step 5: 시트 상태 조회
  log('\n--- Step 5: 시트 상태 조회 ---');
  var status = api_getSheetStatus('refund', target);
  assert(status.rawExists, 'rawExists = true');
  assert(status.hasSnapshot, 'hasSnapshot = true');
  assert(status.rawRows > 0, 'rawRows > 0 (' + status.rawRows + ')');
  assert(!status.compressed, '아직 압축 안 됨');

  // Step 6: 압축 시뮬레이션 (실제 압축은 파괴적이라 안 함, 대신 함수 시그니처만 확인)
  log('\n--- Step 6: 압축 함수 signature 확인 (실행 X) ---');
  assert(typeof api_compressSheet === 'function', 'api_compressSheet 함수 존재');
  assert(typeof api_getDashboardDataOrSnapshot === 'function', '라우터 함수 존재');

  // Step 7: 파괴적 시뮬레이션 (라우터 fallback 확인) — 존재하지 않는 시트명으로 호출
  log('\n--- Step 7: 라우터 fallback (없는 시트) ---');
  var ghostSheet = '__NONEXISTENT_' + target;
  var ghostRes = api_getDashboardDataOrSnapshot('refund', ghostSheet);
  assert(!!ghostRes && ghostRes.error, '없는 시트 → 스냅샷도 없음 → error 반환');

  log('\n=== 완료 ===');
  log('통과: ' + pass + ' / 실패: ' + fail);

  return {
    ok: fail === 0,
    pass: pass,
    fail: fail,
    report: report,
    dashOriginalSize: JSON.stringify(dashOriginal).length,
    snapshotSaveResult: saveRes
  };
}

// 특정 시트로 바로 이동하는 URL (gid 포함)
function api_getSheetUrl(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(sheetName);
  if (!sh) return { ok: false, error: '시트 없음: ' + sheetName };
  var url = ss.getUrl().replace(/\/edit.*$/, '/edit#gid=' + sh.getSheetId());
  return { ok: true, url: url };
}

// 대시보드 데이터 기반 마크다운 리포트 생성 (Slack 붙여넣기용)
function api_generateReportText(channelKey, sheetName) {
  var data = api_getDashboardData(channelKey, sheetName);
  if (!data) return { ok: false, error: '데이터 없음' };
  var lines = [];
  var chLabel = data.channelLabel || channelKey;
  var scope = sheetName || '전체 누적';
  var k = data.kpi;
  lines.push('📊 [' + chLabel + '] ' + scope + ' VOC 리포트');
  lines.push('생성: ' + data.generatedAt);
  lines.push('');
  lines.push('■ 전체 현황');
  lines.push('  · 총 상담 ' + k.totalChats + '건');
  lines.push('  · 태그 입력률 ' + (k.tagInputRate || 0) + '% (' + k.chatsWithTag + '/' + k.totalChats + ')');
  lines.push('  · 해지 관련 문의 ' + k.cancelChatCount + '건');
  lines.push('  · 미트 활용 ' + k.meetChatCount + '건 (' + (k.meetChatRate || 0) + '%)');
  lines.push('');
  if (data.topTags && data.topTags.length) {
    lines.push('■ TOP 태그');
    data.topTags.slice(0, 10).forEach(function(t) { lines.push('  · ' + t.name + ': ' + t.count + '건'); });
    lines.push('');
  }
  if (data.cancelTags && data.cancelTags.length) {
    lines.push('■ 해지 상세 태그');
    data.cancelTags.forEach(function(t) { lines.push('  · ' + t.name + ': ' + t.count + '건'); });
    lines.push('');
  }
  if (data.dowDistribution) {
    lines.push('■ 요일별 문의');
    data.dowDistribution.forEach(function(d) { lines.push('  · ' + d.name + ': ' + d.count + '건'); });
    lines.push('');
  }
  if (data.teamList && data.teamList.length) {
    lines.push('■ 팀별 문의 건수 (미트 병기)');
    data.teamList.forEach(function(t) {
      lines.push('  · ' + t.name + ': ' + t.count + '건' + (t.meet ? ' (미트 ' + t.meet + ')' : ''));
    });
    lines.push('');
  }
  if (data.flow && data.flow.totalWithFlow) {
    lines.push('■ 상담 흐름');
    lines.push('  · 완결 데이터 ' + data.flow.totalWithFlow + '/' + data.flow.totalChats + ' (' + (data.flow.coverage || 0) + '%)');
    lines.push('  · 자체 종결율 ' + (100 - (data.flow.handoffRate || 0)) + '% · 타팀 이관 ' + data.flow.handoffCount + '건');
    lines.push('');
  }
  if (data.cancelMatch && data.cancelMatch.ok) {
    var cm = data.cancelMatch.totals;
    lines.push('■ 해지 매칭 분석 (채널톡 태그 vs 실제 해지)');
    lines.push('  · 실제 해지 ' + cm.actualCancels + '건 / 태그 대화 ' + cm.taggedChats + '건 / 매칭 ' + cm.matched + '건');
    lines.push('  · 태그 커버리지 ' + cm.tagCoverage + '% · 태그 정확도 ' + cm.tagPrecision + '%');
    lines.push('  · 태그 누락(태깅 안 됨) ' + cm.wfOnly + '건 · 태그만(실제 해지 없음) ' + cm.tagOnly + '건');
    lines.push('');
  }
  return { ok: true, text: lines.join('\n') };
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

  // 매니저 → 팀 매핑 로드 (_매니저_팀매핑 시트)
  var mgrToTeam = loadManagerTeamMap(ss);

  var totalChats = 0;
  var totalTags = 0;
  var chatsWithTag = 0;
  var tagCounts = {};
  var inflowCounts = {};
  var managerCounts = {};
  var ratingSum = 0;
  var ratingCount = 0;
  var ratingDist = {};
  var respondents = [];

  // 요일별 · 팀별 · 담당자별 (전체 및 미트 활용)
  var byDow = { '월': 0, '화': 0, '수': 0, '목': 0, '금': 0, '토': 0, '일': 0 };
  var dowKr = ['일', '월', '화', '수', '목', '금', '토'];
  var byTeam = {};
  var byMgr = {};
  var byTeamMeet = {};
  var byMgrMeet = {};

  // 진입/종결 팀 흐름 분석 (첫응대자 팀 → 최종응대자 팀)
  var byEntryTeam = {};   // 첫응대자 팀 카운트
  var byCloserTeam = {};  // 최종응대자 팀 카운트
  var flowMatrix = {};    // 'entry→closer' -> count
  var flowChatsCollected = 0;
  var flowHandoffCount = 0; // entry !== closer 대화 수

  // 해지 관련 태그 집계 (태그명에 '해지' 또는 '폐업' 포함)
  var cancelTagCounts = {};
  var cancelChatCount = 0;
  var CANCEL_RE = /(해지|폐업|해약)/;
  var cancelTaggedBsnos = {}; // bsno -> { rep, bizName, tags, link, startTime }

  // 팀 미매칭 진단 (매니저 이름은 있는데 _매니저_팀매핑에 미등록)
  var unmatchedByName = {}; // rawName -> { count, normalized }

  // 환급 채널 전용: ALF 세그먼트 · 취소·탈퇴 태그
  var alfClosedCount = 0;     // 담당자 필드 공백 = ALF 종결
  var agentAssignedCount = 0; // 담당자 필드 존재 = 상담사 배정
  var dailyAlfMap = {};       // 'YYYY-MM-DD' → { closed, assigned }
  var CANCEL_ONLY_RE = /취소/;
  var WITHDRAW_RE = /탈퇴/;
  var refundCancelCount = 0;
  var refundWithdrawCount = 0;
  var refundCancelDaily = {};
  var refundWithdrawDaily = {};

  // 담당자 없음 / 미트 인바운드 / 팀 미매칭 대화 상세 (검증용, 상위 200건까지)
  var noMgrChats = [];
  var meetInboundChats = [];
  var unmatchedChats = [];

  // enrichment 필요 카운트 (기장담당자 empty AND 상담담당자 empty)
  var emptyResponderCount = 0;

  outputSheets.forEach(function(sh) {
    var last = sh.getLastRow();
    if (last < 2) return;
    var numCols = Math.max(sh.getLastColumn(), CFG.OUT_HEADERS.length);
    var vals = sh.getRange(2, 1, last - 1, numCols).getValues();
    for (var i = 0; i < vals.length; i++) {
      var row = vals[i];
      totalChats++;

      var repName = String(row[1] || '').trim();
      var bsnoRaw = normalizeBsno(row[2]);
      var bizName = String(row[3] || '').trim();
      var bkpManager = String(row[4] || '').trim();     // 기장담당자 (CarePro 배정)
      var startTime = row[6];
      var deskLink = String(row[7] || '').trim();
      var inflow = String(row[8] || '').trim() || '미확인';
      var tagsRaw = String(row[9] || '').trim();
      var rating = row[11];
      var reviewedAt = String(row[12] || '').trim();
      var firstResponder = String(row[13] || '').trim();  // 첫응대자 (진입)
      var lastResponder = String(row[14] || '').trim();   // 최종응대자 (종결)
      var chatResponder = lastResponder || firstResponder; // fallback 통합값
      // enrichment 필요 판정: (응대자 정보 아예 없음) 또는 (첫만 있고 최종 없음: 레거시 or 미완결)
      if ((!firstResponder && !lastResponder) || (firstResponder && !lastResponder)) emptyResponderCount++;

      // 팀 매핑: 기장담당자 우선, 없으면 상담담당자 fallback
      var manager = bkpManager || chatResponder;
      var managerSource = bkpManager ? 'bkp' : (chatResponder ? 'chat' : '');
      var teamKey;
      if (manager) {
        managerCounts[manager] = (managerCounts[manager] || 0) + 1;
        byMgr[manager] = (byMgr[manager] || 0) + 1;
        var normMgr = normalizeMgrName(manager);
        var mappedTeam = mgrToTeam[normMgr];
        if (mappedTeam) {
          teamKey = mappedTeam;
        } else {
          teamKey = '(팀 미매칭)';
          if (!unmatchedByName[manager]) unmatchedByName[manager] = { count: 0, normalized: normMgr };
          unmatchedByName[manager].count++;
          if (unmatchedChats.length < 200) {
            unmatchedChats.push({
              startTime: startTime instanceof Date ? Utilities.formatDate(startTime, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') : String(startTime || ''),
              rep: repName,
              bizName: bizName,
              manager: manager,
              managerSource: managerSource,
              inflow: inflow,
              tags: tagsRaw,
              link: deskLink,
              chatId: extractChatIdFromLink(deskLink)
            });
          }
        }
      } else {
        // 매칭 실패 케이스: 미트 인바운드는 별도 라벨링·리스트, 그 외는 담당자 없음
        var isMeetInbound = (inflow === '미트');
        teamKey = isMeetInbound ? '미트(인바운드)' : '(담당자 없음)';
        var chatEntry = {
          startTime: startTime instanceof Date ? Utilities.formatDate(startTime, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') : String(startTime || ''),
          rep: repName,
          bizName: bizName,
          inflow: inflow,
          tags: tagsRaw,
          link: deskLink,
          chatId: extractChatIdFromLink(deskLink)
        };
        if (isMeetInbound) {
          if (meetInboundChats.length < 200) meetInboundChats.push(chatEntry);
        } else {
          if (noMgrChats.length < 200) noMgrChats.push(chatEntry);
        }
      }
      byTeam[teamKey] = (byTeam[teamKey] || 0) + 1;

      // 환급 채널 전용: ALF 세그먼트 판정 (담당자 필드 공백 여부)
      var dateKey = '';
      if (startTime) {
        var d0 = (startTime instanceof Date) ? startTime : new Date(startTime);
        if (!isNaN(d0.getTime())) dateKey = Utilities.formatDate(d0, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }
      if (channelKey === 'refund') {
        var isAlfClosed = !bkpManager && !chatResponder;
        if (isAlfClosed) alfClosedCount++;
        else agentAssignedCount++;
        if (dateKey) {
          if (!dailyAlfMap[dateKey]) dailyAlfMap[dateKey] = { closed: 0, assigned: 0 };
          if (isAlfClosed) dailyAlfMap[dateKey].closed++;
          else dailyAlfMap[dateKey].assigned++;
        }
      }

      // 진입/종결 팀 흐름 분석 (첫응대자 팀 → 최종응대자 팀)
      // 둘 다 있는 경우만 매트릭스에 반영. 하나만 있으면 정보 부족으로 별도 카운트.
      if (firstResponder && lastResponder) {
        var entryTeam = mgrToTeam[normalizeMgrName(firstResponder)] || '(팀 미매칭)';
        var closerTeam = mgrToTeam[normalizeMgrName(lastResponder)] || '(팀 미매칭)';
        byEntryTeam[entryTeam] = (byEntryTeam[entryTeam] || 0) + 1;
        byCloserTeam[closerTeam] = (byCloserTeam[closerTeam] || 0) + 1;
        var fkey = entryTeam + '→' + closerTeam;
        flowMatrix[fkey] = (flowMatrix[fkey] || 0) + 1;
        flowChatsCollected++;
        if (entryTeam !== closerTeam) flowHandoffCount++;
      }

      // 미트 활용
      if (inflow === '미트') {
        byTeamMeet[teamKey] = (byTeamMeet[teamKey] || 0) + 1;
        if (manager) byMgrMeet[manager] = (byMgrMeet[manager] || 0) + 1;
      }

      inflowCounts[inflow] = (inflowCounts[inflow] || 0) + 1;

      // 요일별
      if (startTime) {
        var d = (startTime instanceof Date) ? startTime : new Date(startTime);
        if (!isNaN(d.getTime())) byDow[dowKr[d.getDay()]]++;
      }

      if (tagsRaw) {
        chatsWithTag++;
        var tags = tagsRaw.split(',').map(function(t) { return t.trim(); }).filter(Boolean);
        totalTags += tags.length;
        var hitCancel = false;
        var hitCancelOnly = false;   // 환급: '취소' 단독 (탈퇴 아님)
        var hitWithdraw = false;      // 환급: '탈퇴'
        for (var t = 0; t < tags.length; t++) {
          var tag = tags[t];
          tagCounts[tag] = (tagCounts[tag] || 0) + 1;
          if (CANCEL_RE.test(tag)) {
            cancelTagCounts[tag] = (cancelTagCounts[tag] || 0) + 1;
            hitCancel = true;
          }
          if (CANCEL_ONLY_RE.test(tag) && !WITHDRAW_RE.test(tag)) hitCancelOnly = true;
          if (WITHDRAW_RE.test(tag)) hitWithdraw = true;
        }
        if (hitCancel) {
          cancelChatCount++;
          if (bsnoRaw && !cancelTaggedBsnos[bsnoRaw]) {
            cancelTaggedBsnos[bsnoRaw] = {
              rep: repName, bizName: bizName, tags: tagsRaw, link: deskLink,
              startTime: startTime instanceof Date ? Utilities.formatDate(startTime, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') : String(startTime || ''),
              bkpManager: bkpManager
            };
          }
        }
        // 환급 채널: 취소·탈퇴 각각 집계 + 일별
        if (channelKey === 'refund') {
          if (hitCancelOnly) {
            refundCancelCount++;
            if (dateKey) refundCancelDaily[dateKey] = (refundCancelDaily[dateKey] || 0) + 1;
          }
          if (hitWithdraw) {
            refundWithdrawCount++;
            if (dateKey) refundWithdrawDaily[dateKey] = (refundWithdrawDaily[dateKey] || 0) + 1;
          }
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
          startTime: startTime instanceof Date ? Utilities.formatDate(startTime, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') : String(startTime || ''),
          reviewedAt: reviewedAt,
          link: deskLink
        });
      }
    }
  });

  var topTags = Object.keys(tagCounts).map(function(k) {
    return { name: k, count: tagCounts[k] };
  }).sort(function(a, b) { return b.count - a.count; }).slice(0, 10);

  // ---- 실제 해지 워크플로 매칭 (케어 채널만)
  var cancelMatch = null;
  if (channelKey === 'care') {
    cancelMatch = buildCancelMatchAnalysis(sheetName, cancelTaggedBsnos);
  }

  var inflowList = Object.keys(inflowCounts).map(function(k) {
    return { name: k, count: inflowCounts[k] };
  }).sort(function(a, b) { return b.count - a.count; });

  var managerList = Object.keys(managerCounts).map(function(k) {
    return { name: k, count: managerCounts[k] };
  }).sort(function(a, b) { return b.count - a.count; });

  // 팀별 · 담당자별 리스트 (미트 활용 수 병합)
  var teamList = Object.keys(byTeam).map(function(k) {
    return { name: k, count: byTeam[k], meet: byTeamMeet[k] || 0 };
  }).sort(function(a, b) { return b.count - a.count; });

  var managerListMeet = managerList.map(function(m) {
    var norm = normalizeMgrName(m.name);
    var team = mgrToTeam[norm] || '';
    return {
      name: displayMgrName(m.name),
      rawName: m.name,
      team: team,
      count: m.count,
      meet: byMgrMeet[m.name] || 0
    };
  });

  var cancelTagList = Object.keys(cancelTagCounts).map(function(k) {
    return { name: k, count: cancelTagCounts[k] };
  }).sort(function(a, b) { return b.count - a.count; });

  var dowList = ['월', '화', '수', '목', '금', '토', '일'].map(function(d) {
    return { name: d, count: byDow[d] };
  });

  var meetTotal = inflowCounts['미트'] || 0;

  respondents.sort(function(a, b) { return b.rating - a.rating; });

  var history = getExecutionHistory(ss, channelKey);
  var executedCount = history.length;
  var executedSum = history.reduce(function(acc, h) { return acc + (h.collected || 0); }, 0);

  // ---- LLM 분류 결과 (Phase 2, sheetName이 지정된 경우만)
  var vocClassification = null;
  if (sheetName) {
    var vocSheet = ss.getSheetByName('_VOC_분류_' + sheetName);
    if (vocSheet && vocSheet.getLastRow() >= 2) {
      var lastR = vocSheet.getLastRow();
      var vocVals = vocSheet.getRange(2, 1, lastR - 1, 7).getValues();
      var catCounts = {};
      var samplesByCat = {};
      vocVals.forEach(function(vr) {
        var cat = String(vr[1] || '').trim();
        if (!cat) return;
        catCounts[cat] = (catCounts[cat] || 0) + 1;
        if (!samplesByCat[cat]) samplesByCat[cat] = [];
        if (samplesByCat[cat].length < 5) {
          samplesByCat[cat].push({
            quote: String(vr[2] || '').trim(),
            date: String(vr[3] || '').trim(),
            rep: String(vr[4] || '').trim(),
            manager: String(vr[5] || '').trim(),
            tags: String(vr[6] || '').trim()
          });
        }
      });
      var catList = Object.keys(catCounts).map(function(k) {
        return { name: k, count: catCounts[k], samples: samplesByCat[k] || [] };
      }).sort(function(a, b) { return b.count - a.count; });
      vocClassification = {
        totalClassified: vocVals.length,
        categories: catList
      };
    }
  }

  // 리포트 시트 존재 여부 (있으면 대시보드에서 바로 열기 가능)
  var reportSheetName = sheetName ? ('_리포트_' + sheetName) : '';
  var reportExists = reportSheetName && ss.getSheetByName(reportSheetName) ? true : false;

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
      ratingResponseRate: totalChats > 0 ? Math.round(ratingCount / totalChats * 1000) / 10 : 0,
      chatsWithTag: chatsWithTag,
      tagInputRate: totalChats > 0 ? Math.round(chatsWithTag / totalChats * 1000) / 10 : 0,
      cancelChatCount: cancelChatCount,
      meetChatCount: meetTotal,
      meetChatRate: totalChats > 0 ? Math.round(meetTotal / totalChats * 1000) / 10 : 0,
      emptyResponderCount: emptyResponderCount
    },
    topTags: topTags,
    cancelTags: cancelTagList,
    dowDistribution: dowList,
    teamList: teamList,
    managerListMeet: managerListMeet.slice(0, 30),
    inflowDistribution: inflowList,
    managerList: managerList.slice(0, 20),
    ratingDistribution: ratingDist,
    respondents: respondents.slice(0, 100),
    history: history.slice(0, 10),
    vocClassification: vocClassification,
    reportSheet: reportExists ? reportSheetName : '',
    unmatchedManagers: Object.keys(unmatchedByName).map(function(k) {
      return { name: k, normalized: unmatchedByName[k].normalized, count: unmatchedByName[k].count };
    }).sort(function(a, b) { return b.count - a.count; }),
    noMgrChats: noMgrChats,
    meetInboundChats: meetInboundChats,
    unmatchedChats: unmatchedChats,
    cancelMatch: cancelMatch,
    refund: (channelKey === 'refund') ? {
      alfSegment: {
        alfClosed: alfClosedCount,
        agentAssigned: agentAssignedCount,
        total: alfClosedCount + agentAssignedCount,
        alfClosedRate: (alfClosedCount + agentAssignedCount) > 0 ? Math.round(alfClosedCount / (alfClosedCount + agentAssignedCount) * 1000) / 10 : 0,
        agentAssignedRate: (alfClosedCount + agentAssignedCount) > 0 ? Math.round(agentAssignedCount / (alfClosedCount + agentAssignedCount) * 1000) / 10 : 0,
        dailyTrend: Object.keys(dailyAlfMap).sort().map(function(d) {
          return { date: d, closed: dailyAlfMap[d].closed, assigned: dailyAlfMap[d].assigned };
        })
      },
      cancelWithdraw: {
        cancelCount: refundCancelCount,
        withdrawCount: refundWithdrawCount,
        cancelDaily: Object.keys(refundCancelDaily).sort().map(function(d) { return { date: d, count: refundCancelDaily[d] }; }),
        withdrawDaily: Object.keys(refundWithdrawDaily).sort().map(function(d) { return { date: d, count: refundWithdrawDaily[d] }; })
      }
    } : null,
    flow: {
      entryTeams: Object.keys(byEntryTeam).map(function(k) { return { name: k, count: byEntryTeam[k] }; }).sort(function(a, b) { return b.count - a.count; }),
      closerTeams: Object.keys(byCloserTeam).map(function(k) { return { name: k, count: byCloserTeam[k] }; }).sort(function(a, b) { return b.count - a.count; }),
      matrix: Object.keys(flowMatrix).map(function(k) {
        var parts = k.split('→');
        return { entry: parts[0], closer: parts[1], count: flowMatrix[k] };
      }).sort(function(a, b) { return b.count - a.count; }),
      totalWithFlow: flowChatsCollected,
      totalChats: totalChats,
      coverage: totalChats > 0 ? Math.round(flowChatsCollected / totalChats * 1000) / 10 : 0,
      handoffCount: flowHandoffCount,
      handoffRate: flowChatsCollected > 0 ? Math.round(flowHandoffCount / flowChatsCollected * 1000) / 10 : 0
    }
  };
}

// 하드코딩 오버라이드: CarePro 원본 이름 표기가 채널톡 매핑 시트와 달라 매칭 실패하는 케이스.
// 시트가 리셋돼도 살아남고, 시트 매핑보다 우선순위 높음. 정규화된 이름(normalizeMgrName 결과) 기준.
// NOTE: 저장소 버전에서 실제 매니저 이름은 마스킹.
// 오버라이드 맵은 채널톡의 팀 정보와 실제 소속이 어긋난 소수 케이스를 이름 기준으로 강제 매칭하기 위해 사용.
var MGR_TEAM_OVERRIDES = {
  '<MGR_A_NAME>': '<TEAM_A>',
  '<MGR_B_NAME>': '<CX_TEAM>'
  // ...
};

// 이름에서 부수 표기(공백/직급/괄호/젠트/팀 접두사) 제거해 표시용 원이름만 남김
function displayMgrName(raw) {
  if (!raw) return '';
  var s = String(raw).trim();
  s = s.replace(/^\d+팀\s+/, '');
  s = s.replace(/^협세\s+/, '');
  s = s.replace(/^젠트\s+/, '');
  s = s.replace(/\s*\([^)]+\)$/, '');
  s = s.replace(/\s+(매니저|팀장|사원|선임|주임|세무사|부장|총괄|대리)$/, '');
  return s.trim();
}

// 실제 해지 데이터 vs 채널톡 해지 태그 매칭 분석
// - sheetName의 기간(YYYYMMDD~YYYYMMDD) 안의 해지처리일자만 대상
// - 매칭 키: 사업자번호 (bsno)
// - cancelTaggedBsnos: { bsno -> {rep, bizName, tags, link, startTime, bkpManager} }
function buildCancelMatchAnalysis(sheetName, cancelTaggedBsnos) {
  var period = null;
  if (sheetName) {
    var pm = sheetName.match(/^(\d{4})(\d{2})(\d{2})~(\d{4})(\d{2})(\d{2})_/);
    if (pm) {
      var fromMs = new Date(Number(pm[1]), Number(pm[2]) - 1, Number(pm[3]), 0, 0, 0).getTime();
      var toMs = new Date(Number(pm[4]), Number(pm[5]) - 1, Number(pm[6]), 23, 59, 59).getTime();
      period = { fromMs: fromMs, toMs: toMs, fromYmd: pm[1] + '-' + pm[2] + '-' + pm[3], toYmd: pm[4] + '-' + pm[5] + '-' + pm[6] };
    }
  }

  var wfAll = loadCancelWorkflow();
  if (!wfAll.length) return { ok: false, error: '해지 워크플로 시트 로드 실패 또는 데이터 없음' };

  // 팀 매핑 로드 (담당자 이름 → 팀)
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var teamMap = loadManagerTeamMap(ss);
  function teamOf(name) {
    if (!name) return '';
    var norm = normalizeMgrName(name);
    return teamMap[norm] || '';
  }

  // 기간 내 해지 이벤트만 필터
  var wf = wfAll.filter(function(r) {
    if (!period) return true;
    var d = r.processedAt || r.requestedAt;
    if (!d) return false;
    var t = d.getTime();
    return t >= period.fromMs && t <= period.toMs;
  });

  // 사업자번호 → 워크플로 레코드
  var wfByBsno = {};
  wf.forEach(function(r) {
    if (r.bsno) wfByBsno[r.bsno] = r;
    if (r.bsnoAlt) wfByBsno[r.bsnoAlt] = r;
  });

  // 카테고리별 매칭
  var matched = []; // 태그 있고 실제 해지 있음
  var tagOnly = []; // 태그 있는데 실제 해지 없음
  var wfOnly = [];  // 실제 해지인데 태그 없음

  Object.keys(cancelTaggedBsnos).forEach(function(bsno) {
    var chatInfo = cancelTaggedBsnos[bsno];
    var wfRec = wfByBsno[bsno];
    if (wfRec) {
      var wfMgr = displayMgrName(wfRec.bkpMgr);
      matched.push({
        bsno: bsno, rep: chatInfo.rep, bizName: chatInfo.bizName,
        tags: chatInfo.tags, chatLink: chatInfo.link,
        wfReason: wfRec.reasonCat, wfRequester: wfRec.requester,
        wfBkpMgr: wfMgr, wfBkpMgrTeam: teamOf(wfMgr),
        wfProcessedAt: wfRec.processedAt ? Utilities.formatDate(wfRec.processedAt, Session.getScriptTimeZone(), 'yyyy-MM-dd') : ''
      });
    } else {
      var chatMgr = displayMgrName(chatInfo.bkpManager);
      tagOnly.push({
        bsno: bsno, rep: chatInfo.rep, bizName: chatInfo.bizName,
        tags: chatInfo.tags, chatLink: chatInfo.link, startTime: chatInfo.startTime,
        bkpManager: chatMgr, bkpManagerTeam: teamOf(chatMgr)
      });
    }
  });

  wf.forEach(function(r) {
    if (!cancelTaggedBsnos[r.bsno] && (!r.bsnoAlt || !cancelTaggedBsnos[r.bsnoAlt])) {
      var wfMgr2 = displayMgrName(r.bkpMgr);
      wfOnly.push({
        bsno: r.bsno, rep: r.name, wfReason: r.reasonCat, wfRequester: r.requester,
        wfBkpMgr: wfMgr2, wfBkpMgrTeam: teamOf(wfMgr2),
        chatLink: r.chatLink,
        wfProcessedAt: r.processedAt ? Utilities.formatDate(r.processedAt, Session.getScriptTimeZone(), 'yyyy-MM-dd') : ''
      });
    }
  });

  // 실제 해지 breakdown: 사유별, 요청자별, 기장담당자별 (이름 정규화 후 집계)
  // 총합이 전체 wf 건수와 일치하도록 빈 값은 '(미기재)' 버킷으로 카운트
  var byReason = {};
  var byRequester = {};
  var byWfBkpMgr = {}; // name -> { count, team }
  wf.forEach(function(r) {
    var reasonKey = r.reasonCat || '(미기재)';
    byReason[reasonKey] = (byReason[reasonKey] || 0) + 1;

    var reqKey = r.requester || '(미기재)';
    byRequester[reqKey] = (byRequester[reqKey] || 0) + 1;

    var nm = r.bkpMgr ? displayMgrName(r.bkpMgr) : '(미기재)';
    if (!byWfBkpMgr[nm]) byWfBkpMgr[nm] = { count: 0, team: r.bkpMgr ? teamOf(nm) : '' };
    byWfBkpMgr[nm].count++;
  });

  // 태그 누락(wfOnly) 담당자별 breakdown (이름 정규화 후)
  var tagMissingByBkpMgr = {}; // name -> { count, team }
  wfOnly.forEach(function(r) {
    var k = r.wfBkpMgr || '(담당자 없음)';
    if (!tagMissingByBkpMgr[k]) tagMissingByBkpMgr[k] = { count: 0, team: r.wfBkpMgrTeam || '' };
    tagMissingByBkpMgr[k].count++;
  });

  return {
    ok: true,
    period: period ? { fromYmd: period.fromYmd, toYmd: period.toYmd } : null,
    totals: {
      taggedChats: Object.keys(cancelTaggedBsnos).length,
      actualCancels: wf.length,
      matched: matched.length,
      tagOnly: tagOnly.length,
      wfOnly: wfOnly.length,
      tagCoverage: wf.length > 0 ? Math.round(matched.length / wf.length * 1000) / 10 : 0,
      tagPrecision: Object.keys(cancelTaggedBsnos).length > 0 ? Math.round(matched.length / Object.keys(cancelTaggedBsnos).length * 1000) / 10 : 0
    },
    byReason: Object.keys(byReason).map(function(k) { return { name: k, count: byReason[k] }; }).sort(function(a, b) { return b.count - a.count; }),
    byRequester: Object.keys(byRequester).map(function(k) { return { name: k, count: byRequester[k] }; }).sort(function(a, b) { return b.count - a.count; }),
    byWfBkpMgr: Object.keys(byWfBkpMgr).map(function(k) { return { name: k, count: byWfBkpMgr[k].count, team: byWfBkpMgr[k].team }; }).sort(function(a, b) { return b.count - a.count; }),
    tagMissingByBkpMgr: Object.keys(tagMissingByBkpMgr).map(function(k) { return { name: k, count: tagMissingByBkpMgr[k].count, team: tagMissingByBkpMgr[k].team }; }).sort(function(a, b) { return b.count - a.count; }),
    matchedList: matched.slice(0, 100),
    tagOnlyList: tagOnly.slice(0, 100),
    wfOnlyList: wfOnly.slice(0, 100)
  };
}

// 외부 해지 워크플로 시트 로드
// 반환: [{ ts, name, rpnTin, bsno, bsnoAlt, success, chatLink, requestedAt, processedAt, firstPaidAt, reasonCat, reasonDetail, bkpMgr, requester }]
function loadCancelWorkflow() {
  try {
    var ext = SpreadsheetApp.openById(CANCEL_WF_SHEET_ID);
    var sh = CANCEL_WF_TAB ? ext.getSheetByName(CANCEL_WF_TAB) : ext.getSheets()[0];
    if (!sh) return [];
    var last = sh.getLastRow();
    if (last < 2) return [];
    var vals = sh.getRange(2, 1, last - 1, 14).getValues();
    var out = [];
    for (var i = 0; i < vals.length; i++) {
      var r = vals[i];
      var bsno = normalizeBsno(r[3]);
      if (!bsno) continue;
      out.push({
        ts: r[0] instanceof Date ? r[0] : (r[0] ? new Date(r[0]) : null),
        name: String(r[1] || '').trim(),
        rpnTin: String(r[2] || '').trim(),
        bsno: bsno,
        bsnoAlt: normalizeBsno(r[4]),
        success: String(r[5] || '').trim(),
        chatLink: String(r[6] || '').trim(),
        requestedAt: r[7] instanceof Date ? r[7] : (r[7] ? new Date(r[7]) : null),
        processedAt: r[8] instanceof Date ? r[8] : (r[8] ? new Date(r[8]) : null),
        firstPaidAt: r[9] instanceof Date ? r[9] : (r[9] ? new Date(r[9]) : null),
        reasonCat: String(r[10] || '').trim(),
        reasonDetail: String(r[11] || '').trim(),
        bkpMgr: String(r[12] || '').trim(),
        requester: normalizeRequester(r[13])
      });
    }
    return out;
  } catch (e) {
    Logger.log('loadCancelWorkflow error: ' + e.message);
    return [];
  }
}

function normalizeBsno(v) {
  if (v === null || v === undefined) return '';
  return String(v).replace(/[^0-9]/g, '');
}

// 요청자 이름 정규화: '@sally', '@Sally', 'sally', 'Sally' → 'sally'
function normalizeRequester(v) {
  if (!v) return '';
  var s = String(v).trim().replace(/^@/, '').toLowerCase();
  return s;
}

/* =========================
 * 매니저-팀 매핑 관리 API
 * ========================= */

// 매핑 시트 + 그룹 목록 반환
function api_getManagerMappings() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shMgrs = ss.getSheetByName('_매니저_팀매핑');
  var managers = [];
  if (shMgrs) {
    var last = shMgrs.getLastRow();
    if (last >= 2) {
      var vals = shMgrs.getRange(2, 1, last - 1, 4).getValues();
      for (var i = 0; i < vals.length; i++) {
        var name = String(vals[i][1] || '').trim();
        var email = String(vals[i][2] || '').trim();
        var teamsStr = String(vals[i][3] || '').trim();
        if (!name && !email && !teamsStr) continue;
        managers.push({
          rowIdx: i + 2, // 실제 시트 row 번호 (1-based)
          id: String(vals[i][0] || ''),
          name: name,
          email: email,
          teams: teamsStr ? teamsStr.split(',').map(function(t) { return t.trim(); }).filter(Boolean) : []
        });
      }
    }
  }

  // 그룹 리스트: _그룹_리스트 시트 (title 컬럼) + 매핑 시트에서 사용된 팀 합집합
  var shGroups = ss.getSheetByName('_그룹_리스트');
  var groupSet = {};
  if (shGroups) {
    var glast = shGroups.getLastRow();
    if (glast >= 2) {
      var gvals = shGroups.getRange(2, 2, glast - 1, 1).getValues(); // col 2 = title
      gvals.forEach(function(r) {
        var t = String(r[0] || '').trim();
        if (t) groupSet[t] = true;
      });
    }
  }
  managers.forEach(function(m) { m.teams.forEach(function(t) { groupSet[t] = true; }); });
  var groups = Object.keys(groupSet).sort();

  return { ok: true, managers: managers, groups: groups };
}

function api_saveManagerMapping(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shMgrs = ss.getSheetByName('_매니저_팀매핑');
  if (!shMgrs) {
    shMgrs = ss.insertSheet('_매니저_팀매핑');
    shMgrs.getRange(1, 1, 1, 4).setValues([['id', 'name', 'email', '소속 그룹들 (title, 콤마구분)']]).setFontWeight('bold');
  }
  var name = String((payload && payload.name) || '').trim();
  var email = String((payload && payload.email) || '').trim().toLowerCase();
  var teams = Array.isArray(payload && payload.teams) ? payload.teams.map(function(t) { return String(t).trim(); }).filter(Boolean) : [];
  if (!name) return { ok: false, error: '이름은 필수입니다' };
  var teamsStr = teams.join(', ');
  var rowIdx = Number(payload && payload.rowIdx) || 0;

  // 중복 검사 (기존 행 업데이트가 아니면)
  var last = shMgrs.getLastRow();
  if (rowIdx < 2 && last >= 2) {
    var existing = shMgrs.getRange(2, 1, last - 1, 4).getValues();
    for (var i = 0; i < existing.length; i++) {
      var xName = String(existing[i][1] || '').trim();
      var xEmail = String(existing[i][2] || '').trim().toLowerCase();
      if (name && xName === name) {
        return { ok: false, error: '이미 등록된 이름입니다: ' + name, existingRow: i + 2, existingName: xName, existingEmail: xEmail };
      }
      if (email && xEmail === email) {
        return { ok: false, error: '이미 등록된 이메일입니다: ' + email, existingRow: i + 2, existingName: xName, existingEmail: xEmail };
      }
    }
  }

  if (rowIdx > 1) {
    shMgrs.getRange(rowIdx, 2, 1, 3).setValues([[name, email, teamsStr]]);
  } else {
    shMgrs.appendRow(['', name, email, teamsStr]);
    rowIdx = shMgrs.getLastRow();
  }
  return { ok: true, rowIdx: rowIdx };
}

function api_deleteManagerMapping(payload) {
  var rowIdx = Number(payload && payload.rowIdx) || 0;
  if (rowIdx < 2) return { ok: false, error: '잘못된 row' };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shMgrs = ss.getSheetByName('_매니저_팀매핑');
  if (!shMgrs) return { ok: false, error: '매핑 시트 없음' };
  if (rowIdx > shMgrs.getLastRow()) return { ok: false, error: 'row 범위 밖' };
  shMgrs.deleteRow(rowIdx);
  return { ok: true };
}

function api_addGroup(payload) {
  var title = String((payload && payload.title) || '').trim();
  if (!title) return { ok: false, error: '그룹명 필수' };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shGroups = ss.getSheetByName('_그룹_리스트');
  if (!shGroups) {
    shGroups = ss.insertSheet('_그룹_리스트');
    shGroups.getRange(1, 1, 1, 4).setValues([['id', 'title', 'scope', 'memberCount']]).setFontWeight('bold');
  }
  // 중복 체크
  var last = shGroups.getLastRow();
  if (last >= 2) {
    var vals = shGroups.getRange(2, 2, last - 1, 1).getValues();
    for (var i = 0; i < vals.length; i++) {
      if (String(vals[i][0] || '').trim() === title) return { ok: false, error: '이미 존재하는 그룹' };
    }
  }
  shGroups.appendRow(['', title, 'manual', 0]);
  return { ok: true, title: title };
}

// _매니저_팀매핑 시트에서 매니저 name(정규화) → 소속 그룹(1번째) 맵 반환.
// _매니저_팀매핑_수동 시트가 있으면 그것도 병합 (수기 편집이 자동 갱신에 안 날아가게).
function loadManagerTeamMap(ss) {
  var map = {};

  function absorb(sh, priority) {
    if (!sh) return;
    var last = sh.getLastRow();
    if (last < 2) return;
    var vals = sh.getRange(2, 1, last - 1, 4).getValues();
    for (var i = 0; i < vals.length; i++) {
      var mName = String(vals[i][1] || '').trim();
      var teams = String(vals[i][3] || '').trim();
      if (!mName || !teams) continue;
      var primary = teams.split(',')[0].trim();
      // 팀명 통폐합: 케어운영팀·케어전환·신고 → 케어CX (실제 업무 담당 통합)
      if (primary === '케어운영팀' || primary === '케어전환' || primary === '신고') primary = '케어CX';
      var norm = normalizeMgrName(mName);
      if (!norm) continue;
      if (priority || !map[norm]) map[norm] = primary;
    }
  }

  absorb(ss.getSheetByName('_매니저_팀매핑'), false);
  absorb(ss.getSheetByName('_매니저_팀매핑_수동'), true); // 수기 시트가 있으면 덮어씀

  // 코드 레벨 오버라이드 (최우선)
  Object.keys(MGR_TEAM_OVERRIDES).forEach(function(k) {
    map[k] = MGR_TEAM_OVERRIDES[k];
  });

  return map;
}

/**
 * 만족도 응답자 조회 (이벤트 단위)
 * - 채널의 satisfactionGroupId 그룹 메시지를 스캔해 각 만족도 응답 이벤트 추출
 * - chatId를 중복 제거하지 않음 → 워크플로가 여러 번 발동되면 각 응답을 별도 이벤트로 취급
 * - sheetName이 있으면 기간(YYYYMMDD~YYYYMMDD)을 파싱해서 이벤트에 inPeriod 플래그 부여
 * - 대화 재문의로 응답 시점이 대화 시작 시점과 크게 어긋나는 케이스도 정확히 반영
 */
function api_getSatisfactionRespondents(channelKey, sheetName) {
  var ch = CHANNELS[channelKey];
  if (!ch) return { ok: false, error: '알 수 없는 채널' };
  if (!ch.satisfactionGroupId) return { ok: false, error: '이 채널은 만족도 조사가 없습니다.' };
  var keys = getApiKeys(channelKey);
  if (!keys) return { ok: false, error: 'API 키 없음' };

  var headers = buildHeaders(keys.key, keys.secret);
  var events = []; // { chatId, respondedAtMs }
  var since = '';
  var pagesScanned = 0;
  var earliestMs = null;
  var latestMs = null;
  var last7dCount = 0;
  var last30dCount = 0;
  var now = new Date().getTime();
  var sevenDaysAgo = now - 7 * 24 * 60 * 60 * 1000;
  var thirtyDaysAgo = now - 30 * 24 * 60 * 60 * 1000;
  var tz = Session.getScriptTimeZone();

  // 시트 이름에서 기간(YYYYMMDD~YYYYMMDD) 파싱
  var period = null;
  if (sheetName) {
    var pm = sheetName.match(/^(\d{4})(\d{2})(\d{2})~(\d{4})(\d{2})(\d{2})_/);
    if (pm) {
      var fromDate = new Date(Number(pm[1]), Number(pm[2]) - 1, Number(pm[3]), 0, 0, 0);
      var toDate = new Date(Number(pm[4]), Number(pm[5]) - 1, Number(pm[6]), 23, 59, 59);
      period = { fromMs: fromDate.getTime(), toMs: toDate.getTime() };
    }
  }

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
      var seenInMsg = {};
      var match;
      while ((match = re.exec(combined)) !== null) {
        var chatId = match[1];
        if (seenInMsg[chatId]) continue; // 같은 메시지 내 중복 링크는 스킵
        seenInMsg[chatId] = true;
        var ts = Number(m.createdAt || 0);
        events.push({ chatId: chatId, respondedAtMs: ts });
        if (ts) {
          if (earliestMs === null || ts < earliestMs) earliestMs = ts;
          if (latestMs === null || ts > latestMs) latestMs = ts;
          if (ts >= sevenDaysAgo) last7dCount++;
          if (ts >= thirtyDaysAgo) last30dCount++;
        }
      }
    });

    if (!data.next) break;
    since = data.next;
  }

  // 시트 매핑
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets;
  if (sheetName) {
    var s = ss.getSheetByName(sheetName);
    sheets = s ? [s] : [];
  } else {
    var suffix = '_' + ch.label;
    sheets = ss.getSheets().filter(function(sh) {
      var name = sh.getName();
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

  // 이벤트를 시간 오래된 순 정렬 후 chatId별 순번 계산 (재문의 이벤트 표시용)
  events.sort(function(a, b) { return (a.respondedAtMs || 0) - (b.respondedAtMs || 0); });
  var eventCountPerChat = {};
  events.forEach(function(e) {
    e.chatEventIndex = (eventCountPerChat[e.chatId] || 0) + 1;
    eventCountPerChat[e.chatId] = e.chatEventIndex;
  });
  events.forEach(function(e) { e.chatEventTotal = eventCountPerChat[e.chatId]; });

  var matched = [];
  var uniqueChatIds = {};
  var inPeriodCount = 0;
  var outOfPeriodCount = 0;
  var inPeriodUniqueChats = {};
  events.forEach(function(ev) {
    if (!chatDataMap[ev.chatId]) return;
    var row = chatDataMap[ev.chatId];
    uniqueChatIds[ev.chatId] = true;
    var inPeriod = true;
    if (period) inPeriod = ev.respondedAtMs >= period.fromMs && ev.respondedAtMs <= period.toMs;
    if (inPeriod) {
      inPeriodCount++;
      inPeriodUniqueChats[ev.chatId] = true;
    } else {
      outOfPeriodCount++;
    }
    matched.push({
      eventId: ev.chatId + '#' + ev.chatEventIndex,
      chatId: ev.chatId,
      chatEventIndex: ev.chatEventIndex,
      chatEventTotal: ev.chatEventTotal,
      name: String(row[1] || ''),
      bizName: String(row[3] || ''),
      manager: String(row[4] || ''),
      startTime: String(row[6] || ''),
      link: String(row[7] || ''),
      inflow: String(row[8] || ''),
      tags: String(row[9] || ''),
      respondedAt: ev.respondedAtMs ? Utilities.formatDate(new Date(ev.respondedAtMs), tz, 'yyyy-MM-dd HH:mm') : '',
      respondedAtMs: ev.respondedAtMs || 0,
      inPeriod: inPeriod
    });
  });

  // 응답시각 최신순 정렬
  matched.sort(function(a, b) { return (b.respondedAtMs || 0) - (a.respondedAtMs || 0); });

  var uniqueChatCount = Object.keys(uniqueChatIds).length;
  var inPeriodChatCount = Object.keys(inPeriodUniqueChats).length;

  return {
    ok: true,
    channel: channelKey,
    groupId: ch.satisfactionGroupId,
    pagesScanned: pagesScanned,
    totalRespondentsScanned: events.length,
    uniqueChatCount: uniqueChatCount,
    matchedInSheets: matched.length,
    inPeriodEvents: inPeriodCount,
    outOfPeriodEvents: outOfPeriodCount,
    inPeriodChatCount: inPeriodChatCount,
    period: period ? {
      fromYmd: Utilities.formatDate(new Date(period.fromMs), tz, 'yyyy-MM-dd'),
      toYmd: Utilities.formatDate(new Date(period.toMs), tz, 'yyyy-MM-dd')
    } : null,
    totalChatsInScope: totalChatsInScope,
    responseRate: totalChatsInScope > 0 ? Math.round(uniqueChatCount / totalChatsInScope * 1000) / 10 : 0,
    inPeriodResponseRate: totalChatsInScope > 0 ? Math.round(inPeriodChatCount / totalChatsInScope * 1000) / 10 : 0,
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
  var vals = sh.getRange(1, 1, last, 7).getValues();
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
    // 자동 분석 완료 마커는 이력 목록에서 제외 (파이프라인 상태로 별도 표시됨)
    var firstCol = String(row[0]);
    if (firstCol.indexOf('[자동') === 0 || firstCol.indexOf('[LLM') === 0 || firstCol.indexOf('[리포트') === 0) continue;
    var period = String(row[1] || '');
    var sheetName = '';
    var m = period.match(/(\d{8})\s*~\s*(\d{8})/);
    if (m) sheetName = m[1] + '~' + m[2] + '_' + chLabel;
    var candidates = Number(row[2] || 0);
    var collected = Number(row[3] || 0);
    var skips = null;
    try { if (row[6]) skips = JSON.parse(row[6]); } catch (e) {}
    out.push({
      completedAt: formatMaybeDate(row[0]),
      period: period,
      sheetName: sheetName,
      candidates: candidates,
      collected: collected,
      failed: Number(row[4] || 0),
      uncollected: Math.max(0, candidates - collected),
      status: String(row[5] || ''),
      skips: skips,
      successRate: (candidates > 0) ? Math.round(collected / candidates * 1000) / 10 : 0
    });
  }
  out.reverse();
  return out;
}

function api_getChannelHistory(channelKey) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var history = getExecutionHistory(ss, channelKey);
  // 레거시 이력(skips 미저장) fallback: 진행상황 시트 top의 최근 실행 skip 값을 매칭
  try {
    var last = api_getLastProgress(channelKey);
    if (last && last.skips && last.period) {
      history.forEach(function(h) {
        if (!h.skips && h.period === last.period) {
          h.skips = last.skips;
        }
      });
    }
  } catch (e) {}
  return history;
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

  var vals = sh.getRange(1, 1, last, 7).getValues();
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
  var targetSheetName = '';
  var relatedDeleted = [];
  if (deleteSheet) {
    var chLabel = CHANNELS[channelKey].label;
    var m = period.match(/(\d{8})\s*~\s*(\d{8})/);
    if (m) {
      targetSheetName = m[1] + '~' + m[2] + '_' + chLabel;
      var outputSh = ss.getSheetByName(targetSheetName);
      if (outputSh) {
        ss.deleteSheet(outputSh);
        sheetDeleted = true;
      }
      // 부수 시트도 함께 삭제 (_리포트_, _VOC_분류_)
      ['_리포트_' + targetSheetName, '_VOC_분류_' + targetSheetName].forEach(function(sub) {
        var subSh = ss.getSheetByName(sub);
        if (subSh) {
          try { ss.deleteSheet(subSh); relatedDeleted.push(sub); } catch (e) {}
        }
      });
    }
  }

  return {
    ok: true,
    sheetDeleted: sheetDeleted,
    sheetName: targetSheetName,
    relatedDeleted: relatedDeleted
  };
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
  // 메타베이스는 케어에만 필수. 환급은 스킵 (매칭 정보 없이 채널톡 원본만 사용)
  if (channelKey === 'care') {
    try {
      mustSheet(ss, CFG.SHEET_META);
    } catch (e) {
      return { ok: false, error: '"메타베이스" 시트가 없습니다. 케어 채널 수집 전에 시트를 만들어 주세요.' };
    }
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
  var outputSheetName = job.outputSheet;
  var shOut = mustOutputSheet(ss, outputSheetName);
  // 메타베이스는 케어 채널에서만 사용. 환급은 스킵.
  var meta = { byPhone: {} };
  if (channelKey === 'care') {
    var shMeta = ss.getSheetByName(CFG.SHEET_META);
    if (shMeta) {
      try {
        meta = buildMetaIndex(shMeta);
      } catch (e) {
        Logger.log('[META] care metabase index build failed, using empty: ' + e.message);
      }
    } else {
      Logger.log('[META] care: 메타베이스 시트 없음, 빈 인덱스로 진행');
    }
  }
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

  // 매니저 id → name 캐시 (배치당 1회 조회)
  var managersById = fetchManagersMap(headers);

  // 만족도 폼 응답 통합 수집 준비 (채널에 만족도 그룹 있을 때만)
  var satisfactionBatchRows = [];
  var existingSatKeys = {};
  var satSh = null;
  if (CHANNELS[channelKey] && CHANNELS[channelKey].satisfactionGroupId) {
    satSh = ss.getSheetByName(WEBHOOK_SHEET) || ss.insertSheet(WEBHOOK_SHEET);
    if (satSh.getLastRow() === 0) writeSheetValues(satSh, 1, 1, [WEBHOOK_HEADERS]);
    // 중복 방지 인덱스: chatId + '|' + submittedAt
    if (satSh.getLastRow() > 1) {
      var satVals = satSh.getRange(2, 1, satSh.getLastRow() - 1, WEBHOOK_HEADERS.length).getValues();
      for (var sk = 0; sk < satVals.length; sk++) {
        var scid = String(satVals[sk][1] || '').trim();
        if (!scid) continue;
        var pRaw = satVals[sk][6];
        var subA = '';
        try {
          var pObj = typeof pRaw === 'string' ? JSON.parse(pRaw) : (pRaw || {});
          subA = pObj.submittedAt ? String(pObj.submittedAt) : '';
        } catch (e) {}
        if (subA) existingSatKeys[scid + '|' + subA] = true;
      }
    }
  }

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
    // 폰 매칭 실패 시 대표자 이름으로 fallback (단일 매칭만 신뢰)
    if (!metaHit && chat.userName && meta.byName) {
      var nameHits = meta.byName[chat.userName];
      if (nameHits && nameHits.length === 1) metaHit = nameHits[0];
      // 다수 매칭이면 모두 동일 기장담당자인지 확인해 신뢰 판단
      else if (nameHits && nameHits.length > 1) {
        var allSameMgr = nameHits.every(function(h) { return h.careproMgr === nameHits[0].careproMgr; });
        if (allSameMgr) metaHit = nameHits[0];
      }
    }

    // 첫응대자: 시간상 첫 외부 응답 매니저 (진입)
    // 최종응대자: 시간상 마지막 외부 응답 매니저 (종결). 없으면 assignee 사용.
    var firstId = String(result.firstResponderId || '');
    var lastId = String(result.lastResponderId || result.assigneeId || result.primaryResponderId || '');
    var firstName = firstId ? (managersById[firstId] || '') : '';
    var lastName = lastId ? (managersById[lastId] || '') : '';

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
      result.reviewedAt ? formatTs(result.reviewedAt) : '',
      firstName,
      lastName
    ]);

    // 만족도 폼 응답 즉시 만족도_웹훅에 큐잉 (중복 스킵)
    if (satSh && result.satisfactionResponses && result.satisfactionResponses.length) {
      var chatName = metaHit ? (metaHit.userName || chat.userName || '') : (chat.userName || '');
      for (var sr = 0; sr < result.satisfactionResponses.length; sr++) {
        var rp = result.satisfactionResponses[sr];
        var key = chat.id + '|' + rp.submittedAt;
        if (existingSatKeys[key]) continue;
        existingSatKeys[key] = true;
        satisfactionBatchRows.push([
          new Date(rp.submittedAt || rp.createdAt),
          chat.id,
          rp.score !== null && rp.score !== undefined ? String(rp.score) : '',
          rp.factor || '',
          rp.etc || '',
          chatName,
          JSON.stringify({
            source: 'collection',
            submittedAt: rp.submittedAt,
            messageId: rp.messageId,
            factorLabel: rp.factorLabel || '',
            factorOptions: rp.factorOptions || []
          })
        ]);
      }
    }
  }

  appendOutputRows(shOut, batchRows);
  if (satSh && satisfactionBatchRows.length) {
    writeSheetValues(satSh, satSh.getLastRow() + 1, 1, satisfactionBatchRows);
  }
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
    // ==== 자동 분석 파이프라인 hook (케어 채널만) ====
    try {
      startAutoAnalysisPipeline(channelKey, outputSheetName);
    } catch (e) {
      Logger.log('[AUTO] hook failed: ' + e.message);
    }
    // ==== 스냅샷 자동 저장: 다음 수집 시 이전 시트 압축·삭제해도 대시보드 유지 ====
    try {
      var snapRes = api_saveDashboardSnapshot(channelKey, outputSheetName);
      Logger.log('[SNAPSHOT] ' + outputSheetName + ' → ' + (snapRes.ok ? snapRes.jsonBytes + ' bytes, chunks=' + snapRes.chunks : 'FAILED: ' + snapRes.error));
    } catch (e) {
      Logger.log('[SNAPSHOT] hook failed: ' + e.message);
    }
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

// 이상값(timestamp 오염 등) 방어: 0 이상 정수만 허용
function sanitizeSkipCount(v) {
  var n = Number(v);
  if (!isFinite(n) || isNaN(n) || n < 0 || n > 1e7) return 0;
  return Math.floor(n);
}
function sanitizeSkips(s) {
  return {
    autoClosed: sanitizeSkipCount(s && s.autoClosed),
    outOfRange: sanitizeSkipCount(s && s.outOfRange),
    emptyText: sanitizeSkipCount(s && s.emptyText),
    fetchError: sanitizeSkipCount(s && s.fetchError)
  };
}

function loadSkipCounts(channelKey) {
  var raw = PropertiesService.getScriptProperties().getProperty(skipProp(channelKey));
  var out = { autoClosed: 0, outOfRange: 0, emptyText: 0, fetchError: 0 };
  if (!raw) return out;
  try {
    var parsed = JSON.parse(raw);
    Object.keys(parsed).forEach(function(k) { out[k] = sanitizeSkipCount(parsed[k]); });
  } catch (e) {}
  return out;
}

function saveSkipCounts(channelKey, skips) {
  PropertiesService.getScriptProperties().setProperty(skipProp(channelKey), JSON.stringify(sanitizeSkips(skips)));
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

// 채널톡 desk 링크에서 chatId 추출
function extractChatIdFromLink(link) {
  if (!link) return '';
  var s = String(link).split('?')[0].split('#')[0];
  var parts = s.split('/').filter(function(x) { return x; });
  return parts.length ? parts[parts.length - 1] : '';
}

// 채널톡 매니저 id → name 맵 (전체 스캔)
function fetchManagersMap(headers) {
  var map = {};
  // state 파라미터를 여러 값으로 시도해 active/dormant/inactive 매니저 모두 수집
  var states = ['', 'active', 'dormant', 'inactive'];
  states.forEach(function(state) {
    var since = '';
    var guard = 0;
    while (guard++ < 30) {
      var params = { limit: 100 };
      if (state) params.state = state;
      if (since) params.since = since;
      try {
        var data = ctGet('/managers', params, headers);
        var mgrs = data.managers || [];
        mgrs.forEach(function(m) {
          if (m.id && m.name && !map[String(m.id)]) map[String(m.id)] = String(m.name);
        });
        if (!data.next) break;
        since = data.next;
      } catch (e) {
        Logger.log('fetchManagersMap[' + state + '] error: ' + e.message);
        break;
      }
    }
  });
  return map;
}

// 특정 매니저 id 한 명만 조회 (map에 없을 때 fallback)
function fetchManagerNameById(headers, mgrId) {
  if (!mgrId) return '';
  try {
    var data = ctGet('/managers/' + mgrId, {}, headers);
    var m = data.manager || data;
    return m && m.name ? String(m.name) : '';
  } catch (e) {
    return '';
  }
}

function fetchTranscriptAndInflow(chatId, headers, startMs, endMs) {
  var slice = fetchTranscriptSliceInRange(chatId, headers, startMs, endMs);
  if (!slice) return null;
  if (!slice.transcript) return slice; // skipReason 포함
  var detail = fetchUserChatDetail(chatId, headers);
  slice.inflow = detail.inflow;
  slice.rating = detail.rating;
  slice.reviewedAt = detail.reviewedAt;
  slice.assigneeId = detail.assigneeId || '';
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
      reviewedAt: review.reviewedAt,
      assigneeId: extractAssigneeId(chat)
    };
  } catch (e) {
    return { inflow: '채널톡', rating: null, reviewedAt: null, assigneeId: '' };
  }
}

// 채널톡 chat 객체에서 담당자 id 추출 (필드 위치가 API 버전마다 다름)
function extractAssigneeId(chat) {
  if (!chat) return '';
  var candidates = [
    chat.assigneeId,
    chat.managerId,
    chat.assignedManagerId,
    chat.hostId,
    chat.owner && chat.owner.id,
    chat.assignee && chat.assignee.id
  ];
  for (var i = 0; i < candidates.length; i++) {
    if (candidates[i]) return String(candidates[i]);
  }
  var lists = [chat.managerIds, chat.assignedManagerIds, chat.hostIds, chat.assigneeIds];
  for (var j = 0; j < lists.length; j++) {
    if (Array.isArray(lists[j]) && lists[j].length) return String(lists[j][0]);
  }
  return '';
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
 * 기존 수집 시트의 기장담당자(col 4) 배치 채우기 - 이름 기반 CarePro 역조회
 * 폰번호 매칭 실패로 기장담당자 셀이 비어있는 행에 대해 대표자 이름으로 CarePro 시트 검색
 */
function api_enrichBkpManagers(channelKey, sheetName) {
  if (!CHANNELS[channelKey]) return { ok: false, error: '알 수 없는 채널' };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(sheetName);
  if (!sh) return { ok: false, error: '시트 없음: ' + sheetName };
  var shMeta = ss.getSheetByName('메타베이스');
  if (!shMeta) return { ok: false, error: '메타베이스 시트 없음' };
  var meta;
  try { meta = buildMetaIndex(shMeta); } catch (e) { return { ok: false, error: 'meta 인덱스 실패: ' + e.message }; }
  if (!meta.byName) return { ok: false, error: 'meta.byName 없음 (구버전 인덱스)' };

  var lastRow = sh.getLastRow();
  if (lastRow < 2) return { ok: false, error: '데이터 없음' };
  var numCols = Math.max(sh.getLastColumn(), CFG.OUT_HEADERS.length);
  var vals = sh.getRange(2, 1, lastRow - 1, numCols).getValues();

  var updated = 0, scanned = 0, skipped = 0;
  for (var i = 0; i < vals.length; i++) {
    scanned++;
    var row = vals[i];
    var existingMgr = String(row[4] || '').trim();
    if (existingMgr) { skipped++; continue; }
    var repName = String(row[1] || '').trim();
    if (!repName) continue;
    var hits = meta.byName[repName];
    if (!hits || !hits.length) continue;
    var pick = null;
    if (hits.length === 1) pick = hits[0];
    else {
      var same = hits.every(function(h) { return h.careproMgr === hits[0].careproMgr; });
      if (same) pick = hits[0];
    }
    if (!pick || !pick.careproMgr) continue;
    // col 순서: 1=Rpntin, 2=대표자, 3=사업자번호, 4=사업장명, 5=기장담당자, 6=현재상태
    sh.getRange(i + 2, 1).setValue(pick.rpnTin || row[0] || '');
    sh.getRange(i + 2, 3).setValue(pick.bsno || row[2] || '');
    sh.getRange(i + 2, 4).setValue(pick.tnmNm || row[3] || '');
    sh.getRange(i + 2, 5).setValue(pick.careproMgr);
    sh.getRange(i + 2, 6).setValue(pick.bkpStatus || row[5] || '');
    updated++;
  }

  return { ok: true, scanned: scanned, updated: updated, skipped: skipped, total: vals.length };
}

/**
 * 기존 수집 시트의 상담담당자(col N, index 13) 배치 채우기
 * 빈 셀에 대해 채널톡 assignee 조회 → 매니저 이름 저장
 * 시간 초과 방지: 배치당 최대 5분 실행, 처리한 만큼 진행상황 저장 (재실행하면 이어감)
 */
function api_enrichResponders(channelKey, sheetName) {
  if (!CHANNELS[channelKey]) return { ok: false, error: '알 수 없는 채널' };
  var keys = getApiKeys(channelKey);
  if (!keys) return { ok: false, error: 'API 키 없음' };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(sheetName);
  if (!sh) return { ok: false, error: '시트 없음: ' + sheetName };

  var headers = buildHeaders(keys.key, keys.secret);
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return { ok: false, error: '데이터 없음' };
  var numCols = Math.max(sh.getLastColumn(), CFG.OUT_HEADERS.length);

  // col 8 (index 7) = 채널톡링크, col 14 (index 13) = 상담담당자
  var range = sh.getRange(2, 1, lastRow - 1, numCols);
  var vals = range.getValues();

  // 상담담당자 컬럼 없으면 헤더에 추가
  if (numCols < CFG.OUT_HEADERS.length) {
    sh.getRange(1, 1, 1, CFG.OUT_HEADERS.length).setValues([CFG.OUT_HEADERS]);
  }

  var managersById = fetchManagersMap(headers);

  var started = new Date().getTime();
  var MAX_MS = 5 * 60 * 1000;
  var updated = 0;
  var scanned = 0;
  var skipped = 0;
  var errors = 0;

  var propKey = 'ENRICH_RESUME_' + sheetName;
  var startFrom = Number(PropertiesService.getScriptProperties().getProperty(propKey) || 0);

  for (var i = startFrom; i < vals.length; i++) {
    if (new Date().getTime() - started > MAX_MS) break;
    scanned++;
    var row = vals[i];
    var existingFirst = String(row[13] || '').trim();
    var existingLast = String(row[14] || '').trim();
    // 둘 다 이미 있으면 skip (신 포맷). 하나만 있으면 레거시 데이터일 수 있으니 재처리.
    if (existingFirst && existingLast) { skipped++; continue; }

    var link = String(row[7] || '').trim();
    var chatId = extractChatIdFromLink(link);
    if (!chatId) { errors++; continue; }

    try {
      var firstName = '';
      var lastName = '';
      // 메시지 스캔으로 첫/마지막 외부 매니저 파악
      var msgData = ctGet('/user-chats/' + chatId + '/messages', { limit: 500, sortOrder: 'asc' }, headers);
      var msgs = msgData.messages || [];
      var firstPid = '', lastPid = '';
      msgs.forEach(function(m) {
        var pt = String(m.personType || '').toLowerCase();
        if (pt !== 'manager') return;
        var pid = String(m.personId || '');
        if (!pid) return;
        var vis = getVisibilityLabel(m);
        if (vis !== '외부') return;
        if (!firstPid) firstPid = pid;
        lastPid = pid;
      });
      firstName = firstPid ? (managersById[firstPid] || fetchManagerNameById(headers, firstPid)) : '';
      lastName = lastPid ? (managersById[lastPid] || fetchManagerNameById(headers, lastPid)) : '';
      // 캐시에도 저장 (같은 배치에서 재조회 방지)
      if (firstPid && firstName && !managersById[firstPid]) managersById[firstPid] = firstName;
      if (lastPid && lastName && !managersById[lastPid]) managersById[lastPid] = lastName;

      // 메시지 기반 실패 시 chat.assigneeId로 최종응대자 fallback
      if (!lastName) {
        try {
          var chatData = ctGet('/user-chats/' + chatId, {}, headers);
          var chat = chatData.userChat || chatData.chat || {};
          var assigneeId = extractAssigneeId(chat);
          if (assigneeId) lastName = managersById[assigneeId] || '';
          if (!firstName && assigneeId) firstName = lastName;
        } catch (e3) {}
      }

      // 레거시 데이터 감지: col 14가 채워져 있는데 col 15는 비어 있음 → 옛 "상담담당자"(primary)라 신뢰 못함.
      // 이 경우 col 14도 실제 "첫응대자"로 덮어씀. (existingFirst && !existingLast) 조건
      var isLegacy = existingFirst && !existingLast;
      var didUpdate = false;
      if ((!existingFirst || isLegacy) && firstName) { sh.getRange(i + 2, 14).setValue(firstName); didUpdate = true; }
      if (!existingLast && lastName) { sh.getRange(i + 2, 15).setValue(lastName); didUpdate = true; }
      if (didUpdate) updated++;
    } catch (e) {
      errors++;
    }

    PropertiesService.getScriptProperties().setProperty(propKey, String(i + 1));
    Utilities.sleep(150); // rate limit 방어
  }

  var doneAll = (startFrom + scanned) >= vals.length;
  if (doneAll) PropertiesService.getScriptProperties().deleteProperty(propKey);

  return {
    ok: true,
    scanned: scanned,
    updated: updated,
    skipped: skipped,
    errors: errors,
    total: vals.length,
    processedTo: startFrom + scanned,
    done: doneAll,
    resumeAt: doneAll ? null : (startFrom + scanned)
  };
}

/**
 * 담당자 진단: 특정 chatId 의 assignee/manager 정보 + 메시지 기반 응대자 감지 결과 반환
 * 프론트에서 이 결과를 보고 매칭 실패 원인 파악
 */
function api_diagnoseChatResponder(channelKey, chatId) {
  if (!CHANNELS[channelKey]) return { ok: false, error: '알 수 없는 채널' };
  var keys = getApiKeys(channelKey);
  if (!keys) return { ok: false, error: 'API 키 없음' };
  var headers = buildHeaders(keys.key, keys.secret);
  try {
    var data = ctGet('/user-chats/' + chatId, {}, headers);
    var chat = data.userChat || data.chat || {};
    var mgrs = fetchManagersMap(headers);
    var assigneeId = extractAssigneeId(chat);
    var mgrsById = mgrs;

    // 최근 100 메시지 스캔해 응대자 감지 (isPrivate/personType별)
    var msgData = ctGet('/user-chats/' + chatId + '/messages', { limit: 100, sortOrder: 'desc' }, headers);
    var msgs = (msgData.messages || []);
    var msgManagers = {}; // personId -> { total, external, internal, sampleTexts }
    msgs.forEach(function(m) {
      var pt = String(m.personType || '').toLowerCase();
      if (pt !== 'manager') return;
      var pid = String(m.personId || '');
      if (!pid) return;
      if (!msgManagers[pid]) msgManagers[pid] = { total: 0, external: 0, internal: 0, sampleTexts: [] };
      msgManagers[pid].total++;
      var vis = getVisibilityLabel(m);
      if (vis === '외부') msgManagers[pid].external++;
      else msgManagers[pid].internal++;
      if (msgManagers[pid].sampleTexts.length < 2) {
        msgManagers[pid].sampleTexts.push({ vis: vis, isPrivate: m.isPrivate, text: (m.plainText || '').slice(0, 60) });
      }
    });

    var msgManagersOut = Object.keys(msgManagers).map(function(pid) {
      return {
        personId: pid,
        name: mgrsById[pid] || '(managers map에 없음)',
        total: msgManagers[pid].total,
        external: msgManagers[pid].external,
        internal: msgManagers[pid].internal,
        sampleTexts: msgManagers[pid].sampleTexts
      };
    }).sort(function(a, b) { return b.external - a.external; });

    return {
      ok: true,
      chatId: chatId,
      rawChatKeys: Object.keys(chat),
      assigneeExtraction: {
        assigneeId: chat.assigneeId || null,
        managerId: chat.managerId || null,
        assignedManagerId: chat.assignedManagerId || null,
        hostId: chat.hostId || null,
        managerIds: chat.managerIds || null,
        assignedManagerIds: chat.assignedManagerIds || null,
        followerIds: chat.followerIds || null,
        extractedId: assigneeId,
        extractedName: assigneeId ? (mgrsById[assigneeId] || '(managers map에 없음)') : '(없음)'
      },
      messageBasedResponders: msgManagersOut,
      managersMapSize: Object.keys(mgrsById).length,
      recommendation: assigneeId
        ? '✓ assignee 추출 성공: ' + (mgrsById[assigneeId] || assigneeId)
        : (msgManagersOut.length ? '⚠ assignee 필드 없음, 메시지 기반 응대자 사용 권장: ' + msgManagersOut[0].name : '✗ 응대자 정보 전혀 없음')
    };
  } catch (e) {
    return { ok: false, error: String(e.message || e) };
  }
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
  if (!gid) return { ok: false, error: 'group ID 를 입력하세요 (예: <GROUP_ID> 또는 일반상담_만족도-<GROUP_ID>)' };

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

  // 응대자 추적: 외부(고객 대상) 메시지를 보낸 매니저 personId를 시간순으로 수집
  var responderCounts = {}; // personId -> count of external messages
  var firstResponderId = '';
  var firstResponderMs = Infinity;
  var lastResponderId = '';
  var lastResponderMs = -Infinity;

  // 만족도 폼 응답 수집 (수집 시점에 통합) — 기간 필터 안 함 (폼 응답은 상담 종료 후에도 도착 가능)
  var satisfactionFormMsgs = [];

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

      // 폼 응답 감지 (기간 무관하게 전량 수집 → 상담 종료 후 응답도 캡처)
      if (m.form && m.form.type === 'custom' && m.form.submittedAt) {
        satisfactionFormMsgs.push(m);
      }

      var createdMs = Number(m.createdAt || 0);
      if (!createdMs) continue;
      if (createdMs < oldestInPage) oldestInPage = createdMs;
      if (createdMs > endMs) continue;
      if (createdMs < startMs) continue;

      hadInRangeMsg = true;
      var text = extractMessageText(m);
      if (!text) continue;

      var vis = getVisibilityLabel(m);
      var pType = String(m.personType || '').toLowerCase();
      var personId = String(m.personId || '');
      if (pType === 'manager' && vis === '외부' && personId) {
        responderCounts[personId] = (responderCounts[personId] || 0) + 1;
        if (createdMs < firstResponderMs) {
          firstResponderMs = createdMs;
          firstResponderId = personId;
        }
        if (createdMs > lastResponderMs) {
          lastResponderMs = createdMs;
          lastResponderId = personId;
        }
      }

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

  // 폼 응답 파싱 (만족도 폼만 필터)
  var satisfactionResponses = extractSatisfactionFromChatMessages(satisfactionFormMsgs);

  if (!inPeriod.length) {
    var reason = hadInRangeMsg ? 'emptyText' : (totalFetched > 0 ? 'outOfRange' : 'autoClosed');
    return { transcript: null, skipReason: reason, satisfactionResponses: satisfactionResponses };
  }

  var firstMsgMs = inPeriod[0].ms;
  for (var a = 0; a < inPeriod.length; a++) {
    if (inPeriod[a].ms < firstMsgMs) firstMsgMs = inPeriod[a].ms;
  }

  // 주요 응대자: 외부 메시지를 가장 많이 보낸 매니저 (동률이면 첫 응대자)
  var primaryResponderId = firstResponderId;
  var maxCount = 0;
  Object.keys(responderCounts).forEach(function(pid) {
    if (responderCounts[pid] > maxCount) {
      maxCount = responderCounts[pid];
      primaryResponderId = pid;
    }
  });

  inPeriod.sort(function(x, y) { return x.ms - y.ms; });
  var out = [];
  for (var c = 0; c < inPeriod.length; c++) out.push(inPeriod[c].line);
  return {
    transcript: out.join('\n'),
    firstMsgMs: firstMsgMs,
    firstResponderId: firstResponderId,
    lastResponderId: lastResponderId,
    primaryResponderId: primaryResponderId,
    satisfactionResponses: satisfactionResponses
  };
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
/**
 * 메타베이스 파일 업로드 (csv / json): 시트 전체 교체
 * - 헤더 검증 (필수 컬럼 존재 확인)
 * - Script Properties에 최근 갱신일 기록
 * payload: { format: 'csv'|'json', content: string }
 */
function api_uploadMetabaseFile(payload) {
  if (!payload || !payload.content) return { ok: false, error: '파일 데이터 없음' };
  var format = String(payload.format || 'csv').toLowerCase();
  var content = String(payload.content);
  if (content.length > 50 * 1024 * 1024) return { ok: false, error: '파일 크기가 너무 큽니다 (50MB 초과)' };

  var rows;
  try {
    if (format === 'json') {
      var arr = JSON.parse(content);
      if (!Array.isArray(arr) || !arr.length) return { ok: false, error: 'JSON은 객체 배열이어야 함' };
      var jsonHeaders = Object.keys(arr[0]);
      rows = [jsonHeaders];
      arr.forEach(function(obj) {
        rows.push(jsonHeaders.map(function(h) {
          var v = obj[h];
          return v === null || v === undefined ? '' : v;
        }));
      });
    } else {
      // csv
      rows = Utilities.parseCsv(content);
    }
  } catch (e) {
    return { ok: false, error: format + ' 파싱 실패: ' + e.message };
  }
  if (!rows || rows.length < 2) return { ok: false, error: '빈 파일 또는 헤더만 존재' };

  // 헤더 검증
  var headers = rows[0].map(function(v) { return String(v || '').trim(); });
  var required = [
    CFG.META_HEADERS.RPN_TIN,
    CFG.META_HEADERS.CONTACT_PHONE,
    CFG.META_HEADERS.BSNO,
    CFG.META_HEADERS.BKP_STATUS,
    CFG.META_HEADERS.TNM_NM,
    CFG.META_HEADERS.CAREPRO_MGR
  ];
  var missing = required.filter(function(k) { return headers.indexOf(k) < 0; });
  if (missing.length) return { ok: false, error: '필수 헤더 누락: ' + missing.join(', '), headersFound: headers };

  // 시트 교체
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(CFG.SHEET_META);
  if (!sh) sh = ss.insertSheet(CFG.SHEET_META);
  sh.clearContents();
  var numCols = headers.length;
  // 각 행 컬럼 수를 맞춤 (부족하면 빈 값 padding)
  var normalized = rows.map(function(r) {
    var arr = r.slice(0, numCols);
    while (arr.length < numCols) arr.push('');
    return arr;
  });
  sh.getRange(1, 1, normalized.length, numCols).setValues(normalized);

  // 갱신 메타 저장
  var props = PropertiesService.getScriptProperties();
  var updatedAt = new Date();
  props.setProperty('META_UPLOADED_AT', updatedAt.toISOString());
  props.setProperty('META_UPLOADED_ROWS', String(normalized.length - 1));

  // 업로드 직후 모든 케어 시트에 대해 기장담당자 이름 매칭 자동 실행 (백그라운드성 - 5분 실행 제한 내)
  var enrichSummary = null;
  try {
    enrichSummary = enrichAllCareBkpManagers();
  } catch (e) {
    Logger.log('post-upload enrich failed: ' + e.message);
  }

  return {
    ok: true,
    rows: normalized.length - 1,
    headers: headers.length,
    sizeKB: Math.round(content.length / 1024),
    updatedAt: Utilities.formatDate(updatedAt, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'),
    enrichSummary: enrichSummary
  };
}

// 모든 케어 시트에 대해 이름 기반 기장담당자 매칭 실행 (업로드 시 자동 호출)
function enrichAllCareBkpManagers() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var suffix = '_' + CHANNELS['care'].label;
  var sheets = ss.getSheets().filter(function(s) {
    var name = s.getName();
    return /^\d{8}~\d{8}_/.test(name) && name.substr(name.length - suffix.length) === suffix;
  });
  var summary = { sheetsProcessed: 0, totalUpdated: 0, sheetResults: [] };
  var started = new Date().getTime();
  var MAX_MS = 4.5 * 60 * 1000;
  for (var i = 0; i < sheets.length; i++) {
    if (new Date().getTime() - started > MAX_MS) break;
    var name = sheets[i].getName();
    try {
      var res = api_enrichBkpManagers('care', name);
      if (res && res.ok) {
        summary.sheetsProcessed++;
        summary.totalUpdated += res.updated;
        summary.sheetResults.push({ sheet: name, updated: res.updated, scanned: res.scanned });
      }
    } catch (e) {
      Logger.log('enrichAllCareBkpManagers[' + name + '] error: ' + e.message);
    }
  }
  return summary;
}

// CSAT 통계 업로드: `Workflow Daily Stats` 행을 _CSAT_통계 시트에 저장
function api_uploadCsatData(payload) {
  if (!payload) return { ok: false, error: 'payload 없음' };
  var rows;
  if (payload.format === 'rows' && Array.isArray(payload.rows)) {
    rows = payload.rows;
  } else if (payload.format === 'csv' && payload.content) {
    try { rows = Utilities.parseCsv(payload.content); } catch (e) { return { ok: false, error: 'CSV 파싱 실패: ' + e.message }; }
  } else {
    return { ok: false, error: '지원하지 않는 포맷' };
  }
  if (!rows || rows.length < 2) return { ok: false, error: '빈 데이터' };

  var headers = rows[0].map(function(v) { return String(v || '').trim(); });
  var required = ['date', 'view', 'end'];
  var missing = required.filter(function(k) { return headers.indexOf(k) < 0; });
  if (missing.length) return { ok: false, error: '필수 컬럼 누락: ' + missing.join(', ') + ' (workbook의 "Workflow Daily Stats" 시트 확인)' };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('_CSAT_통계') || ss.insertSheet('_CSAT_통계');
  sh.clearContents();
  var numCols = headers.length;
  var normalized = rows.map(function(r) {
    var arr = (r || []).slice(0, numCols);
    while (arr.length < numCols) arr.push('');
    return arr;
  });
  sh.getRange(1, 1, normalized.length, numCols).setValues(normalized);

  PropertiesService.getScriptProperties().setProperty('CSAT_UPLOADED_AT', new Date().toISOString());
  PropertiesService.getScriptProperties().setProperty('CSAT_UPLOADED_ROWS', String(normalized.length - 1));

  return { ok: true, rows: normalized.length - 1 };
}

function api_getCsatInfo() {
  var props = PropertiesService.getScriptProperties();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('_CSAT_통계');
  var uploadedAt = props.getProperty('CSAT_UPLOADED_AT');
  return {
    hasSheet: !!sh,
    rowCount: sh ? Math.max(0, sh.getLastRow() - 1) : 0,
    uploadedAt: uploadedAt ? Utilities.formatDate(new Date(uploadedAt), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') : null,
    ageDays: uploadedAt ? Math.floor((new Date().getTime() - new Date(uploadedAt).getTime()) / 86400000) : null
  };
}

// 메타베이스 갱신 정보 조회
function api_getMetaInfo() {
  var props = PropertiesService.getScriptProperties();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(CFG.SHEET_META);
  var rowCount = sh ? Math.max(0, sh.getLastRow() - 1) : 0;
  var uploadedAt = props.getProperty('META_UPLOADED_AT');
  var uploadedRows = props.getProperty('META_UPLOADED_ROWS');
  return {
    hasSheet: !!sh,
    rowCount: rowCount,
    uploadedAt: uploadedAt ? Utilities.formatDate(new Date(uploadedAt), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') : null,
    uploadedRows: uploadedRows ? Number(uploadedRows) : null,
    ageDays: uploadedAt ? Math.floor((new Date().getTime() - new Date(uploadedAt).getTime()) / 86400000) : null
  };
}

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
    careproMgr: h.indexOf(CFG.META_HEADERS.CAREPRO_MGR),
    careproMgrEmail: h.indexOf(CFG.META_HEADERS.CAREPRO_MGR_EMAIL)
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
  var byName = {};   // 대표자 이름 → 항목 (동명이인 다수면 배열)
  for (var i = 1; i < vals.length; i++) {
    var row = vals[i];
    var phone = normalizePhone(row[idx.phone]);
    var record = {
      rpnTin: String(row[idx.rpnTin] || '').trim(),
      userName: idx.userName >= 0 ? String(row[idx.userName] || '').trim() : '',
      bsno: String(row[idx.bsno] || '').trim(),
      bkpStatus: String(row[idx.bkpStatus] || '').trim(),
      tnmNm: String(row[idx.tnmNm] || '').trim(),
      careproMgr: String(row[idx.careproMgr] || '').trim(),
      careproMgrEmail: idx.careproMgrEmail >= 0 ? String(row[idx.careproMgrEmail] || '').trim().toLowerCase() : ''
    };
    if (phone && !byPhone[phone]) byPhone[phone] = record;
    var nm = record.userName;
    if (nm) {
      if (!byName[nm]) byName[nm] = [];
      byName[nm].push(record);
    }
  }
  return { byPhone: byPhone, byName: byName };
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

  var skips = info.skips ? sanitizeSkips(info.skips) : null;
  if (skips) {
    rows.push(['스킵 - ' + SKIP_LABELS.autoClosed, skips.autoClosed]);
    rows.push(['스킵 - ' + SKIP_LABELS.outOfRange, skips.outOfRange]);
    rows.push(['스킵 - ' + SKIP_LABELS.emptyText, skips.emptyText]);
    rows.push(['스킵 - ' + SKIP_LABELS.fetchError, skips.fetchError]);
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
  var skipsJson = info.skips ? JSON.stringify(info.skips) : '';
  var dataRow = [[
    completedAt,
    info.period || '',
    info.candidates || 0,
    info.collected || 0,
    info.failed || 0,
    info.status || '완료',
    skipsJson
  ]];

  var histHeaderRow = findProgressHistoryHeaderRow(sh);
  if (histHeaderRow < 0) {
    var startRow = sh.getLastRow() + 2;
    writeSheetValues(sh, startRow, 1, [['[실행 완료 이력]', '']]);
    writeSheetValues(sh, startRow + 1, 1, [['완료시각', '기간', '후보', '확정수집', '실패', '상태', '스킵JSON']]);
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
    'Channel-Version': '2026-06-01',
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

/**
 * 활성 탭(기간별 출력 시트)의 메타베이스 매칭 컬럼(A~F)만 다시 채움.
 * 채널톡 재수집 없이 채팅별 사용자 정보만 다시 조회해서 메타베이스와 매칭.
 * 활성 탭의 이름 suffix로 채널(케어/환급) 자동 판별.
 */
function refreshMetaOnActiveOutputSheet() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getActiveSheet();
  var name = sh.getName();

  var channelKey = null;
  Object.keys(CHANNELS).forEach(function(k) {
    var suffix = '_' + CHANNELS[k].label;
    if (name.length > suffix.length && name.substr(name.length - suffix.length) === suffix
        && /^\d{8}~\d{8}_/.test(name)) {
      channelKey = k;
    }
  });
  if (!channelKey) {
    ui.alert('현재 활성 시트가 기간별 출력 탭이 아닙니다.\n예: "20260824~20260826_케어" 형식의 탭을 활성화한 후 다시 실행하세요.');
    return;
  }

  var res = refreshMetaOnOutputSheet(channelKey, name);
  if (!res.ok) {
    ui.alert('실패: ' + res.error);
    return;
  }
  ui.alert(
    '메타베이스 매칭 갱신 완료 (' + CHANNELS[channelKey].label + ')\n' +
    '- 시트: ' + name + '\n' +
    '- 대상 행: ' + res.total + '\n' +
    '- 매칭 성공: ' + res.matched + '\n' +
    '- 매칭 실패: ' + res.unmatched + '\n' +
    '- 소요 시간: ' + res.elapsedSec + '초'
  );
}

function refreshMetaOnOutputSheet(channelKey, sheetName) {
  var ch = CHANNELS[channelKey];
  if (!ch) return { ok: false, error: '알 수 없는 채널: ' + channelKey };
  var keys = getApiKeys(channelKey);
  if (!keys) return { ok: false, error: ch.label + ' API 키가 없습니다' };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shOut = ss.getSheetByName(sheetName);
  if (!shOut) return { ok: false, error: '출력 시트 없음: ' + sheetName };

  var shMeta = ss.getSheetByName(CFG.SHEET_META);
  if (!shMeta) return { ok: false, error: '"' + CFG.SHEET_META + '" 시트가 없습니다' };

  var meta;
  try {
    meta = buildMetaIndex(shMeta);
  } catch (e) {
    return { ok: false, error: String(e.message || e) };
  }

  var last = shOut.getLastRow();
  if (last < 2) return { ok: false, error: '데이터 행 없음' };

  var numCols = CFG.OUT_HEADERS.length;
  var vals = shOut.getRange(2, 1, last - 1, numCols).getValues();
  var headers = buildHeaders(keys.key, keys.secret);
  var started = new Date().getTime();
  var matched = 0;
  var unmatched = 0;
  var CHUNK = 20;

  for (var i = 0; i < vals.length; i += CHUNK) {
    var chunkRows = vals.slice(i, i + CHUNK);
    var reqMap = [];
    chunkRows.forEach(function(row, k) {
      var link = String(row[7] || '');
      var m = link.match(/user-chats\/([a-f0-9]{20,40})/);
      if (!m) {
        reqMap.push({ localIdx: k, req: null });
        return;
      }
      reqMap.push({
        localIdx: k,
        req: {
          url: CFG.CHANNEL_BASE + '/user-chats/' + m[1],
          method: 'get',
          headers: headers,
          muteHttpExceptions: true
        }
      });
    });
    var validReqs = reqMap.filter(function(x) { return x.req !== null; }).map(function(x) { return x.req; });
    if (validReqs.length === 0) {
      chunkRows.forEach(function() { unmatched++; });
      continue;
    }

    var responses;
    try {
      responses = UrlFetchApp.fetchAll(validReqs);
    } catch (e) {
      chunkRows.forEach(function() { unmatched++; });
      continue;
    }

    var respIdx = 0;
    reqMap.forEach(function(entry) {
      var globalIdx = i + entry.localIdx;
      if (entry.req === null) {
        unmatched++;
        return;
      }
      var res = responses[respIdx++];
      var code = res.getResponseCode();
      if (code < 200 || code >= 300) {
        unmatched++;
        return;
      }
      var data;
      try {
        data = JSON.parse(res.getContentText() || '{}');
      } catch (e) {
        unmatched++;
        return;
      }
      var user = pickUserFromApiDetail(data);
      var profile = (user && user.profile) || {};
      var phone = normalizePhone(profile.mobileNumber || user.mobileNumber || profile.mobile || user.mobile || '');
      var hit = phone ? meta.byPhone[phone] : null;
      if (hit) {
        vals[globalIdx][0] = hit.rpnTin || '';
        vals[globalIdx][1] = hit.userName || vals[globalIdx][1] || '';
        vals[globalIdx][2] = hit.bsno || '';
        vals[globalIdx][3] = hit.tnmNm || '';
        vals[globalIdx][4] = hit.careproMgr || '';
        vals[globalIdx][5] = hit.bkpStatus || '';
        matched++;
      } else {
        unmatched++;
      }
    });
  }

  var updateVals = vals.map(function(r) { return r.slice(0, 6); });
  shOut.getRange(2, 1, updateVals.length, 6).setValues(updateVals);

  var elapsedSec = Math.round((new Date().getTime() - started) / 1000);
  return {
    ok: true,
    channel: channelKey,
    total: vals.length,
    matched: matched,
    unmatched: unmatched,
    elapsedSec: elapsedSec
  };
}

/**
 * 케어 채널 팀/매니저 매핑 수기 세팅 (2026-08-31 채널톡 팀 설정 스냅샷)
 * `_그룹_리스트`(id, title, scope, memberCount)와 `_매니저_팀매핑`(id, name, email, 소속 그룹들)
 * 두 탭을 기존 컬럼 구조 그대로 덮어씀. 스크린샷 기준 팀별 멤버 데이터를 반영.
 * runManagerTeamMapping()으로 환급 채널 API 데이터 다시 뽑고 싶으면 그 함수 실행하면 덮어씀.
 */
function setupManagerTeamMappingManual() {
  // NOTE: 실제 매니저·팀 데이터는 저장소에서 마스킹되었습니다.
  //       원본은 스크립트 실행 시 스프레드시트의 `_그룹_리스트` / `_매니저_팀매핑` 탭을 채우며,
  //       샘플 구조만 예시로 남겨둡니다.
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // _그룹_리스트: (id 없음), title, scope, memberCount
  var groupHeaders = ['id', 'title', 'scope', 'memberCount'];
  var groupRows = [
    ['', '<TEAM_A>', 'public', 10],
    ['', '<TEAM_B>', 'public', 10],
    ['', '<CX_TEAM>', 'public', 6],
    ['', '<PARTNER_TAX_TEAM>', 'public', 6]
  ];

  var shGroups = ss.getSheetByName('_그룹_리스트') || ss.insertSheet('_그룹_리스트');
  shGroups.clearContents();
  shGroups.getRange(1, 1, 1, groupHeaders.length).setValues([groupHeaders]).setFontWeight('bold');
  shGroups.getRange(2, 1, groupRows.length, groupHeaders.length).setValues(groupRows);
  shGroups.setFrozenRows(1);
  shGroups.autoResizeColumns(1, groupHeaders.length);

  // _매니저_팀매핑: (id 없음), name, email, 소속 그룹들 (title, 콤마구분)
  var mgrHeaders = ['id', 'name', 'email', '소속 그룹들 (title, 콤마구분)'];
  var mgrRows = [
    ['', '<MGR_1>', '<MGR_1_EMAIL>', '<TEAM_A>'],
    ['', '<MGR_2>', '<MGR_2_EMAIL>', '<TEAM_B>'],
    ['', '<MGR_3>', '<MGR_3_EMAIL>', '<CX_TEAM>, <CX_CONVERT_TEAM>'],
    ['', '<MGR_4>', '<MGR_4_EMAIL>', '<PARTNER_TAX_TEAM>']
    // ... 실제 저장소 원본은 마스킹 (원본 스크립트에는 약 60여 명의 매니저·팀 매핑 포함)
  ];

  var shMgrs = ss.getSheetByName('_매니저_팀매핑') || ss.insertSheet('_매니저_팀매핑');
  shMgrs.clearContents();
  shMgrs.getRange(1, 1, 1, mgrHeaders.length).setValues([mgrHeaders]).setFontWeight('bold');
  shMgrs.getRange(2, 1, mgrRows.length, mgrHeaders.length).setValues(mgrRows);
  shMgrs.setFrozenRows(1);
  shMgrs.autoResizeColumns(1, mgrHeaders.length);

  SpreadsheetApp.getUi().alert(
    '수기 세팅 완료 (2026-08-31 케어 스냅샷)\n' +
    '- _그룹_리스트: ' + groupRows.length + '팀\n' +
    '- _매니저_팀매핑: ' + mgrRows.length + '명\n' +
    '- 2팀 인원 수 채널톡 표시(12명) 대비 스크린샷 식별 10명. 이수진/유영진 겸직 확인 필요'
  );
}

/* =========================
 * Groq VOC 분류
 * ========================= */
var GROQ_CFG = {
  API_URL: 'https://api.groq.com/openai/v1/chat/completions',
  MODELS_URL: 'https://api.groq.com/openai/v1/models',
  DEFAULT_MODEL: 'llama-3.1-8b-instant',
  PROP_KEY: 'GROQ_API_KEY',
  PROP_MODEL: 'GROQ_MODEL',
  PILOT_SIZE: 10
};

function getGroqModel() {
  return PropertiesService.getScriptProperties().getProperty(GROQ_CFG.PROP_MODEL) || GROQ_CFG.DEFAULT_MODEL;
}

/**
 * 현재 계정에서 사용 가능한 Groq 모델 목록을 팝업으로 보여줌.
 */
function listGroqModels() {
  var ui = SpreadsheetApp.getUi();
  var key = getGroqApiKey();
  if (!key) { ui.alert('먼저 Groq API 키를 설정하세요.'); return; }
  var res = UrlFetchApp.fetch(GROQ_CFG.MODELS_URL, {
    method: 'get',
    headers: { 'Authorization': 'Bearer ' + key },
    muteHttpExceptions: true
  });
  if (res.getResponseCode() < 200 || res.getResponseCode() >= 300) {
    ui.alert('모델 목록 조회 실패 HTTP ' + res.getResponseCode() + '\n' + res.getContentText().slice(0, 500));
    return;
  }
  var data = JSON.parse(res.getContentText());
  var models = (data.data || []).map(function(m) { return m.id; }).sort();
  ui.alert('사용 가능 모델 (' + models.length + '개)\n\n' + models.join('\n') + '\n\n현재 설정: ' + getGroqModel());
}

function setGroqApiKey() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt('Groq API Key', 'gsk_... 형식', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  var key = String(res.getResponseText() || '').trim();
  if (!key) { ui.alert('키가 비었습니다.'); return; }
  PropertiesService.getScriptProperties().setProperty(GROQ_CFG.PROP_KEY, key);
  ui.alert('Groq API 키 저장 완료 (첫 4자: ' + key.slice(0, 4) + '...)');
}

function getGroqApiKey() {
  return PropertiesService.getScriptProperties().getProperty(GROQ_CFG.PROP_KEY);
}

/**
 * 활성 탭에서 첫 10건의 전사문을 미태그 분류 프롬프트로 분류.
 * 결과를 `_VOC_파일럿_결과` 탭에 기록.
 */
function runVocClassificationPilot() {
  var ui = SpreadsheetApp.getUi();
  var apiKey = getGroqApiKey();
  if (!apiKey) { ui.alert('먼저 "Groq API 키 설정"을 실행하세요.'); return; }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getActiveSheet();
  var name = sh.getName();
  if (!/^\d{8}~\d{8}_/.test(name)) {
    ui.alert('현재 활성 시트가 기간별 출력 탭이 아닙니다.\n예: "20260824~20260830_케어"');
    return;
  }

  var last = sh.getLastRow();
  if (last < 2) { ui.alert('데이터 행 없음'); return; }

  var take = Math.min(GROQ_CFG.PILOT_SIZE, last - 1);
  var vals = sh.getRange(2, 1, take, CFG.OUT_HEADERS.length).getValues();

  // 파일럿 입력 구성: 상담일자, 태그, 전사문
  var lines = [];
  vals.forEach(function(row, i) {
    var startTime = String(row[6] || '');
    var tags = String(row[9] || '(없음)');
    var transcript = String(row[10] || '').replace(/\s+/g, ' ').slice(0, 2000); // 컨텍스트 절약
    lines.push(
      '--- 상담 #' + (i + 1) + ' ---\n' +
      '상담일자: ' + startTime + '\n' +
      '태그: ' + (tags || '(없음)') + '\n' +
      '전사문: ' + transcript
    );
  });

  var systemPrompt = '당신은 비즈넵케어 VOC 분석 담당자입니다. 아래 상담 로그에서 태그가 비어있는 건만 골라 전사문을 읽고 고정 카테고리로 분류하세요.\n\n' +
    '카테고리 (이 목록 외 신설 금지):\n' +
    '1. 부가세(추정) - 매입/매출 자료, 신고, 공제, 조기환급, 표준증명\n' +
    '2. 원천세/급여(추정) - 급여자료, 4대보험, 퇴사, 프리랜서 3.3%\n' +
    '3. 종소세(추정) - 경비처리, 공제요건, 예상세액, 환급\n' +
    '4. 기장대리 - 서비스 범위, 계좌/카드 등록, 세무상담 연결, 일반 대리업무\n' +
    '5. 해지/폐업 관련 - 해지 요청, 휴폐업, 재수임 문의\n' +
    '6. 이용료/결제 - 이용료 안내, 결제 취소/변경, 조정료 불만\n' +
    '7. 세금계산서 - 발행/수취/지연/누락\n' +
    '8. 자료제출/서류 - 서류·증빙 단순 주고받기\n' +
    '9. 상담연결/문의 - 담당자 변경, 단순 문의처 확인\n' +
    '10. 기타/단순응대 - 계정/로그인, 인사성 문의 등 세무 내용 없음\n' +
    '11. 미분류(내용 불충분) - 대화 짧거나 맥락 부족\n\n' +
    '출력 형식 (JSON 배열 하나만, 다른 텍스트 금지):\n' +
    '[{"index": 1, "카테고리": "부가세(추정)", "대표발화": "15단어 이내 원문 발췌"}, ...]';

  var model = getGroqModel();
  var body = {
    model: model,
    messages: [
      { role: 'system', content: systemPrompt },
      { role: 'user', content: '아래 상담 로그를 분류하세요.\n\n' + lines.join('\n\n') }
    ],
    temperature: 0.2,
    response_format: { type: 'json_object' }
  };

  var started = new Date().getTime();
  var res;
  try {
    res = UrlFetchApp.fetch(GROQ_CFG.API_URL, {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': 'Bearer ' + apiKey },
      payload: JSON.stringify(body),
      muteHttpExceptions: true
    });
  } catch (e) {
    ui.alert('API 호출 실패: ' + e.message);
    return;
  }

  var code = res.getResponseCode();
  var text = res.getContentText() || '';
  if (code < 200 || code >= 300) {
    ui.alert('Groq API 오류 HTTP ' + code + '\n' + text.slice(0, 500));
    return;
  }

  var data;
  try { data = JSON.parse(text); } catch (e) {
    ui.alert('응답 파싱 실패: ' + text.slice(0, 500));
    return;
  }

  var content = (((data.choices || [])[0] || {}).message || {}).content || '';
  var usage = data.usage || {};
  var parsed;
  try {
    var obj = JSON.parse(content);
    parsed = Array.isArray(obj) ? obj : (obj.results || obj.data || Object.values(obj).find(Array.isArray));
  } catch (e) {
    ui.alert('LLM 응답 JSON 파싱 실패:\n' + content.slice(0, 800));
    return;
  }
  if (!Array.isArray(parsed)) {
    ui.alert('LLM 응답이 배열 아님:\n' + content.slice(0, 800));
    return;
  }

  // 결과 탭에 기록
  var resultSheet = ss.getSheetByName('_VOC_파일럿_결과') || ss.insertSheet('_VOC_파일럿_결과');
  resultSheet.clearContents();
  var headers = ['index', '카테고리', '대표발화', '원본_상담일자', '원본_태그', '원본_대표자'];
  resultSheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  var rows = parsed.map(function(item, i) {
    var origIdx = (item.index || (i + 1)) - 1;
    var orig = vals[origIdx] || [];
    return [
      item.index || (i + 1),
      item['카테고리'] || item.category || '',
      item['대표발화'] || item.quote || '',
      String(orig[6] || ''),
      String(orig[9] || ''),
      String(orig[1] || '')
    ];
  });
  if (rows.length) {
    resultSheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
  resultSheet.setFrozenRows(1);
  resultSheet.autoResizeColumns(1, headers.length);

  var elapsedSec = Math.round((new Date().getTime() - started) / 1000);
  ui.alert(
    'VOC 분류 파일럿 완료\n' +
    '- 모델: ' + model + '\n' +
    '- 입력: ' + take + '건 → 결과: ' + rows.length + '건\n' +
    '- 인풋 토큰: ' + (usage.prompt_tokens || 0) + ', 아웃풋 토큰: ' + (usage.completion_tokens || 0) + '\n' +
    '- 소요: ' + elapsedSec + '초\n' +
    '- 탭: _VOC_파일럿_결과'
  );
}

/* =========================
 * VOC 전체 배치 분류 (Phase 2)
 * ========================= */
var VOC_BATCH_CFG = {
  BATCH_SIZE: 30,
  SLEEP_MS: 2500,   // 24 RPM 안전 마진 (30 RPM 한계 대비)
  MAX_TRANSCRIPT_CHARS: 1500,
  MAX_RUN_MS: 5 * 60 * 1000, // Apps Script 6분 timeout 대비 5분에 중단
  RESUME_PROP_PREFIX: 'VOC_LAST_ROW_'
};

function runVocClassificationFullBatch() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getActiveSheet();
  var name = sh.getName();
  if (!/^\d{8}~\d{8}_/.test(name)) {
    ui.alert('활성 탭이 기간별 출력 탭이 아닙니다.');
    return;
  }
  var result = runVocClassificationCore(name);
  if (!result.ok) { ui.alert('실패: ' + result.error); return; }
  ui.alert(result.message);
}

/**
 * Core: sheetName을 받아서 LLM 배치 분류 수행. UI 없음. 트리거·메뉴 양쪽에서 호출 가능.
 * 반환: { ok, done, processed, total, classified, tokensIn, tokensOut, elapsedSec, message, error }
 */
function runVocClassificationCore(sheetName) {
  var apiKey = getGroqApiKey();
  if (!apiKey) return { ok: false, error: 'Groq API 키 미설정' };
  var model = getGroqModel();

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(sheetName);
  if (!sh) return { ok: false, error: '시트 없음: ' + sheetName };

  var last = sh.getLastRow();
  if (last < 2) return { ok: false, error: '데이터 없음' };
  var totalRows = last - 1;

  var propKey = VOC_BATCH_CFG.RESUME_PROP_PREFIX + sheetName;
  var resumeFrom = Number(PropertiesService.getScriptProperties().getProperty(propKey) || '0');
  if (resumeFrom >= totalRows) {
    resumeFrom = 0;
    PropertiesService.getScriptProperties().deleteProperty(propKey);
  }

  var vals = sh.getRange(2, 1, totalRows, CFG.OUT_HEADERS.length).getValues();

  var resultSheetName = '_VOC_분류_' + sheetName;
  var resultSheet = ss.getSheetByName(resultSheetName);
  var headers = ['원본_row', '카테고리', '대표발화', '상담일자', '대표자', '기장담당자', '태그'];
  if (!resultSheet) {
    resultSheet = ss.insertSheet(resultSheetName);
    resultSheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    resultSheet.setFrozenRows(1);
  } else if (resumeFrom === 0) {
    resultSheet.clearContents();
    resultSheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    resultSheet.setFrozenRows(1);
  }

  var started = new Date().getTime();
  var totalClassified = 0;
  var totalTokensIn = 0;
  var totalTokensOut = 0;
  var i = resumeFrom;
  var batchesRun = 0;

  while (i < totalRows) {
    if ((new Date().getTime() - started) > VOC_BATCH_CFG.MAX_RUN_MS) break;
    var chunk = vals.slice(i, i + VOC_BATCH_CFG.BATCH_SIZE);
    var result = classifyBatchGroq(chunk, i, apiKey, model);
    batchesRun++;

    if (result.ok && result.classified && result.classified.length) {
      var rows = result.classified.map(function(item) {
        var origRow = vals[item.origRowIdx];
        return [
          item.origRowIdx + 2,
          item.category,
          item.quote,
          String(origRow[6] || ''),
          String(origRow[1] || ''),
          String(origRow[4] || ''),
          String(origRow[9] || '')
        ];
      });
      resultSheet.getRange(resultSheet.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
      totalClassified += rows.length;
    }
    totalTokensIn += result.tokensIn || 0;
    totalTokensOut += result.tokensOut || 0;

    i += VOC_BATCH_CFG.BATCH_SIZE;
    PropertiesService.getScriptProperties().setProperty(propKey, String(i));

    try {
      ss.toast('배치 ' + batchesRun + ' — 진행 ' + Math.min(i, totalRows) + '/' + totalRows + ' (분류 ' + totalClassified + '건)', 'VOC 분류', 30);
    } catch (e) {}

    if (i < totalRows) Utilities.sleep(VOC_BATCH_CFG.SLEEP_MS);
  }

  var elapsedSec = Math.round((new Date().getTime() - started) / 1000);
  var doneAll = i >= totalRows;
  if (doneAll) PropertiesService.getScriptProperties().deleteProperty(propKey);

  var msg = (doneAll ? '✅ VOC 전체 배치 분류 완료\n' : '⏸ 5분 timeout — 재실행 or 자동 트리거로 이어서 진행\n') +
    '- 대상: ' + totalRows + '건\n' +
    '- 처리 위치: ' + Math.min(i, totalRows) + '/' + totalRows + '\n' +
    '- 이번 분류: ' + totalClassified + '건 (배치 ' + batchesRun + '개)\n' +
    '- 토큰 IN/OUT: ' + totalTokensIn + ' / ' + totalTokensOut + '\n' +
    '- 소요: ' + elapsedSec + '초\n' +
    '- 탭: ' + resultSheetName;

  return {
    ok: true,
    done: doneAll,
    processed: Math.min(i, totalRows),
    total: totalRows,
    classified: totalClassified,
    tokensIn: totalTokensIn,
    tokensOut: totalTokensOut,
    elapsedSec: elapsedSec,
    resultSheetName: resultSheetName,
    message: msg
  };
}

function classifyBatchGroq(chunk, startIdx, apiKey, model) {
  var lines = [];
  chunk.forEach(function(row, i) {
    var startTime = String(row[6] || '');
    var tags = String(row[9] || '');
    var transcript = String(row[10] || '').replace(/\s+/g, ' ').slice(0, VOC_BATCH_CFG.MAX_TRANSCRIPT_CHARS);
    lines.push(
      '--- 상담 #' + (i + 1) + ' ---\n' +
      '상담일자: ' + startTime + '\n' +
      '태그: ' + (tags || '(없음)') + '\n' +
      '전사문: ' + transcript
    );
  });

  var systemPrompt = '당신은 비즈넵케어 VOC 분석 담당자입니다. 아래 상담 로그의 전사문을 읽고 각 상담을 아래 11개 고정 카테고리 중 하나로만 분류하세요. 태그 유무 무관하게 모든 상담을 처리합니다.\n\n' +
    '카테고리 (신설·변형 금지):\n' +
    '1. 부가세(추정) - 매입/매출 자료, 신고, 공제, 조기환급, 표준증명\n' +
    '2. 원천세/급여(추정) - 급여자료, 4대보험, 퇴사, 프리랜서 3.3%\n' +
    '3. 종소세(추정) - 경비처리, 공제요건, 예상세액, 환급\n' +
    '4. 기장대리 - 서비스 범위, 계좌/카드 등록, 세무상담 연결, 일반 대리업무\n' +
    '5. 해지/폐업 관련 - 해지 요청, 휴폐업, 재수임 문의\n' +
    '6. 이용료/결제 - 이용료 안내, 결제 취소/변경, 조정료 불만\n' +
    '7. 세금계산서 - 발행/수취/지연/누락\n' +
    '8. 자료제출/서류 - 서류·증빙 단순 주고받기\n' +
    '9. 상담연결/문의 - 담당자 변경, 단순 문의처 확인\n' +
    '10. 기타/단순응대 - 계정/로그인, 인사성 문의 등 세무 내용 없음\n' +
    '11. 미분류(내용 불충분) - 대화 짧거나 맥락 부족\n\n' +
    '출력 형식 (JSON만, 다른 텍스트 금지):\n' +
    '{"results": [{"index": <상담 #번호>, "카테고리": "<위 목록 정확한 이름>", "대표발화": "<원문 15단어 이내 발췌>"}, ...]}\n' +
    '입력된 모든 상담을 반드시 결과에 포함시키세요.';

  var body = {
    model: model,
    messages: [
      { role: 'system', content: systemPrompt },
      { role: 'user', content: '상담 로그:\n\n' + lines.join('\n\n') }
    ],
    temperature: 0.1,
    response_format: { type: 'json_object' }
  };

  var res;
  try {
    res = UrlFetchApp.fetch(GROQ_CFG.API_URL, {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': 'Bearer ' + apiKey },
      payload: JSON.stringify(body),
      muteHttpExceptions: true
    });
  } catch (e) {
    return { ok: false, error: e.message };
  }

  if (res.getResponseCode() < 200 || res.getResponseCode() >= 300) {
    return { ok: false, error: 'HTTP ' + res.getResponseCode() + ': ' + res.getContentText().slice(0, 200) };
  }

  var data;
  try { data = JSON.parse(res.getContentText()); } catch (e) {
    return { ok: false, error: 'response JSON parse failed' };
  }

  var usage = data.usage || {};
  var content = (((data.choices || [])[0] || {}).message || {}).content || '';
  var parsed;
  try {
    var obj = JSON.parse(content);
    parsed = obj.results || (Array.isArray(obj) ? obj : []);
  } catch (e) {
    return { ok: false, error: 'llm output parse failed', tokensIn: usage.prompt_tokens, tokensOut: usage.completion_tokens };
  }
  if (!Array.isArray(parsed)) parsed = [];

  var classified = parsed.map(function(item) {
    var relIdx = Number(item.index || 0) - 1;
    if (relIdx < 0 || relIdx >= chunk.length) return null;
    return {
      origRowIdx: startIdx + relIdx,
      category: String(item['카테고리'] || item.category || '').trim(),
      quote: String(item['대표발화'] || item.quote || '').trim()
    };
  }).filter(function(x) { return x !== null; });

  return {
    ok: true,
    classified: classified,
    tokensIn: usage.prompt_tokens || 0,
    tokensOut: usage.completion_tokens || 0
  };
}

/* =========================
 * 리포트 생성 (Phase 1 + 2 통합)
 * ========================= */
function buildCareReport() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getActiveSheet();
  var name = sh.getName();
  if (!/^\d{8}~\d{8}_/.test(name)) {
    ui.alert('활성 탭이 기간별 출력 탭이 아닙니다.');
    return;
  }
  var result = buildCareReportCore(name);
  if (!result.ok) { ui.alert('실패: ' + result.error); return; }
  ui.alert(result.message);
}

/**
 * Core: sheetName을 받아서 리포트 생성. UI 없음. 트리거·메뉴 양쪽에서 호출 가능.
 */
function buildCareReportCore(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(sheetName);
  if (!sh) return { ok: false, error: '시트 없음: ' + sheetName };
  var name = sheetName;

  var last = sh.getLastRow();
  if (last < 2) return { ok: false, error: '데이터 없음' };
  var vals = sh.getRange(2, 1, last - 1, CFG.OUT_HEADERS.length).getValues();
  var total = vals.length;

  // ---- 매니저 → 팀 매핑 로드
  var shMgrs = ss.getSheetByName('_매니저_팀매핑');
  var mgrByName = {};
  if (shMgrs) {
    var mgrVals = shMgrs.getDataRange().getValues();
    for (var m = 1; m < mgrVals.length; m++) {
      var mName = String(mgrVals[m][1] || '').trim();
      var teams = String(mgrVals[m][3] || '').trim();
      if (!mName || !teams) continue;
      var primary = teams.split(',')[0].trim();
      var norm = normalizeMgrName(mName);
      if (norm && !mgrByName[norm]) mgrByName[norm] = primary;
    }
  }

  // ---- 집계 (Phase 1)
  var byTeam = {};
  var byMgr = {};
  var byDow = { '월': 0, '화': 0, '수': 0, '목': 0, '금': 0, '토': 0, '일': 0 };
  var dowKr = ['일', '월', '화', '수', '목', '금', '토'];
  var unmatchedMgrCount = 0;
  var noMgrCount = 0;

  vals.forEach(function(row) {
    var mgr = String(row[4] || '').trim();
    var startTime = row[6];

    if (mgr) {
      byMgr[mgr] = (byMgr[mgr] || 0) + 1;
      var norm = normalizeMgrName(mgr);
      var team = mgrByName[norm];
      if (!team) {
        unmatchedMgrCount++;
        team = '(팀 미매칭)';
      }
      byTeam[team] = (byTeam[team] || 0) + 1;
    } else {
      noMgrCount++;
      byTeam['(담당자 없음)'] = (byTeam['(담당자 없음)'] || 0) + 1;
    }

    if (startTime) {
      var d = (startTime instanceof Date) ? startTime : new Date(startTime);
      if (!isNaN(d.getTime())) byDow[dowKr[d.getDay()]]++;
    }
  });

  // ---- LLM 분류 결과 로드 (Phase 2, 있으면)
  var byCategory = null;
  var cancelCount = null;
  var vocSheet = ss.getSheetByName('_VOC_분류_' + name);
  if (vocSheet && vocSheet.getLastRow() >= 2) {
    byCategory = {};
    var vocVals = vocSheet.getRange(2, 1, vocSheet.getLastRow() - 1, 2).getValues();
    vocVals.forEach(function(vr) {
      var cat = String(vr[1] || '').trim();
      if (cat) byCategory[cat] = (byCategory[cat] || 0) + 1;
    });
    cancelCount = byCategory['해지/폐업 관련'] || 0;
  }

  // ---- 리포트 시트 작성
  var reportSheetName = '_리포트_' + name;
  var reportSheet = ss.getSheetByName(reportSheetName) || ss.insertSheet(reportSheetName);
  reportSheet.clearContents();

  var out = [];
  var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  out.push(['[비즈넵케어 VOC 리포트]', name]);
  out.push(['생성일시', now]);
  out.push([]);

  out.push(['[1. 총 상담 건수]']);
  out.push(['총 건수', total]);
  out.push(['담당자 매칭 있음', total - noMgrCount]);
  out.push(['담당자 없음(미매칭 전사)', noMgrCount]);
  out.push([]);

  out.push(['[2. 요일별 문의 수]']);
  out.push(['요일', '건수']);
  ['월', '화', '수', '목', '금', '토', '일'].forEach(function(d) { out.push([d, byDow[d]]); });
  out.push([]);

  out.push(['[3. 팀별 문의 건수]']);
  out.push(['팀', '건수']);
  Object.keys(byTeam).map(function(k) { return [k, byTeam[k]]; })
    .sort(function(a, b) { return b[1] - a[1]; })
    .forEach(function(r) { out.push(r); });
  out.push([]);

  out.push(['[4. 담당자별 문의 건수]']);
  out.push(['담당자', '건수']);
  Object.keys(byMgr).map(function(k) { return [k, byMgr[k]]; })
    .sort(function(a, b) { return b[1] - a[1]; })
    .forEach(function(r) { out.push(r); });
  out.push([]);

  if (byCategory) {
    out.push(['[5. LLM 카테고리별 건수]']);
    out.push(['카테고리', '건수']);
    Object.keys(byCategory).map(function(k) { return [k, byCategory[k]]; })
      .sort(function(a, b) { return b[1] - a[1]; })
      .forEach(function(r) { out.push(r); });
    out.push([]);
    out.push(['[6. 해지 관련 문의 건수]']);
    out.push(['해지/폐업 카테고리 건수', cancelCount]);
  } else {
    out.push(['[5-6. LLM 카테고리·해지]']);
    out.push(['(LLM 분류 미실행 — `활성 탭 전체 VOC 분류` 먼저 실행하세요)']);
  }

  var padded = out.map(function(r) {
    var arr = r.slice();
    while (arr.length < 2) arr.push('');
    return arr;
  });
  reportSheet.getRange(1, 1, padded.length, 2).setValues(padded);
  reportSheet.setColumnWidth(1, 260);
  reportSheet.setColumnWidth(2, 140);

  var msg = '리포트 생성 완료\n' +
    '- 총 ' + total + '건\n' +
    '- 팀 ' + Object.keys(byTeam).length + '개 / 담당자 ' + Object.keys(byMgr).length + '명\n' +
    '- 담당자 미매칭(팀 매핑 실패): ' + unmatchedMgrCount + '건\n' +
    '- 담당자 없음(전사문 매칭 실패): ' + noMgrCount + '건\n' +
    (byCategory ? ('- LLM 카테고리: ' + Object.keys(byCategory).length + '개 / 해지 관련: ' + cancelCount + '건\n') : '- LLM 분류: 미실행\n') +
    '- 탭: ' + reportSheetName;

  return {
    ok: true,
    total: total,
    reportSheetName: reportSheetName,
    llmDone: !!byCategory,
    message: msg
  };
}

/**
 * 매니저 이름 정규화: 접두사(N팀, 젠트) + 접미사(매니저/팀장/사원/선임/주임/세무사) + 괄호 제거
 * 대소문자 무시.
 * 예: "5팀 이수진 팀장" → "이수진"
 *     "젠트 리나(CX)" → "리나"
 *     "Jane" → "jane"
 */
function normalizeMgrName(raw) {
  if (!raw) return '';
  var s = String(raw).trim();
  s = s.replace(/^@/, '');           // 워크플로 시트의 '@이름' 표기 제거
  s = s.replace(/^\d+팀\s+/, '');
  s = s.replace(/^협세\s+/, '');
  s = s.replace(/^젠트\s+/, '');
  // (CX), (가이드) 같은 접미 괄호 — 공백 유무 관계없이 제거
  s = s.replace(/\s*\([^)]+\)$/, '');
  s = s.replace(/\s+(매니저|팀장|사원|선임|주임|세무사|프로)$/, '');
  return s.trim().toLowerCase();
}

/* =========================
 * 자동 분석 파이프라인 (수집 완료 → LLM → 리포트 → 완료 알림)
 * ========================= */
var AUTO_PIPELINE_CFG = {
  SHEET_PROP: 'AUTO_PIPELINE_SHEET',
  STAGE_PROP: 'AUTO_PIPELINE_STAGE',
  STARTED_AT_PROP: 'AUTO_PIPELINE_STARTED_AT',
  HANDLER: 'runAutoPipelineStep',
  INITIAL_DELAY_SEC: 30,
  LLM_CONTINUE_DELAY_SEC: 60,
  REPORT_DELAY_SEC: 10
};

function startAutoAnalysisPipeline(channelKey, outputSheetName) {
  if (channelKey !== 'care') {
    Logger.log('[AUTO] skipping non-care channel: ' + channelKey);
    return;
  }
  var props = PropertiesService.getScriptProperties();
  props.setProperty(AUTO_PIPELINE_CFG.SHEET_PROP, outputSheetName);
  props.setProperty(AUTO_PIPELINE_CFG.STAGE_PROP, 'LLM');
  props.setProperty(AUTO_PIPELINE_CFG.STARTED_AT_PROP, new Date().toISOString());
  scheduleAutoPipelineTrigger(AUTO_PIPELINE_CFG.INITIAL_DELAY_SEC);
  Logger.log('[AUTO] pipeline started for sheet: ' + outputSheetName);
}

function scheduleAutoPipelineTrigger(delaySec) {
  clearAutoPipelineTriggers();
  ScriptApp.newTrigger(AUTO_PIPELINE_CFG.HANDLER)
    .timeBased()
    .after(Math.max(1, delaySec) * 1000)
    .create();
}

function clearAutoPipelineTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === AUTO_PIPELINE_CFG.HANDLER) {
      ScriptApp.deleteTrigger(t);
    }
  });
}

function clearAutoPipelineState() {
  clearAutoPipelineTriggers();
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty(AUTO_PIPELINE_CFG.SHEET_PROP);
  props.deleteProperty(AUTO_PIPELINE_CFG.STAGE_PROP);
  props.deleteProperty(AUTO_PIPELINE_CFG.STARTED_AT_PROP);
}

/**
 * 트리거 진입점. Apps Script가 정해진 시간에 이 함수를 호출.
 * 현재 stage에 따라 LLM 배치 → REPORT 순으로 진행. 필요 시 다음 트리거 스케줄.
 */
function runAutoPipelineStep() {
  var props = PropertiesService.getScriptProperties();
  var sheetName = props.getProperty(AUTO_PIPELINE_CFG.SHEET_PROP);
  var stage = props.getProperty(AUTO_PIPELINE_CFG.STAGE_PROP);
  if (!sheetName || !stage) {
    Logger.log('[AUTO] no state, clearing triggers');
    clearAutoPipelineTriggers();
    return;
  }
  Logger.log('[AUTO] step stage=' + stage + ' sheet=' + sheetName);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName(sheetName)) {
    Logger.log('[AUTO] target sheet missing, aborting: ' + sheetName);
    clearAutoPipelineState();
    return;
  }

  if (stage === 'LLM') {
    var result = runVocClassificationCore(sheetName);
    if (!result.ok) {
      Logger.log('[AUTO] LLM error: ' + result.error);
      // 실패해도 리포트는 만들 수 있으니 REPORT stage로 넘어감
      props.setProperty(AUTO_PIPELINE_CFG.STAGE_PROP, 'REPORT');
      scheduleAutoPipelineTrigger(AUTO_PIPELINE_CFG.REPORT_DELAY_SEC);
      return;
    }
    if (result.done) {
      Logger.log('[AUTO] LLM complete, scheduling REPORT');
      props.setProperty(AUTO_PIPELINE_CFG.STAGE_PROP, 'REPORT');
      scheduleAutoPipelineTrigger(AUTO_PIPELINE_CFG.REPORT_DELAY_SEC);
    } else {
      Logger.log('[AUTO] LLM partial ' + result.processed + '/' + result.total + ', continuing');
      scheduleAutoPipelineTrigger(AUTO_PIPELINE_CFG.LLM_CONTINUE_DELAY_SEC);
    }
  } else if (stage === 'REPORT') {
    var reportRes = buildCareReportCore(sheetName);
    Logger.log('[AUTO] report done: ' + (reportRes.ok ? reportRes.reportSheetName : reportRes.error));
    clearAutoPipelineState();
    // 최종 완료 상태 흔적 남기기 (진행상황 탭 하단에 append)
    try {
      var progressSheet = ss.getSheetByName(CFG.SHEET_PROGRESS_PREFIX + 'care');
      if (progressSheet) {
        var lastRow = progressSheet.getLastRow();
        progressSheet.getRange(lastRow + 1, 1, 1, 2).setValues([['[자동 분석 완료]', reportRes.reportSheetName || '실패: ' + reportRes.error]]);
      }
    } catch (e) {}
  } else {
    Logger.log('[AUTO] unknown stage: ' + stage);
    clearAutoPipelineState();
  }
}

/**
 * 웹앱에서 자동 분석 파이프라인 시작 (대상 시트 명시)
 */
function api_startAutoPipeline(outputSheetName) {
  if (!outputSheetName) return { ok: false, error: '대상 시트가 없습니다.' };
  if (!/^\d{8}~\d{8}_/.test(outputSheetName)) return { ok: false, error: '기간별 출력 시트가 아닙니다: ' + outputSheetName };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName(outputSheetName)) return { ok: false, error: '시트를 찾을 수 없습니다: ' + outputSheetName };
  try {
    startAutoAnalysisPipeline('care', outputSheetName);
    return { ok: true, message: '자동 분석 파이프라인 시작', targetSheet: outputSheetName };
  } catch (e) {
    return { ok: false, error: String(e && e.message || e) };
  }
}

function startAutoAnalysisPipelineFromActive() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getActiveSheet();
  var name = sh.getName();
  if (!/^\d{8}~\d{8}_/.test(name)) {
    ui.alert('활성 탭이 기간별 출력 탭이 아닙니다.');
    return;
  }
  startAutoAnalysisPipeline('care', name);
  ui.alert(
    '자동 분석 파이프라인 시작\n' +
    '- 대상 탭: ' + name + '\n' +
    '- ' + AUTO_PIPELINE_CFG.INITIAL_DELAY_SEC + '초 후 LLM 분류 자동 시작\n' +
    '- LLM 5분 timeout 시 자동으로 1분 뒤 이어서 처리\n' +
    '- LLM 완료 후 자동으로 리포트 탭 생성\n\n' +
    '진행상황: 하단 토스트 알림 + 진행상황_care 탭'
  );
}

function cancelAutoAnalysisPipeline() {
  var ui = SpreadsheetApp.getUi();
  var props = PropertiesService.getScriptProperties();
  var sheetName = props.getProperty(AUTO_PIPELINE_CFG.SHEET_PROP);
  var stage = props.getProperty(AUTO_PIPELINE_CFG.STAGE_PROP);
  clearAutoPipelineState();
  ui.alert('자동 파이프라인 취소' + (sheetName ? '\n- 대상: ' + sheetName + ' (stage: ' + stage + ')' : ' (실행 중인 파이프라인 없음)'));
}

function showAutoAnalysisPipelineStatus() {
  var ui = SpreadsheetApp.getUi();
  var props = PropertiesService.getScriptProperties();
  var sheetName = props.getProperty(AUTO_PIPELINE_CFG.SHEET_PROP);
  var stage = props.getProperty(AUTO_PIPELINE_CFG.STAGE_PROP);
  var startedAt = props.getProperty(AUTO_PIPELINE_CFG.STARTED_AT_PROP);
  var triggers = ScriptApp.getProjectTriggers().filter(function(t) {
    return t.getHandlerFunction() === AUTO_PIPELINE_CFG.HANDLER;
  });
  if (!sheetName) {
    ui.alert('자동 파이프라인 실행 중 아님');
    return;
  }
  ui.alert(
    '자동 파이프라인 상태\n' +
    '- 대상 탭: ' + sheetName + '\n' +
    '- 현재 단계: ' + stage + '\n' +
    '- 시작 시각: ' + (startedAt || '?') + '\n' +
    '- 대기 트리거: ' + triggers.length + '개'
  );
}
