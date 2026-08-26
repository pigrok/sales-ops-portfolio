/*************************************************
 * 채널톡 대화 전사 + 메타베이스 매칭
 * - 진행상황 탭에 실행 상태/완료값 저장
 * - 시간 초과 시 자동 이어서 수집 (트리거)
 *************************************************/

var CFG = {
  SHEET_META: '메타베이스',
  SHEET_OUT_LEGACY: '채널톡 대화 전사문',
  SHEET_REPORT: '리포트 요약',
  SHEET_REPORT_LEGACY: '리포트요약',
  SHEET_PROGRESS: '진행상황',
  SHEET_JOB: '_job_candidates',

  CHANNEL_BASE: 'https://api.channel.io/open/v5',
  DESK_BASE: 'https://desk.channel.io/YOUR_WORKSPACE/user-chats/',
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
    '전사문'
  ]
};

var JOB_KEYS = {
  STATUS: 'CT_JOB_STATUS',
  START_MS: 'CT_JOB_START_MS',
  END_MS: 'CT_JOB_END_MS',
  NEXT_IDX: 'CT_JOB_NEXT_IDX',
  TOTAL: 'CT_JOB_TOTAL',
  FAILED: 'CT_JOB_FAILED',
  STARTED_AT: 'CT_JOB_STARTED_AT',
  OUTPUT_SHEET: 'CT_JOB_OUTPUT_SHEET'
};

/* =========================
 * 메뉴
 * ========================= */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('채널톡 전사')
    .addItem('API 키 설정(1회)', 'setChannelTalkKeys')
    .addItem('기간 지정 후 수집', 'runChannelTalkExport')
    .addItem('이어서 수집 (수동)', 'runChannelTalkExportResumeManual')
    .addItem('리포트 요약 다시 만들기', 'rebuildReportSummaryFromOutput')
    .addItem('인입경로만 다시 채우기', 'refreshInflowOnOutputSheet')
    .addItem('수집 중지', 'cancelChannelTalkJob')
    .addToUi();
}

/* =========================
 * API 키
 * ========================= */
function setChannelTalkKeys() {
  var ui = SpreadsheetApp.getUi();
  var keyRes = ui.prompt('ChannelTalk Access Key', ui.ButtonSet.OK_CANCEL);
  if (keyRes.getSelectedButton() !== ui.Button.OK) return;

  var secRes = ui.prompt('ChannelTalk Access Secret', ui.ButtonSet.OK_CANCEL);
  if (secRes.getSelectedButton() !== ui.Button.OK) return;

  var key = String(keyRes.getResponseText() || '').trim();
  var sec = String(secRes.getResponseText() || '').trim();
  if (!key || !sec) {
    ui.alert('키가 비어 있습니다.');
    return;
  }

  PropertiesService.getScriptProperties()
    .setProperty('CHANNELTALK_ACCESS_KEY', key)
    .setProperty('CHANNELTALK_ACCESS_SECRET', sec);
  ui.alert('저장 완료');
}

/* =========================
 * 메인 / 이어서 실행
 * ========================= */
function runChannelTalkExport() {
  var ui = SpreadsheetApp.getUi();

  var pending = getJobState();
  if (pending && pending.status === 'running') {
    var ans = ui.alert(
      '이전 수집이 진행 중입니다 (' + pending.nextIdx + '/' + pending.total + ').\n이어서 계속할까요?',
      ui.ButtonSet.YES_NO
    );
    if (ans === ui.Button.YES) {
      runChannelTalkExportResume(false);
      return;
    }
    cancelChannelTalkJobSilent();
  }

  var s = ui.prompt('시작일 (YYYYMMDD)', '예: 20260601', ui.ButtonSet.OK_CANCEL);
  if (s.getSelectedButton() !== ui.Button.OK) return;
  var e = ui.prompt('종료일 (YYYYMMDD)', '예: 20260630', ui.ButtonSet.OK_CANCEL);
  if (e.getSelectedButton() !== ui.Button.OK) return;

  var startDate = parseYmd(s.getResponseText());
  var endDate = parseYmd(e.getResponseText());
  if (!startDate || !endDate) {
    ui.alert('날짜 형식 오류: YYYYMMDD');
    return;
  }
  if (startDate > endDate) {
    ui.alert('시작일이 종료일보다 늦습니다.');
    return;
  }

  if (!getApiKeys()) {
    ui.alert('먼저 API 키 설정(1회)을 실행하세요.');
    return;
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  mustSheet(ss, CFG.SHEET_META);
  ensureSheet(ss, CFG.SHEET_PROGRESS);

  var startMs = startDate.getTime();
  var endMs = endOfDay(endDate).getTime();
  var keys = getApiKeys();
  var headers = buildHeaders(keys.key, keys.secret);

  SpreadsheetApp.getActiveSpreadsheet().toast('대화방 목록 조회 중...', '채널톡 전사', 30);
  var candidates = fetchChatCandidates(headers, startMs, endMs);
  var outputSheetName = buildPeriodSheetName(startMs, endMs);

  initJob(ss, candidates, startMs, endMs, outputSheetName);
  initOutputSheet(ss, outputSheetName);

  writeProgress(ss, {
    status: '진행중',
    filter: 'updatedAt·메시지 모두 기간 내',
    period: formatYmd(startDate) + ' ~ ' + formatYmd(endDate),
    outputSheet: outputSheetName,
    candidates: candidates.length,
    scanned: 0,
    collected: 0,
    failed: 0,
    nextIdx: 0,
    progressPct: '0%',
    note: '수집 시작'
  });

  var result = runChannelTalkExportResume(false);
  if (result && result.done) {
    ui.alert(buildDoneMessage(result));
  } else if (result && !result.done) {
    ui.alert(
      '부분 완료 (시간 제한)\n' +
      '- 진행: ' + result.nextIdx + '/' + result.total + '\n' +
      '- 확정 수집: ' + result.collected + '건\n' +
      '- 약 1분 후 자동 이어서 수집됩니다.\n' +
      '- 진행상황 탭을 확인하세요.'
    );
  }
}

function runChannelTalkExportResumeManual() {
  var job = getJobState();
  if (!job || job.status !== 'running') {
    SpreadsheetApp.getUi().alert('이어서 할 작업이 없습니다.「기간 지정 후 수집」을 먼저 실행하세요.');
    return;
  }
  var result = runChannelTalkExportResume(false);
  if (result && result.done) {
    SpreadsheetApp.getUi().alert(buildDoneMessage(result));
  } else if (result && !result.done) {
    SpreadsheetApp.getUi().alert(
      '부분 완료\n' +
      '- 진행: ' + result.nextIdx + '/' + result.total + '\n' +
      '- 확정 수집: ' + result.collected + '건\n' +
      '- 약 1분 후 자동 이어서 수집됩니다.\n' +
      '- 진행상황 탭을 확인하세요.'
    );
  }
}

/** 트리거/메뉴 공통 이어서 실행 */
function runChannelTalkExportResume(fromTrigger) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    if (!fromTrigger) {
      SpreadsheetApp.getUi().alert(
        '이미 다른 수집이 실행 중입니다.\n잠시 후 진행상황 탭을 확인하거나, 자동 이어서 수집이 끝날 때까지 기다려 주세요.'
      );
    }
    return null;
  }

  try {
    return runChannelTalkExportResumeLocked(fromTrigger);
  } finally {
    lock.releaseLock();
  }
}

function runChannelTalkExportResumeLocked(fromTrigger) {
  var started = new Date().getTime();
  var deadline = started + CFG.MAX_RUN_MS;
  var job = getJobState();

  if (!job || job.status !== 'running') {
    clearResumeTriggers();
    return null;
  }

  var keys = getApiKeys();
  if (!keys) {
    writeProgress(SpreadsheetApp.getActiveSpreadsheet(), {
      status: '오류',
      note: 'API 키 없음'
    });
    return null;
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shMeta = mustSheet(ss, CFG.SHEET_META);
  var outputSheetName = resolveJobOutputSheet(job);
  var shOut = mustOutputSheet(ss, outputSheetName);
  var meta = buildMetaIndex(shMeta);
  var headers = buildHeaders(keys.key, keys.secret);
  var candidates = loadJobCandidates(ss);

  if (!candidates.length) {
    cancelChannelTalkJobSilent();
    return null;
  }

  if (job.total && job.total !== candidates.length) {
    PropertiesService.getScriptProperties().setProperty(JOB_KEYS.TOTAL, String(candidates.length));
  }

  var nextIdx = job.nextIdx;
  var failed = job.failed;
  var batchRows = [];
  var collectedBefore = Math.max(0, shOut.getLastRow() - 1);

  while (nextIdx < candidates.length) {
    if (new Date().getTime() > deadline) break;

    var chat = candidates[nextIdx];
    nextIdx++;

    if (nextIdx === 1 || nextIdx % CFG.PROGRESS_EVERY === 0) {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        '전사 수집 ' + nextIdx + '/' + candidates.length + ' (확정 ' + (collectedBefore + batchRows.length) + '건)',
        '채널톡 전사',
        30
      );
    }

    var result = null;
    try {
      result = fetchTranscriptAndInflow(chat.id, headers, job.startMs, job.endMs);
    } catch (fetchErr) {
      failed++;
      continue;
    }
    if (!result || !result.transcript) continue;

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
      CFG.DESK_BASE + chat.id,
      result.inflow || '채널톡',
      (chat.tags || []).join(', '),
      truncateText(result.transcript, CFG.MAX_TRANSCRIPT_CHAR)
    ]);
  }

  appendOutputRows(shOut, batchRows);
  saveJobProgress(nextIdx, failed);

  var collected = Math.max(0, shOut.getLastRow() - 1);
  var elapsedSec = Math.round((new Date().getTime() - started) / 1000);
  var done = nextIdx >= candidates.length;
  var pct = candidates.length ? Math.round((nextIdx / candidates.length) * 100) + '%' : '100%';
  var periodLabel = formatYmd(new Date(job.startMs)) + ' ~ ' + formatYmd(new Date(job.endMs));

  if (done) {
    clearResumeTriggers();
    clearJobState();
    deleteJobCandidatesSheet(ss);
    var allRows = readOutputDataRows(shOut);
    buildReportSummary(ss, allRows, outputSheetName);
    writeProgress(ss, {
      status: '완료',
      filter: 'updatedAt·메시지 모두 기간 내',
      period: periodLabel,
      outputSheet: outputSheetName,
      candidates: candidates.length,
      scanned: nextIdx,
      collected: collected,
      failed: failed,
      nextIdx: nextIdx,
      progressPct: '100%',
      elapsedSec: elapsedSec,
      note: '최종 실행 완료'
    });
    appendProgressHistory(ss, {
      completedAt: new Date(),
      period: periodLabel,
      candidates: candidates.length,
      collected: collected,
      failed: failed,
      status: '완료'
    });
  } else {
    scheduleResumeTrigger();
    writeProgress(ss, {
      status: '이어서 대기',
      filter: 'updatedAt·메시지 모두 기간 내',
      period: periodLabel,
      outputSheet: outputSheetName,
      candidates: candidates.length,
      scanned: nextIdx,
      collected: collected,
      failed: failed,
      nextIdx: nextIdx,
      progressPct: pct,
      elapsedSec: elapsedSec,
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
  cancelChannelTalkJobSilent();
  SpreadsheetApp.getUi().alert('수집 작업을 중지했습니다.');
}

function cancelChannelTalkJobSilent() {
  clearResumeTriggers();
  clearJobState();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  deleteJobCandidatesSheet(ss);
  writeProgress(ss, {
    status: '중지됨',
    note: '사용자 또는 새 작업으로 중지'
  });
}

function buildDoneMessage(result) {
  return (
    '완료\n' +
    '- 출력 탭: ' + (result.outputSheet || '') + '\n' +
    '- 후보: ' + result.total + '건\n' +
    '- 처리: ' + result.nextIdx + '건\n' +
    '- 확정 수집: ' + result.collected + '건\n' +
    (result.failed ? ('- 전사 실패(건너뜀): ' + result.failed + '건\n') : '') +
    '- 리포트: ' + buildReportSheetName(result.outputSheet || '') + '\n' +
    '- 진행상황 탭에 최종 결과가 저장되었습니다.'
  );
}

/* =========================
 * Job 상태
 * ========================= */
function initJob(ss, candidates, startMs, endMs, outputSheet) {
  clearResumeTriggers();
  saveJobCandidates(ss, candidates);
  var props = PropertiesService.getScriptProperties();
  props.setProperty(JOB_KEYS.STATUS, 'running');
  props.setProperty(JOB_KEYS.START_MS, String(startMs));
  props.setProperty(JOB_KEYS.END_MS, String(endMs));
  props.setProperty(JOB_KEYS.NEXT_IDX, '0');
  props.setProperty(JOB_KEYS.TOTAL, String(candidates.length));
  props.setProperty(JOB_KEYS.FAILED, '0');
  props.setProperty(JOB_KEYS.STARTED_AT, String(new Date().getTime()));
  props.setProperty(JOB_KEYS.OUTPUT_SHEET, sanitizeSheetName(outputSheet));
}

function saveJobProgress(nextIdx, failed) {
  var props = PropertiesService.getScriptProperties();
  props.setProperty(JOB_KEYS.NEXT_IDX, String(nextIdx));
  props.setProperty(JOB_KEYS.FAILED, String(failed));
}

function getJobState() {
  var props = PropertiesService.getScriptProperties();
  var status = props.getProperty(JOB_KEYS.STATUS);
  if (!status) return null;
  return {
    status: status,
    startMs: Number(props.getProperty(JOB_KEYS.START_MS) || 0),
    endMs: Number(props.getProperty(JOB_KEYS.END_MS) || 0),
    nextIdx: Number(props.getProperty(JOB_KEYS.NEXT_IDX) || 0),
    total: Number(props.getProperty(JOB_KEYS.TOTAL) || 0),
    failed: Number(props.getProperty(JOB_KEYS.FAILED) || 0),
    startedAt: Number(props.getProperty(JOB_KEYS.STARTED_AT) || 0),
    outputSheet: props.getProperty(JOB_KEYS.OUTPUT_SHEET) || ''
  };
}

function clearJobState() {
  var props = PropertiesService.getScriptProperties();
  Object.keys(JOB_KEYS).forEach(function(k) {
    props.deleteProperty(JOB_KEYS[k]);
  });
}

function saveJobCandidates(ss, candidates) {
  var sh = ensureSheet(ss, CFG.SHEET_JOB);
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
  try {
    sh.hideSheet();
  } catch (hideErr) {
    // 이미 숨김 상태면 무시
  }
}

function loadJobCandidates(ss) {
  var sh = ss.getSheetByName(CFG.SHEET_JOB);
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

function deleteJobCandidatesSheet(ss) {
  var sh = ss.getSheetByName(CFG.SHEET_JOB);
  if (sh) ss.deleteSheet(sh);
}

function scheduleResumeTrigger() {
  clearResumeTriggers();
  ScriptApp.newTrigger('runChannelTalkExportResume')
    .timeBased()
    .after(CFG.RESUME_TRIGGER_MS)
    .create();
}

function clearResumeTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'runChannelTalkExportResume') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

function getApiKeys() {
  var props = PropertiesService.getScriptProperties();
  var key = props.getProperty('CHANNELTALK_ACCESS_KEY');
  var secret = props.getProperty('CHANNELTALK_ACCESS_SECRET');
  if (!key || !secret) return null;
  return { key: key, secret: secret };
}

/* =========================
 * ChannelTalk 수집
 * ========================= */
function fetchChatCandidates(headers, startMs, endMs) {
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
  var inflow = fetchUserChatInflow(chatId, headers);
  var slice = fetchTranscriptSliceInRange(chatId, headers, startMs, endMs);
  if (!slice) return null;
  slice.inflow = inflow;
  return slice;
}

function fetchUserChatInflow(chatId, headers) {
  try {
    var data = ctGet('/user-chats/' + chatId, {}, headers);
    var chat = data.userChat || data.chat || {};
    var user = pickUserFromApiDetail(data);
    return resolveInflow(user, chat);
  } catch (e) {
    return '채널톡';
  }
}

function pickUserFromApiDetail(data) {
  if (!data) return {};
  if (data.user && typeof data.user === 'object' && !Array.isArray(data.user)) return data.user;
  if (data.users && data.users.length) return data.users[0];
  return {};
}

function extractChatIdFromDeskLink(link) {
  var m = String(link || '').match(/user-chats\/([^/?#]+)/i);
  return m ? m[1] : '';
}

function refreshInflowOnOutputSheet() {
  var keys = getApiKeys();
  if (!keys) {
    SpreadsheetApp.getUi().alert('먼저 API 키 설정(1회)을 실행하세요.');
    return;
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shOut = resolveTargetOutputSheet(ss);
  if (!shOut || shOut.getLastRow() < 2) {
    return;
  }

  var headers = buildHeaders(keys.key, keys.secret);
  var rows = readOutputDataRows(shOut);
  var updated = 0;

  for (var i = 0; i < rows.length; i++) {
    var chatId = extractChatIdFromDeskLink(rows[i][7]);
    if (!chatId) continue;
    rows[i][8] = fetchUserChatInflow(chatId, headers);
    updated++;
    if (updated % 20 === 0) {
      ss.toast('인입경로 갱신 ' + updated + '/' + rows.length, '채널톡 전사', 15);
    }
  }

  writeSheetValues(shOut, 2, 1, rows);
  buildReportSummary(ss, rows, shOut.getName());
  SpreadsheetApp.getUi().alert(
    '인입경로 갱신 완료\n' +
    '- 탭: ' + shOut.getName() + '\n' +
    '- 처리: ' + updated + '건'
  );
}

function fetchTranscriptSliceInRange(chatId, headers, startMs, endMs) {
  var since = '';
  var guard = 0;
  var inPeriod = [];

  while (guard++ < 200) {
    var params = { sortOrder: 'desc', limit: CFG.MESSAGE_PAGE_LIMIT };
    if (since) params.since = since;

    var data = ctGet('/user-chats/' + chatId + '/messages', params, headers);
    var msgs = data.messages || [];
    if (!msgs.length) break;

    var oldestInPage = Infinity;

    for (var i = 0; i < msgs.length; i++) {
      var m = msgs[i];
      var createdMs = Number(m.createdAt || 0);
      if (!createdMs) continue;
      if (createdMs < oldestInPage) oldestInPage = createdMs;
      if (createdMs > endMs) continue;
      if (createdMs < startMs) continue;

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

  if (!inPeriod.length) return null;

  var firstMsgMs = inPeriod[0].ms;
  var lastMsgMs = inPeriod[0].ms;

  for (var a = 0; a < inPeriod.length; a++) {
    var item = inPeriod[a];
    if (item.ms < firstMsgMs) firstMsgMs = item.ms;
    if (item.ms > lastMsgMs) lastMsgMs = item.ms;
  }

  var slice = [];
  for (var b = 0; b < inPeriod.length; b++) {
    if (inPeriod[b].ms >= firstMsgMs && inPeriod[b].ms <= lastMsgMs) {
      slice.push(inPeriod[b]);
    }
  }

  slice.sort(function(x, y) { return x.ms - y.ms; });
  var out = [];
  for (var c = 0; c < slice.length; c++) out.push(slice[c].line);
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
 * 출력 / 진행상황 / 리포트
 * ========================= */
function writeProgress(ss, info) {
  var sh = ensureSheet(ss, CFG.SHEET_PROGRESS);
  var savedHistory = extractProgressHistoryRows(sh);
  sh.clearContents();

  var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  var rows = [
    ['현재 작업 상태', ''],
    ['상태', info.status || ''],
    ['필터 기준', info.filter || ''],
    ['기간', info.period || ''],
    ['출력 탭', info.outputSheet || ''],
    ['후보 대화방', info.candidates !== undefined ? info.candidates : ''],
    ['처리한 대화방', info.scanned !== undefined ? info.scanned : ''],
    ['다음 시작 인덱스', info.nextIdx !== undefined ? info.nextIdx : ''],
    ['진행률', info.progressPct || ''],
    ['확정 수집', info.collected !== undefined ? info.collected : ''],
    ['전사 실패(건너뜀)', info.failed !== undefined ? info.failed : ''],
    ['이번 실행 소요(초)', info.elapsedSec !== undefined ? info.elapsedSec : ''],
    ['비고', info.note || ''],
    ['갱신시각', now]
  ];

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

function appendProgressHistory(ss, info) {
  var sh = ensureSheet(ss, CFG.SHEET_PROGRESS);
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

function buildPeriodSheetName(startMs, endMs) {
  return sanitizeSheetName(formatYmd(new Date(startMs)) + '~' + formatYmd(new Date(endMs)));
}

function buildReportSheetName(outputSheetName) {
  return sanitizeSheetName(outputSheetName) + '_요약';
}

function sanitizeSheetName(name) {
  var s = String(name || '').trim();
  if (!s) return '전사문';
  return s.replace(/[\[\]\:\*\?\/\\]/g, '').substring(0, 100);
}

function resolveJobOutputSheet(job) {
  if (job && job.outputSheet) return job.outputSheet;
  return CFG.SHEET_OUT_LEGACY;
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
  if (!sh) {
    sh = ensureOutputSheet(ss, sheetName, false);
  }
  return sh;
}

function isOutputSheet(sh) {
  if (!sh || sh.getLastRow() < 1) return false;
  var header = readSheetValues(sh, 1, 1, 1, CFG.OUT_HEADERS.length)[0];
  if (!header) return false;
  for (var i = 0; i < CFG.OUT_HEADERS.length; i++) {
    if (String(header[i] || '') !== CFG.OUT_HEADERS[i]) return false;
  }
  return true;
}

function resolveTargetOutputSheet(ss) {
  var active = ss.getActiveSheet();
  if (active && isOutputSheet(active)) return active;

  var job = getJobState();
  if (job && job.outputSheet) {
    var shJob = ss.getSheetByName(job.outputSheet);
    if (shJob) return shJob;
  }

  var shLegacy = ss.getSheetByName(CFG.SHEET_OUT_LEGACY);
  if (shLegacy && shLegacy.getLastRow() >= 2) return shLegacy;

  SpreadsheetApp.getUi().alert(
    '전사문 탭을 선택한 뒤 다시 실행하세요.\n' +
    '(예: 20260601~20260630)'
  );
  return null;
}

function appendOutputRows(shOut, rows) {
  if (!rows || !rows.length) return;
  var start = shOut.getLastRow() + 1;
  writeSheetValues(shOut, start, 1, rows);
}

function readOutputDataRows(shOut) {
  var last = shOut.getLastRow();
  if (last < 2) return [];
  return readSheetValues(shOut, 2, 1, last - 1, CFG.OUT_HEADERS.length);
}

function rebuildReportSummaryFromOutput() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shOut = resolveTargetOutputSheet(ss);
  if (!shOut || shOut.getLastRow() < 2) {
    return;
  }
  var rows = readOutputDataRows(shOut);
  buildReportSummary(ss, rows, shOut.getName());
  SpreadsheetApp.getUi().alert(
    '리포트 요약을 다시 만들었습니다.\n' +
    '- 탭: ' + buildReportSheetName(shOut.getName()) + '\n' +
    '- 전체 대화 수: ' + rows.length + '건'
  );
}

function getReportSheetForOutput(ss, outputSheetName) {
  return ensureSheet(ss, buildReportSheetName(outputSheetName));
}

function buildReportSummary(ss, rows, outputSheetName) {
  var sh = outputSheetName
    ? getReportSheetForOutput(ss, outputSheetName)
    : ensureSheet(ss, CFG.SHEET_REPORT);
  sh.clearContents();

  var IDX = { INFLOW: 8, TAGS: 9 };
  var totalChats = rows.length;

  var inflowCnt = {};
  var tagCnt = {};
  var totalTagSlots = 0;
  var maxTagsOnChat = 0;

  for (var i = 0; i < rows.length; i++) {
    var inflow = String(rows[i][IDX.INFLOW] || '미확인').trim() || '미확인';
    inflowCnt[inflow] = (inflowCnt[inflow] || 0) + 1;

    var tagsRaw = String(rows[i][IDX.TAGS] || '').trim();
    var tags = tagsRaw
      ? tagsRaw.split(',').map(function(t) { return t.trim(); }).filter(Boolean)
      : [];
    totalTagSlots += tags.length;
    if (tags.length > maxTagsOnChat) maxTagsOnChat = tags.length;

    for (var t = 0; t < tags.length; t++) {
      tagCnt[tags[t]] = (tagCnt[tags[t]] || 0) + 1;
    }
  }

  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  var out = [];

  out.push(['리포트 생성일시', today]);
  out.push([]);
  out.push(['전체 대화 수', totalChats]);
  out.push([]);

  out.push(['[인입경로별 수]']);
  out.push(['인입경로', '대화 수']);
  var inflowSorted = Object.keys(inflowCnt).map(function(k) {
    return [k, inflowCnt[k]];
  }).sort(function(a, b) { return b[1] - a[1]; });
  for (var a = 0; a < inflowSorted.length; a++) out.push(inflowSorted[a]);
  if (!inflowSorted.length) out.push(['미확인', 0]);

  out.push([]);
  out.push(['전체 태그 수', totalTagSlots]);
  out.push(['최대 태그 수', maxTagsOnChat]);
  out.push([]);

  out.push(['[태그별 수]']);
  out.push(['태그', '대화 수']);
  var tagSorted = Object.keys(tagCnt).map(function(k) {
    return [k, tagCnt[k]];
  }).sort(function(a, b) { return b[1] - a[1]; });
  for (var b = 0; b < tagSorted.length; b++) out.push(tagSorted[b]);
  if (!tagSorted.length) out.push(['(태그 없음)', 0]);

  var values = out.map(function(r) {
    return [r[0] !== undefined ? r[0] : '', r[1] !== undefined ? r[1] : ''];
  });
  writeSheetValues(sh, 1, 1, values);
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

      if (code === 429) {
        Utilities.sleep(Math.min(30000, Math.pow(2, i) * 1000));
        continue;
      }
      if (code >= 500) {
        Utilities.sleep(Math.min(12000, Math.pow(2, i) * 700));
        continue;
      }
      if (code >= 400) {
        throw new Error('API 오류 ' + code + ' ' + path + ': ' + body.slice(0, 300));
      }

      try {
        return JSON.parse(body || '{}');
      } catch (parseErr) {
        if (i === CFG.API_RETRY - 1) {
          throw new Error('JSON 파싱 실패 ' + path + ' / 길이=' + body.length);
        }
        Utilities.sleep(Math.min(12000, Math.pow(2, i) * 700));
      }
    } catch (fetchErr) {
      lastErr = String(fetchErr.message || fetchErr);
      if (i === CFG.API_RETRY - 1) {
        throw new Error('API 연결 실패 ' + path + ': ' + lastErr);
      }
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
    if ((b.type === 'text' || b.type === 'code') && b.value) {
      parts.push(String(b.value));
    }
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
    if (!row || !row.length) {
      row = [''];
    }
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

    if (chunk.length !== (rowEnd - rowStart + 1)) {
      throw new Error('시트 쓰기 범위 오류: rows=' + chunk.length + ', range=' + (rowEnd - rowStart + 1));
    }

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
