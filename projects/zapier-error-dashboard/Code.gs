// ============================================================
// CONFIG
// ============================================================
const CONFIG = {
  SHEET_NAME:    'ZapErrors',
  LOG_SHEET:     'WebhookLog',
  DELETE_LOG:    'DeleteLog',
  SLACK_WEBHOOK_URL: PropertiesService.getScriptProperties().getProperty('SLACK_WEBHOOK_URL') || '',
  GROQ_API_KEY:      PropertiesService.getScriptProperties().getProperty('GROQ_API_KEY') || '',
};

// Pipedrive 계정 (Zap 이름으로 자동 판별)
function getPipedriveAccount(zapName) {
  const p = PropertiesService.getScriptProperties();
  if (zapName.includes('케어')) return {
    domain: p.getProperty('PD_CARE_DOMAIN'),
    token:  p.getProperty('PD_CARE_TOKEN'),
    name:   '케어'
  };
  if (zapName.includes('법인') && !zapName.includes('개인')) return {
    domain: p.getProperty('PD_CORPORATE_DOMAIN'),
    token:  p.getProperty('PD_CORPORATE_TOKEN'),
    name:   '법인'
  };
  return {
    domain: p.getProperty('PD_INDIVIDUAL_DOMAIN'),
    token:  p.getProperty('PD_INDIVIDUAL_TOKEN'),
    name:   '개인'
  };
}

// Sheets 컬럼 (14열, M열 고객전화 제거 - ISMS)
const COL = {
  timestamp: 1, zapName: 2, stepName: 3, errorMessage: 4,
  zapId: 5, taskId: 6, analysis: 7, zapLink: 8,
  status: 9, refundId: 10, actionTime: 11,
  qaResult: 12, templateId: 13, dealId: 14,
};

// ============================================================
// Web App
// ============================================================

function _getZapierSession() {
  const c = CacheService.getScriptCache();
  return {
    zapsession: c.get('ZAPIER_SESSION') || '',
    csrftoken:  c.get('ZAPIER_CSRF')    || '',
    accountId:  c.get('ZAPIER_ACCOUNT_ID') || '3996284',
  };
}

function _checkToken(token) {
  const valid = PropertiesService.getScriptProperties().getProperty('ZAPIER_API_KEY');
  return valid && token === valid;
}

function _unauthorizedPage() {
  return HtmlService.createHtmlOutput(
    '<div style="font-family:-apple-system,sans-serif;display:flex;align-items:center;justify-content:center;height:100vh;margin:0;background:#f8f9fa">' +
    '<div style="text-align:center;padding:40px">' +
    '<div style="font-size:48px;margin-bottom:16px">🔒</div>' +
    '<h2 style="color:#111;margin:0 0 8px">접근 권한이 없습니다</h2>' +
    '<p style="color:#6b7280;margin:0">올바른 접근 URL을 사용해주세요.</p>' +
    '</div></div>'
  );
}

function doGet(e) {
  const action = e?.parameter?.action;
  const token  = e?.parameter?.token;

  if (action === 'getErrorsMissingData') {
    if (!_checkToken(token)) {
      return ContentService.createTextOutput(JSON.stringify({ error: 'unauthorized' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService.createTextOutput(JSON.stringify(getErrorsMissingData()))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (!_checkToken(token)) return _unauthorizedPage();

  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Zapier Error Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

function doPost(e) {
  const rawContents = e?.postData?.contents || '';

  try {
    const raw = JSON.parse(rawContents);

    // ── 토큰 인증 — URL 쿼리파라미터 또는 body.token ────────────────────
    const token = e?.parameter?.token || raw.token || '';
    if (!_checkToken(token)) {
      return ContentService.createTextOutput(JSON.stringify({ error: 'unauthorized' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ── Chrome 익스텐션: step 데이터 업데이트 + 자동 QA ─────────────────
    if (raw.action === 'updateErrorStepData') {
      if (raw.taskId) {
        updateErrorStepDataByTaskId(raw.taskId, '', raw.templateId || '', raw.dealId || '');
        autoQAByTaskId(raw.taskId, raw.phone || ''); // phone은 파라미터로만, 시트 미저장
      } else if (raw.rowId) {
        updateErrorStepData(Number(raw.rowId), '', raw.templateId || '', raw.dealId || '');
      }
      return ContentService.createTextOutput(JSON.stringify({ success: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ── Chrome 확장프로그램 쿠키 동기화 (세션토큰 로그 기록 없이 처리) ──────
    if (raw.action === 'syncCookies') {
      const cache = CacheService.getScriptCache();
      if (raw.zapsession) cache.put('ZAPIER_SESSION',    raw.zapsession, 21600);
      if (raw.csrftoken)  cache.put('ZAPIER_CSRF',       raw.csrftoken,  21600);
      if (raw.sessionJwt) cache.put('ZAPIER_JWT',        raw.sessionJwt, 21600);
      if (raw.accountId)  cache.put('ZAPIER_ACCOUNT_ID', raw.accountId,  21600);
      if (raw.ssoid)      cache.put('ZAPIER_SSOID',      raw.ssoid,      21600);
      PropertiesService.getScriptProperties().setProperty('ZAPIER_COOKIE_SYNCED_AT', new Date().toISOString());
      return ContentService.createTextOutput(JSON.stringify({ success: true, action: 'cookies_saved' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    _logWebhook('POST', _sanitizeLog(raw));

    // Extension 릴레이: taskId가 이미 존재하면 stepName/errorMessage 업데이트 후 종료
    if (raw.source === 'extension_poll' && (raw.task_id || raw.taskId)) {
      const tid = raw.task_id || raw.taskId;
      const existingRow = _findRowByTaskId(tid);
      if (existingRow > 0) {
        const sheet = getSheet();
        const stepName = raw.step_name || raw.stepName || '';
        const errMsg   = raw.error_message || raw.errorMessage || '';
        if (stepName) sheet.getRange(existingRow, COL.stepName).setValue(stepName);
        if (errMsg && errMsg !== '에러 발생 — Zapier 히스토리 확인') {
          sheet.getRange(existingRow, COL.errorMessage).setValue(errMsg);
          const analysis = analyzeWithGroq({ zapName: raw.zap_name || '', stepName, errorMessage: errMsg });
          sheet.getRange(existingRow, COL.analysis).setValue(analysis);
        }
        _logWebhook('EXT_UPDATE', `row=${existingRow} taskId=${tid}`);
        return ContentService.createTextOutput(JSON.stringify({ success: true, type: 'updated' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }

    const rawPhone = raw.customer_phone || raw.customerPhone || '';
    const error = {
      timestamp:     new Date().toISOString(),
      zapName:       raw.zap_name      || raw.zapName      || 'Unknown Zap',
      stepName:      raw.step_name     || raw.stepName     || raw.service || 'Unknown Step',
      errorMessage:  raw.error_message || raw.errorMessage || raw.message || 'No message',
      zapId:         raw.zap_id        || raw.zapId        || '',
      taskId:        raw.task_id       || raw.taskId       || '',
      zapLink:       raw.zap_link      || raw.zapLink      || '',
      customerPhone: rawPhone,
      customerPhoneMasked: maskPhone(rawPhone),
      templateId:    raw.template_id   || raw.templateId   || '',
      dealId:        raw.deal_id       || raw.dealId       || '',
      refundId:      raw.refund_id     || raw.refundId     || '',
    };

    // ── 인증 오류 감지 (RefreshAuthError 등) ──────────────────────────
    const errLower = error.errorMessage.toLowerCase();
    const isAuthError = errLower.includes('refreshautherror')
      || errLower.includes('refresh_auth_error')
      || (errLower.includes('auth') && errLower.includes('expired'))
      || errLower.includes('reconnect your');

    if (isAuthError) {
      const rowIndex = logToSheet(error);
      const authNote = '🔑 인증 오류: Zapier My Apps에서 해당 계정을 재연결해야 합니다.\nzapier.com/app/connections → 해당 앱 → Reconnect';
      updateAnalysis(rowIndex, authNote);
      const qaNote = { summary: '⚠️ OAuth 토큰 만료 - QA 불필요. Zapier My Apps에서 재연결 후 보류 해제하세요.' };
      updateQAResult(rowIndex, qaNote);
      getSheet().getRange(rowIndex, COL.status).setValue('보류');
      getSheet().getRange(rowIndex, COL.actionTime).setValue(
        new Date().toLocaleString('ko-KR', { timeZone: 'Asia/Seoul' })
      );
      error.analysis = authNote;
      error.qa = { action: '보류', checked: false, summary: qaNote.summary };
      sendSlack(error);
      return ContentService.createTextOutput(JSON.stringify({ success: true, type: 'auth_error' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ── 일반 오류 처리 ───────────────────────────────────────────────
    const rowIndex = logToSheet(error);
    const [analysis, qa] = [analyzeWithGroq(error), runQA(error)];

    updateAnalysis(rowIndex, analysis);
    updateQAResult(rowIndex, qa);

    if (qa.checked) {
      getSheet().getRange(rowIndex, COL.status).setValue(qa.action);
      getSheet().getRange(rowIndex, COL.actionTime).setValue(
        new Date().toLocaleString('ko-KR', { timeZone: 'Asia/Seoul' })
      );
    }

    error.analysis = analysis;
    error.qa = qa;
    sendSlack(error);

    return ContentService.createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    _logWebhook('ERROR', err.message);
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================================
// Sheets
// ============================================================

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
    const headers = [
      '타임스탬프','잽 이름','오류 단계','오류 메시지','Zap ID','Task ID',
      '원인 분석','Zap 링크','상태','refundId','처리 시간',
      '처리 확인','템플릿 ID','딜 ID'
    ];
    sheet.appendRow(headers);
    const h = sheet.getRange(1, 1, 1, headers.length);
    h.setBackground('#1a73e8').setFontColor('#fff').setFontWeight('bold');
    sheet.setFrozenRows(1);
    [160,200,160,260,100,100,360,70,70,110,150,320,100,100]
      .forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  }
  return sheet;
}

function getDeleteLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.DELETE_LOG);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.DELETE_LOG);
    const headers = [
      '삭제시간','원본 타임스탬프','잽 이름','오류 단계','오류 메시지',
      'Zap ID','Task ID','원상태','refundId','처리 시간','QA 결과','템플릿 ID','딜 ID'
    ];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setBackground('#7f1d1d').setFontColor('#fff').setFontWeight('bold');
    sheet.setFrozenRows(1);
    [160,160,200,160,260,100,100,70,110,150,320,100,100].forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  }
  return sheet;
}

function getDeleteLog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.DELETE_LOG);
  if (!sheet || sheet.getLastRow() <= 1) return [];
  const last = sheet.getLastRow();
  return sheet.getRange(2, 1, last - 1, 14).getValues()
    .map((r, i) => ({
      id:           i + 2,
      deletedAt:    r[0] ? new Date(r[0]).toLocaleString('ko-KR', { timeZone: 'Asia/Seoul' }) : '',
      timestamp:    r[1] ? new Date(r[1]).toLocaleString('ko-KR', { timeZone: 'Asia/Seoul' }) : '',
      zapName:      r[2] || '',
      stepName:     r[3] || '',
      errorMessage: r[4] || '',
      zapId:        r[5] || '',
      taskId:       r[6] || '',
      status:       r[7] || '',
      refundId:     r[8] || '',
      actionTime:   r[9] || '',
      qaResult:     r[10] || '',
      customerPhone:r[11] || '',
      templateId:   r[12] || '',
      dealId:       r[13] || '',
    }))
    .reverse();
}

function _findRowByTaskId(taskId) {
  if (!taskId) return -1;
  const sheet = getSheet();
  const last = sheet.getLastRow();
  if (last <= 1) return -1;
  // Read zapName (col 2) through taskId (col 6) in one batch
  const data = sheet.getRange(2, 1, last - 1, 6).getValues();
  let firstMatch = -1;
  let realMatch  = -1;
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][5]) !== String(taskId)) continue;
    const rowNum  = i + 2;
    const zapName = String(data[i][1] || '');
    if (firstMatch === -1) firstMatch = rowNum;
    // Prefer rows with a real zap name over "Unknown Zap" ghost rows
    if (realMatch === -1 && zapName && zapName !== 'Unknown Zap') realMatch = rowNum;
  }
  return realMatch > 0 ? realMatch : firstMatch;
}

// "Unknown Zap" 유령 행 정리: 실제 행이 존재하는 중복 taskId의 Unknown 행 삭제
function purgeGhostRows() {
  const sheet = getSheet();
  const last = sheet.getLastRow();
  if (last <= 1) return { removed: 0 };
  const data = sheet.getRange(2, 1, last - 1, 6).getValues();

  // taskId별로 행들을 그룹화
  const byTaskId = {};
  data.forEach((row, i) => {
    const taskId  = String(row[5] || '').trim();
    const zapName = String(row[1] || '');
    if (!taskId) return;
    if (!byTaskId[taskId]) byTaskId[taskId] = [];
    byTaskId[taskId].push({ rowNum: i + 2, zapName });
  });

  const toDelete = [];
  for (const [, rows] of Object.entries(byTaskId)) {
    if (rows.length <= 1) continue;
    const hasReal = rows.some(r => r.zapName && r.zapName !== 'Unknown Zap');
    if (!hasReal) continue;
    rows.filter(r => r.zapName === 'Unknown Zap').forEach(r => toDelete.push(r.rowNum));
  }

  // 아래 행부터 삭제해야 인덱스 안 밀림
  toDelete.sort((a, b) => b - a).forEach(r => sheet.deleteRow(r));
  _logWebhook('GHOST_PURGE', `유령 행 ${toDelete.length}건 삭제`);
  return { removed: toDelete.length };
}

// ZapErrors 시트 전체 초기화 (헤더 1행 유지)
function clearAllErrors() {
  const sheet = getSheet();
  const last = sheet.getLastRow();
  if (last > 1) sheet.deleteRows(2, last - 1);
  _logWebhook('CLEAR_ALL', `전체 초기화 — ${last - 1}행 삭제`);
  return { success: true, deleted: Math.max(0, last - 1) };
}

function maskPhone(phone) {
  if (!phone) return '';
  const digits = phone.replace(/[^0-9]/g, '');
  if (digits.length < 8) return phone;
  return digits.slice(0, 3) + '-****-' + digits.slice(-4);
}

function logToSheet(e) {
  const sheet = getSheet();
  sheet.appendRow([
    e.timestamp, e.zapName, e.stepName, e.errorMessage,
    e.zapId, e.taskId, '분석 중...', e.zapLink || '',
    '신규', e.refundId || '', '', '확인 대기중',
    e.templateId || '', e.dealId || ''  // 14열, 전화번호 미저장 (ISMS)
  ]);
  return sheet.getLastRow();
}

function updateAnalysis(row, analysis) {
  getSheet().getRange(row, COL.analysis).setValue(analysis);
}

function updateQAResult(row, qa) {
  getSheet().getRange(row, COL.qaResult).setValue(qa.summary || '');
}

function purgeRows(rowIds) {
  const sheet = getSheet();
  const logSheet = getDeleteLogSheet();
  const now = new Date();

  // 삭제 전 전체 행 데이터 읽기 + DeleteLog 기록
  const rowData = rowIds.map(id => sheet.getRange(id, 1, 1, 14).getValues()[0]);
  rowData.forEach(r => {
    logSheet.appendRow([
      now, r[0], r[1], r[2], r[3], r[4], r[5], r[8], r[9], r[10], r[11], r[12], r[13]
    ]);
  });

  // taskId 수집 (Zapier 삭제용)
  const taskIds = rowData.map(r => String(r[5] || '').trim()).filter(t => t);

  // Zapier에서 삭제 시도 (세션 있을 때만)
  const { zapsession, csrftoken, accountId } = _getZapierSession();

  if (zapsession && taskIds.length) {
    const gql = `mutation DeleteRuns($accountId: ID!, $filters: RunFilter!, $runIds: [ID!]) {
      deleteRuns(accountId: $accountId, filters: $filters, runIds: $runIds) {
        channel failures globalFailures globalFailureType isMassActionsDisabled pending __typename
      }
    }`;
    const now = new Date();
    const past90 = new Date(now.getTime() - 90 * 24 * 60 * 60 * 1000);
    try {
      const res = UrlFetchApp.fetch('https://zapier.com/api/reporting/graphql', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Cookie': `zapsession=${zapsession}; csrftoken=${csrftoken}`,
          'X-CSRFToken': csrftoken,
          'Referer': 'https://zapier.com/app/history',
          'Origin': 'https://zapier.com',
        },
        payload: JSON.stringify({
          operationName: 'DeleteRuns',
          query: gql,
          variables: {
            accountId: accountId,
            filters: { apps: [], customuserIds: [], folderIds: [], zapIds: [],
              status: ['error'],
              periodEnd: now.toISOString(), periodStart: past90.toISOString() },
            runIds: taskIds
          }
        }),
        muteHttpExceptions: true
      });
      _logWebhook('DELETE_ZAP', `HTTP ${res.getResponseCode()} | ${taskIds.join(',')} | ${res.getContentText().substring(0, 200)}`);
    } catch (err) {
      _logWebhook('DELETE_ZAP_ERR', err.message);
    }
  }

  // 시트에서 행 삭제 (아래부터 삭제해야 인덱스 안 밀림)
  rowIds.sort((a, b) => b - a);
  rowIds.forEach(id => sheet.deleteRow(id));
  return { success: true, deleted: rowIds.length };
}

function updateRowStatus(rowId, status, refundId) {
  const sheet = getSheet();
  sheet.getRange(rowId, COL.status).setValue(status);
  if (refundId) sheet.getRange(rowId, COL.refundId).setValue(refundId);
  sheet.getRange(rowId, COL.actionTime).setValue(
    new Date().toLocaleString('ko-KR', { timeZone: 'Asia/Seoul' })
  );
  return { success: true };
}

// ============================================================
// QA 자동화
// ============================================================

function runQA(error) {
  const hasPhone = !!error.customerPhone;
  const hasDeal  = !!error.dealId;
  const step     = (error.stepName || '').toLowerCase();
  const errMsg   = (error.errorMessage || '').toLowerCase();

  // 에러 단계 타입 판별
  const isGmail     = step.includes('googlemail') || step.includes('gmail');
  const isSolapi    = step.includes('solapi') || step.includes('alimtalk') || step.includes('kakao');
  const isPipedrive = step.includes('pipedrive');
  const isFilter    = step.includes('filter') || errMsg.includes('custom filters blocked');
  const isWebhook   = step.includes('webhook');
  const isSubZap    = step.includes('subzap') || step.includes('sub_zap');

  const lines = [];
  const results = {};

  // ── ① 에러 단계 설명 ─────────────────────────────────────────
  if (isGmail) {
    lines.push('📧 Gmail 발송 실패');
    lines.push('   Gmail 계정 연결 상태 및 발송 필터 조건 확인 필요');
    lines.push('   Zapier → My Apps → Gmail → Reconnect 확인 권장');
  } else if (isFilter) {
    lines.push('🔀 Zapier 필터 조건 미충족');
    lines.push('   해당 딜은 이 단계의 조건에 해당하지 않아 정상 건너뜀');
    lines.push('   실제 오류가 아닐 수 있음 — Zapier 히스토리 직접 확인');
  } else if (isWebhook) {
    lines.push('🌐 웹훅 전송 실패');
    lines.push('   외부 서비스(수신 서버) 응답 오류 또는 연결 문제');
  } else if (isSubZap) {
    lines.push('⚡ 서브잽 실행 실패');
    lines.push('   하위 Zap 로그 별도 확인 필요');
  } else if (isSolapi) {
    lines.push('📱 알림톡(솔라피) 발송 오류');
  } else if (isPipedrive) {
    lines.push('🗂 파이프드라이브 연동 오류');
    const isAuthErr   = errMsg.includes('refreshautherror') || errMsg.includes('auth');
    const isServerErr = errMsg.includes('500') || errMsg.includes('504') || errMsg.includes('hydrating');
    if (isAuthErr)   lines.push('   🔑 인증 만료 — Zapier My Apps → Pipedrive → Reconnect 후 재실행 필요');
    if (isServerErr && !isAuthErr) lines.push('   🌐 Pipedrive 서버 오류 (500/504) — 재실행으로 해결되는 경우가 많음');
  }

  // ── ② 체크: 전화번호 있으면 솔라피 항상, dealId 있으면 파이프드라이브 항상 ──
  // 알림톡 발송 여부 확인
  if (hasPhone) {
    const sr = checkSolapi(error.customerPhone, error.templateId || null, error.timestamp);
    results.solapi = sr;
    const tplHeader = error.templateId ? ` [오류 템플릿: ${error.templateId}]` : '';
    if (sr.checked) {
      if (sr.sent) {
        const t = sr.sentAt ? new Date(sr.sentAt).toLocaleString('ko-KR', { timeZone: 'Asia/Seoul', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit' }) : '';
        const extTag = sr.extended ? ' ⚠️ 재실행/수동처리로 발송 추정' : '';
        lines.push(`\n📱 솔라피${tplHeader}`);
        lines.push(`   ✅ 발송됨 (${t}) | ${_solapiStatus(sr.status)} | ${sr.count||1}건${extTag}`);
        if (sr.templateId) lines.push(`   발송 템플릿: ${sr.templateId}`);
      } else {
        const rangeTag = sr.searchRange ? ` [조회: ${sr.searchRange}]` : '';
        const totalTag = (sr.totalFound > 0) ? ` — 해당 번호 다른 메시지 ${sr.totalFound}건` : '';
        lines.push(`\n📱 솔라피${tplHeader}`);
        lines.push(`   ❌ 미발송${rangeTag}${totalTag}`);
        // 디버그: 조회된 템플릿 ID 목록
        if (sr.foundTemplateIds && sr.foundTemplateIds.length > 0) {
          const uniq = [...new Set(sr.foundTemplateIds)];
          lines.push(`   조회된 템플릿: ${uniq.slice(0,5).join(', ')}${uniq.length > 5 ? ' …' : ''}`);
        }
      }
    } else {
      lines.push(`\n📱 솔라피: ⚠️ 확인 실패 (${sr.reason})`);
    }
  }

  // 파이프드라이브 — dealId 있으면 항상 체크, 없으면 Pipedrive 오류 시 자동 탐색
  {
    let dealId = error.dealId;
    if (!dealId && isPipedrive && error.timestamp) {
      const recent = _findRecentPipedriveDeal(error.zapName, error.timestamp);
      if (recent.length === 1) {
        dealId = String(recent[0].id);
        lines.push(`\n🔍 딜 자동매칭: ${recent[0].title || ''} (ID: ${dealId})`);
      } else if (recent.length > 1) {
        lines.push(`\n🔍 최근 딜 ${recent.length}건 발견 — 확인 버튼에서 딜 ID 직접 입력 권장`);
        lines.push('   ' + recent.slice(0, 3).map(d => `${d.id}: ${(d.title||'').slice(0,20)}`).join(' / '));
      }
    }
    if (dealId) {
      const pr = checkPipedrive(error.zapName, dealId);
      results.pipedrive = pr;
      if (pr.checked) {
        const statusKo = { open: '진행중', won: '✅ 성사', lost: '❌ 실패', deleted: '삭제' }[pr.dealStatus] || pr.dealStatus || '';
        lines.push(`\n🗂 파이프드라이브 [${pr.account}]`);
        lines.push(`   딜: ${pr.dealTitle}`);
        lines.push(`   현재 단계: ${pr.stageName}`);
        lines.push(`   딜 상태: ${statusKo}`);
        lines.push(pr.updated ? '   ✅ 처리 완료' : '   ❌ 미처리 — 재실행 또는 수동 처리 필요');
      } else {
        lines.push(`\n🗂 파이프드라이브: ⚠️ 확인 실패 (${pr.reason})`);
      }
    } else if (isPipedrive && !error.dealId) {
      lines.push('\n⚠️ 딜 ID 없음 — 확인 버튼에서 딜 ID 직접 입력하세요');
    }
  }

  // ── ③ 결제완료 알림톡 — 오류 노드 무관, 전화번호 있으면 항상 체크 ──
  if (hasPhone) {
    const completionTemplates = _getCompletionTemplates();
    if (completionTemplates.length > 0) {
      // Solapi 체크 결과 재사용 (이미 조회한 경우) or 별도 조회
      const sr2 = results.solapi || checkSolapi(error.customerPhone, null, error.timestamp);
      if (sr2.checked) {
        if (sr2.completionResult) {
          const cr = sr2.completionResult;
          const ct = cr.sentAt ? new Date(cr.sentAt).toLocaleString('ko-KR', { timeZone: 'Asia/Seoul', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit' }) : '';
          lines.push(`\n✅ 결제완료 알림톡 발송 확인 (${ct}) | ${_solapiStatus(cr.status)}`);
          results.completionSent = true;
        } else {
          const totalFound = sr2.totalFound || 0;
          const foundTpls  = sr2.foundTemplateIds || [];
          const targetTpls = completionTemplates.join(', ');
          lines.push(`\n❌ 결제완료 알림톡 미발송`);
          lines.push(`   확인 대상 템플릿: ${targetTpls}`);
          lines.push(`   조회 메시지: ${totalFound}건`);
          if (foundTpls.length > 0) {
            const uniq = [...new Set(foundTpls)];
            lines.push(`   조회된 템플릿: ${uniq.slice(0,6).join(', ')}${uniq.length > 6 ? ' …' : ''}`);
          }
        }
      }
    }
  }

  // ── ④ 데이터 없을 때 안내 ────────────────────────────────────
  if (!hasPhone && !error.dealId && !isFilter && !isGmail && !isPipedrive) {
    lines.push('\n⚠️ 확인 버튼으로 전화번호 / 딜 ID 직접 입력하세요');
  }

  // ── ⑤ 판단 ──────────────────────────────────────────────────
  const solapiOk = results.solapi?.sent === true;
  const pdOk     = results.pipedrive?.updated === true;
  const solapiChecked = !!results.solapi?.checked;
  const pdChecked     = !!results.pipedrive?.checked;

  let action = '보류';
  let verdict = '';

  if (results.completionSent) {
    action = '삭제 필요'; verdict = '→ 결제완료 알림톡 발송 확인 (삭제 필요)';
  } else if (isFilter) {
    action = '삭제 필요'; verdict = '→ 필터 정상 동작 (삭제 필요)';
  } else if (solapiChecked && pdChecked) {
    if (solapiOk && pdOk)        { action = '삭제 필요';   verdict = '→ 처리 완료 (삭제 필요)'; }
    else if (!solapiOk && !pdOk) { action = '재실행 필요'; verdict = '→ 미처리 (전체 재실행 권장)'; }
    else                         { action = '재실행 필요'; verdict = '→ 부분 처리 (에러 단계부터 재실행 권장)'; }
  } else if (pdChecked) {
    action = pdOk ? '삭제 필요' : '재실행 필요';
    verdict = pdOk ? '→ Pipedrive 처리 완료 (삭제 필요)' : '→ Pipedrive 미처리 (재실행 필요)';
  } else if (solapiChecked) {
    action = solapiOk ? '삭제 필요' : '재실행 필요';
    verdict = solapiOk ? '→ 발송 확인 (삭제 필요)' : '→ 미발송 (재실행 필요)';
  } else if (isGmail) {
    action = '보류'; verdict = '→ Gmail 수동 확인 필요 (보류)';
  } else if (isPipedrive && !pdChecked) {
    // dealId 없어도 에러 패턴으로 판단 가능한 Pipedrive 오류
    const isAuthErr   = errMsg.includes('refreshautherror') || errMsg.includes('auth');
    const isServerErr = errMsg.includes('500') || errMsg.includes('504') || errMsg.includes('hydrating');
    if (isAuthErr || isServerErr) {
      action = '재실행 필요';
      verdict = isAuthErr
        ? '→ Pipedrive 인증 만료 — Reconnect 후 재실행 필요'
        : '→ Pipedrive 서버 오류 — 재실행 필요';
    }
  }

  if (verdict) lines.push(`\n⚡ 판단: ${verdict}`);

  const isPdPatternChecked = isPipedrive && !pdChecked && !!verdict;
  const hasAnyCheck = solapiChecked || pdChecked || isFilter || isPdPatternChecked;
  return {
    checked: hasAnyCheck,
    action,
    summary: lines.join('\n'),
  };
}

function _solapiStatus(code) {
  const map = { '2000': '정상수신', '3000': '미수신', '4000': '발송완료', '5000': '전송실패', '7000': '수신불가' };
  return map[String(code)] || code || '';
}

// 시간 기반 Pipedrive 최근 딜 자동 검색 (±15분)
function _findRecentPipedriveDeal(zapName, timestamp) {
  const pd = getPipedriveAccount(zapName);
  if (!pd.token || !pd.domain) return [];
  try {
    const ts = new Date(timestamp);
    const since = new Date(ts.getTime() - 15 * 60 * 1000).toISOString().replace('T', ' ').slice(0, 19);
    const cleanDomain = pd.domain.replace(/^https?:\/\//, '').replace(/\/+$/, '');
    const url = `https://${cleanDomain}/api/v1/deals?updated_since=${encodeURIComponent(since)}&sort=update_time+DESC&limit=5&api_token=${pd.token}`;
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return [];
    return JSON.parse(res.getContentText()).data || [];
  } catch(e) { return []; }
}

// ============================================================
// Solapi API
// ============================================================

function getSolapiAuthHeader() {
  const p = PropertiesService.getScriptProperties();
  const apiKey    = p.getProperty('SOLAPI_API_KEY');
  const apiSecret = p.getProperty('SOLAPI_API_SECRET');
  if (!apiKey || !apiSecret) return null;

  const date = new Date().toISOString();
  const salt = Utilities.getUuid().replace(/-/g, '');
  const sigBytes = Utilities.computeHmacSha256Signature(date + salt, apiSecret);
  const signature = sigBytes.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
  return `HMAC-SHA256 apiKey=${apiKey}, date=${date}, salt=${salt}, signature=${signature}`;
}

// 결제완료 등 "발송 확인되면 삭제 가능" 템플릿 목록 (Script Properties: COMPLETION_TEMPLATES)
function _getCompletionTemplates() {
  const raw = PropertiesService.getScriptProperties().getProperty('COMPLETION_TEMPLATES') || '';
  return raw.split(',').map(s => s.trim()).filter(Boolean);
}

function checkSolapi(phone, templateId, errorTimestamp) {
  const auth = getSolapiAuthHeader();
  if (!auth) return { checked: false, reason: 'Solapi 키 미설정' };

  let cleanPhone = phone.replace(/[^0-9]/g, '');
  // 10자리이고 0으로 시작 안 하면 앞에 0 추가 (구글시트 숫자 저장 시 소실된 경우)
  if (cleanPhone.length === 10 && !cleanPhone.startsWith('0')) cleanPhone = '0' + cleanPhone;
  // 국제번호 형식 821XXXXXXXXX → 01XXXXXXXXX
  if (cleanPhone.length === 12 && cleanPhone.startsWith('82')) cleanPhone = '0' + cleanPhone.slice(2);
  if (cleanPhone.length === 11 && cleanPhone.startsWith('820')) cleanPhone = '0' + cleanPhone.slice(2);
  const completionTemplates = _getCompletionTemplates();

  // 국내형 01X → 국제형 82X 변환
  const intlPhone = cleanPhone.startsWith('0') ? '82' + cleanPhone.slice(1) : cleanPhone;

  function _query(startIso, endIso, phoneNum) {
    const qs = `limit=100&to=${phoneNum}&dateCreatedFrom=${encodeURIComponent(startIso)}&dateCreatedTo=${encodeURIComponent(endIso)}`;
    const res = UrlFetchApp.fetch(`https://api.solapi.com/messages/v4/list?${qs}`, {
      headers: { 'Authorization': auth }, muteHttpExceptions: true
    });
    const code = res.getResponseCode();
    const body = res.getContentText();
    _logWebhook('SOLAPI_DEBUG', `HTTP ${code} | to=${phoneNum} | ${body.slice(0, 300)}`);
    const json = JSON.parse(body);
    const raw = json.messageList || json.fileList || json.messages || json.data;
    if (!raw) return [];
    return Array.isArray(raw) ? raw : Object.values(raw);
  }

  function _getMsgTemplateId(m) {
    return String(m.kakaoOptions?.templateId || m.kakaoOptions?.templateCode
      || m.templateId || m.templateCode || '');
  }

  function _pickBest(messages) {
    if (!templateId) return messages;
    const matched = messages.filter(m => _getMsgTemplateId(m) === String(templateId));
    return matched.length > 0 ? matched : messages;
  }

  try {
    const errorDate = new Date(errorTimestamp);

    // 오류 전후 각 30일 조회 (총 60일 창)
    const start = new Date(errorDate.getTime() - 30 * 86400000).toISOString();
    const end   = new Date(Math.min(errorDate.getTime() + 30 * 86400000, Date.now())).toISOString();
    const range = `${_fmtKST(start)} ~ ${_fmtKST(end)}`;
    // 국내형(01X)으로 먼저 조회, 없으면 국제형(82X)으로 재조회
    let msgs = _query(start, end, cleanPhone);
    if (msgs.length === 0 && intlPhone !== cleanPhone) {
      msgs = _query(start, end, intlPhone);
    }

    // 완료 템플릿 발송 여부 별도 확인 (에러 템플릿과 무관하게 체크)
    let completionResult = null;
    const foundTemplateIds = msgs.map(m => _getMsgTemplateId(m)).filter(Boolean);
    if (completionTemplates.length && msgs.length > 0) {
      const cm = msgs.find(m => completionTemplates.includes(_getMsgTemplateId(m)));
      if (cm) {
        const sentAt = cm.dateCreated || cm.registeredDate || '';
        completionResult = {
          templateId: _getMsgTemplateId(cm),
          sentAt,
          status: cm.statusCode || cm.status || '',
        };
      }
    }

    if (msgs.length > 0) {
      const picked = _pickBest(msgs);
      if (picked.length > 0) {
        const msg = picked[0];
        const sentAt = msg.dateCreated || msg.registeredDate || '';
        const extended = sentAt && new Date(sentAt) > new Date(errorDate.getTime() + 2 * 3600000);
        return {
          checked: true, sent: true,
          sentAt, status: msg.statusCode || msg.status || '',
          messageId: msg.messageId || '',
          templateId: _getMsgTemplateId(msg),
          count: picked.length,
          searchRange: range,
          extended,
          completionResult,
          foundTemplateIds,
        };
      }
    }

    return { checked: true, sent: false, searchRange: range, totalFound: msgs.length, completionResult, foundTemplateIds };

  } catch (err) {
    return { checked: false, reason: err.message };
  }
}

function _fmtKST(iso) {
  return new Date(iso).toLocaleString('ko-KR', { timeZone: 'Asia/Seoul', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit' });
}

// ============================================================
// Pipedrive API
// ============================================================

function checkPipedrive(zapName, dealId, context) {
  const pd = getPipedriveAccount(zapName);
  if (!pd.token || !pd.domain) return { checked: false, reason: '파이프드라이브 키 미설정' };

  try {
    const cleanDomain = pd.domain.replace(/^https?:\/\//, '').replace(/\/+$/, '');
    const dealUrl  = `https://${cleanDomain}/api/v1/deals/${dealId}?api_token=${pd.token}`;

    // 딜 조회 (단건, 스테이지는 딜 응답의 stage_id로 후속 요청)
    const dealRes = UrlFetchApp.fetch(dealUrl, { muteHttpExceptions: true });
    const json = JSON.parse(dealRes.getContentText());

    if (!json.success || !json.data) {
      return { checked: false, reason: `API 오류: ${json.error || '딜 없음'}` };
    }

    const deal = json.data;

    // 딜 + 스테이지 병렬 조회 (fetchAll)
    let stageName = '';
    if (deal.stage_id) {
      try {
        const stageUrl = `https://${cleanDomain}/api/v1/stages/${deal.stage_id}?api_token=${pd.token}`;
        const [stageRes] = UrlFetchApp.fetchAll([{ url: stageUrl, muteHttpExceptions: true }]);
        stageName = JSON.parse(stageRes.getContentText()).data?.name || '';
      } catch(_) {}
    }

    const verdict = judgePipedriveByAI(zapName, deal);

    return {
      checked:    true,
      updated:    verdict.updated,
      value:      verdict.value,
      reason:     verdict.reason,
      account:    pd.name,
      dealTitle:  deal.title  || '',
      dealStatus: deal.status || '',
      stageName,
      stage:      deal.stage_id || '',
    };

  } catch (err) {
    return { checked: false, reason: err.message };
  }
}

function judgePipedriveByAI(zapName, deal) {
  if (!CONFIG.GROQ_API_KEY) {
    // 기본 판단: won(성사) 상태이거나 최근 7일 이내 업데이트면 처리완료
    const won = deal.status === 'won';
    const updateTime = new Date(deal.update_time);
    const sevenDaysAgo = new Date(Date.now() - 7 * 86400000);
    const updated = won || updateTime >= sevenDaysAgo;
    const reason = won ? '딜 상태: 성사' : (updated ? '최근 7일 이내 업데이트' : '딜 미업데이트');
    return { updated, value: deal.update_time, reason };
  }

  try {
    // 딜에서 주요 커스텀 필드와 업데이트 시간만 추출
    const summary = {
      title:       deal.title,
      stage_id:    deal.stage_id,
      update_time: deal.update_time,
      status:      deal.status,
    };
    // 커스텀 필드 중 null 아닌 것만 포함
    Object.entries(deal).forEach(([k, v]) => {
      if (k.length === 40 && v !== null && v !== '') summary[k] = v; // Pipedrive 커스텀 필드는 해시값
    });

    const prompt = `Zapier 자동화 잽 이름: "${zapName}"
이 잽은 파이프드라이브 딜을 업데이트하는 작업입니다.
아래 딜 데이터를 보고 이 잽이 의도한 작업(필드 업데이트 또는 스테이지 변경)이 이미 완료되었는지 판단하세요.

딜 데이터:
${JSON.stringify(summary, null, 2)}

다음 JSON 형식으로만 답하세요:
{"updated": true/false, "value": "근거값", "reason": "판단 이유 1문장"}`;

    const res = UrlFetchApp.fetch('https://api.groq.com/openai/v1/chat/completions', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + CONFIG.GROQ_API_KEY },
      payload: JSON.stringify({
        model: 'llama-3.3-70b-versatile',
        messages: [{ role: 'user', content: prompt }],
        max_tokens: 150, temperature: 0.1
      }),
      muteHttpExceptions: true
    });

    const content = JSON.parse(res.getContentText()).choices?.[0]?.message?.content || '{}';
    const cleaned = content.replace(/```json|```/g, '').trim();
    return JSON.parse(cleaned);

  } catch (err) {
    return { updated: false, value: '', reason: `AI 판단 실패: ${err.message}` };
  }
}

// ============================================================
// Zapier 재실행
// ============================================================

function bulkUpdateStatus(rowIds, status) {
  const sheet = getSheet();
  const now = new Date().toLocaleString('ko-KR', { timeZone: 'Asia/Seoul' });
  rowIds.forEach(rowId => {
    try {
      sheet.getRange(rowId, COL.status).setValue(status);
      sheet.getRange(rowId, COL.actionTime).setValue(now);
    } catch(_) {}
  });
  return { success: true };
}

function bulkReplayZapRuns(rowIds, replayType) {
  const sheet = getSheet();
  const { zapsession, csrftoken, accountId } = _getZapierSession();
  if (!zapsession) return { success: false, successCount: 0, failedCount: rowIds.length, error: '세션 없음' };

  // rowId → taskId 매핑
  const items = rowIds.map(rowId => {
    const row = sheet.getRange(rowId, 1, 1, 14).getValues()[0];
    return { rowId, taskId: String(row[5] || '').trim() };
  }).filter(it => it.taskId);

  const noTask = rowIds.length - items.length;
  const taskIds = items.map(it => it.taskId);

  const headers = {
    'Content-Type': 'application/json',
    'Cookie': `zapsession=${zapsession}; csrftoken=${csrftoken}`,
    'X-CSRFToken': csrftoken,
    'Referer': 'https://zapier.com/app/history',
    'Origin': 'https://zapier.com',
  };

  let ok = false;
  try {
    if (replayType === 'from_error') {
      // REST bulk API — 한 번에 전체 전송
      const res = UrlFetchApp.fetch('https://zapier.com/api/zap-history/v2/runs/bulk-replay', {
        method: 'POST', headers,
        payload: JSON.stringify({ account_id: Number(accountId), zap_run_ids: taskIds }),
        muteHttpExceptions: true
      });
      const code = res.getResponseCode();
      ok = code >= 200 && code < 300;
      _logWebhook(ok ? 'BULK_REPLAY' : 'BULK_REPLAY_ERR',
        `HTTP ${code} | ${taskIds.length}건 | ${ok ? '' : res.getContentText().slice(0, 200)}`);
    } else {
      // GQL — runIds 배열로 한 번에 전송
      const gql = `mutation ReplayRuns($accountId: ID!, $filters: RunFilter!, $runIds: [ID!]) {
        replayRuns(accountId: $accountId, filters: $filters, runIds: $runIds) {
          channel failures globalFailures globalFailureType isMassActionsDisabled pending __typename
        }
      }`;
      const now = new Date();
      const past90 = new Date(now.getTime() - 90 * 24 * 60 * 60 * 1000);
      const res = UrlFetchApp.fetch('https://zapier.com/api/reporting/graphql', {
        method: 'POST', headers: { ...headers, 'X-Requested-With': 'XMLHttpRequest' },
        payload: JSON.stringify({
          operationName: 'ReplayRuns', query: gql,
          variables: {
            accountId,
            filters: { apps: [], customuserIds: [], folderIds: [],
              periodEnd: now.toISOString(), periodStart: past90.toISOString() },
            runIds: taskIds
          }
        }),
        muteHttpExceptions: true
      });
      const code = res.getResponseCode();
      ok = code >= 200 && code < 300;
      if (ok) {
        try {
          const body = JSON.parse(res.getContentText());
          const f = body?.data?.replayRuns?.failures || [];
          const gf = body?.data?.replayRuns?.globalFailures || [];
          if (f.length || gf.length || body?.errors?.length) ok = false;
        } catch(_) {}
      }
      _logWebhook(ok ? 'BULK_REPLAY' : 'BULK_REPLAY_ERR',
        `HTTP ${code} | GQL | ${taskIds.length}건`);
    }
  } catch(e) {
    _logWebhook('BULK_REPLAY_ERR', e.message);
    return { success: false, successCount: 0, failedCount: rowIds.length, error: e.message };
  }

  // 시트 일괄 업데이트
  const newStatus = ok ? '재실행 완료' : '재실행 실패';
  const now = new Date().toLocaleString('ko-KR', { timeZone: 'Asia/Seoul' });
  items.forEach(({ rowId }) => {
    sheet.getRange(rowId, COL.status).setValue(newStatus);
    sheet.getRange(rowId, COL.actionTime).setValue(now);
  });

  const successCount = ok ? items.length : 0;
  const failedCount  = (ok ? 0 : items.length) + noTask;
  return { success: true, successCount, failedCount };
}

function replayZapRun(rowId, replayType) {
  const sheet = getSheet();
  const row = sheet.getRange(rowId, 1, 1, 14).getValues()[0];
  const taskId = String(row[5] || '').trim(); // Zap Run ID

  if (!taskId) return { success: false, error: 'Task ID(Zap Run ID)가 없습니다. 웹훅 매핑에서 task_id를 확인하세요.' };

  const { zapsession, csrftoken, accountId } = _getZapierSession();

  if (!zapsession) return { success: false, error: '세션 없음. Chrome 확장프로그램에서 동기화 필요.' };

  // 에러 스텝부터 재실행 → REST API
  if (replayType === 'from_error') {
    try {
      const res = UrlFetchApp.fetch('https://zapier.com/api/zap-history/v2/runs/bulk-replay', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Cookie': `zapsession=${zapsession}; csrftoken=${csrftoken}`,
          'X-CSRFToken': csrftoken,
          'Referer': 'https://zapier.com/app/history',
          'Origin': 'https://zapier.com',
        },
        payload: JSON.stringify({
          account_id: Number(accountId),
          zap_run_ids: [taskId]
        }),
        muteHttpExceptions: true
      });

      const code = res.getResponseCode();
      const bodyText = res.getContentText();
      const ok = code >= 200 && code < 300;
      const newStatus = ok ? '재실행 완료' : '재실행 실패';
      sheet.getRange(rowId, COL.status).setValue(newStatus);
      sheet.getRange(rowId, COL.actionTime).setValue(
        new Date().toLocaleString('ko-KR', { timeZone: 'Asia/Seoul' })
      );
      _logWebhook(ok ? 'REPLAY' : 'REPLAY_ERR',
        `HTTP ${code} | taskId:${taskId}${ok ? '' : ' | ' + bodyText.substring(0, 200)}`);
      return { success: ok, type: 'api_from_error', httpCode: code, status: newStatus };
    } catch (err) {
      return { success: false, error: err.message };
    }
  }

  // 전체 잽 재실행 → Zapier GraphQL (세션 쿠키 기반)

  const gql = `mutation ReplayRuns($accountId: ID!, $filters: RunFilter!, $runIds: [ID!]) {
    replayRuns(accountId: $accountId, filters: $filters, runIds: $runIds) {
      channel failures globalFailures globalFailureType isMassActionsDisabled pending __typename
    }
  }`;

  const now = new Date();
  const past90 = new Date(now.getTime() - 90 * 24 * 60 * 60 * 1000);

  try {
    const res = UrlFetchApp.fetch('https://zapier.com/api/reporting/graphql', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Cookie': `zapsession=${zapsession}; csrftoken=${csrftoken}`,
        'X-CSRFToken': csrftoken,
        'Referer': 'https://zapier.com/app/history',
        'Origin': 'https://zapier.com',
        'X-Requested-With': 'XMLHttpRequest'
      },
      payload: JSON.stringify({
        operationName: 'ReplayRuns',
        query: gql,
        variables: {
          accountId: accountId,
          filters: { apps: [], customuserIds: [], folderIds: [],
            periodEnd: now.toISOString(), periodStart: past90.toISOString() },
          runIds: [taskId]
        }
      }),
      muteHttpExceptions: true
    });

    const code = res.getResponseCode();
    const bodyText = res.getContentText();
    let ok = code >= 200 && code < 300;
    if (ok) {
      try {
        const body = JSON.parse(bodyText);
        const failures       = body?.data?.replayRuns?.failures       || [];
        const globalFailures = body?.data?.replayRuns?.globalFailures  || [];
        if (failures.length > 0 || globalFailures.length > 0 || body?.errors?.length > 0) ok = false;
      } catch(_) {}
    }
    const newStatus = ok ? '재실행 완료' : '재실행 실패';
    sheet.getRange(rowId, COL.status).setValue(newStatus);
    sheet.getRange(rowId, COL.actionTime).setValue(
      new Date().toLocaleString('ko-KR', { timeZone: 'Asia/Seoul' })
    );
    _logWebhook(ok ? 'REPLAY' : 'REPLAY_ERR',
      `HTTP ${code} | taskId:${taskId}${ok ? '' : ' | ' + bodyText.substring(0, 200)}`);
    return { success: ok, type: 'api', httpCode: code, status: newStatus };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// ============================================================
// QA 실행 (수동 / 자동 공통)
// ============================================================

function _buildErrorFromRow(row) {
  return {
    timestamp:     row[0] ? new Date(row[0]).toISOString() : new Date().toISOString(),
    zapName:       row[1] || '',
    stepName:      row[2] || '',
    errorMessage:  row[3] || '',
    templateId:    row[12] || '',
    dealId:        row[13] || '',
  };
}

// 수동 QA (대시보드 버튼)
function runQAForRow(rowId, manualPhone, manualDealId, manualTemplateId) {
  manualPhone      = manualPhone      || '';
  manualDealId     = manualDealId     || '';
  manualTemplateId = manualTemplateId || '';
  const sheet = getSheet();
  const row = sheet.getRange(rowId, 1, 1, 14).getValues()[0];
  const error = _buildErrorFromRow(row);

  if (manualPhone)      { error.customerPhone = manualPhone; } // 시트 미저장 (ISMS)
  if (manualDealId)     { error.dealId = manualDealId; sheet.getRange(rowId, COL.dealId).setValue(manualDealId); }
  if (manualTemplateId) { error.templateId = manualTemplateId; sheet.getRange(rowId, COL.templateId).setValue(manualTemplateId); }

  const qa = runQA(error);
  // 수동 QA: 결과 앞에 👤 마킹
  qa.summary = '👤 수동 QA\n' + (qa.summary || '');
  updateQAResult(rowId, qa);
  if (qa.action && ['재실행 필요', '삭제 필요', '보류'].includes(qa.action)) {
    updateRowStatus(rowId, qa.action, '');
  }
  return { success: true, qa };
}

// 자동 QA (크롬 익스텐션 → step 데이터 확보 후 호출)
function autoQAByTaskId(taskId, phone) {
  const rowId = _findRowByTaskId(taskId);
  if (rowId < 2) return { success: false, reason: 'taskId not found' };

  const sheet = getSheet();
  const row = sheet.getRange(rowId, 1, 1, 14).getValues()[0];
  const error = _buildErrorFromRow(row);

  // 전화번호는 파라미터로만 — 시트에는 없음 (ISMS)
  if (phone) error.customerPhone = phone;

  // 필터/Gmail 오류는 phone/dealId 없이도 판단 가능; 나머지는 익스텐션 데이터 필요
  const isFilterOrGmail = /custom filter|blocked this zap/i.test(error.errorMessage) || /gmail/i.test(error.stepName);
  if (!error.customerPhone && !error.dealId && !isFilterOrGmail) {
    // 스피너가 계속 뜨지 않도록 terminal 상태로 종료
    getSheet().getRange(rowId, COL.qaResult).setValue('수동 확인 필요 (데이터 없음)');
    return { success: false, reason: '데이터 없음' };
  }

  const qa = runQA(error);
  qa.summary = '🤖 자동 QA\n' + (qa.summary || '');
  updateQAResult(rowId, qa);
  if (qa.action && ['재실행 필요', '삭제 필요', '보류'].includes(qa.action)) {
    updateRowStatus(rowId, qa.action, '');
  }
  _logWebhook('AUTO_QA', `rowId=${rowId} checked=${qa.checked} action=${qa.action}`);
  return { success: true, qa };
}

// ============================================================
// Groq AI 분석
// ============================================================

function analyzeWithGroq(e) {
  if (!CONFIG.GROQ_API_KEY) return '(API 키 미설정)';

  const prompt = `Zapier 자동화에서 오류가 발생했습니다. 한국어로 분석해주세요.

잽 이름: ${e.zapName}
오류 단계: ${e.stepName}
오류 메시지: ${e.errorMessage}

다음 형식으로만 답하세요:
🔍 원인: (1-2문장)
✅ 해결책: (번호 목록 2-3가지)
⚠️ 긴급도: 낮음 / 보통 / 높음 중 하나`;

  try {
    const res = UrlFetchApp.fetch('https://api.groq.com/openai/v1/chat/completions', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + CONFIG.GROQ_API_KEY },
      payload: JSON.stringify({
        model: 'llama-3.3-70b-versatile',
        messages: [{ role: 'user', content: prompt }],
        max_tokens: 500, temperature: 0.2
      }),
      muteHttpExceptions: true
    });
    const json = JSON.parse(res.getContentText());
    return json.choices?.[0]?.message?.content || '분석 오류: Groq 응답 없음';
  } catch (err) {
    return `분석 오류: ${err.message}`;
  }
}

// ============================================================
// Slack
// ============================================================

function sendSlack(e) {
  if (!CONFIG.SLACK_WEBHOOK_URL) return;
  const qa = e.qa || {};
  const qaText = qa.summary ? `*✅ 처리 확인*\n${qa.summary}` : '';
  const actionEmoji = { '재실행 필요': '🔴', '재실행 완료': '🟢', '삭제 필요': '✅', '삭제 완료': '✅', '보류': '🟡' };

  UrlFetchApp.fetch(CONFIG.SLACK_WEBHOOK_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    payload: JSON.stringify({
      text: `${actionEmoji[qa.action] || '⚡'} Zapier 오류: *${e.zapName}* → ${qa.action || '신규'}`,
      blocks: [
        { type: 'header', text: { type: 'plain_text', text: '⚡ Zapier 오류 감지', emoji: true } },
        { type: 'section', fields: [
          { type: 'mrkdwn', text: `*잽 이름*\n${e.zapName}` },
          { type: 'mrkdwn', text: `*오류 단계*\n${e.stepName}` }
        ]},
        { type: 'section', text: { type: 'mrkdwn', text: `*오류 메시지*\n\`\`\`${e.errorMessage.substring(0, 300)}\`\`\`` } },
        ...(qaText ? [{ type: 'section', text: { type: 'mrkdwn', text: qaText } }] : []),
        { type: 'section', text: { type: 'mrkdwn', text: `*🔍 원인 분석*\n${e.analysis}` } },
        ...(e.zapLink ? [{ type: 'section', text: { type: 'mrkdwn', text: `<${e.zapLink}|🔗 Zapier에서 재실행>` } }] : []),
        { type: 'context', elements: [{ type: 'mrkdwn', text: `🕐 ${e.timestamp}` }] }
      ]
    }),
    muteHttpExceptions: true
  });
}

// ============================================================
// 대시보드 데이터
// ============================================================

function getErrorData() {
  const sheet = getSheet();
  const last = sheet.getLastRow();
  if (last <= 1) return [];
  return sheet.getRange(2, 1, last - 1, 14).getValues()
    .map((r, i) => {
      const rawDate = r[0] ? new Date(r[0]) : null;
      return {
        id:           i + 2,
        timestamp:    rawDate && !isNaN(rawDate) ? rawDate.toISOString() : '',
        zapName:      r[1] || '', stepName: r[2] || '',
        errorMessage: r[3] || '', zapId: r[4] || '', taskId: r[5] || '',
        analysis:     r[6] || '',
        zapLink: (r[4] && r[5])
          ? `https://zapier.com/editor/${r[4]}/run/${r[5]}?sidebar=runs&filterRun=true`
          : (r[7] || ''),
        status:       r[8] || '신규', refundId: r[9] || '', actionTime: r[10] || '',
        qaResult:     r[11] || '',
        templateId:   r[12] || '', dealId: r[13] || '',
      };
    })
    .reverse();
}

// 크롬 익스텐션: step 데이터 없는 오류 목록 반환 (최근 7일, 최대 20건)
function getErrorsMissingData() {
  const sheet = getSheet();
  const last = sheet.getLastRow();
  if (last <= 1) return [];
  const rows = sheet.getRange(2, 1, last - 1, 14).getValues();
  const cutoff = Date.now() - 7 * 86400000;
  const result = [];
  rows.forEach((r, i) => {
    const ts = r[0] ? new Date(r[0]).getTime() : 0;
    if (ts < cutoff) return;
    const taskId = String(r[5] || '').trim();
    const tplId  = String(r[12] || '').trim(); // templateId (M열)
    if (!taskId) return;
    if (tplId) return; // templateId 있으면 이미 enrich됨
    result.push({ rowId: i + 2, taskId, zapId: String(r[4] || '') });
    if (result.length >= 20) return;
  });
  return result;
}

// 크롬 익스텐션: step 데이터 업데이트
function updateErrorStepData(rowId, _phone, templateId, dealId) {
  // _phone 파라미터는 ISMS 대응으로 시트에 저장하지 않음 — QA 시 파라미터로만 사용
  const sheet = getSheet();
  const row = Number(rowId);
  if (!row || row < 2) return { success: false, reason: 'invalid rowId' };
  const existing = sheet.getRange(row, COL.templateId, 1, 2).getValues()[0];
  if (templateId && !existing[0]) sheet.getRange(row, COL.templateId).setValue(templateId);
  if (dealId     && !existing[1]) sheet.getRange(row, COL.dealId).setValue(dealId);
  return { success: true };
}

// taskId로 step 데이터 업데이트 (크롬 익스텐션 push용)
function updateErrorStepDataByTaskId(taskId, phone, templateId, dealId) {
  const rowId = _findRowByTaskId(taskId);
  if (!rowId) return { success: false, reason: 'taskId not found' };
  return updateErrorStepData(rowId, phone, templateId, dealId);
}

function getStats() {
  const sheet = getSheet();
  const last = sheet.getLastRow();
  if (last <= 1) return { total: 0, today: 0, uniqueZaps: 0, topZap: '-', zapNames: [] };

  const data = sheet.getRange(2, 1, last - 1, 2).getValues();
  const todayStart = new Date(); todayStart.setHours(0, 0, 0, 0);
  const zapCounts = {}; let todayCount = 0;

  data.forEach(r => {
    if (new Date(r[0]) >= todayStart) todayCount++;
    const name = r[1] || 'Unknown';
    zapCounts[name] = (zapCounts[name] || 0) + 1;
  });

  const topEntry = Object.entries(zapCounts).sort((a, b) => b[1] - a[1])[0];
  const delLog = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.DELETE_LOG);
  const purgedCount = delLog ? Math.max(0, delLog.getLastRow() - 1) : 0;
  return {
    total: data.length, today: todayCount,
    uniqueZaps: Object.keys(zapCounts).length,
    topZap: topEntry ? `${topEntry[0]} (${topEntry[1]}건)` : '-',
    zapNames: Object.keys(zapCounts).sort(),
    purgedCount
  };
}

// ============================================================
// 웹훅 로그 (수신 진단용)
// ============================================================

function _logWebhook(type, content) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let log = ss.getSheetByName(CONFIG.LOG_SHEET);
    if (!log) {
      log = ss.insertSheet(CONFIG.LOG_SHEET);
      log.appendRow(['시간', '유형', '내용']);
      log.setFrozenRows(1);
      log.setColumnWidth(1, 150); log.setColumnWidth(2, 70); log.setColumnWidth(3, 600);
    }
    log.appendRow([new Date(), type, content]);
    const last = log.getLastRow();
    if (last > 201) log.deleteRows(2, last - 201); // 최근 200건만 유지
  } catch (_) {}
}

// ── 시트 컬럼 마이그레이션 (한 번만 실행) ─────────────────────────
// GAS 편집기에서 직접 실행: 기존 M열(고객 전화) 삭제 + 헤더 업데이트
function fixSheetColumns() {
  const sheet = getSheet();
  const mainHeaders = [
    '타임스탬프','잽 이름','오류 단계','오류 메시지','Zap ID','Task ID',
    '원인 분석','Zap 링크','상태','refundId','처리 시간',
    '처리 확인','템플릿 ID','딜 ID'
  ];
  // M열(13번) 헤더가 '고객 전화'이면 삭제
  const colMHeader = sheet.getRange(1, 13).getValue();
  if (colMHeader === '고객 전화') {
    sheet.deleteColumn(13);
    Logger.log('메인 시트: M열(고객 전화) 삭제 완료');
  }
  // 헤더 전체 갱신
  sheet.getRange(1, 1, 1, mainHeaders.length)
    .setValues([mainHeaders])
    .setBackground('#1a73e8').setFontColor('#fff').setFontWeight('bold');
  [160,200,160,260,100,100,360,70,70,110,150,320,100,100]
    .forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  // DeleteLog 시트도 동일하게 처리
  const delSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.DELETE_LOG);
  if (delSheet) {
    const delHeaders = [
      '삭제시간','원본 타임스탬프','잽 이름','오류 단계','오류 메시지',
      'Zap ID','Task ID','원상태','refundId','처리 시간','QA 결과','템플릿 ID','딜 ID'
    ];
    const delColHeader = delSheet.getRange(1, 12).getValue();
    if (delColHeader === '고객 전화') {
      delSheet.deleteColumn(12);
      Logger.log('DeleteLog 시트: 고객 전화 열 삭제 완료');
    }
    delSheet.getRange(1, 1, 1, delHeaders.length)
      .setValues([delHeaders])
      .setBackground('#7f1d1d').setFontColor('#fff').setFontWeight('bold');
  }
  Logger.log('fixSheetColumns 완료 — 이제 H~N열까지 헤더가 표시됩니다');
}

function getWebhookLog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = ss.getSheetByName(CONFIG.LOG_SHEET);
  if (!log || log.getLastRow() <= 1) return [];
  const last = log.getLastRow();
  const start = Math.max(2, last - 29);
  return log.getRange(start, 1, last - start + 1, 3).getValues()
    .reverse()
    .map(r => ({
      time:    r[0] ? new Date(r[0]).toLocaleString('ko-KR', { timeZone: 'Asia/Seoul' }) : '',
      type:    r[1] || '',
      content: String(r[2] || '').substring(0, 200),
    }));
}

function getSystemStatus() {
  const log = getWebhookLog();
  const last = log.find(r => r.type === 'POST');
  const errors = log.filter(r => r.type === 'ERROR');
  return {
    lastReceived: last ? last.time : '없음',
    totalLogs: log.length,
    recentErrors: errors.slice(0, 3).map(r => ({ time: r.time, msg: r.content })),
  };
}

// ============================================================
// Zapier GraphQL 폴링 (Zapier Manager 버그 우회)
// ============================================================

const RUN_DETAIL_GQL = `query RunDetail($runId: ID!) {
  zapRun(id: $runId) {
    id isRunning startTime status zapVersionName zapVersionId
    hasPaths heldReason heldStaleAuth canEdit canReplay triggerType
    billableCount billingAlignedBillableCount billingAlignmentEnabled billingAlignmentFailed
    steps {
      action app childrenHavePaths dataTruncated description
      error { title __typename }
      frameId id input invocationId meta order output params parentId
      rescheduledTime startTime status title type __typename
    }
    timezone
    zap { id title sourceUrl kind __typename }
    __typename
  }
}`;

// fetchRunDetail: GAS 서버(Google IP)에서는 Zapier zapRun 상세 API 접근 불가
// - GraphQL zapRun → 401 (IP 기반 세션 검증)
// - REST /api/v4/* → 404 (UUID 형식 미지원)
// 향후 Chrome Extension 릴레이 방식으로 대체 가능
function fetchRunDetail(_taskId, _zapsession, _csrftoken) {
  return null;
}

const POLL_GQL = `fragment ZapRun on ZapRun {
  id startTime status
  steps { error { title __typename } status title input output __typename }
  zap { id title sourceUrl __typename }
  __typename
}
query ZapRuns($accountId: ID!, $status: [String!], $limit: Int, $offset: Int, $periodStart: String, $periodEnd: String, $sortBy: String) {
  zapRuns(accountId: $accountId, status: $status, limit: $limit, offset: $offset, periodStart: $periodStart, periodEnd: $periodEnd, sortBy: $sortBy, apps: [], customuserIds: [], folderIds: [], zapIds: []) {
    edges { ...ZapRun __typename }
    totalCount
    __typename
  }
}`;

// ============================================================
// 스텝 input 데이터에서 고객 정보 추출
// ============================================================

function _tryParseJson(v) {
  if (!v) return null;
  if (typeof v === 'object') return v;
  try { return JSON.parse(v); } catch (_) { return null; }
}

// 중첩 객체 평탄화 (depth 3 이하)
function _flatten(obj, depth) {
  const result = {};
  if (!obj || typeof obj !== 'object') return result;
  for (const [k, v] of Object.entries(obj)) {
    result[k.toLowerCase()] = v;
    if (depth > 0 && v && typeof v === 'object' && !Array.isArray(v)) {
      const sub = _flatten(v, depth - 1);
      for (const [sk, sv] of Object.entries(sub)) result[sk] = sv;
    }
  }
  return result;
}

// 값에서 01X 형식 전화번호 추출 (JSON 배열/중첩 구조 포함)
function _extractPhone(v) {
  if (v === null || v === undefined || v === '') return '';
  if (Array.isArray(v)) {
    for (const item of v) {
      const p = _extractPhone(item);
      if (p) return p;
    }
    return '';
  }
  if (typeof v === 'object') {
    for (const key of ['value', 'phone', 'number', 'tel']) {
      if (v[key]) { const p = _extractPhone(v[key]); if (p) return p; }
    }
    for (const val of Object.values(v)) {
      const p = _extractPhone(val);
      if (p) return p;
    }
    return '';
  }
  const s = String(v);
  const m = s.match(/\b(01[016789]\d{7,8})\b/);
  if (m) return m[1];
  if (s.length > 2 && (s[0] === '[' || s[0] === '{')) {
    try { return _extractPhone(JSON.parse(s)); } catch(_) {}
  }
  return '';
}

function extractCustomerDataFromSteps(steps) {
  let phone = '', templateId = '', dealId = '';

  for (const step of (steps || [])) {
    // input + output 둘 다 검색 (POLL_GQL은 input만, runQA 재조회는 둘 다)
    const sources = [step.input, step.output].map(_tryParseJson).filter(Boolean);
    if (!sources.length) continue;

    const flat = {};
    for (const src of sources) Object.assign(flat, _flatten(src, 3));

    // ── 전화번호: 01X 패턴 (JSON 배열/중첩 포함) ─────────────
    if (!phone) {
      for (const v of Object.values(flat)) {
        const p = _extractPhone(v);
        if (p) { phone = p; break; }
      }
    }

    const title = (step.title || '').toLowerCase();

    // ── 템플릿 ID: Solapi 계열 스텝 ──────────────────────────
    if (!templateId && (title.includes('solapi') || title.includes('alimtalk') || title.includes('kakao'))) {
      const tmplKeys = ['templateid', 'template_id', 'templatecode', 'template_code',
                        'kakaotemplatedid', 'kakao_template_id', 'kakaotemplatecode'];
      for (const k of tmplKeys) {
        if (flat[k]) { templateId = String(flat[k]); break; }
      }
      if (!templateId) {
        for (const [k, v] of Object.entries(flat)) {
          if (k.includes('template') && v) { templateId = String(v); break; }
        }
      }
    }

    // ── 딜 ID: Pipedrive 계열 스텝 (input 또는 output 모두 확인) ─
    if (!dealId && (title.includes('pipedrive') || title.includes('pipe'))) {
      for (const k of ['deal_id', 'dealid', 'id']) {
        const v = String(flat[k] || '');
        if (/^\d+$/.test(v) && v.length >= 3) { dealId = v; break; }
      }
    }
  }

  return { customerPhone: phone, templateId, dealId };
}

// 특정 run ID의 step 데이터를 Zapier에서 재조회 (POLL_GQL 동일 구조 사용)
function _fetchRunSteps(taskId, errorTimestampIso) {
  const { zapsession, csrftoken, accountId } = _getZapierSession();
  if (!zapsession || !taskId) return null;

  const ts = errorTimestampIso ? new Date(errorTimestampIso).getTime() : Date.now();
  const periodStart = new Date(ts - 8 * 60 * 60 * 1000).toISOString();
  const periodEnd   = new Date(ts + 8 * 60 * 60 * 1000).toISOString();

  const hdrs = {
    'Content-Type': 'application/json',
    'Cookie': 'zapsession=' + zapsession + (csrftoken ? '; csrftoken=' + csrftoken : ''),
    'X-CSRFToken': csrftoken || '',
    'Referer': 'https://zapier.com/app/history',
    'Origin': 'https://zapier.com',
  };

  const baseVars = { accountId, status: ['error'], limit: 100, offset: 0, periodStart, periodEnd, sortBy: '-start_time' };

  function tryQuery(stepsFragment, label) {
    const gql = `fragment ZapRun on ZapRun {
  id ${stepsFragment} __typename
}
query ZapRuns($accountId:ID!, $status:[String!], $limit:Int, $offset:Int, $periodStart:String, $periodEnd:String, $sortBy:String) {
  zapRuns(accountId:$accountId, status:$status, limit:$limit, offset:$offset, periodStart:$periodStart, periodEnd:$periodEnd, sortBy:$sortBy, apps:[], customuserIds:[], folderIds:[], zapIds:[]) {
    edges { ...ZapRun __typename }
    __typename
  }
}`;
    try {
      const res = UrlFetchApp.fetch('https://zapier.com/api/reporting/graphql', {
        method: 'POST', headers: hdrs,
        payload: JSON.stringify({ query: gql, variables: baseVars }),
        muteHttpExceptions: true,
      });
      const code = res.getResponseCode();
      if (code !== 200) { _logWebhook('FETCH_RUN_DBG', `${label} HTTP ${code}`); return null; }
      const d = JSON.parse(res.getContentText());
      const edges = d?.data?.zapRuns?.edges || [];
      _logWebhook('FETCH_RUN_DBG', `${label} edges=${edges.length}`);
      const run = edges.find(e => e.id === taskId);  // POLL_GQL: edge IS ZapRun (no .node wrapper)
      return run?.steps || null;
    } catch(e) { _logWebhook('FETCH_RUN_DBG', `${label} ERR ${e.message}`); return null; }
  }

  // 1차: output 포함 (phone은 Pipedrive step output에 있음)
  const s1 = tryQuery('steps { title status input output __typename }', '+output');
  if (s1) { _logWebhook('FETCH_RUN', `OK +output steps=${s1.length}`); return s1; }

  // 2차: input만 (output이 400 유발 시 fallback)
  const s2 = tryQuery('steps { title status input __typename }', '+input');
  if (s2) { _logWebhook('FETCH_RUN', `OK +input steps=${s2.length}`); return s2; }

  _logWebhook('FETCH_RUN', `not found taskId=${taskId}`);
  return null;
}

// 특정 기간의 에러를 수집해서 시트에 추가 (LAST_POLL_TIME 업데이트 없음)
function pollRange(periodStart, periodEnd) {
  const { zapsession, csrftoken, accountId } = _getZapierSession();

  if (!zapsession) return { success: false, error: '세션 없음 — Chrome Extension 동기화 필요' };

  const sheet = getSheet();
  const existingIds = new Set();
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, COL.taskId, lastRow - 1, 1).getValues()
      .forEach(([v]) => { if (v) existingIds.add(String(v)); });
  }

  const headers = {
    'Content-Type': 'application/json',
    'Cookie': `zapsession=${zapsession}; csrftoken=${csrftoken}`,
    'X-CSRFToken': csrftoken,
    'Referer': 'https://zapier.com/app/history',
    'Origin': 'https://zapier.com',
  };

  let offset = 0;
  const limit = 50;
  let added = 0, skipped = 0, totalFetched = 0;

  while (true) {
    try {
      const res = UrlFetchApp.fetch('https://zapier.com/api/reporting/graphql', {
        method: 'POST',
        headers,
        payload: JSON.stringify({
          operationName: 'ZapRuns',
          query: POLL_GQL,
          variables: {
            accountId,
            status: ['error'],
            limit, offset,
            periodStart, periodEnd,
            sortBy: '-start_time',
          }
        }),
        muteHttpExceptions: true,
      });

      const code = res.getResponseCode();
      if (code !== 200) return { success: false, error: `HTTP ${code}` };

      const body = JSON.parse(res.getContentText());
      if (body.errors) return { success: false, error: JSON.stringify(body.errors).substring(0, 300) };

      const edges = body?.data?.zapRuns?.edges || [];
      const totalCount = body?.data?.zapRuns?.totalCount || 0;
      totalFetched += edges.length;

      for (const run of edges) {
        const taskId = String(run.id || '');
        if (!taskId) continue;
        if (existingIds.has(taskId)) { skipped++; continue; }
        // Zap 정보 없는 경우(삭제된 Zap 등) ghost 행 방지
        if (!run.zap?.title) { skipped++; continue; }

        const startedAt = run.startTime ? new Date(run.startTime) : null;
        if (!startedAt || isNaN(startedAt.getTime())) continue;

        const steps = run.steps || [];
        const errorStep = steps.find(s => s.status === 'error');
        const stepName = errorStep?.title || '';
        const errMsg = errorStep?.error?.title
          || (stepName ? `${stepName} 실패 — Zapier 히스토리에서 상세 확인` : '에러 발생 — Zapier 히스토리 확인');

        // 모든 스텝의 input에서 고객 데이터 추출
        const extracted = extractCustomerDataFromSteps(steps);

        const error = {
          timestamp: startedAt.toISOString(),
          zapName: run.zap.title,
          stepName, errorMessage: errMsg,
          zapId: String(run.zap.id || ''),
          taskId,
          zapLink: run.zap.id ? `https://zapier.com/editor/${run.zap.id}/run/${taskId}?sidebar=runs&filterRun=true` : `https://zapier.com/app/history/usage/${taskId}`,
          customerPhone:       extracted.customerPhone,
          customerPhoneMasked: maskPhone(extracted.customerPhone),
          templateId:          extracted.templateId,
          dealId:              extracted.dealId,
          refundId:            '',
        };

        const rowIndex = logToSheet(error);
        const analysis = analyzeWithGroq(error);
        updateAnalysis(rowIndex, analysis);
        error.analysis = analysis;

        const qa = runQA(error);
        updateQAResult(rowIndex, qa);
        if (qa.checked && ['재실행 필요','삭제 필요','보류'].includes(qa.action)) {
          updateRowStatus(rowIndex, qa.action, '');
        }

        sendSlack(error);

        existingIds.add(taskId);
        added++;
      }

      offset += limit;
      if (edges.length < limit || (totalCount > 0 && offset >= totalCount)) break;

    } catch (err) {
      return { success: false, error: err.message };
    }
  }

  _logWebhook('RANGE_POLL', `기간조회 ${periodStart.substring(0,10)}~${periodEnd.substring(0,10)} | 총 ${totalFetched}건 확인 / 신규 ${added}건 추가 / 중복 ${skipped}건`);
  return { success: true, added, skipped, total: totalFetched };
}

function pollZapierErrors() {
  const { zapsession, csrftoken, accountId } = _getZapierSession();

  if (!zapsession) {
    _logWebhook('POLL_ERR', '세션 없음 — Chrome Extension 동기화 필요');
    return;
  }

  const p = PropertiesService.getScriptProperties();
  const lastPollStr = p.getProperty('LAST_POLL_TIME');
  const since = lastPollStr ? new Date(lastPollStr) : new Date(Date.now() - 60 * 60 * 1000);
  const now = new Date();

  const sheet = getSheet();
  const existingIds = new Set();
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, COL.taskId, lastRow - 1, 1).getValues()
      .forEach(([v]) => { if (v) existingIds.add(String(v)); });
  }

  const headers = {
    'Content-Type': 'application/json',
    'Cookie': `zapsession=${zapsession}; csrftoken=${csrftoken}`,
    'X-CSRFToken': csrftoken,
    'Referer': 'https://zapier.com/app/history',
    'Origin': 'https://zapier.com',
  };

  let offset = 0;
  const limit = 50;
  let added = 0, totalFetched = 0;

  while (true) {
    try {
      const res = UrlFetchApp.fetch('https://zapier.com/api/reporting/graphql', {
        method: 'POST',
        headers,
        payload: JSON.stringify({
          operationName: 'ZapRuns',
          query: POLL_GQL,
          variables: {
            accountId,
            status: ['error'],
            limit,
            offset,
            periodStart: since.toISOString(),
            periodEnd:   now.toISOString(),
            sortBy: '-start_time',
          }
        }),
        muteHttpExceptions: true,
      });

      const code = res.getResponseCode();
      if (code !== 200) {
        _logWebhook('POLL_ERR', `HTTP ${code}: ${res.getContentText().substring(0, 300)}`);
        break;
      }

      const body = JSON.parse(res.getContentText());

      // GraphQL 에러 로그
      if (body.errors) {
        _logWebhook('POLL_ERR', `GQL errors: ${JSON.stringify(body.errors).substring(0, 400)}`);
        break;
      }

      const edges = body?.data?.zapRuns?.edges || [];
      const totalCount = body?.data?.zapRuns?.totalCount || 0;

      totalFetched += edges.length;

      for (const run of edges) {
        const taskId = String(run.id || '');
        if (!taskId || existingIds.has(taskId)) continue;
        // Zap 정보 없는 경우(삭제된 Zap 등) ghost 행 방지
        if (!run.zap?.title) continue;

        const startedAt = run.startTime ? new Date(run.startTime) : null;
        if (!startedAt || isNaN(startedAt.getTime())) continue;

        const steps = run.steps || [];
        const errorStep = steps.find(s => s.status === 'error');
        const stepName = errorStep?.title || '';
        const errMsg = errorStep?.error?.title
          || (stepName ? `${stepName} 실패 — Zapier 히스토리에서 상세 확인` : '에러 발생 — Zapier 히스토리 확인');

        // 모든 스텝의 input에서 고객 데이터 추출
        const extracted = extractCustomerDataFromSteps(steps);

        const error = {
          timestamp:           startedAt.toISOString(),
          zapName:             run.zap.title,
          stepName,
          errorMessage:        errMsg,
          zapId:               String(run.zap.id || ''),
          taskId,
          zapLink:             run.zap.id ? `https://zapier.com/editor/${run.zap.id}/run/${taskId}?sidebar=runs&filterRun=true` : `https://zapier.com/app/history/usage/${taskId}`,
          customerPhone:       extracted.customerPhone,
          customerPhoneMasked: maskPhone(extracted.customerPhone),
          templateId:          extracted.templateId,
          dealId:              extracted.dealId,
          refundId:            '',
        };

        const rowIndex = logToSheet(error);
        const analysis = analyzeWithGroq(error);
        updateAnalysis(rowIndex, analysis);
        error.analysis = analysis;

        // webhook 경로와 동일하게 QA 즉시 실행
        const qa = runQA(error);
        updateQAResult(rowIndex, qa);
        if (qa.checked && ['재실행 필요','삭제 필요','보류'].includes(qa.action)) {
          updateRowStatus(rowIndex, qa.action, '');
        }

        sendSlack(error);

        existingIds.add(taskId);
        added++;
      }

      offset += limit;
      if (edges.length < limit || (totalCount > 0 && offset >= totalCount)) break;

    } catch (err) {
      _logWebhook('POLL_ERR', err.message);
      break;
    }
  }

  p.setProperty('LAST_POLL_TIME', now.toISOString());
  _logWebhook('POLL', `GraphQL ZapRuns 확인 ${totalFetched}건 / 신규 추가 ${added}건`);
}

// ── 로그 저장 전 개인정보 마스킹 (ISMS) ────────────────────────────
function _sanitizeLog(raw) {
  const MASK_KEYS = ['phone','customer_phone','customerPhone','token',
                     'zapsession','csrftoken','sessionJwt','ssoid'];
  const safe = Object.assign({}, raw);
  MASK_KEYS.forEach(k => { if (safe[k]) safe[k] = '***'; });
  return JSON.stringify(safe).substring(0, 500);
}

// ── DeleteLog 90일 초과 행 자동 삭제 (ISMS 보존기간) ─────────────────
function purgeOldDeleteLog() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.DELETE_LOG);
  if (!sheet || sheet.getLastRow() <= 1) return;
  const cutoff = Date.now() - 90 * 86400000;
  const dates = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  const toDelete = [];
  dates.forEach((r, i) => {
    const ts = r[0] ? new Date(r[0]).getTime() : 0;
    if (ts && ts < cutoff) toDelete.push(i + 2);
  });
  toDelete.reverse().forEach(row => sheet.deleteRow(row));
  if (toDelete.length) _logWebhook('PURGE_LOG', `DeleteLog 90일 초과 ${toDelete.length}건 삭제`);
}

// DeleteLog 자동 삭제 트리거 등록 (한 번만 실행)
function setupPurgeDeleteLogTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'purgeOldDeleteLog')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('purgeOldDeleteLog').timeBased().everyDays(1).atHour(3).create();
  Logger.log('DeleteLog 자동 삭제 트리거 등록 완료 (매일 새벽 3시)');
}

// Apps Script 에디터에서 한 번 실행해 5분 트리거 등록
function setupPollTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'pollZapierErrors')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('pollZapierErrors')
    .timeBased()
    .everyMinutes(5)
    .create();

  Logger.log('폴링 트리거 등록 완료 (5분 간격)');
}

// 트리거 제거
function removePollTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'pollZapierErrors')
    .forEach(t => ScriptApp.deleteTrigger(t));

  Logger.log('폴링 트리거 제거 완료');
}

// 시트 초기화 + 최근 7일 폴링 (일반 리셋용)
function resetAndPoll() {
  const sheet = getSheet();
  if (sheet.getLastRow() > 1) sheet.deleteRows(2, sheet.getLastRow() - 1);
  PropertiesService.getScriptProperties().setProperty('LAST_POLL_TIME', new Date(Date.now() - 7 * 24 * 60 * 60 * 1000).toISOString());
  pollZapierErrors();
}

// 시트 초기화 + 최근 30일 풀 시드 (최초 세팅 시 1회만 실행)
function seedPoll30d() {
  const sheet = getSheet();
  if (sheet.getLastRow() > 1) sheet.deleteRows(2, sheet.getLastRow() - 1);
  PropertiesService.getScriptProperties().setProperty('LAST_POLL_TIME', new Date(Date.now() - 30 * 24 * 60 * 60 * 1000).toISOString());
  pollZapierErrors();
}

// ============================================================
// 테스트: 특정 run의 output 데이터 확인 (에디터에서 실행)
// ============================================================
function testStepsOutput() {
  const TARGET_TASK_ID = '003cfa7c-6058-a62a-7a8f-4630d9ed6254';
  const TARGET_TS = '2026-07-31T01:50:50.193Z';

  const steps = _fetchRunSteps(TARGET_TASK_ID, TARGET_TS);
  if (!steps) { Logger.log('steps 없음 — 조회 실패'); return; }

  Logger.log('총 steps: ' + steps.length);
  steps.forEach((step, i) => {
    const hasOut = !!step.output;
    const hasIn  = !!step.input;
    if (step.title && step.title.toLowerCase().includes('pipe')) {
      Logger.log(`\n[${i}] ${step.title} status=${step.status}`);
      if (hasIn)  Logger.log(`  INPUT (앞300): ${String(step.input).substring(0,300)}`);
      if (hasOut) Logger.log(`  OUTPUT(앞300): ${String(step.output).substring(0,300)}`);
    }
  });

  const extracted = extractCustomerDataFromSteps(steps);
  Logger.log(`\n→ 추출결과: phone=${extracted.customerPhone||'없음'} deal=${extracted.dealId||'없음'}`);
}

// ============================================================
// 테스트: steps input 데이터 확인 (에디터에서 실행)
// ============================================================
function testStepsInput() {
  const { zapsession, csrftoken, accountId } = _getZapierSession();

  const headers = {
    'Content-Type': 'application/json',
    'Cookie': `zapsession=${zapsession}; csrftoken=${csrftoken}`,
    'X-CSRFToken': csrftoken,
    'Referer': 'https://zapier.com/app/history',
    'Origin': 'https://zapier.com',
  };

  const since = new Date(Date.now() - 30 * 24 * 60 * 60 * 1000).toISOString();
  const now   = new Date().toISOString();

  const res = UrlFetchApp.fetch('https://zapier.com/api/reporting/graphql', {
    method: 'POST', headers,
    payload: JSON.stringify({
      operationName: 'ZapRuns',
      query: POLL_GQL,
      variables: { accountId, status: ['error'], limit: 3, offset: 0, periodStart: since, periodEnd: now, sortBy: '-start_time' }
    }),
    muteHttpExceptions: true,
  });

  Logger.log('HTTP: ' + res.getResponseCode());
  const body = JSON.parse(res.getContentText());
  if (body.errors) { Logger.log('GQL errors: ' + JSON.stringify(body.errors)); return; }

  const edges = body?.data?.zapRuns?.edges || [];
  Logger.log('에러 run ' + edges.length + '건');

  edges.forEach(run => {
    Logger.log('\n── ' + run.zap?.title + ' (' + run.id + ')');
    (run.steps || []).forEach((step, i) => {
      const extracted = extractCustomerDataFromSteps([step]);
      Logger.log(`  [${i}] ${step.title} status=${step.status} phone=${extracted.customerPhone||'없음'} tmpl=${extracted.templateId||'없음'} deal=${extracted.dealId||'없음'}`);
      if (step.input) Logger.log(`      input(앞200자): ${JSON.stringify(step.input).substring(0, 200)}`);
    });
    const all = extractCustomerDataFromSteps(run.steps || []);
    Logger.log(`  → 전체 추출: phone=${all.customerPhone||'없음'} templateId=${all.templateId||'없음'} dealId=${all.dealId||'없음'}`);
  });
}

// ============================================================
// 진단: "개인 0" 같이 안 잡히는 에러의 실제 status 확인
// Apps Script 에디터에서 직접 실행
// ============================================================
function diagMissingErrors() {
  const { zapsession, csrftoken, accountId } = _getZapierSession();

  const headers = {
    'Content-Type': 'application/json',
    'Cookie': `zapsession=${zapsession}; csrftoken=${csrftoken}`,
    'X-CSRFToken': csrftoken,
    'Referer': 'https://zapier.com/app/history',
    'Origin': 'https://zapier.com',
  };

  const statuses = ['error', 'errored', 'held', 'delayed', 'throttled', 'paused'];

  Logger.log('=== status 별 건수 조회 (최근 30일) ===');
  const since = new Date(Date.now() - 30 * 24 * 60 * 60 * 1000).toISOString();
  const now   = new Date().toISOString();

  for (const s of statuses) {
    try {
      const res = UrlFetchApp.fetch('https://zapier.com/api/reporting/graphql', {
        method: 'POST', headers,
        payload: JSON.stringify({
          operationName: 'ZapRuns',
          query: `query ZapRuns($accountId: ID!, $status: [String!], $periodStart: String, $periodEnd: String) {
  zapRuns(accountId: $accountId, status: $status, limit: 5, offset: 0, periodStart: $periodStart, periodEnd: $periodEnd, sortBy: "-start_time", apps: [], customuserIds: [], folderIds: [], zapIds: []) {
    edges { id status startTime zap { title __typename } steps { status title error { title __typename } __typename } __typename }
    totalCount __typename
  }
}`,
          variables: { accountId, status: [s], periodStart: since, periodEnd: now }
        }),
        muteHttpExceptions: true,
      });
      const body = JSON.parse(res.getContentText());
      const total = body?.data?.zapRuns?.totalCount || 0;
      const edges = body?.data?.zapRuns?.edges || [];
      Logger.log(`status="${s}" → totalCount=${total}`);
      edges.forEach(r => {
        const errStep = (r.steps || []).find(st => st.status === 'error' || st.error?.title);
        Logger.log(`  zap="${r.zap?.title}" id=${r.id} apiStatus=${r.status} errStep="${errStep?.title||'없음'}" errMsg="${errStep?.error?.title||'없음'}"`);
      });
    } catch(e) {
      Logger.log(`status="${s}" → 오류: ${e.message}`);
    }
    Utilities.sleep(300);
  }

  // status 필터 없이 전체 (오늘)
  Logger.log('\n=== status 필터 없음 — 오늘 전체 runs (최근 20건) ===');
  try {
    const todayStart = new Date(); todayStart.setHours(0,0,0,0);
    const res = UrlFetchApp.fetch('https://zapier.com/api/reporting/graphql', {
      method: 'POST', headers,
      payload: JSON.stringify({
        operationName: 'ZapRunsAll',
        query: `query ZapRunsAll($accountId: ID!, $periodStart: String, $periodEnd: String) {
  zapRuns(accountId: $accountId, status: [], limit: 20, offset: 0, periodStart: $periodStart, periodEnd: $periodEnd, sortBy: "-start_time", apps: [], customuserIds: [], folderIds: [], zapIds: []) {
    edges { id status startTime zap { title __typename } __typename }
    totalCount __typename
  }
}`,
        variables: { accountId, periodStart: todayStart.toISOString(), periodEnd: now }
      }),
      muteHttpExceptions: true,
    });
    const body = JSON.parse(res.getContentText());
    Logger.log(`totalCount(오늘): ${body?.data?.zapRuns?.totalCount}`);
    (body?.data?.zapRuns?.edges || []).forEach(r => {
      Logger.log(`  status="${r.status}" zap="${r.zap?.title}" id=${r.id}`);
    });
  } catch(e) {
    Logger.log('전체 조회 오류: ' + e.message);
  }
}

// ============================================================
// [방법 3 테스트] ZapRuns 리스트에 steps 포함 + REST 엔드포인트
// ============================================================
function testMethod3() {
  const { zapsession, csrftoken, accountId } = _getZapierSession();

  const headers = {
    'Content-Type': 'application/json',
    'Cookie': `zapsession=${zapsession}; csrftoken=${csrftoken}`,
    'X-CSRFToken': csrftoken,
    'Referer': 'https://zapier.com/app/history',
    'Origin': 'https://zapier.com',
  };

  const since = new Date(Date.now() - 30 * 24 * 60 * 60 * 1000).toISOString();
  const now   = new Date().toISOString();

  // ── 테스트 -1: 오늘 전체 run 상태 조회 (status 필터 없음) ────
  try {
    const todayStart = new Date(); todayStart.setHours(0,0,0,0);
    const resAll = UrlFetchApp.fetch('https://zapier.com/api/reporting/graphql', {
      method: 'POST', headers,
      payload: JSON.stringify({
        operationName: 'ZapRunsAll',
        query: `query ZapRunsAll($accountId: ID!, $limit: Int, $periodStart: String, $periodEnd: String) {
  zapRuns(accountId: $accountId, status: [], limit: $limit, offset: 0, periodStart: $periodStart, periodEnd: $periodEnd, sortBy: "-start_time", apps: [], customuserIds: [], folderIds: [], zapIds: []) {
    edges { id status startTime zap { id title __typename } __typename }
    totalCount __typename
  }
}`,
        variables: { accountId, limit: 20, periodStart: todayStart.toISOString(), periodEnd: now }
      }),
      muteHttpExceptions: true,
    });
    const bodyAll = JSON.parse(resAll.getContentText());
    Logger.log('=== 테스트 -1: 오늘 전체 runs ===');
    Logger.log('HTTP: ' + resAll.getResponseCode());
    Logger.log('totalCount: ' + bodyAll?.data?.zapRuns?.totalCount);
    const edges = bodyAll?.data?.zapRuns?.edges || [];
    edges.forEach(r => Logger.log(`status="${r.status}" zap="${r.zap?.title}" id=${r.id}`));
    if (bodyAll?.errors) Logger.log('errors: ' + JSON.stringify(bodyAll.errors));
  } catch(e) {
    Logger.log('테스트 -1 실패: ' + e.message);
  }

  // ── 테스트 0: 현재 세션의 계정 목록 조회 ──────────────────────
  try {
    const resAcct = UrlFetchApp.fetch('https://zapier.com/api/reporting/graphql', {
      method: 'POST', headers,
      payload: JSON.stringify({
        operationName: 'Accounts',
        query: `query Accounts { accounts { id name __typename } }`,
        variables: {}
      }),
      muteHttpExceptions: true,
    });
    Logger.log('=== 테스트 0: 계정 목록 ===');
    Logger.log('HTTP: ' + resAcct.getResponseCode());
    Logger.log(resAcct.getContentText().substring(0, 1000));
  } catch(e) {
    Logger.log('테스트 0 실패: ' + e.message);
  }

  // ── 테스트 A: ZapRuns 리스트에 steps 필드 추가 ──────────────
  const GQL_WITH_STEPS = `fragment ZapRun on ZapRun {
  id startTime status
  steps { error { title __typename } status title __typename }
  zap { id title sourceUrl __typename }
  __typename
}
query ZapRuns($accountId: ID!, $status: [String!], $limit: Int, $offset: Int, $periodStart: String, $periodEnd: String, $sortBy: String) {
  zapRuns(accountId: $accountId, status: $status, limit: $limit, offset: $offset, periodStart: $periodStart, periodEnd: $periodEnd, sortBy: $sortBy, apps: [], customuserIds: [], folderIds: [], zapIds: []) {
    edges { ...ZapRun __typename }
    totalCount __typename
  }
}`;

  try {
    const resA = UrlFetchApp.fetch('https://zapier.com/api/reporting/graphql', {
      method: 'POST', headers,
      payload: JSON.stringify({
        operationName: 'ZapRuns',
        query: GQL_WITH_STEPS,
        variables: { accountId, status: ['error'], limit: 3, offset: 0,
          periodStart: since, periodEnd: now, sortBy: '-start_time' }
      }),
      muteHttpExceptions: true,
    });
    const bodyA = JSON.parse(resA.getContentText());
    const firstRun = bodyA?.data?.zapRuns?.edges?.[0];
    Logger.log('=== 테스트 A: ZapRuns + steps ===');
    Logger.log('HTTP: ' + resA.getResponseCode());
    Logger.log('사용된 accountId: ' + accountId);
    Logger.log('totalCount: ' + bodyA?.data?.zapRuns?.totalCount);
    Logger.log('edges 수: ' + (bodyA?.data?.zapRuns?.edges?.length ?? 'N/A'));
    Logger.log('GQL errors: ' + JSON.stringify(bodyA?.errors));
    Logger.log('첫 번째 run: ' + JSON.stringify(firstRun, null, 2));
    if (firstRun?.steps?.length) {
      Logger.log('✅ steps 있음! 에러 스텝: ' + JSON.stringify(firstRun.steps.find(s => s.status === 'error')));
    } else {
      Logger.log('❌ steps 없음 또는 빈 배열');
    }
  } catch(e) {
    Logger.log('테스트 A 실패: ' + e.message);
  }

  // ── 테스트 B: REST /api/zap-history/v2/runs (zap_id 기반) ──
  // 먼저 리스트에서 zap_id 하나 가져오기
  try {
    const resBase = UrlFetchApp.fetch('https://zapier.com/api/reporting/graphql', {
      method: 'POST', headers,
      payload: JSON.stringify({
        operationName: 'ZapRuns',
        query: `query ZapRuns($accountId: ID!, $status: [String!], $limit: Int, $periodStart: String, $periodEnd: String) {
  zapRuns(accountId: $accountId, status: $status, limit: $limit, offset: 0, periodStart: $periodStart, periodEnd: $periodEnd, sortBy: "-start_time", apps: [], customuserIds: [], folderIds: [], zapIds: []) {
    edges { id zap { id __typename } __typename }
  }
}`,
        variables: { accountId, status: ['error'], limit: 1, periodStart: since, periodEnd: now }
      }),
      muteHttpExceptions: true,
    });
    const baseBody = JSON.parse(resBase.getContentText());
    const run = baseBody?.data?.zapRuns?.edges?.[0];
    if (!run) { Logger.log('테스트 B: 에러 run 없음'); return; }

    const taskId = run.id;
    const zapId  = run.zap?.id;
    Logger.log('=== 테스트 B: REST /api/zap-history/v2/runs ===');
    Logger.log(`taskId=${taskId}, zapId=${zapId}`);

    // B-1: zap_id 기반 REST
    const resB1 = UrlFetchApp.fetch(
      `https://zapier.com/api/zap-history/v2/runs?zap_id=${zapId}&status=error&limit=5`,
      { method: 'GET', headers: { 'Cookie': `zapsession=${zapsession}; csrftoken=${csrftoken}`, 'X-CSRFToken': csrftoken },
        muteHttpExceptions: true }
    );
    Logger.log('B-1 HTTP: ' + resB1.getResponseCode());
    Logger.log('B-1 body: ' + resB1.getContentText().substring(0, 800));

    // B-2: run ID 직접 조회
    const resB2 = UrlFetchApp.fetch(
      `https://zapier.com/api/zap-history/v2/runs/${taskId}`,
      { method: 'GET', headers: { 'Cookie': `zapsession=${zapsession}; csrftoken=${csrftoken}`, 'X-CSRFToken': csrftoken },
        muteHttpExceptions: true }
    );
    Logger.log('B-2 HTTP: ' + resB2.getResponseCode());
    Logger.log('B-2 body: ' + resB2.getContentText().substring(0, 800));

  } catch(e) {
    Logger.log('테스트 B 실패: ' + e.message);
  }
}

// ============================================================
// 유틸
// ============================================================

function reanalyzeAll() {
  const sheet = getSheet();
  const last = sheet.getLastRow();
  if (last <= 1) return;
  const rows = sheet.getRange(2, 1, last - 1, 7).getValues();
  let count = 0;
  rows.forEach((r, i) => {
    const a = r[6];
    if (!a || a === '분석 중...' || a === '(API 키 미설정)' || a.startsWith('분석 오류')) {
      const result = analyzeWithGroq({ zapName: r[1], stepName: r[2], errorMessage: r[3] });
      sheet.getRange(i + 2, COL.analysis).setValue(result);
      count++; Utilities.sleep(500);
    }
  });
  Logger.log(`재분석 완료: ${count}건`);
}

function importFromCsv() {
  const files = DriveApp.getFilesByName('zapier_history.csv');
  if (!files.hasNext()) { Logger.log('zapier_history.csv 없음'); return; }

  const csv = files.next().getBlob().getDataAsString();
  const rows = Utilities.parseCsv(csv);
  const header = rows[0];
  const IDX = {
    ident: header.indexOf('ident'), date: header.indexOf('date'),
    object_title: header.indexOf('object_title'), object_id: header.indexOf('object_id'),
    status: header.indexOf('status'), message: header.indexOf('message'),
  };

  const props = PropertiesService.getScriptProperties();
  const processed = JSON.parse(props.getProperty('importedRows') || '[]');
  let count = 0;

  rows.slice(1).forEach(row => {
    const status = (row[IDX.status] || '').toLowerCase();
    if (status !== 'error' && status !== 'errored') return;
    const rowKey = row[IDX.ident] || '';
    if (!rowKey || processed.includes(rowKey)) return;

    const rawMsg = row[IDX.message] || '';
    const error = {
      timestamp:    row[IDX.date] ? new Date(row[IDX.date]).toISOString() : new Date().toISOString(),
      zapName:      row[IDX.object_title] || 'Unknown Zap',
      stepName:     row[IDX.object_id]    || 'Unknown Step',
      errorMessage: rawMsg || '오류 메시지 없음 (CSV 미포함)',
      zapId: row[IDX.object_id] || '', taskId: rowKey,
      zapLink: rowKey ? `https://zapier.com/app/history/usage/${rowKey}` : '',
      customerPhone: '', templateId: '', dealId: '',
    };

    const rowIndex = logToSheet(error);
    const analysis = analyzeWithGroq(error);
    updateAnalysis(rowIndex, analysis);
    processed.push(rowKey); count++;
    Utilities.sleep(500);
  });

  props.setProperty('importedRows', JSON.stringify(processed.slice(-500)));
  Logger.log(`CSV 임포트 완료: ${count}건`);
}

function checkCsvHeaders() {
  const files = DriveApp.getFilesByName('zapier_history.csv');
  if (!files.hasNext()) { Logger.log('파일 없음'); return; }
  const rows = Utilities.parseCsv(files.next().getBlob().getDataAsString());
  Logger.log('헤더: ' + JSON.stringify(rows[0]));
  Logger.log('1행: ' + JSON.stringify(rows[1]));
}

function testManual() {
  doPost({
    postData: {
      contents: JSON.stringify({
        zap_name: '[테스트] 개인 알림톡 발송',
        step_name: 'Solapi: Send Alimtalk',
        error_message: 'Template not found: invalid template ID',
        task_id: 'test_001',
        zap_link: 'https://zapier.com/app/history',
        customer_phone: '010-XXXX-XXXX',
        template_id: 'TEST_TEMPLATE_01',
        deal_id: '123',
      })
    }
  });
}

// ============================================================
// 기존 고착 행 일괄 재QA (GAS 에디터에서 한 번 실행)
// qaResult가 '확인 대기중' / 'QA 대기중' / '분석 중...' 상태인 모든 행을 재처리
// ============================================================
function reprocessStuckRows() {
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log('처리할 행 없음'); return; }

  const data = sheet.getRange(2, 1, lastRow - 1, 14).getValues();
  const STUCK = ['확인 대기중', 'QA 대기중', '수동 확인 필요 (데이터 없음)'];

  let processed = 0, skipped = 0, errors = 0;

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const rowId = i + 2;
    const qaResult = String(row[COL.qaResult - 1] || '');
    const analysis = String(row[COL.analysis - 1] || '');

    const isStuck = STUCK.includes(qaResult) || qaResult === '' || analysis === '분석 중...';
    if (!isStuck) { skipped++; continue; }

    try {
      const error = _buildErrorFromRow(row);

      // 분석도 안 됐으면 재분석
      if (!analysis || analysis === '분석 중...' || analysis.startsWith('분석 오류')) {
        const newAnalysis = analyzeWithGroq(error);
        updateAnalysis(rowId, newAnalysis);
        error.analysis = newAnalysis;
        Utilities.sleep(500);
      }

      const qa = runQA(error);
      updateQAResult(rowId, qa);
      if (qa.checked && ['재실행 필요', '삭제 필요', '보류'].includes(qa.action)) {
        updateRowStatus(rowId, qa.action, '');
      }

      Logger.log(`[${rowId}] ${error.zapName || ''} → checked:${qa.checked} action:${qa.action}`);
      processed++;
      Utilities.sleep(300);
    } catch (e) {
      Logger.log(`[${rowId}] 오류: ${e.message}`);
      errors++;
    }
  }

  const msg = `재QA 완료 — 처리: ${processed}건 / 건너뜀: ${skipped}건 / 오류: ${errors}건`;
  Logger.log(msg);
  _logWebhook('REPROCESS', msg);
}
