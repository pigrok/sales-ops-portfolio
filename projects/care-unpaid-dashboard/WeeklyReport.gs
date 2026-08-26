/**
 * Bookkeeping Subscription SaaS — 주간 미납 리포트 Slack 자동 발송
 *
 * 매주 화요일 17시, 지정 채널에 주간 미납 통계 자동 전송.
 * ────────────────────────────────────────────────────────────────
 *
 * 데이터 규약
 * - 모든 지표는 시트 표시값 그대로 사용 (재계산 금지)
 * - 대표자명이 빈 값이거나 제외 대상인 행은 모든 통계에서 제외
 * - 대상월 = 오늘 기준 지난달 (4월이면 3월, 1월이면 전년 12월)
 * - 이번 주 = 이번 주 월요일 00:00 ~ 오늘 23:59
 *
 * 함수 목록 (모두 weeklyReport_ 접두사로 격리, 기존 스크립트와 충돌 없음)
 * - weeklyReport_run()              : 트리거가 호출하는 메인 함수
 * - weeklyReport_test()             : 발송 없이 메시지 미리보기 (로그 출력)
 * - weeklyReport_setupWebhookUrl()  : Webhook URL 1회 등록 헬퍼
 * - weeklyReport_buildMessage_()    : 시트 → 메시지 텍스트 (내부)
 * - weeklyReport_postToSlack_()     : Webhook 발송 (내부)
 */

// ====== 설정 ======
const WEEKLY_REPORT_CONFIG = {
  CURRENT_TAB: '미납현황',
  DASHBOARD_TAB: '대시보드',
  // 월별 탭은 'YYYY-MM' 형식으로 자동 매칭
  EXCLUDE_NAMES: ['', 'EXAMPLE_EXCLUDED_NAME'],
};

// ====== 메인 ======
function weeklyReport_run() {
  // 실제 발송 — 발송 후 현재 미납자 스냅샷 저장 (다음 주 비교 기준)
  const message = weeklyReport_buildMessage_(true);
  weeklyReport_postToSlack_(message);
  console.log('[weeklyReport] 발송 성공');
}

function weeklyReport_test() {
  // 테스트 — 스냅샷 갱신 안 함
  const message = weeklyReport_buildMessage_(false);
  console.log('=== 메시지 미리보기 ===');
  console.log(message);
}

function weeklyReport_resetSnapshot() {
  const props = PropertiesService.getScriptProperties();
  const all = props.getProperties();
  let n = 0;
  Object.keys(all).forEach(k => {
    if (k.indexOf('WEEKLY_REPORT_SNAPSHOT_') === 0) {
      props.deleteProperty(k);
      n++;
    }
  });
  console.log('[weeklyReport] 스냅샷 ' + n + '개 삭제됨');
}

function weeklyReport_setupWebhookUrl() {
  // ⚠ 사용 후에는 url을 빈 문자열로 되돌리고 저장하세요 (코드에 평문 보관 방지)
  const url = '';
  if (!url) {
    throw new Error('weeklyReport_setupWebhookUrl: url 변수에 webhook URL을 입력 후 다시 실행하세요.');
  }
  PropertiesService.getScriptProperties().setProperty('WEEKLY_REPORT_WEBHOOK_URL', url);
  console.log('[weeklyReport] Webhook URL 저장 완료');
}

// ====== 메시지 빌드 ======
function weeklyReport_buildMessage_(updateSnapshot) {
  const ss = SpreadsheetApp.getActive();
  const tz = Session.getScriptTimeZone() || 'Asia/Seoul';
  const today = new Date();

  // 1) 대상월 (지난달) 계산
  const curMonth = today.getMonth() + 1;
  const curYear = today.getFullYear();
  let targetMonth = curMonth - 1;
  let targetYear = curYear;
  if (targetMonth === 0) { targetMonth = 12; targetYear -= 1; }
  const prevTabName = targetYear + '-' + String(targetMonth).padStart(2, '0');

  // 2) 대시보드 지표 추출 (표시값 그대로)
  const dash = ss.getSheetByName(WEEKLY_REPORT_CONFIG.DASHBOARD_TAB);
  if (!dash) throw new Error('대시보드 탭을 찾을 수 없음: ' + WEEKLY_REPORT_CONFIG.DASHBOARD_TAB);
  const dashDisp = dash.getDataRange().getDisplayValues();

  let dashHeaderIdx = -1;
  const dashCols = {};
  // 대시보드에 다른 테이블이 있어서 '모수' 등 컬럼 이름이 중복될 수 있음.
  // → '귀속월' 컬럼 위치를 기준으로 오른쪽으로 빈 셀/다른 테이블 헤더까지만 매핑.
  const TABLE_BOUNDARY = new Set(['고객구간', '구분', '미납 현황 대시보드']);
  for (let i = 0; i < dashDisp.length; i++) {
    const row = dashDisp[i].map(c => String(c).trim());
    const guiso = row.indexOf('귀속월');
    if (guiso !== -1 && row.indexOf('최초미납률') !== -1) {
      dashHeaderIdx = i;
      for (let j = guiso; j < row.length; j++) {
        const name = row[j];
        if (!name) break;
        if (TABLE_BOUNDARY.has(name)) break;
        dashCols[name] = j;
      }
      break;
    }
  }
  if (dashHeaderIdx === -1) throw new Error('대시보드에서 귀속월/최초미납률 헤더를 찾을 수 없음');

  let metrics = null;
  for (let i = dashHeaderIdx + 1; i < dashDisp.length; i++) {
    const ymStr = String(dashDisp[i][dashCols['귀속월']] || '').trim();
    if (!ymStr) continue;
    const m = ymStr.match(/^(\d{4})[-./]\s*0?(\d{1,2})$/);
    if (!m) continue;
    if (Number(m[1]) === targetYear && Number(m[2]) === targetMonth) {
      metrics = {
        모수: dashDisp[i][dashCols['모수']],
        최초미납수: dashDisp[i][dashCols['최초미납수']],
        최초미납률: dashDisp[i][dashCols['최초미납률']],
        최종미납수: dashDisp[i][dashCols['최종미납수']],
        최종미납률: dashDisp[i][dashCols['최종미납률']],
        회수율: dashDisp[i][dashCols['회수율']],
      };
      break;
    }
  }
  if (!metrics) throw new Error('대시보드에서 ' + prevTabName + ' 행을 찾을 수 없음');

  // 3) 미납현황 탭에서 주의 고객 / 사유 TOP5 산출
  const curSh = ss.getSheetByName(WEEKLY_REPORT_CONFIG.CURRENT_TAB);
  if (!curSh) throw new Error('미납현황 탭을 찾을 수 없음: ' + WEEKLY_REPORT_CONFIG.CURRENT_TAB);
  const curDisp = curSh.getDataRange().getDisplayValues();

  let curHeaderIdx = -1;
  const curCols = {};
  for (let i = 0; i < curDisp.length; i++) {
    const row = curDisp[i].map(c => String(c).trim());
    const ownerCol = row.indexOf('대표자명');
    if (ownerCol !== -1 && row.indexOf('미납개월수') !== -1) {
      curHeaderIdx = i;
      for (let j = ownerCol; j < row.length; j++) {
        const name = row[j];
        if (!name) continue;
        if (j !== ownerCol && (name === '대표자명' || name === '이름' || name === '강제해지일')) break;
        if (!(name in curCols)) curCols[name] = j;
      }
      break;
    }
  }
  if (curHeaderIdx === -1) throw new Error('미납현황 탭에서 헤더를 찾을 수 없음');

  const exclude = new Set(WEEKLY_REPORT_CONFIG.EXCLUDE_NAMES.map(s => s.trim()));
  const curRows = [];
  let blankRun = 0;
  for (let i = curHeaderIdx + 1; i < curDisp.length; i++) {
    const owner = String(curDisp[i][curCols['대표자명']] || '').trim();
    if (owner === '대표자명' || owner === '이름' || owner === '강제해지일') break;
    if (!owner) {
      blankRun++;
      if (blankRun >= 3) break;
      continue;
    }
    blankRun = 0;
    if (exclude.has(owner)) continue;
    curRows.push(curDisp[i]);
  }

  // 주의 고객 (3개월↑)
  const attentionAll = curRows.map(r => ({
    대표자명: String(r[curCols['대표자명']] || '').trim(),
    상호: String(r[curCols['상호']] || '').trim(),
    미납개월수: Number(String(r[curCols['미납개월수']] || '0').replace(/[^0-9.-]/g, '')) || 0,
    담당자: String(r[curCols['담당자']] || '').trim(),
  }));
  const attention = attentionAll
    .filter(x => x.미납개월수 >= 3)
    .sort((a, b) => b.미납개월수 - a.미납개월수);

  // 4) 이번 주 신규 미납 = 직전 스냅샷에 없던 사업자번호
  const prevSh = ss.getSheetByName(prevTabName);
  if (!prevSh) throw new Error('월별 탭을 찾을 수 없음: ' + prevTabName);
  const prevDisp = prevSh.getDataRange().getDisplayValues();

  let prevHeaderIdx = -1;
  const prevCols = {};
  for (let i = 0; i < prevDisp.length; i++) {
    const row = prevDisp[i].map(c => String(c).trim());
    const ownerCol = row.indexOf('대표자명');
    if (ownerCol !== -1 && row.indexOf('최근확인일') !== -1) {
      prevHeaderIdx = i;
      for (let j = ownerCol; j < row.length; j++) {
        const name = row[j];
        if (!name) continue;
        if (j !== ownerCol && (name === '이름' || name === '대표자명' || name === '강제해지일')) break;
        if (!(name in prevCols)) prevCols[name] = j;
      }
      break;
    }
  }
  if (prevHeaderIdx === -1) throw new Error('월별 탭 ' + prevTabName + ' 에서 헤더를 찾을 수 없음');

  const todayStr = Utilities.formatDate(today, tz, 'yyyy-MM-dd');

  const currentKeys = [];
  const currentDetail = {};
  const prevReasonCounts = {};
  let prevBlankRun = 0;
  for (let i = prevHeaderIdx + 1; i < prevDisp.length; i++) {
    const owner = String(prevDisp[i][prevCols['대표자명']] || '').trim();
    if (owner === '대표자명' || owner === '이름' || owner === '강제해지일') break;
    if (!owner) {
      prevBlankRun++;
      if (prevBlankRun >= 3) break;
      continue;
    }
    prevBlankRun = 0;
    if (exclude.has(owner)) continue;

    const status = String(prevDisp[i][prevCols['상태']] || '').trim();
    if (status === '미납') {
      const bizno = String(prevDisp[i][prevCols['사업자번호']] || '').trim();
      const sangho = String(prevDisp[i][prevCols['상호']] || '').trim();
      const key = owner + '|' + bizno;
      currentKeys.push(key);
      currentDetail[key] = owner + ' (' + sangho + ')';
      const reason = String(prevDisp[i][prevCols['사유']] || '').trim();
      if (reason) prevReasonCounts[reason] = (prevReasonCounts[reason] || 0) + 1;
    }
  }

  // 신규 미납 매칭
  const snapshotKey = 'WEEKLY_REPORT_SNAPSHOT_' + prevTabName;
  const prevSnapshotJson = PropertiesService.getScriptProperties().getProperty(snapshotKey);
  let prevSnapshot = null;
  if (prevSnapshotJson) {
    try { prevSnapshot = new Set(JSON.parse(prevSnapshotJson)); } catch (e) { prevSnapshot = null; }
  }

  let newCount = 0;
  const newList = [];
  if (prevSnapshot) {
    currentKeys.forEach(k => {
      if (!prevSnapshot.has(k)) { newCount++; newList.push(currentDetail[k]); }
    });
  }

  if (updateSnapshot) {
    PropertiesService.getScriptProperties().setProperty(snapshotKey, JSON.stringify(currentKeys));
  }

  const reasonTop5 = Object.keys(prevReasonCounts)
    .map(k => ({ 사유: k, 건수: prevReasonCounts[k] }))
    .sort((a, b) => b.건수 - a.건수 || a.사유.localeCompare(b.사유, 'ko-KR'))
    .slice(0, 5);

  // 5) 메시지 텍스트 조립 (Slack mrkdwn: *bold*)
  const NBSP = ' ';

  const lines = [];
  lines.push('*📊 주간 미납 현황 (' + todayStr + ')*');
  lines.push('⸻');
  lines.push('*📅 ' + targetMonth + '월 미납 기준 (모수: ' + _wr_num_(metrics.모수) + ')*');
  lines.push('▫️ 최초 미납률: ' + _wr_str(metrics.최초미납률) + ' (' + _wr_num_(metrics.최초미납수) + '건)');
  lines.push('▫️ 현재 미납률: ' + _wr_str(metrics.최종미납률) + ' (' + _wr_num_(metrics.최종미납수) + '건)');
  lines.push('▫️ 회수율: ' + _wr_str(metrics.회수율));
  lines.push('▫️ 이번 주 신규 미납: ' + newCount + '건 🆕');
  lines.push('⸻');
  lines.push('*⚠️ 주의 고객 (3개월↑ 미납)*');
  if (attention.length === 0) {
    lines.push('- (해당 없음)');
  } else {
    attention.forEach(a => {
      lines.push('- ' + a.대표자명 + ' (' + a.상호 + ') - ' + a.미납개월수 + '개월, 담당: ' + a.담당자);
    });
  }
  lines.push(NBSP);
  lines.push('⸻');
  lines.push('*📌 미납 사유 TOP 5*');
  reasonTop5.forEach((e, i) => {
    lines.push((i + 1) + '. ' + e.사유 + ' - ' + e.건수 + '건');
  });

  return lines.join('\n');
}

function _wr_str(v) {
  if (v === null || v === undefined) return '';
  return String(v).trim();
}

function _wr_num_(v) {
  if (v === null || v === undefined || v === '') return '';
  if (typeof v === 'number') return v.toLocaleString('ko-KR');
  const s = String(v).trim();
  if (/^\d+$/.test(s)) return Number(s).toLocaleString('ko-KR');
  return s;
}

function weeklyReport_postToSlack_(message) {
  const url = PropertiesService.getScriptProperties().getProperty('WEEKLY_REPORT_WEBHOOK_URL');
  if (!url) throw new Error('Script Property "WEEKLY_REPORT_WEBHOOK_URL" 미설정. weeklyReport_setupWebhookUrl 또는 [프로젝트 설정 → 스크립트 속성] 에서 등록 필요.');
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ text: message }),
    muteHttpExceptions: true,
  });
  const code = res.getResponseCode();
  const body = res.getContentText();
  if (code !== 200) {
    throw new Error('Webhook 발송 실패: ' + code + ' / ' + body);
  }
}
