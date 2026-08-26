const FILTER_CONFIG = [
  { name: '[zap 미실행] 4-0. 작성완료_록', label: '4-작성완료', triggerStage: '작성완료' },
  { name: '[zap 미실행] 5-0. 신고완료_록', label: '5-신고완료', triggerStage: '신고완료' },
  { name: '[zap 미실행] 6-0. 환급결정_록', label: '6-환급결정', triggerStage: '환급결정' },
];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Pipedrive 필터 오류 대시보드')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ─── 대시보드 데이터 (프론트에서 호출) ─────────────────────────────────────
function getDashboardData() {
  const token = PropertiesService.getScriptProperties().getProperty('PIPEDRIVE_TOKEN');
  const filters  = getFilters_(token);
  const stageMap = getStageMap_(token);                      // { id → name }
  const nameToId = buildNameToId_(stageMap);                 // { name → id }

  return FILTER_CONFIG.map((cfg, filterIdx) => {
    const filter = filters.find(f => f.name === cfg.name);
    if (!filter) return { ...cfg, filterIdx, deals: [], error: '필터를 찾을 수 없음' };

    const deals = getDeals_(token, filter.id);
    const triggerStageId = nameToId[cfg.triggerStage];

    const rows = deals.map(deal => {
      try {
        const flowEvents   = getDealFlow_(token, deal.id);
        const stageInfo    = findStageEntry_(flowEvents, triggerStageId);
        const personId     = deal.person_id?.value ?? deal.person_id ?? null;
        const phone        = personId ? getPersonPhone_(token, personId) : null;
        const solapiResult = (phone && stageInfo?.timestamp)
          ? checkSolapi_(phone, stageInfo.timestamp)
          : { sent: null, time: null, type: null, error: phone ? '단계 진입 기록 없음' : '연락처 없음' };

        return {
          id: deal.id,
          title: deal.title,
          currentStageId: deal.stage_id,
          currentStageName: stageMap[deal.stage_id] || '?',
          triggerStageId,
          prevStageId: stageInfo?.prevStageId || null,
          stageEntryTime: stageInfo?.timestamp || null,
          phone,
          solapiSent: solapiResult.sent,
          solapiTime: solapiResult.time,
          solapiType: solapiResult.type,
          solapiError: solapiResult.error,
          pipedriveUrl: `https://app.pipedrive.com/deal/${deal.id}`,
        };
      } catch (e) {
        const safeMsg = e.message.replace(/api_token=[^&\s]*/g, 'api_token=***');
        return { id: deal.id, title: deal.title, pipedriveUrl: `https://app.pipedrive.com/deal/${deal.id}`, error: safeMsg };
      }
    });

    return { ...cfg, filterIdx, deals: rows };
  });
}

// ─── 재발송 트리거 (프론트에서 호출) ────────────────────────────────────────
function retriggerAlimtalk(dealId, filterIdx) {
  const token  = PropertiesService.getScriptProperties().getProperty('PIPEDRIVE_TOKEN');
  const cfg    = FILTER_CONFIG[filterIdx];
  const nameToId = buildNameToId_(getStageMap_(token));
  const triggerStageId = nameToId[cfg.triggerStage];

  const deal = JSON.parse(
    UrlFetchApp.fetch(`https://api.pipedrive.com/v1/deals/${dealId}?api_token=${token}`).getContentText()
  ).data;
  const currentStageId = deal.stage_id;

  if (currentStageId === triggerStageId) {
    // 이미 트리거 단계 → 이전 단계로 빠졌다가 재진입
    const stageInfo = findStageEntry_(getDealFlow_(token, dealId), triggerStageId);
    if (!stageInfo?.prevStageId) throw new Error('히스토리에서 이전 단계를 찾을 수 없습니다');
    moveDeal_(token, dealId, stageInfo.prevStageId);
    Utilities.sleep(2000);
    moveDeal_(token, dealId, triggerStageId);
  } else {
    // 다른 단계 → 트리거 단계로 이동 후 복귀
    moveDeal_(token, dealId, triggerStageId);
    Utilities.sleep(2000);
    moveDeal_(token, dealId, currentStageId);
  }

  return { success: true };
}

// ─── Pipedrive 헬퍼 ──────────────────────────────────────────────────────────
function getFilters_(token) {
  return JSON.parse(
    UrlFetchApp.fetch(`https://api.pipedrive.com/v1/filters?type=deals&api_token=${token}`).getContentText()
  ).data || [];
}

function getStageMap_(token) {
  const map = {};
  (JSON.parse(
    UrlFetchApp.fetch(`https://api.pipedrive.com/v1/stages?api_token=${token}`).getContentText()
  ).data || []).forEach(s => { map[s.id] = s.name; });
  return map;
}

function buildNameToId_(stageMap) {
  const m = {};
  Object.entries(stageMap).forEach(([id, name]) => { m[name] = Number(id); });
  return m;
}

function getDeals_(token, filterId) {
  const deals = [];
  let start = 0;
  while (true) {
    const data = JSON.parse(
      UrlFetchApp.fetch(`https://api.pipedrive.com/v1/deals?filter_id=${filterId}&start=${start}&limit=100&api_token=${token}`).getContentText()
    );
    if (!data.data) break;
    deals.push(...data.data);
    if (!data.additional_data?.pagination?.more_items_in_collection) break;
    start += 100;
  }
  return deals;
}

function getDealFlow_(token, dealId) {
  const events = [];
  let start = 0;
  while (start < 500) {
    const data = JSON.parse(
      UrlFetchApp.fetch(`https://api.pipedrive.com/v1/deals/${dealId}/flow?start=${start}&limit=100&api_token=${token}`).getContentText()
    );
    if (!data.data) break;
    events.push(...data.data);
    if (!data.additional_data?.pagination?.more_items_in_collection) break;
    start += 100;
  }
  return events;
}

// flow에서 특정 stageId 진입 이벤트 찾기 (최근 기준)
function findStageEntry_(events, stageId) {
  for (const e of events) {
    if (e.object === 'dealChange' && e.data?.field_key === 'stage_id' &&
        Number(e.data?.new_value) === stageId) {
      return {
        timestamp: e.timestamp || e.log_time,
        prevStageId: Number(e.data?.old_value) || null,
      };
    }
  }
  return null;
}

function getPersonPhone_(token, personId) {
  const phones = JSON.parse(
    UrlFetchApp.fetch(`https://api.pipedrive.com/v1/persons/${personId}?api_token=${token}`).getContentText()
  ).data?.phone || [];
  if (!phones.length) return null;
  return normalizePhone_(phones[0].value);
}

function normalizePhone_(raw) {
  if (!raw) return null;
  let d = raw.replace(/\D/g, '');
  if (d.startsWith('82') && d.length >= 12) d = '0' + d.slice(2);
  return d || null;
}

function moveDeal_(token, dealId, stageId) {
  UrlFetchApp.fetch(`https://api.pipedrive.com/v1/deals/${dealId}?api_token=${token}`, {
    method: 'patch',
    contentType: 'application/json',
    payload: JSON.stringify({ stage_id: stageId }),
  });
}

// ─── Solapi 헬퍼 ─────────────────────────────────────────────────────────────
function solapiAuth_() {
  const props = PropertiesService.getScriptProperties();
  const key    = props.getProperty('SOLAPI_API_KEY');
  const secret = props.getProperty('SOLAPI_API_SECRET');
  const date   = new Date().toISOString();
  const salt   = Math.random().toString(36).substring(2, 16);
  const sig    = Utilities.computeHmacSha256Signature(date + salt, secret)
    .map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
  return `HMAC-SHA256 apiKey=${key}, date=${date}, salt=${salt}, signature=${sig}`;
}

function checkSolapi_(phone, stageTimestamp) {
  try {
    const center = new Date(stageTimestamp.replace(' ', 'T') + '+09:00'); // Pipedrive = KST
    const from   = new Date(center.getTime() - 30 * 60 * 1000).toISOString();
    const to     = new Date(center.getTime() + 30 * 60 * 1000).toISOString();

    // type 파라미터 없이 조회 → SMS·ATA·LMS 전체 체크
    const url = `https://api.solapi.com/messages/v4/list?to=${phone}` +
      `&dateCreatedFrom=${encodeURIComponent(from)}&dateCreatedTo=${encodeURIComponent(to)}&limit=20`;

    const res  = UrlFetchApp.fetch(url, { headers: { Authorization: solapiAuth_() }, muteHttpExceptions: true });
    const data = JSON.parse(res.getContentText());
    const msgs = Object.values(data.messageList || {});

    if (!msgs.length) return { sent: false, time: null, type: null, error: null };

    const ok = msgs.find(m => !String(m.statusCode).startsWith('5'));
    if (ok) return { sent: true, time: ok.dateCreated, type: ok.type || null, error: null };

    return { sent: false, time: null, type: null, error: msgs[0].statusMessage || '발송 실패' };
  } catch (e) {
    return { sent: null, time: null, error: e.message };
  }
}
