const GAS_URL   = 'https://script.google.com/macros/s/AKfycbzvDwmyMFPLcT0kUKF3RWAdBNxqdfug26sMK2ruI2WfpbRdeFzAN80tnVAG-WlMO6Jy/exec';
const GAS_TOKEN = 'ea5ffb9d120a06a171de11fe7ab976c6';

// ── 쿠키 헬퍼 ────────────────────────────────────────────────────
async function getZapierCookies() {
  const names = ['zapsession', 'csrftoken', '__Host-session.jwt', 'currentAccountId', 'ssoid'];
  const result = {};
  for (const name of names) {
    const c = await chrome.cookies.get({ url: 'https://zapier.com', name });
    if (c) result[name] = c.value;
  }
  return result;
}

// ── 쿠키 동기화 ──────────────────────────────────────────────────
async function syncCookies() {
  const cookies = await getZapierCookies();
  if (!cookies['zapsession']) {
    console.log('[ZapSync] Zapier 로그인 필요');
    await chrome.storage.local.set({ lastSync: null, status: 'not_logged_in' });
    return;
  }
  try {
    const payload = JSON.stringify({
      action: 'syncCookies',
      token:      GAS_TOKEN,
      zapsession: cookies['zapsession'],
      csrftoken:  cookies['csrftoken']         || '',
      sessionJwt: cookies['__Host-session.jwt'] || '',
      accountId:  cookies['currentAccountId']   || '',
      ssoid:      cookies['ssoid']              || '',
    });
    await fetch(GAS_URL, {
      method: 'POST', mode: 'no-cors',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: payload,
    });
    const now = new Date().toLocaleString('ko-KR', { timeZone: 'Asia/Seoul' });
    console.log('[ZapSync] 쿠키 동기화 완료:', now);
    await chrome.storage.local.set({ lastSync: now, status: 'ok' });
  } catch (err) {
    console.error('[ZapSync] 쿠키 동기화 오류:', err);
    await chrome.storage.local.set({ status: 'error' });
  }
}

// ── Zapier: 최근 오류 런 ID 목록 (브라우저 세션) ─────────────────
const LIST_GQL = `fragment ZapRun on ZapRun {
  id startTime status
  zap { id title __typename }
  __typename
}
query ZapRuns($accountId:ID!, $status:[String!], $limit:Int, $offset:Int, $periodStart:String, $sortBy:String) {
  zapRuns(accountId:$accountId, status:$status, limit:$limit, offset:$offset,
          periodStart:$periodStart, sortBy:$sortBy,
          apps:[], customuserIds:[], folderIds:[], zapIds:[]) {
    edges { ...ZapRun __typename }
    totalCount __typename
  }
}`;

// ── Zapier: 단건 run 상세 (step input/output 포함) ────────────────
const DETAIL_GQL = `query ZapRun($id:ID!) {
  zapRun(id:$id) {
    id
    steps { id title status input output error { title } __typename }
    __typename
  }
}`;

async function fetchErrorRunIds(accountId) {
  const periodStart = new Date(Date.now() - 30 * 86400000).toISOString(); // 30일
  console.log('[ZapSync] 런 목록 조회 accountId:', accountId);
  try {
    const res = await fetch('https://zapier.com/api/reporting/graphql', {
      method: 'POST', credentials: 'include',
      headers: { 'Content-Type': 'application/json', 'Accept': 'application/json' },
      body: JSON.stringify({
        query: LIST_GQL,
        variables: { accountId, status: ['error'], limit: 30, offset: 0, periodStart, sortBy: 'startTime' },
      }),
    });
    console.log('[ZapSync] 런 목록 응답:', res.status);
    if (!res.ok) { console.error('[ZapSync] 런 목록 오류:', res.status); return []; }
    const json = await res.json();
    if (json.errors) console.error('[ZapSync] 런 목록 GQL errors:', JSON.stringify(json.errors).slice(0,200));
    const edges = json?.data?.zapRuns?.edges || [];
    console.log('[ZapSync] 런 목록:', edges.length + '건');
    return edges;
  } catch (e) { console.error('[ZapSync] fetchErrorRunIds 오류:', e); return []; }
}

async function fetchRunDetail(taskId) {
  try {
    const res = await fetch('https://zapier.com/api/reporting/graphql', {
      method: 'POST', credentials: 'include',
      headers: { 'Content-Type': 'application/json', 'Accept': 'application/json' },
      body: JSON.stringify({ query: DETAIL_GQL, variables: { id: taskId } }),
    });
    if (!res.ok) return null;
    const json = await res.json();
    return json?.data?.zapRun?.steps || null;
  } catch (e) { console.error('[ZapSync] fetchRunDetail 오류:', e); return null; }
}

// ── step 데이터에서 phone / templateId / dealId 추출 ─────────────
function flattenObj(obj, depth, prefix) {
  if (depth <= 0 || !obj || typeof obj !== 'object') return {};
  prefix = prefix || '';
  const result = {};
  for (const [k, v] of Object.entries(obj)) {
    const key = prefix ? `${prefix}.${k}` : k;
    if (Array.isArray(v)) {
      v.forEach((item, i) => {
        if (item && typeof item === 'object') Object.assign(result, flattenObj(item, depth - 1, `${key}[${i}]`));
        else result[`${key}[${i}]`] = item;
      });
    } else if (v && typeof v === 'object') {
      Object.assign(result, flattenObj(v, depth - 1, key));
    } else {
      result[key] = v;
    }
  }
  return result;
}

function extractFromSteps(steps) {
  let phone = '', templateId = '', dealId = '';
  for (const step of (steps || [])) {
    const stepTitle = (step.title || '').toLowerCase();
    const isPipedrive = stepTitle.includes('pipedrive');

    for (const src of [step.input, step.output]) {
      if (!src) continue;
      let obj;
      try { obj = typeof src === 'string' ? JSON.parse(src) : src; } catch(_) { continue; }
      const flat = flattenObj(obj, 4);
      for (const [k, v] of Object.entries(flat)) {
        const s = String(v || '').trim();
        const kl = k.toLowerCase();

        // 전화번호: 숫자만 추출 후 패턴 매칭
        if (!phone) {
          const digits = s.replace(/[^0-9]/g, '');
          if (/^(01[016789]\d{7,8})$/.test(digits)) phone = digits;
        }

        // 템플릿ID: Kakao 형식(KA01TP...) 값 패턴 OR 키 이름 기반
        if (!templateId) {
          if (/^KA01[A-Za-z0-9]{6,}/.test(s)) {
            templateId = s;
          } else if ((kl.includes('template') || kl.includes('templateid') || kl.includes('templatecode'))
                     && s.length > 5 && s !== 'null' && s !== 'undefined') {
            templateId = s;
          }
        }

        // 딜ID: 키가 dealid/deal_id이거나, Pipedrive 스텝의 id 필드
        if (!dealId && /^\d{3,8}$/.test(s)) {
          if (kl.includes('dealid') || kl.includes('deal_id')) {
            dealId = s;
          } else if (isPipedrive && (kl === 'id' || kl.endsWith('.id') || kl === 'data.id')) {
            dealId = s;
          }
        }
      }
    }
    if (phone && templateId && dealId) break;
  }
  return { phone, templateId, dealId };
}

// ── push-based 데이터 보강: Zapier → GAS ─────────────────────────
// 이미 GAS에 전송한 taskId 캐시 (서비스워커 재시작 시 초기화되는 것은 무방)
const _enrichedCache = new Set();

async function enrichRecentErrors() {
  const cookies = await getZapierCookies();
  const accountId = cookies['currentAccountId'];
  if (!accountId) { console.log('[ZapSync] accountId 없음, 스킵'); return; }

  const runs = await fetchErrorRunIds(accountId);
  if (!runs.length) { console.log('[ZapSync] 최근 오류 런 없음'); return; }

  console.log(`[ZapSync] 오류 런 ${runs.length}건 step 조회 시작`);
  let updated = 0;

  for (const run of runs) {
    const taskId = run.id;

    // 이미 처리한 taskId는 재전송 안 함
    if (_enrichedCache.has(taskId)) continue;

    const steps = await fetchRunDetail(taskId);
    if (!steps) { await new Promise(r => setTimeout(r, 200)); continue; }

    const { phone, templateId, dealId } = extractFromSteps(steps);

    // 데이터 없어도 GAS 호출 — GAS가 패턴 기반으로 판단 (Pipedrive 인증오류 등)
    console.log(`[ZapSync] taskId=${taskId} phone=${!!phone} tpl=${!!templateId} deal=${!!dealId}`);

    await fetch(GAS_URL, {
      method: 'POST', mode: 'no-cors',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ action: 'updateErrorStepData', token: GAS_TOKEN, taskId, phone, templateId, dealId }),
    });
    _enrichedCache.add(taskId);
    updated++;
    await new Promise(r => setTimeout(r, 300));
  }
  console.log(`[ZapSync] 보강 완료: ${updated}건 업데이트`);
}

// ── 알람 등록 ────────────────────────────────────────────────────
chrome.alarms.create('syncCookies',  { periodInMinutes: 30 });
chrome.alarms.create('enrichSteps',  { periodInMinutes: 1 });

chrome.alarms.onAlarm.addListener((alarm) => {
  if (alarm.name === 'syncCookies') syncCookies();
  if (alarm.name === 'enrichSteps')  enrichRecentErrors();
});

chrome.runtime.onInstalled.addListener(() => { syncCookies(); enrichRecentErrors(); });
chrome.runtime.onStartup.addListener(()    => { syncCookies(); enrichRecentErrors(); });

chrome.runtime.onMessage.addListener((msg, sender, reply) => {
  if (msg.action === 'manualSync') {
    syncCookies().then(() => enrichRecentErrors()).then(() => reply({ ok: true }));
    return true;
  }
  if (msg.action === 'enrichNow') {
    enrichRecentErrors().then(() => reply({ ok: true }));
    return true;
  }
});
