// 대시보드(GAS) → 익스텐션 background 중계 (postMessage 방식)
window.addEventListener('message', (e) => {
  if (e.source !== window || e.data?.type !== 'ZENT_ENRICH_NOW') return;
  chrome.runtime.sendMessage({ action: 'enrichNow' });
});

// Zapier 탭에서 실행 — background의 gqlFetch 요청을 same-origin으로 중계
chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  if (msg.action !== 'gqlFetch') return;

  const csrftoken = document.cookie.split('; ')
    .find(r => r.startsWith('csrftoken='))?.split('=')[1] || '';

  fetch('https://zapier.com/api/reporting/graphql', {
    method: 'POST',
    credentials: 'same-origin',
    headers: {
      'Content-Type': 'application/json',
      'X-CSRFToken': csrftoken,
      'apollographql-client-name': 'reporting',
    },
    body: JSON.stringify({
      operationName: msg.operationName,
      query: msg.query,
      variables: msg.variables,
    }),
  })
    .then(r => r.json())
    .then(data => sendResponse({ ok: true, data }))
    .catch(err => sendResponse({ ok: false, error: err.message }));

  return true; // async 유지
});
