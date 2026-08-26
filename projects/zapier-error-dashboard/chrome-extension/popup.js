chrome.storage.local.get(['lastSync', 'status'], (data) => {
  const el = document.getElementById('status');
  const timeEl = document.getElementById('lastSync');

  if (data.status === 'ok') {
    el.className = 'status ok';
    el.textContent = '✅ 동기화 정상';
  } else if (data.status === 'not_logged_in') {
    el.className = 'status error';
    el.textContent = '❌ Zapier 로그인 필요';
  } else if (data.status === 'error') {
    el.className = 'status error';
    el.textContent = '⚠️ 동기화 실패';
  } else {
    el.className = 'status unknown';
    el.textContent = '⏳ 아직 동기화 안 됨';
  }

  if (data.lastSync) timeEl.textContent = '마지막: ' + data.lastSync;
});

document.getElementById('syncBtn').addEventListener('click', () => {
  document.getElementById('status').textContent = '동기화 중...';
  chrome.runtime.sendMessage({ action: 'manualSync' }, () => {
    chrome.storage.local.get(['lastSync', 'status'], (data) => {
      const el = document.getElementById('status');
      if (data.status === 'ok') {
        el.className = 'status ok';
        el.textContent = '✅ 동기화 완료';
      } else {
        el.className = 'status error';
        el.textContent = '❌ 실패 - Zapier 로그인 확인';
      }
    });
  });
});
