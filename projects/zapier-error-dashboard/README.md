# Zapier Error Dashboard

Zapier Zap 오류 로그를 자동 수집하고 LLM으로 유형을 분류해 슬랙 노티를 최소화하는 사내 대시보드.

## 문제
Zap 98건을 운영하는 상황에서 매일 Task History를 눈으로 훑는 방식은 노이즈가 심하고 놓치기 쉬웠습니다. 오류 유형이 다양했습니다 — API 키 만료, Rate limit, 필드 매핑 오류, 페이로드 스키마 변경 등.

## 기능
- Zapier 실패 페이로드를 웹훅으로 수신 → 시트 로깅
- LLM에 오류 메시지 넘겨 유형·원인 자동 분류
- 신규·재발 유형만 슬랙 노티, 반복 유형은 자동 처리
- 대응 상태 대시보드 (대기 · 처리 중 · 완료)
- 서비스별 CRM 계정 라우팅
- Solapi 발송 상태 조회
- Chrome Extension으로 Zapier UI 페이로드 캡처 보조

## 스택
Google Apps Script (Web App) · LLM (Groq) · Pipedrive API · Solapi API · Slack Webhook · Chrome Extension

## 파일
- `Code.gs` — 서버 로직 (2,375줄)
- `index.html` — 대시보드 UI
- `chrome-extension/` — Zapier UI 보조 확장 (background·content·popup)
- `appsscript.json` — 매니페스트

## 보안
- 웹훅 인증: `_checkToken()` 사전 공유 토큰 검증
- 로그 마스킹: `maskPhone()` 전화번호 자동 마스킹
- API 키: Pipedrive · LLM · Zapier 관련 인증 정보 모두 Script Properties 관리
