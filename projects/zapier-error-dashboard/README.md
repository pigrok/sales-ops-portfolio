# Zapier Error Dashboard

Zapier Zap 오류 로그를 자동 수집·분석·대응하는 사내 대시보드. **LLM으로 오류 유형을 자동 분류**하고 우선순위별로 슬랙 노티를 보냅니다.

## 문제

Zapier Zap **98건**을 운영하는 상황에서 매일 아침 Task History를 눈으로 훑으며 실패 원인을 파악해야 했습니다. 오류 유형이 다양하고(API 키 만료·Rate limit·필드 매핑 오류·페이로드 스키마 변경 등) 노이즈가 심해서 진짜 대응이 필요한 오류를 놓치기 쉬웠습니다.

## 해결

Zapier에서 오류 페이로드를 웹훅으로 받아 시트에 로깅하고, LLM에 넘겨 유형별로 자동 분류. 신규·재발 유형은 슬랙에 노티, 반복 유형은 자동 처리. 대응 상태를 대시보드에서 관리.

## 아키텍처

```
Zapier Zap 실패
      ↓  Webhook
GAS doPost() → logToSheet()
      ↓
LLM 분석 (runQA)  →  updateAnalysis()
  · 오류 유형 분류
  · 원인 요약
  · 재발 여부 판단
      ↓
Slack 알림 (신규·재발 케이스만)
      ↓
Web Dashboard (doGet)
  · 유형별 집계
  · 처리 상태 관리
  · Pipedrive 딜 링크 연동
```

## 스택

Google Apps Script (Web App) · LLM (Groq) · Pipedrive API · Solapi API · Slack Webhook · Chrome Extension (Zapier 페이로드 캡처 보조)

## 규모

- Code.gs 약 **98KB / 2,375줄**
- Chrome Extension 별도 (Zapier UI에서 오류 페이로드 캡처 보조)
- 누적 오류 처리 **3,500+ 건, 미처리 0**

## 주요 함수

- `doPost()` — Zapier 웹훅 수신 · 인증 토큰 검증
- `logToSheet()` — 오류 로그 시트 저장
- `runQA()` — LLM에 오류 메시지 넘겨 유형 분류 · 원인 요약
- `updateAnalysis()` — 분석 결과를 시트에 반영
- `updateRowStatus()` — 대응 상태 업데이트 (대기·처리 중·완료)
- `purgeGhostRows()` — 유령 로그 정리
- `maskPhone()` — 로그에 담긴 전화번호 마스킹 (개인정보 보호)
- `_solapiStatus()` — Solapi 발송 상태 코드 해석

## 보안 설계

- **인증**: `_checkToken()` — Zapier 웹훅에 사전 공유 토큰 필수
- **개인정보 마스킹**: 로그에 담긴 전화번호는 `maskPhone()`으로 자동 마스킹
- **API 키**: Pipedrive(3개 계정), Groq, Zapier CSRF 모두 Script Properties
- **감사 로그**: 오류 발생·분석·처리 이력 시트에 시계열 기록

## 결과

- 슬랙 노티 발송량: 하루 40건 → 5건 (신규·재발 유형만)
- 오류 대응 시간: 케이스당 평균 15분 → 3분
- Pipedrive API v1→v2 마이그레이션 **98건 무중단 전환**
- 단일 대량 장애(1,500건) **5시간 내 정상화**

## 스크린샷

- [ ] Zapier Task History (오류 대응 이력 3,500+건)
- [ ] Zapier 대시보드 Zap 목록 (98개)
- [ ] 사내 오류 분석 대시보드 (유형별 집계)
- [ ] 슬랙 자동 노티 예시 (마스킹)
- [ ] Chrome extension UI

## 관련 자료

- 노션 프로젝트 페이지: [환급 Zapier 오류 대응 운영](../../docs/related-project-notes.md)
- Confluence 티켓 참조: REFOP-270 (내부)
