# Pipedrive Filter Dashboard

Pipedrive 필터 결과를 웹 대시보드로 보여주고, 각 딜의 알림톡 발송 상태를 확인·재발송하는 사내 도구.

## 문제

Pipedrive 필터별 딜 리스트를 확인하려면 매번 CRM에 로그인해야 했고, 각 딜의 알림톡 발송 이력·Solapi 발송 상태(도달·미도달)도 별도로 확인해야 했습니다. 반복 클릭이 많은 업무.

## 해결

GAS Web App으로 필터별 딜을 한 화면에 시각화하고, Solapi API로 실제 발송 상태를 조회. 미도달 딜은 원클릭으로 알림톡 재발송 트리거.

## 아키텍처

```
GAS Web App (doGet)
      ↓
Pipedrive API v1
  · getFilters_    (필터 목록)
  · getStageMap_   (단계 매핑)
  · getDeals_      (필터별 딜)
  · getDealFlow_   (딜 이벤트 스트림)
  · getPersonPhone_ (연락처 조회)
      ↓
Solapi API (checkSolapi_) → 발송 상태
      ↓
Web Dashboard (Index.html)
  · 필터별 딜 요약
  · 발송 상태 뱃지
  · 재발송 버튼
      ↓  onClick
retriggerAlimtalk (Solapi 재발송 + 딜 단계 이동)
```

## 스택

Google Apps Script (Web App) · Pipedrive API · Solapi API

## 규모

작지만 잘 만들어진 도구. Code.gs 약 8KB / 214줄, Index.html 약 9KB / 대시보드 UI.

## 주요 함수

- `getDashboardData()` — 필터별 딜 · 발송 상태 통합 데이터
- `getDealFlow_()` — 딜의 이벤트 스트림에서 특정 단계 진입 시각 추출
- `retriggerAlimtalk()` — Solapi로 재발송 트리거 · 딜 단계 이동
- `moveDeal_()` — Pipedrive 딜 단계 자동 변경
- `checkSolapi_()` — Solapi 발송 상태 조회

## 보안 설계

- **로그 마스킹**: 로그 출력 시 `api_token=***` 자동 마스킹 (라인 53)
- **API 토큰**: Script Properties (`PIPEDRIVE_TOKEN`)
- **전화번호 정규화**: `normalizePhone_()` — 발송 전 국제 표준 변환

## 결과

- 필터별 딜 상태 조회 시간 대폭 단축
- 알림톡 재발송 원클릭화 → CS 담당자 반복 작업 감소

## 스크린샷

- [ ] 대시보드 UI (필터별 딜 요약)
- [ ] 발송 상태 뱃지 (도달·미도달·재발송 필요)
- [ ] 재발송 후 Solapi 상태 변경 로그

## 배포

```bash
./deploy.sh
```
`clasp`로 자동 배포.
