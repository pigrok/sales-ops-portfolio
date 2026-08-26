# Pipedrive Filter Dashboard

Pipedrive 필터별 딜 현황과 Solapi 알림톡 발송 상태를 한 화면에 보여주고, 미도달 딜은 원클릭으로 재발송 트리거하는 사내 도구.

## 문제
필터별 딜 리스트 확인과 알림톡 발송 상태 조회를 위해 매번 여러 시스템을 왕복해야 했습니다.

## 기능
- Pipedrive 필터별 딜 목록 시각화
- 딜 이벤트 스트림에서 특정 단계 진입 시각 추출
- Solapi 발송 상태 조회 (도달·미도달·재발송 필요)
- 원클릭 알림톡 재발송 + Pipedrive 딜 단계 자동 이동
- 전화번호 정규화 (국제 표준)

## 스택
Google Apps Script (Web App) · Pipedrive API · Solapi API

## 파일
- `Code.gs` — 서버 로직 (214줄)
- `Index.html` — 대시보드 UI
- `deploy.sh` — clasp 자동 배포
- `appsscript.json` — 매니페스트

## 보안
- API 토큰: Script Properties (`PIPEDRIVE_TOKEN`)
- 로그 출력 시 `api_token=***` 자동 마스킹
