# Refund KPI Calculator

대시보드 이미지를 첨부하면 Google Drive OCR로 수치를 자동 추출해 KPI를 계산하고, 슬랙 보고문을 자동 생성하는 GAS 도구.

## 문제
매일 아침 대시보드 카드에서 숫자를 눈으로 옮겨 KPI(신고율·총 성공 환급액·커버리지)를 계산하고 슬랙에 붙여야 했습니다.

## 기능
- 이미지 첨부 → Google Drive OCR로 수치 추출
- 대시보드 카드별 금액 자동 파싱
- 신고율 = 신고환급액 / 총 조회환급액
- 총 성공 환급액 = 신청전환 + 취소방어
- 슬랙 보고문 자동 생성
- 데일리 저장·월 마감 조회 (Google Sheets)
- 스키마 변경 대응 정규화 함수

## 스택
Google Apps Script (Web App) · Google Drive API (OCR) · Google Sheets

## 파일
- `Code.gs` — 서버 로직 (469줄)
- `Index.html` — 웹 앱 UI
- `appsscript.json` — 매니페스트 (Drive API 고급 서비스 선언)

## 보안
- `executeAs = 배포자`: 모든 사용자 저장 데이터가 배포자 Drive에 집계됨 → 접근 제어 단일화
- 시트 ID 자동 생성 후 Script Properties(`DATA_SPREADSHEET_ID`) 저장
