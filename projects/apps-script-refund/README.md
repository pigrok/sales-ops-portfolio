# Refund KPI Calculator

대시보드 이미지에서 OCR로 수치를 자동 추출해 환급 KPI를 계산하는 사내 도구.

## 문제

매일 아침 세일즈 담당자가 회사 대시보드의 여러 카드에서 숫자를 하나씩 눈으로 옮겨 KPI(신고율·총 성공 환급액·커버리지)를 계산하고 슬랙에 보고하고 있었습니다. 실수 잦고 반복 노동.

## 해결

대시보드 스크린샷을 업로드하면 Google Drive OCR로 수치를 자동 추출해 KPI를 계산하고, 슬랙 보고 텍스트를 자동 생성. 데일리 저장·월 마감 조회 기능도 함께.

## 아키텍처

```
GAS Web App (doGet)
      ↓  이미지 첨부
Google Drive OCR
  · createImageBlob_
  · driveOcr_
      ↓  raw text
parseFields_ (필드 인덱스 매핑)
  · extractDashboardAmounts_
  · extractCardAmount_
      ↓  KPI 계산
슬랙 보고문 자동 생성
      ↓  선택적 저장
saveDaily → Google Sheets
```

## 계산 로직

- **신고율** = 신고환급액_담당자 / 총 거래 전환 조회환급액
- **총 성공 환급액** = 신청전환 성공 조회환급액 + 취소방어 성공 조회환급액

## 스택

Google Apps Script (Web App) · Google Drive API (OCR) · Google Sheets

## 주요 함수

- `ocrExtract()` — 이미지 OCR 처리 진입점
- `parseFields_()` — OCR 텍스트에서 필드별 인덱스 파싱
- `extractDashboardAmounts_()` / `extractCardAmount_()` — 대시보드 카드별 금액 추출
- `saveDaily()` — 데일리 시트 저장 (같은 날짜 덮어쓰기)
- `getDailySheet_()` / `migrateDailySheets_()` — 시트 스키마 관리·마이그레이션
- `normalizeDailyRow_()` / `normalizeCurrentDailyRow_()` — 스키마 변경 대응 정규화

## 보안 설계

- **executeAs 배포자**: 모든 사용자 저장 데이터가 배포자 Drive에 집계됨 → 접근 제어 단일화
- **시트 ID 자동 생성**: 최초 실행 시 자동 생성, Script Properties(`DATA_SPREADSHEET_ID`)에 저장
- OCR 텍스트만 처리, 원본 이미지는 Drive 임시 폴더에 저장 후 정리

## 결과

- KPI 산출 시간 단축 (수기 계산·복사·붙여넣기 대체)
- 슬랙 보고 표준화

## 배포

1. https://script.google.com 에서 새 프로젝트 생성
2. 파일 3개 붙여넣기: `Code.gs`, `Index.html`, `appsscript.json`
3. Drive API 고급 서비스 추가 (`appsscript.json`에 이미 선언됨)
4. `배포 > 새 배포 > 유형: 웹 앱`
5. 실행 권한: `나(배포자)` / 액세스: `Google 계정을 가진 모든 사용자`

## 주의

OCR 자동추출은 특정 대시보드 타일 배치 순서에 맞춰져 있습니다. 대시보드 레이아웃이 바뀌면 `Code.gs`의 `parseFields_` 인덱스 매핑을 조정해야 합니다.
