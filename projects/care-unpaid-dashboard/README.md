# Care Unpaid Dashboard

세무 기장 구독 서비스의 **주간 미납 관리를 자동화**하는 GAS 사내 도구. 회사 결제 시스템 CSV를 업로드하면 사업자번호 기준 그룹화 → 월별 탭 자동 적재 → 완납·강제해지 상태 자동 감지 → Chart.js 대시보드 → 슬랙 주간 리포트까지 한 파이프라인으로 처리합니다.

## 문제

기장 구독 고객의 미납 관리가 다음 세 가지에서 병목이었습니다:
1. **집계 인프라 부재** — 원본 CSV 뽑아서 매주 수작업 정리
2. **완납·강제해지 상태 판단** — 사업자별 미납개월수를 눈으로 세고 있었음
3. **주간 리포트** — 매주 화요일 팀 회의 자료를 사람이 만들어 슬랙에 붙여야 했음

## 해결

Google Sheets + GAS로 원스톱 시스템 구축:
- CSV 업로드 한 번에 미납현황·월별 탭·블랙리스트·대시보드 전 자동 갱신
- **강제해지 자동 감지**: 3개월+ 미납 → 예정 등록, 실제 해지 확정 시 완료로 자동 승격
- **월별 탭 자동 상태 처리**: 이번 달 CSV에 없으면 완납, canceled 고객이면 강제해지 (버그 픽스 커밋 기록 있음)
- **Chart.js 웹 대시보드**: 웹앱으로 배포 (누구든 링크로 접근)
- **슬랙 주간 리포트**: 매주 화요일 17시 자동 발송, "이번 주 신규 미납"은 스냅샷 비교로 산출

## 아키텍처

```
회사 결제 시스템 CSV
      ↓  주간 업로드 (주간_입력 시트)
GAS syncUnpaidData
      ↓  사업자번호 그룹화
[미납현황] ─→ 월별 탭 (2026-01, 2026-02 ...)
      ↓  강제해지 감지 (canceled 고객 × 3개월+)
[블랙리스트]
   ├─ 예정 = 3개월+ 미납 현재 고객
   └─ 완료 = 실제 강제해지 (예정→완료 자동 승격)
      ↓
[대시보드] 탭 (귀속월별 집계)
      ↓  Chart.js
Web App (doGet) ─→ 팀 접근 URL
      ↓  weekly cron (화 17시)
Slack #internal-care-ops
   ├─ 이번 주 신규 미납 (스냅샷 비교)
   ├─ 3개월+ 주의 고객
   └─ 미납 사유 TOP 5
```

## 스택

Google Apps Script · Google Sheets · Chart.js (Web App) · Slack Webhook · Time-driven Trigger

## 규모

- Code.gs 약 **30KB / 982줄** — 메인 로직 + 대시보드
- WeeklyReport.gs 약 15KB — 슬랙 자동 리포트
- 관리 지표 자동 갱신 시트 5종 (체크인목록·주간_입력·미납현황·대시보드·블랙리스트)

## 주요 함수

**메인 파이프라인**
- `syncUnpaidData()` — 진입점. 대화상자 → CSV 로드 → 그룹화 → 각 탭 갱신
- `updateMonthlySheets()` — 월별 탭 적재 + 완납·강제해지 상태 자동 처리
- `recordBlacklist()` — 강제해지 확정 → 완료 상태, 예정→완료 자동 승격
- `updateForceTerminationList()` — 3개월+ 예정 upsert, 해소 시 자동 제거
- `updateDashboardSheet()` — 귀속월별 집계 + 고객 구간별 통계
- `getCustomerSegment()` — 결제일·계약일 기반 초기/중기/장기 분류

**웹앱**
- `doGet()` — Chart.js 대시보드 JSON payload 생성
- `openDashboard()` — 시트 메뉴에서 웹앱 URL 열기
- `getHtmlTemplate()` — Chart.js + Table UI

**고객 조회**
- `searchCustomer()` — 사업자번호·대표자명 검색 → 현재 미납·월별 이력·블랙리스트 여부

**슬랙 리포트 (WeeklyReport.gs)**
- `weeklyReport_run()` — 시간 트리거로 매주 화요일 17시 실행
- `weeklyReport_test()` — 발송 없이 미리보기 (로그 출력)
- `weeklyReport_buildMessage_()` — 시트 → mrkdwn 텍스트
- 스냅샷 비교로 **신규 미납**만 산출 (Script Property에 저장)

## 보안 설계 (ISMS 준수)

원본 코드에는 회사 특정 사업자번호 5건과 담당자 실명이 하드코딩되어 있었습니다. 이 저장소에는 다음과 같이 마스킹되었습니다:

- **제외 사업자번호 5건**: 실제 값 삭제, 예시 주석으로 대체 (`EXCLUDE_BIZ_NO`)
- **담당자 실명 2건**: `EXCLUDED_MANAGER_A`, `EXCLUDED_MANAGER_B`로 치환
- **회사 스크립트 도메인**: `DASHBOARD_URL`을 Script Properties로 이동
- **슬랙 채널명**: `#internal-care-ops`로 일반화
- 웹훅 URL: 원본에서도 Script Properties 관리 중

**추가 권장**:
- `EXCLUDE_BIZ_NO`·`EXCLUDED_MANAGERS`는 시트로 이동 (코드 배포 없이 관리 가능)
- Chart.js 웹앱 접근은 GAS Web App의 `Execute as owner + Anyone with link` 조합으로 제한

## 결과

- **7월 미납 회수율 86.23% 달성**의 데이터 인프라
- 주간 관리 시간: 수작업 몇 시간 → 클릭 한 번
- 강제해지 대응 리드타임 단축 (예정 자동 감지 → 팀 즉시 인지)
- Metabase 5+ 뷰와 결합해 CS 우선순위 판단 근거 제공

## 스크린샷

- [ ] 미납관리 메뉴 UI (시트 메뉴바)
- [ ] 웹앱 Chart.js 대시보드 (월별 미납 추이 + 회수율 라인)
- [ ] 고객 구간별 상세 테이블
- [ ] 슬랙 주간 리포트 발송 예시 (마스킹)
- [ ] 블랙리스트 예정/완료 자동 승격 결과

## 배포

1. Google Sheets에서 `확장 프로그램 → Apps Script`
2. `Code.gs`, `WeeklyReport.gs` 붙여넣기
3. Script Properties 설정
   - `DASHBOARD_URL`: 웹앱 배포 URL (배포 후 획득)
   - `WEEKLY_REPORT_WEBHOOK_URL`: 슬랙 webhook URL
4. `배포 → 새 배포 → 웹 앱` (Execute as: Me / Access: Anyone with link)
5. 트리거 등록: `weeklyReport_run` — 시간 기반, 주 단위, 화요일 17시

## 관련 프로젝트

- 케어 F그룹 리텐션 대시보드 (Metabase·GAS) — 인터뷰 100건 기반 이탈위험 세그먼트 재정의
