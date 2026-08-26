# Channel Talk VOC Collector

Channel Talk Open API로 대화 전문을 수집해 VOC 분석 시트를 자동 생성하는 GAS 도구. **해지요청 2,829건 전수 분석**의 기반 데이터 수집기.

## 문제

CS 대화가 채널톡에 쌓이는 속도가 팀 분석 속도를 앞질렀습니다. 클레임·해지 사유 파악이 담당자 감각에만 의존하는 상태였고, 데이터 기반 개선 우선순위를 못 잡고 있었습니다.

## 해결

Channel Talk Open API로 지정 기간의 대화 후보 목록을 페이지네이션으로 수집하고, 각 대화의 전문을 export해 시트에 정리. **Resume 기능**으로 대용량 수집 중 실패해도 이어서 진행 가능.

## 아키텍처

```
Channel Talk Open API
      ↓  Access Key + Secret 인증
GAS: fetchChatCandidates (기간별 대화 후보)
      ↓
Job State (page cursor, failed count)  ←  saveJobProgress
      ↓  Time-based Trigger로 재개
runChannelTalkExportResume
      ↓
Google Sheets (대화 전문 · 발화자 구분)
      ↓  다운스트림
Slack 해지 워크플로 시트 결합 → 해지 사유 14 카테고리 분류
```

## 스택

Google Apps Script · Channel Talk Open API · Google Sheets · Time-based Triggers

## 규모

- code.gs 약 **41KB / 1,354줄**
- **2,829건 대화 전수 분석** 실적 (해지요청 프로젝트)

## 주요 함수

- `setChannelTalkKeys()` — Access Key/Secret 안전 저장 (UI 프롬프트 → Properties)
- `runChannelTalkExport()` — 신규 수집 시작
- `runChannelTalkExportResume()` — 중단된 작업 이어서 진행
- `runChannelTalkExportResumeLocked()` — LockService로 동시 실행 방지
- `fetchChatCandidates()` — 페이지네이션 대화 후보 조회
- `initJob()` / `saveJobProgress()` / `getJobState()` — Job 상태 관리
- `scheduleResumeTrigger()` — 자동 재개 트리거 예약
- `cancelChannelTalkJob()` — 중단 처리
- `buildDoneMessage()` — 완료 알림 메시지 생성

## 보안 설계

- **인증**: Access Key + Access Secret UI 프롬프트로 입력받아 Script Properties에만 저장
- **API rate limit 대응**: 지수 백오프 재시도, LockService로 동시 실행 방지
- **Job State 격리**: 중단·재개 상태를 별도 시트로 관리해 원본 데이터 오염 방지
- **개인정보**: 원본 대화에 포함된 고객 정보는 팀 내부 시트에만 저장, 외부 export 시 마스킹

## 결과

- **해지요청 2,829건 전수 분석** 완료
- 클레임성 이탈 **29%** 규명 (819건)
- 사유 **14개 카테고리** 자동 분류
- 담당자 **80명별 클레임률** 도출
- 구조적 개선 과제 **5건 제안** → 경영진 보고

## 파생 프로젝트

이 수집기가 다음 분석·자동화의 데이터 소스가 됐습니다:
- **해지요청 전수 분석 리포트** (Confluence `[Z21-RPT]20260619`)
- 발화자 분리·의도 분류 자동화
- 담당자별 클레임률 대시보드

## 스크린샷

- [ ] Channel Talk 개발자 콘솔 (API 사용 이력)
- [ ] 수집된 대화 시트 예시 (마스킹)
- [ ] Slack 해지 워크플로 시트 결합 결과 (2,829건 노출)
- [ ] 사유 14개 카테고리 분포 차트
- [ ] Confluence 리포트 상단 (본인 명의 노출)
