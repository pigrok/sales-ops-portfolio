# CS Operation Tool

콜 후처리 시간을 줄이는 사내 도구. **통화 녹취 → 자동 전사 → 화자 분리 → CRM 활동 자동 등록**을 한 번에 처리합니다.

## 문제

세일즈 조직에서 콜 이후 CRM 기록 작업이 병목이었습니다. 5인 담당자를 대상으로 실측한 결과 콜 1건당 사후 처리에 **3분 30초**가 소요됐고, 이게 월 **562시간(약 70.3 인일)** 리소스를 잡아먹고 있었습니다. 담당자가 콜에 집중할 시간을 뺏기는 구조.

## 해결

Whisper API로 녹취를 전사하고, LLM으로 화자를 분리한 뒤, Pipedrive Activity API로 자동 등록하는 웹 앱을 GAS로 구축했습니다. 담당자는 녹음 파일을 선택하고 담당 CRM만 지정하면 되고, 나머지는 자동.

## 아키텍처

```
Google Drive (녹취)
      ↓  파일 선택
GAS Web App (사용자 인증·권한 확인)
      ↓
Whisper API (Groq)  ←  transcribeAudioBlob_
      ↓  텍스트
LLM 화자 분리
      ↓
Pipedrive Activity API  ←  createCallActivity
  · 담당자 배정
  · 활동 유형 매핑
  · 노트 첨부
```

핵심 함수 흐름:
- `doGet()` → 웹 UI 서빙 (`Index.html`)
- `getInitData()` → 사용자 권한·담당자 목록 초기 로드
- `listDriveRecordings()` → Drive에서 담당자별 녹취 폴더 조회
- `transcribeCallFromDrive()` → 파일 → Whisper 전사
- `createCallActivity()` → Pipedrive Activity 자동 등록
- `resolvePipedriveToken_()` → 다중 CRM 계정 토큰 분기

## 스택

Google Apps Script (Web App) · Google Drive API · Whisper (Groq) · LLM · Pipedrive API v2

## 보안 설계

- API 키·토큰: 모두 `PropertiesService.getScriptProperties()`로 관리
- 사용자 인증: `USERS` 시트에 등록된 이메일만 접근 허용 (`requireAllowedUser_`)
- 담당자별 CRM 계정 분기 (`resolvePipedriveToken_`) — 다중 CRM 환경 대응
- 실행 로그: 날짜·사용자·통화 정보를 시트에 기록 (감사 대응)

## 결과

- 콜 후 기록 업무 자동화 → 담당자가 콜에 더 오래 집중할 수 있는 구조
- 팀 전체 배포 후 상시 사용 중
- CRM 기록 품질도 개선 (수기 입력 시 누락되던 필드가 자동 채워짐)

## 코드 하이라이트

```javascript
// 다중 Pipedrive 계정 토큰 분기 — 담당 서비스별로 다른 CRM
function resolvePipedriveToken_(targetSystem) {
  const mapJson = getProp_("TARGET_PIPEDRIVE_TOKENS_JSON", "{}");
  const mapObj = JSON.parse(mapJson);
  const key = String(targetSystem || "").trim();
  const token = String(mapObj[key] || "").trim();
  if (!token) throw new Error("TARGET_PIPEDRIVE_TOKENS_JSON에 대상 토큰이 없습니다. target=" + key);
  return token;
}

// Whisper 호출 (multipart 전송)
const apiKey = getProp_("GROQ_API_KEY", "");
const res = multipartFetch_(url, apiKey, fields, audioBlob);
```

## 스크린샷

- [ ] 웹 앱 UI (녹취 파일 선택 화면)
- [ ] 자동 전사 결과 화면
- [ ] Pipedrive에 자동 등록된 Activity 예시 (마스킹 후)
- [ ] 팀 배포 안내 Slack 공지

## 배포

Google Apps Script Web App으로 배포. 사용자 인증은 `USERS` 시트, 담당자 매핑은 `ASSIGNEES` 시트로 관리합니다.

상세 배포 가이드는 아래 참조.

---

## 배포 가이드 (원본)

### 스프레드시트 시트 구성

**`USERS` 시트**
```
email | name | phone | enabled | assignee_user_id | recordings_folder_id
```

**`ASSIGNEES` 시트**
```
user_id | name | target_key | corp_phone
```

- `target_key`: CRM 계정 구분 키 (예: `CARE`)
- `recordings_folder_id` (선택): 담당자별 통화녹음 Google Drive 폴더 ID. 비우면 자동 탐색

### Script Properties 설정

```
GROQ_API_KEY = <Groq API 키>
TARGET_PIPEDRIVE_TOKENS_JSON = {"CARE":"<token>","CORP":"<token>"}
```

### 배포 순서

1. `clasp push`로 코드 배포
2. `배포 > 새 배포 > 유형: 웹 앱`
3. 실행 권한: `나(배포자)`
4. 액세스: `Google 계정을 가진 모든 사용자` (내부에서 `requireAllowedUser_`가 필터링)
