# Channel Talk VOC Collector

Channel Talk Open API로 대화 전문을 안정 수집하는 GAS 도구. 페이지네이션·Resume 기능으로 대용량 수집을 실패 없이 이어갈 수 있습니다.

## 문제
CS 대화가 채널톡에 쌓이는 속도가 팀 분석 속도를 앞질렀습니다. 해지·클레임 사유 파악이 담당자 감에만 의존하는 상태.

## 기능
- Open API 대화 후보 페이지네이션 수집
- 대화 전문 export (발화자·타임스탬프 포함)
- Job State 관리로 중단·재개 지원
- LockService로 동시 실행 방지
- API rate limit 대응 (지수 백오프 재시도)
- Access Key/Secret 안전 저장 (UI 프롬프트 → Properties)

## 스택
Google Apps Script · Channel Talk Open API · Google Sheets · Time-based Triggers

## 파일
- `code.gs` — 서버 로직 (1,354줄)
- `appsscript.json` — 매니페스트

## 보안
- Access Key/Secret: Script Properties 관리
- 원본 대화의 개인정보는 팀 내부 시트에만 저장
