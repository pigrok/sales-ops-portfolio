# AlimTalk Template Manager

카카오 알림톡 템플릿의 **초안 생성 → 검수 요청 → 반려 분석 → 재작성**을 원스톱으로 처리하는 사내 웹 도구. Solapi API와 LLM(Claude·Groq)을 조합해 심사 사이클의 인력 부담을 크게 줄입니다.

## 문제
카카오 알림톡 템플릿은 반려율이 높고 사유 분석·재작성에 시간이 오래 걸립니다. 시즌 개정마다 사업팀 요청부터 등록 완료까지 리드타임이 병목이었습니다.

## 기능
- **알림톡 작성 탭** — 키워드 입력 → LLM 초안 자동 생성 (카카오 가이드라인 사전 체크 포함)
- **반려 분석 탭** — 카카오 반려 사유 파싱 → 개정 방향 자동 제안
- **검수 요청 탭** — Solapi API로 심사 요청 자동화
- 카카오톡 실제 UI 미리보기 (Tailwind CSS)
- 승인 완료 템플릿 재활용 로직 (reusedTemplateId)

## 스택
Web UI (HTML + Tailwind CSS + Vanilla JS) · n8n Webhook (백엔드) · Solapi Kakao Templates API · Anthropic Claude · Groq LLM

## 파일
- `index.html` — 3탭 웹 UI (745줄)
- `.env.example` — 필요 환경 변수 (`SOLAPI_API_KEY` · `SOLAPI_API_SECRET` · `KAKAO_CHANNEL_ID` · `GROQ_API_KEY`)
- n8n 워크플로 3종은 [../../n8n-workflows/](../../n8n-workflows/README.md) 참조 (ISMS 정책상 JSON 원본 미포함)

## 사용 흐름
1. 팀원이 브라우저에서 웹훅 URL 접속 (한 번만 설정)
2. 키워드 입력 → LLM이 승인 템플릿 200+ 참고해 초안 생성
3. 카카오톡 미리보기 확인 → 검수 요청
4. 반려 시 반려 분석 탭에서 사유 파싱 → 개정 → 재요청

## 보안
- API 키(Solapi · Groq · Kakao Channel ID)는 n8n Credential Store 관리
- 카카오 Channel PFID는 원본 코드에서 마스킹 (`KA01PFXXXXXXXXXXXXXXXXXX`)
- 웹훅 URL은 클라이언트 localStorage에만 저장, 서버 전송 없음
