# 마스킹 정책

이 저장소의 모든 코드·문서는 아래 원칙에 따라 회사 특정 정보를 일반화된 용어로 치환했습니다.

## 원칙

- 공개 서비스명(Pipedrive, Channel Talk, Slack, Salesforce 등)은 그대로 노출
- 회사 특정 서비스 브랜드는 기능 중심 일반명으로 치환
- 회사 내부 프로세스 명·팀명·발신 번호는 마스킹
- API 키·토큰·시크릿은 원본에서도 이미 Script Properties로 관리

## 사용된 치환 패턴 (예시)

- 세무 환급 서비스 → `Tax Refund SaaS`
- 세무 기장 구독 서비스 → `Bookkeeping Subscription SaaS`
- 재컨택 캠페인 → `Refund Recontact Campaign`
- 국세청 시스템 → `National Tax System`
- CS 담당자 실명 → `AGENT_NAME`
- 전화번호 → `010-XXXX-XXXX`
- 개인 이메일 → `user@example.com`

## 그대로 유지되는 정보

- 공개 서비스명: Pipedrive · Salesforce · Slack · Channel Talk · Solapi
- 오픈소스·SaaS 도구: Google Apps Script · n8n · Zapier · Whisper · Anthropic Claude
- 저자 본인 정보: 조성록 · GitHub `pigrok` · 이메일 `jsr4613@naver.com`

## 스크린샷 캡처 원칙

**반드시 마스킹**:
- 고객사명·사업자등록번호·전화번호·주소
- Pipedrive 딜 URL의 딜 ID
- Slack 채널명·타팀원 프로필 사진
- API 키·토큰·엔드포인트 URL
- 팀원 실명 (본인 제외)

**노출 OK**:
- 지표 숫자·비율·집계
- 프로젝트명·워크플로명 (마스킹된 형태)
- 본인 명의·직책

## 파일별 정책

- **GAS 코드 (.gs, .html)**: 도메인 용어 마스킹 완료. API 키는 원본부터 Script Properties 사용
- **Python 스크립트 (.py)**: 파일 경로만 남기고 실제 데이터는 로컬 전용
- **n8n 워크플로 JSON**: 저장소 미포함 (개인정보·회사 IP 리스크). 아키텍처 다이어그램으로 대체
- **원본 데이터 (CSV·PDF export)**: 저장소 미포함. 방법론·집계 결과만 서술

## 재사용·라이선스

이 저장소의 코드는 개인 학습·구현물로, 재사용을 위한 라이선스는 별도 문의 부탁드립니다.
