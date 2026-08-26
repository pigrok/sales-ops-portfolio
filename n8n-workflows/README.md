# n8n Self-hosted Workflows

세무 SaaS 조직에서 Zapier로 처리하기 어려운 조건 분기·실시간 이벤트 자동화를 위해 n8n self-hosted 서버를 직접 운영했습니다.

**5개 프로젝트 · 9개 핵심 워크플로 · 15종 계정 통합 운영**

## 왜 self-hosted인가

| 요구 | Zapier | n8n self-hosted |
|------|--------|-----------------|
| 복잡한 조건 분기 | 제약 있음 | 자유 |
| Task 원가 | 사용량 기반 | 서버 원가만 |
| 데이터 잔류 | Zapier 서버 | 자사 서버 |
| Webhook 검증 (HMAC) | 제약 있음 | 완전 커스터마이징 |

높은 처리량과 개인정보 잔류 최소화가 필요해 n8n을 선택했습니다.

## 대표 워크플로

### 1. Contract Termination Detection Pipeline
국세청 시스템 수임해제 이벤트 → HMAC 검증 → 중복 제거 → 담당자 라운드로빈 배정 → Pipedrive 딜 생성 → 슬랙 알림.

### 2. Channel Talk → CRM Sync
채널톡 이벤트 → LLM 의도 분류 (예약/해지/일반) → Pipedrive Activity 또는 해지 파이프라인으로 분기.

### 3. AlimTalk Template Ops
Solapi 알림톡 템플릿 200+ 관리 + Anthropic Claude로 초안 자동 생성 + 반려 사유 파싱·자동 개정.

### 4. HR Notification Automation
Google Calendar 휴가 이벤트 → 최초 공지 + 전날 오전 10시 리마인더 자동 발송.

## 보안 설계 (공통)
- Cloudflare Access + 이메일 도메인 접근 제한
- 웹훅 서명 헤더(HMAC) 필수
- 크레덴셜은 n8n Credential Store, export 시 자동 제외
- 페이로드 원본 24시간 잔류, 이후 요약만 남김

## ISMS 준수 안내

원본 워크플로 JSON은 개인정보(발신 번호)와 회사 마케팅 자산(카카오 알림톡 원문)을 포함하므로 이 저장소에 포함하지 않습니다. 아키텍처 설계 감각을 확인하기 위한 요약만 서술되어 있으며, 면접 시 화면 공유로 실물 시연 가능합니다.
