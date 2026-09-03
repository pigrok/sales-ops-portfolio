# Sales Ops Portfolio — Seongrok Jo (pigrok)

**B2B SaaS 세일즈 성과와 이를 뒷받침한 Sales Ops 자동화·데이터 분석 포트폴리오**

세무 SaaS 조직에서 수행한 세일즈 자동화·데이터 분석·CRM 파이프라인 프로젝트를 담았습니다. 코드·설계·결과 지표까지 실물로 확인할 수 있습니다.

> 이 포트폴리오는 개발 역량 자체보다, 세일즈 현장에서 발생한 반복 업무·후속 관리·고객 VOC 문제를 발견하고 자동화·데이터 분석으로 개선한 과정을 보여주기 위해 정리했습니다.

---

## Contact

- **Name**: 조성록 (Seongrok Jo)
- **Email**: jsr4613@naver.com
- **GitHub**: [pigrok](https://github.com/pigrok)
- **Location**: Seoul

---

## 대표 프로젝트 7개

### 1. [CS Operation Tool](./projects/cs-operation-tool)
**GAS + Whisper + LLM · 단독 개발**

콜 후처리 시간을 줄이는 사내 도구. 통화 녹취 → 자동 전사 → 화자 분리 → CRM 활동 자동 등록.
콜 1건당 사후 처리 **3분 30초 → 자동화**. 팀 배포 후 상시 사용 중.

### 2. [Refund Payment Funnel Analysis (Python)](./projects/refund-payment-funnel)
**Python + Pipedrive · 단독**

6개월 신청→결제 전수 **1,008건 / ₩216억** 분석. 담당자 개입이 D+3 결제율 **+7.2%p**, 성사율 **+4.0%p**, 추심 **−4.0%p**, 결제 경과일 **−1.1일** 개선 실증. → 콜 리소스 배분 정책의 실측 근거로 조직 채택.

### 3. [CX 케어·환급 통합 대시보드](./projects/cx-integrated-dashboard)
**GAS + Open API + Chart.js · 단독** · (구 Channel Talk VOC Dashboard 전면 개편)

환급·케어 두 채널을 한 웹앱에서 수집·분석. 태그·인입경로·조직 성과(진입→종결 히트맵)·해지 태그 vs 실제 해지 매칭·ALF 인입·CSAT까지 통합. 월간 **환급 2,548 / 케어 3,781 대화** 처리 · 매니저-팀 매핑 **56명·11팀** UI 관리 · 자동 이어서 수집 + 스냅샷 시스템.

### 4. [n8n Self-hosted Workflows](./n8n-workflows)
**n8n · self-hosted · 5개 프로젝트 · 9개 핵심 워크플로**

이탈 신호 자동 감지 파이프라인, Channel Talk → CRM Sync, AlimTalk Ops (Solapi + Claude로 초안 자동 생성) 등. 원본 JSON은 ISMS 준수를 위해 미포함, 아키텍처 다이어그램으로 대체.

### 5. [Zapier Error Dashboard](./projects/zapier-error-dashboard)
**GAS + LLM + Chrome Extension · 단독**

Zap 98건 운영 중 오류 로그를 자동 수집 → LLM 유형 분류 → 우선순위별 슬랙 노티. **누적 3,500+ 건 처리 · 미처리 0**. Pipedrive API v1→v2 마이그레이션 98건 무중단 전환. Code.gs 98KB · 2,375줄.

### 6. [Care Unpaid Dashboard](./projects/care-unpaid-dashboard)
**GAS + Chart.js + Slack · 단독**

주간 미납 CSV → 사업자번호 그룹화 → 월별 탭 자동 적재 → 강제해지 자동 감지 → Chart.js 대시보드 → 슬랙 주간 리포트. **7월 미납 회수율 86.23%** 달성의 데이터 인프라.

### 7. [AlimTalk Template Manager](./projects/alimtalk-template-manager)
**Solapi + Claude + n8n · 단독**

카카오 알림톡 심사 사이클(초안 → 검수 요청 → 반려 → 재작성) 원스톱 웹 도구. LLM 초안 자동 생성, 카카오 미리보기 UI, 반려 분석 자동화.

---

## 관련 프로젝트

### 8. [Pipedrive Filter Dashboard](./projects/pipedrive-filter-dashboard)
Pipedrive 필터별 딜 현황 + Solapi 발송 상태 원클릭 재발송.

### 9. [Refund KPI Calculator](./projects/apps-script-refund)
대시보드 이미지 → Drive OCR로 KPI 자동 산출 · 슬랙 보고문 자동 생성 (참고용, 지원 문서에는 미포함).

---

## Skills & Stack

**Sales / CRM**
- Pipedrive (Admin · API v1/v2)
- **Salesforce Sales Cloud** (Opportunity · Task · Reports · Dashboards — 현 파이프라인 운영 중)
- 채널톡 (상담 · Open API · 웹훅)
- Outbound Sales · Lead Qualification · Pipeline Management

**Automation / Ops**
- Google Apps Script · n8n (self-hosted) · Zapier · Chrome Extension
- Slack Webhook · Solapi · Channel Talk (Open API · 웹훅)

**Data / Analysis**
- Python (Pandas · ReportLab) · SQL (Metabase) · Google Sheets

**AI Integration**
- Anthropic Claude · Whisper (STT) · LLM 화자 분리 · 오류 자동 분류
- Claude Code (개발 워크플로)

**Collaboration**
- Jira · Confluence · Notion · Slack

---

## 이 저장소 구조

```
sales-ops-portfolio/
├── README.md                    ← 여기
├── docs/
│   └── MASKING_POLICY.md        ← 마스킹 원칙
├── projects/
│   ├── cs-operation-tool/           ← GAS + Whisper + LLM
│   ├── refund-payment-funnel/       ← Python 분석
│   ├── cx-integrated-dashboard/     ← CX 케어·환급 통합 대시보드 (구 Channel Talk VOC)
│   ├── zapier-error-dashboard/      ← Zapier 오류 대시보드
│   ├── care-unpaid-dashboard/       ← 미납 관리 (Chart.js + Slack)
│   ├── alimtalk-template-manager/   ← Solapi + Claude 심사 사이클
│   ├── pipedrive-filter-dashboard/
│   └── apps-script-refund/          ← KPI 자동 계산 (참고용)
├── n8n-workflows/                   ← n8n 아키텍처 문서
└── assets/
    └── screenshots/             ← 배포 증빙 스크린샷
```

## 보안·마스킹 안내

이 저장소의 코드는 회사 서비스 브랜드·도메인 용어를 일반화된 표현으로 치환했습니다. 상세 원칙은 [MASKING_POLICY](./docs/MASKING_POLICY.md) 참조.

- 모든 API 키·토큰은 원본 코드에서 이미 `PropertiesService.getScriptProperties()`로 관리되고 있어 하드코딩 노출이 없습니다
- n8n 워크플로 원본 JSON은 개인정보(발신 번호)·회사 IP(마케팅 문구) 포함으로 저장소 미포함 — 다이어그램으로 대체
- 원본 데이터(Pipedrive export 등)는 회사 자산이므로 저장소 미포함 — 방법론·집계 결과만 서술

## License

이 저장소의 코드는 조성록 개인이 작성한 자동화 도구로, 회사 자산이 아닌 개인 학습·구현물입니다. 회사 도메인 로직·마케팅 자산은 마스킹되어 있으며, 재사용을 위한 라이선스는 별도 문의 부탁드립니다.
