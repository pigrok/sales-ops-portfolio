# Sales Ops Portfolio — Seongrok Jo (pigrok)

**B2B SaaS 세일즈 · Sales Ops · AX (AI Transformation)** 지원용 포트폴리오

세무 SaaS 조직에서 수행한 세일즈 자동화·데이터 분석·CRM 파이프라인 프로젝트를 담았습니다. 코드·설계·결과 지표까지 실물로 확인할 수 있습니다.

---

## Contact

- **Name**: 조성록 (Seongrok Jo)
- **Email**: jsr4613@naver.com
- **GitHub**: [pigrok](https://github.com/pigrok)
- **Location**: Seoul

---

## Highlights (2024.08 – 재직중)

| 지표 | 수치 | 근거 |
|------|------|------|
| 케어 구독 신규 성사 | **934건** (팀 1위, 2위 661건) | Pipedrive 담당자별 성사 사업장 수 |
| 환급 신청전환 성사 | **385건** · **₩53.4억** | Pipedrive(24.08~25.03) + Salesforce(26.06~) 합산 |
| 환급 취소방어 | **45건** · **₩7.5억** | 동일 합산 |
| Zapier Zap 운영 | **98건 무중단 마이그레이션**, 오류 대응 3,500+ 미처리 0 | REFOP-270 |
| 신청→결제 퍼널 분석 | **1,008건 / ₩216억** (단독) | Python 전수 분석 |
| Channel Talk 해지 전수 분석 | **2,829건** (단독) | Open API + GAS |
| 콜 리소스 실측 | 5인 × 1,600 녹취, **월 562시간** 정량화 | 자체 측정 |

## 서류 자산

- 이력서 · 경력기술서 (준비 중)
- 위 지표의 원본 근거는 회사 CRM(Pipedrive · Salesforce) 대시보드로, 회사 자산이라 공개는 어렵습니다. 이미지로 첨부드립니다.

---

## 대표 프로젝트 5개

### 1. [CS Operation Tool](./projects/cs-operation-tool)
**GAS + Whisper + LLM · 단독 개발**

콜 후처리 시간을 줄이는 사내 도구. 통화 녹취 → 자동 전사 → 화자 분리 → CRM 활동 자동 등록.
콜 1건당 사후 처리 **3분 30초 → 자동화**. 팀 배포 후 상시 사용 중.

### 2. [Refund Payment Funnel Analysis (Python)](./projects/refund-payment-funnel)
**Python + Pipedrive · 단독**

6개월 신청→결제 전수 **1,008건 / ₩216억** 분석. 담당자 개입이 D+3 결제율 **+7.2%p**, 성사율 **+4.0%p**, 추심 **−4.0%p**, 결제 경과일 **−1.1일** 개선 실증. → 콜 리소스 배분 정책의 실측 근거로 조직 채택.

### 3. [Channel Talk VOC Collector](./projects/channeltalk-voc)
**GAS + Open API · 단독**

Channel Talk 대화 전문을 페이지네이션·Resume 기능으로 안정 수집. 이 도구를 기반으로 **해지요청 2,829건 전수 분석** → 클레임성 이탈 **29%** 규명, 담당자 80명별 클레임률 도출.

### 4. [n8n Self-hosted Workflows](./n8n-workflows)
**n8n · self-hosted · 5개 프로젝트 · 9개 핵심 워크플로**

이탈 신호 자동 감지 파이프라인, Channel Talk → CRM Sync, AlimTalk Ops (Solapi + Claude로 초안 자동 생성) 등. 원본 JSON은 ISMS 준수를 위해 미포함, 아키텍처 다이어그램으로 대체.

### 5. [Zapier Error Dashboard](./projects/zapier-error-dashboard)
**GAS + LLM + Chrome Extension · 단독**

Zap 98건 운영 중 오류 로그를 자동 수집 → LLM 유형 분류 → 우선순위별 슬랙 노티. **누적 3,500+ 건 처리 · 미처리 0**. Pipedrive API v1→v2 마이그레이션 98건 무중단 전환. Code.gs 98KB · 2,375줄.

---

## 관련 프로젝트

### 6. [Pipedrive Filter Dashboard](./projects/pipedrive-filter-dashboard)
Pipedrive 필터별 딜 현황 + Solapi 발송 상태 원클릭 재발송.

### 7. [Refund KPI Calculator](./projects/apps-script-refund)
대시보드 이미지 → Drive OCR로 KPI 자동 산출 · 슬랙 보고문 자동 생성.

---

## Skills & Stack

**Sales / CRM**
- Pipedrive (Admin · API v1/v2) · Salesforce Sales Cloud
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
│   ├── cs-operation-tool/       ← GAS + Whisper + LLM
│   ├── refund-payment-funnel/   ← Python 분석
│   ├── channeltalk-voc/         ← Channel Talk API 수집
│   ├── zapier-error-dashboard/  ← Zapier 오류 대시보드
│   ├── pipedrive-filter-dashboard/
│   └── apps-script-refund/      ← KPI 자동 계산
├── n8n-workflows/               ← n8n 아키텍처 문서
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
