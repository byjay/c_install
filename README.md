# HK2401 Cable List · SEcMS 통합 빌드

> Shipboard Electrical cable Management System (SEcMS) · HK2401 호선 케이블 포설 / 결선 / 호선 전화 통합 관리.
> 단독 빌드(Vanilla JS)로 별도 백엔드 없이 동작하며, 추후 [secms.tech](https://secms.tech) 본 시스템과 일체화될 예정입니다.

---

## 핵심 기능

| 영역 | 내용 |
|---|---|
| **Cable List** | Excel 양식 그대로의 편집 그리드, 필터/슬라이서, REV 자동 누적 |
| **자동 매핑** | 어떤 형식의 .xlsx / .xls / .csv / .tsv 가 들어와도 표준 스키마로 자동 변환 (`column-mapper.js`) |
| **다중셀 선택** | Excel-like 드래그 / Shift / Ctrl 클릭 / Ctrl+A / Ctrl+C 클립보드 복사 (`excel-select.js`) |
| **작업 이력** | 모든 변경(가져오기·셀 수정·일괄 변경·호선 추가/이동·등록·삭제·복원·내보내기·로그인)을 시간순으로 영구 보존, TSV/JSONL 내보내기 (`audit-log.js`) |
| **대시보드** | KPI · TYPE/NODE 차트 · 도넛 그래프 |
| **실적보고** | 월간 / 주차별 인쇄용 A4 시트 |
| **달력 뷰** | 포설/결선 일자 캘린더, 특정 날짜 더블클릭 상세 |
| **호선관리** | 다중 호선 등록 / 전환 / 케이블 일괄 입력 |

---

## 기술 스택

- HTML5 / CSS3 / Vanilla JavaScript (ES2020+)
- [SheetJS xlsx](https://github.com/SheetJS/sheetjs) (CDN) — Excel 파일 파싱
- 데이터 저장: `localStorage` (오프라인 동작, 클라이언트 단독)
- 배포: Cloudflare Pages

---

## 폴더 구조

```
hk2401-cable/
├── index.html                # 진입점
├── assets/
│   └── logo.png              # SEcMS 로고
├── css/
│   ├── main.css              # 본 디자인 (Excel 스타일 그리드 / 네이비 / 옐로 액센트)
│   └── enhance.css           # 다중셀 선택 + 작업 이력 모달 스타일
├── data/
│   └── hk2401-cables.tsv     # 기본 Cable List (HK2401 호선)
├── js/
│   ├── store.js              # 상태/저장소 (CABLE_FIELDS, normalizeCable, importCableRows, undo)
│   ├── main.js               # UI 라우팅 / 렌더 / 핸들러
│   ├── column-mapper.js      # 강화된 헤더 자동 매핑 (한국어/영문/약어 통합 사전)
│   ├── audit-log.js          # 작업 이력 영구 저장소 (subscribe / list / export)
│   ├── excel-select.js       # Excel-like 다중셀 선택 + Ctrl+C TSV 복사
│   ├── enhance.js            # 비침투적 통합 레이어 (헤더 버튼 inject / 모달 viewer)
│   └── i18n.js               # (예약) 다국어
├── README.md
├── CHANGELOG.md
└── .gitignore
```

---

## 자동 매핑 동작 요약

`column-mapper.js`는 다음과 같은 다양한 헤더 표기를 모두 같은 표준 필드로 인식합니다.

| 표준 필드 | 인식되는 헤더 (일부) |
|---|---|
| `circuitNo` | CIRCUIT NO. / 회로번호 / CKT NO / 회로 / CIRCUIT |
| `cableType` | CABLE TYPE / 케이블종류 / TYPE / 타입 |
| `installDate` | 포설일자 / 포설 DATE / INSTALL DATE / 포설완료일 / INSTALLED |
| `conFromDate` | 결선 FROM / FROM 결선 / CONNECTION FROM / 결선시작 |
| `conToDate` | 결선 TO / TO 결선 / CONNECTION TO / 결선완료 |

→ **공백 / 줄바꿈 / 슬래시 / 따옴표 / 점**은 무시되며, 부분 일치도 허용 (예: `포설일자_2026` → `installDate`).

---

## 다중셀 선택 단축키

| 동작 | 키/제스처 |
|---|---|
| 단일 셀 편집 진입 | 클릭 |
| 사각형 영역 선택 | 마우스 드래그 |
| 범위 확장 | **Shift** + 클릭 |
| 비연속 셀 추가/제거 | **Ctrl/Cmd** + 클릭 |
| 화면 전체 셀 선택 | **Ctrl/Cmd + A** |
| Excel 형식으로 클립보드 복사 | **Ctrl/Cmd + C** (TSV + HTML 둘 다 제공 → Excel 붙여넣기 호환) |
| 선택 해제 | **Esc** |

---

## 작업 이력 (Audit Log)

상단 우측 **📜 버튼** → 모달 → 모든 변경 이력을 시간순으로 표시합니다.

- 카테고리별 필터 (가져오기 / 셀 수정 / 일괄 수정 / 케이블 추가/삭제/복원 / 호선 등록·이동 / 내보내기 / 시스템)
- 키워드 검색 (action + details JSON 안 검색)
- **TSV** (Excel 호환, UTF-8 BOM) / **JSONL** (외부 분석 도구) 내보내기
- 최대 5,000건 보존 (초과 시 오래된 항목부터 자동 정리)

저장소: `localStorage["hk2401-cable-system:audit-log:v1"]`

---

## 로컬 개발

별도 빌드 단계 없음. 정적 파일이므로 어떤 정적 서버로도 동작합니다.

```bash
# Python 표준 라이브러리
python -m http.server 5500

# 또는 npx serve (Node)
npx serve . -p 5500
```

브라우저에서 <http://localhost:5500/> 접속.

---

## 배포 (Cloudflare Pages)

이 저장소는 GitHub → Cloudflare Pages 자동 배포로 연결되어 있습니다.

- **Production 브랜치**: `main`
- **빌드 커맨드**: 없음 (정적)
- **Build output**: 루트 (`/`)
- **퍼블릭 URL**: <https://hk2401-cable.pages.dev>

`main`에 push 하면 1~2분 내 자동 빌드 / 배포됩니다.

---

## 라이선스 / 문의

- © 2026 SEcMS · Shipboard Electrical cable Management System
- 개발: kbj
- 문의: <admin@secms.tech>
- 본 시스템과의 통합: <https://secms.tech>
