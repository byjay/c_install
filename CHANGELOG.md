# CHANGELOG

본 파일은 HK2401 Cable List 빌드의 모든 의미 있는 변경을 기록합니다.
형식은 [Keep a Changelog](https://keepachangelog.com/ko/1.1.0/)을 따르며,
버전 표기는 [Semantic Versioning](https://semver.org/lang/ko/)을 사용합니다.

## [Unreleased]

## [2.0.0] - 2026-04-29

### 신규 (Added)
- **`column-mapper.js`** · 강화된 헤더 자동 매핑 사전 추가.
  - 한국어 / 영문 / 약어 / 부분일치 모두 지원하는 ALIASES 사전 (12개 표준 필드 × 평균 6개 별칭).
  - `detectHeaderRow()` — 상위 30행 안에서 헤더 후보 행을 자동 추론.
  - `buildIndexMap(headers)` — 헤더 배열에서 표준 필드 인덱스 맵 즉시 생성.
- **`audit-log.js`** · 작업 이력 영구 저장소.
  - 가져오기 / 셀 수정 / 일괄 수정 / 케이블 추가·삭제·복원 / 호선 등록·이동 / 내보내기 / 시스템 등 12개 카테고리.
  - `record()` / `list()` / `subscribe()` / `clear()` / `exportTsv()` / `exportJsonl()` API.
  - 최대 5,000건 자동 보존, `localStorage["hk2401-cable-system:audit-log:v1"]`.
- **`excel-select.js`** · Excel-like 다중셀 선택 모듈.
  - 드래그 사각형 선택 / Shift+클릭 범위 / Ctrl+클릭 비연속 / Ctrl+A 전체 / Ctrl+C TSV 복사 / Esc 해제.
  - HTML 클립보드도 함께 제공하여 Excel · Google Sheets · LibreOffice Calc 모두에 그대로 붙여넣기 가능.
- **`enhance.js`** · 비침투적 통합 레이어.
  - main.js 수정 없이 MutationObserver로 cable 테이블 감지 후 ExcelSelect 자동 attach.
  - 헤더에 `📜 작업이력` 버튼 inject + 검색 / 필터 / 내보내기 모달.
- **`css/enhance.css`** · 다중셀 선택 시각화 + 작업 이력 모달 디자인 (네이비 + 골드 톤, secms.tech 일체).

### 변경 (Changed)
- `store.js`
  - `normalizeHeader()` 가 ColumnMapper가 있으면 위임하도록 변경 (없을 때 fallback).
  - `rowsToCables()` 가 ColumnMapper.detectHeaderRow + buildIndexMap 사용으로 강화 (기존 로직 fallback 보존).
  - `updateCable()` / `importCableRows()` / `addCable()` / `setActiveVessel()` / `addVessel()` 끝에 AuditLog.record() 자동 호출.
- `index.html`
  - 스크립트 로드 순서 정비 (xlsx → column-mapper → audit-log → excel-select → store → main → enhance).
  - 파일 업로드 accept 확장 (`.csv`, `.tsv` 추가).
  - meta description / theme-color / favicon 추가.
  - title에 SEcMS 명시.

### 메모
- 본 버전은 secms.tech 본 시스템과의 메뉴 톤(네이비 #102142 + 골드 #ffd428)을 일치시키는 1차 작업입니다.
- 추후 본 시스템 통합 시 `ws-shell` / `WorkspaceCommandBar` / `WorkspaceRibbon` 패턴으로 마이그레이션 예정.

## [1.0.0] - 2026-04 (INSTALL.zip 베이스)

### 신규 (Added)
- HK2401 호선 단독 빌드 — 호선 선택 / 케이블 리스트 / 대시보드 / 실적보고 / 달력 뷰 / 호선관리 / 관리자 패널.
- Excel 호환 그리드 편집 (contenteditable + 일괄 날짜 입력 + 행 단위 선택).
- TSV 가져오기 / 내보내기 + REV 자동 누적.
- 다중 호선 / 호선 전화부 / 미입력 표시 / 페이지네이션.
