"use strict";

/**
 * SEcMS · HK2401 Cable List · Robust Column Mapper
 *
 * 어떤 형식의 Excel/TSV/CSV가 들어와도 표준 Cable 스키마(SYS, CIRCUIT NO., CABLE TYPE, ...)에
 * 자동 매핑할 수 있도록 강력한 별칭 사전과 헤더 추론 로직을 제공한다.
 *
 * 동작:
 *  1. 헤더 후보 행 탐색 (CIRCUIT/회로/POSITION/등 키워드 매칭)
 *  2. 각 셀에 대해 다중 후보 동의어 비교 (한글/영문/공백/줄바꿈/기호 무시)
 *  3. 결과를 표준 필드 인덱스 맵으로 반환
 *
 * Public API:
 *   ColumnMapper.normalizeHeader(value)
 *   ColumnMapper.detectHeaderRow(rows)
 *   ColumnMapper.buildIndexMap(headers)
 *   ColumnMapper.aliasFor(field)  -> string[]
 */
const ColumnMapper = (() => {
  // 표준 필드별 별칭(영문/한글/약어/공백·줄바꿈·기호 제거 후 비교)
  const ALIASES = {
    sys: [
      "SYS", "SYSTEM", "계통", "시스템", "구분",
    ],
    circuitNo: [
      "CIRCUITNO", "CIRCUIT", "회로번호", "회로NO", "CIRCUIT번호",
      "회선번호", "CKTNO", "CKT", "회로",
    ],
    cableType: [
      "CABLETYPE", "CABLE종류", "케이블종류", "케이블타입", "TYPE",
      "케이블TYPE", "타입",
    ],
    cableDia: [
      "CABLEDIA", "DIA", "직경", "케이블직경", "DIAMETER",
      "케이블DIA", "굵기",
    ],
    fromEquipment: [
      "FROMEQUIPMENT", "FROM기기", "발신기기", "FROMEQUIP", "FROM",
      "FROM장비", "출발", "기점", "출발기기",
    ],
    fromCode: [
      "FROMCODE", "FROM코드", "출발코드", "FROMCD", "FROM부호",
    ],
    fmMargin: [
      "FMMARJIN", "FMMARGIN", "FM여유", "FROMMARGIN", "FROM여유",
      "FROMMARJIN", "FM마진",
    ],
    toEquipment: [
      "TOEQUIPMENT", "TO기기", "수신기기", "TOEQUIP", "TO",
      "TO장비", "도착", "도착기기", "종점",
    ],
    toCode: [
      "TOCODE", "TO코드", "도착코드", "TOCD", "TO부호",
    ],
    toMargin: [
      "TOMARJIN", "TOMARGIN", "TO여유", "TO마진",
    ],
    total: [
      "TOTAL", "총길이", "합계", "전체길이", "LENGTH", "길이", "TOTAL길이",
      "총수량", "TOTAL량",
    ],
    route: [
      "ROUTE", "MERGEDROUTE", "경로", "라우트", "케이블경로",
      "PATH", "통로",
    ],
    node: [
      "NODE구분", "NODE", "노드", "구역", "ZONE", "AREA",
      "NODE분류", "구역구분",
    ],
    installDate: [
      "포설일자", "포설DATE", "INSTALLDATE", "INSTALL", "포설",
      "포설일", "포설날짜", "포설완료일", "INSTALLED", "INSTALLEDDATE",
    ],
    conFromDate: [
      "결선FROM", "FROM결선", "CONFROM", "CONNECTIONFROM",
      "FROM결선일", "결선FROM일자", "결선시작", "FROMCONNECTED",
      "CONNECTFROM",
    ],
    conToDate: [
      "결선TO", "TO결선", "CONTO", "CONNECTIONTO",
      "TO결선일", "결선TO일자", "결선완료", "TOCONNECTED",
      "CONNECTTO",
    ],
    inspection: [
      "검사", "INSPECTION", "검사일", "INSPECTIONDATE",
      "검사완료", "검사결과",
    ],
    rev: [
      "REV", "REVISION", "수정", "비고", "REMARK", "REMARKS", "MEMO",
      "이력", "수정이력", "변경이력", "REV이력",
    ],
    // 추가 메타 (있으면 보존)
    deleted: [
      "삭제", "DELETED", "DEL", "DELETE", "삭제여부", "STATUS",
    ],
  };

  const HEADER_HINTS = ["CIRCUITNO", "CIRCUIT", "회로번호", "CABLETYPE", "CABLE종류", "TYPE"];

  /**
   * 헤더값 정규화: BOM, 공백, 줄바꿈, 따옴표, 점, 슬래시 제거 후 대문자
   */
  function normalizeHeader(value) {
    return String(value || "")
      .replace(/﻿/g, "")
      .replace(/[\r\n"'\s./()\-_]/g, "")
      .toUpperCase();
  }

  /**
   * 헤더 행 자동 탐지
   * 가장 많은 헤더 키워드를 가진 상위 5행 중에서 우선 행을 채택
   */
  function detectHeaderRow(rows) {
    if (!Array.isArray(rows) || !rows.length) return -1;

    let bestIndex = -1;
    let bestScore = 0;
    const limit = Math.min(rows.length, 30); // 최대 상위 30행만 검사
    for (let i = 0; i < limit; i += 1) {
      const row = rows[i];
      if (!Array.isArray(row)) continue;
      let score = 0;
      const seen = new Set();
      row.forEach((cell) => {
        const norm = normalizeHeader(cell);
        if (!norm) return;
        // 직접 hint 매칭
        for (const hint of HEADER_HINTS) {
          if (norm.includes(hint)) score += 2;
        }
        // 별칭 매칭
        Object.keys(ALIASES).forEach((field) => {
          if (seen.has(field)) return;
          const aliases = ALIASES[field];
          for (const alias of aliases) {
            if (norm === normalizeHeader(alias)) {
              score += 1;
              seen.add(field);
              break;
            }
          }
        });
      });
      if (score > bestScore) {
        bestScore = score;
        bestIndex = i;
      }
    }

    // 최소 임계치 (헤더처럼 보이는지)
    return bestScore >= 3 ? bestIndex : -1;
  }

  /**
   * 헤더 배열 → { field: 컬럼인덱스 }
   * 매칭 실패한 필드는 -1
   */
  function buildIndexMap(headers) {
    if (!Array.isArray(headers)) return {};
    const normalized = headers.map(normalizeHeader);
    const map = {};
    Object.keys(ALIASES).forEach((field) => {
      const aliases = ALIASES[field].map(normalizeHeader);
      let foundIndex = -1;
      // 정확 일치 우선
      for (let i = 0; i < normalized.length; i += 1) {
        if (aliases.includes(normalized[i])) {
          foundIndex = i;
          break;
        }
      }
      // 부분 포함도 허용 (예: "포설일자_2026")
      if (foundIndex < 0) {
        for (let i = 0; i < normalized.length; i += 1) {
          for (const alias of aliases) {
            if (alias && normalized[i].includes(alias)) {
              foundIndex = i;
              break;
            }
          }
          if (foundIndex >= 0) break;
        }
      }
      map[field] = foundIndex;
    });
    return map;
  }

  function aliasFor(field) {
    return (ALIASES[field] || []).slice();
  }

  return {
    normalizeHeader,
    detectHeaderRow,
    buildIndexMap,
    aliasFor,
    ALIASES,
    HEADER_HINTS,
  };
})();

window.ColumnMapper = ColumnMapper;
