"use strict";

/**
 * SEcMS · HK2401 Cable List · Audit Log Module
 * 모든 변경 작업(케이블 추가/수정/삭제, Excel 가져오기, 일괄 변경, 호선 작업 등)을
 * 별도 저장소에 영구 보관하고, 시간순 조회/필터/내보내기를 제공한다.
 *
 * Storage:
 *   localStorage["hk2401-cable-system:audit-log:v1"] : Array<AuditEntry>
 *
 * Export formats:
 *   - JSON Lines (.jsonl) : 외부 분석 도구용
 *   - TSV (.tsv) : Excel 호환
 */
const AuditLog = (() => {
  const STORAGE_KEY = "hk2401-cable-system:audit-log:v1";
  const MAX_ENTRIES = 5000;

  const CATEGORY_LABELS = {
    "import": "📂 가져오기",
    "cell-edit": "✏️ 셀 수정",
    "bulk-edit": "📦 일괄 수정",
    "cable-add": "➕ 케이블 추가",
    "cable-delete": "🗑 케이블 삭제",
    "cable-restore": "♻ 케이블 복원",
    "vessel-add": "🚢 호선 등록",
    "vessel-switch": "🔁 호선 이동",
    "phone-add": "📞 전화 등록",
    "phone-delete": "📞 전화 삭제",
    "export": "📥 내보내기",
    "auth": "🔐 로그인/로그아웃",
    "system": "⚙ 시스템",
  };

  let entries = [];
  let listeners = new Set();

  function load() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      entries = raw ? JSON.parse(raw) : [];
      if (!Array.isArray(entries)) entries = [];
    } catch (error) {
      console.warn("Audit log 복구 실패", error);
      entries = [];
    }
  }

  function save() {
    try {
      // 최대치 초과 시 오래된 항목 잘라내기
      if (entries.length > MAX_ENTRIES) {
        entries = entries.slice(-MAX_ENTRIES);
      }
      localStorage.setItem(STORAGE_KEY, JSON.stringify(entries));
    } catch (error) {
      console.warn("Audit log 저장 실패 (storage quota?)", error);
    }
    listeners.forEach((listener) => {
      try { listener(entries); } catch (e) { console.error(e); }
    });
  }

  function makeId() {
    return `log-${Date.now()}-${Math.random().toString(16).slice(2, 8)}`;
  }

  function getActor() {
    try {
      const raw = sessionStorage.getItem("hk2401-cable-system:operator");
      if (raw) {
        const parsed = JSON.parse(raw);
        return parsed.name || "anonymous";
      }
    } catch (e) { /* ignore */ }
    return "anonymous";
  }

  function getActiveVesselId() {
    try {
      if (window.Store && Store.getActiveVessel) {
        return Store.getActiveVessel().id;
      }
    } catch (e) { /* ignore */ }
    return "-";
  }

  /**
   * 새 이력 추가
   * @param {string} category - import|cell-edit|bulk-edit|cable-add|...
   * @param {string} action - 사람이 읽을 수 있는 짧은 설명
   * @param {Object} [details] - 추가 메타데이터 (자유 형식)
   */
  function record(category, action, details = {}) {
    const entry = {
      id: makeId(),
      timestamp: new Date().toISOString(),
      category: String(category || "system"),
      action: String(action || ""),
      actor: getActor(),
      vesselId: getActiveVesselId(),
      details: details || {},
    };
    entries.push(entry);
    save();
    return entry;
  }

  function list(options = {}) {
    const { category = "", actor = "", vesselId = "", from = "", to = "", limit = 0, search = "" } = options;
    let result = entries.slice();
    if (category) result = result.filter((e) => e.category === category);
    if (actor) result = result.filter((e) => e.actor === actor);
    if (vesselId) result = result.filter((e) => e.vesselId === vesselId);
    if (from) result = result.filter((e) => e.timestamp >= from);
    if (to) result = result.filter((e) => e.timestamp <= to);
    if (search) {
      const needle = String(search).toLowerCase();
      result = result.filter((e) =>
        String(e.action).toLowerCase().includes(needle) ||
        JSON.stringify(e.details).toLowerCase().includes(needle)
      );
    }
    result.sort((a, b) => b.timestamp.localeCompare(a.timestamp));
    if (limit > 0) result = result.slice(0, limit);
    return result;
  }

  function clear() {
    entries = [];
    save();
  }

  function subscribe(listener) {
    listeners.add(listener);
    return () => listeners.delete(listener);
  }

  function exportJsonl() {
    return entries
      .slice()
      .sort((a, b) => a.timestamp.localeCompare(b.timestamp))
      .map((e) => JSON.stringify(e))
      .join("\n");
  }

  function exportTsv() {
    const headers = ["TIMESTAMP", "ACTOR", "VESSEL", "CATEGORY", "ACTION", "DETAILS"];
    const escape = (value) => {
      const text = String(value ?? "");
      return /[\t\r\n"]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
    };
    const lines = entries
      .slice()
      .sort((a, b) => a.timestamp.localeCompare(b.timestamp))
      .map((e) => [
        e.timestamp,
        e.actor,
        e.vesselId,
        e.category,
        e.action,
        JSON.stringify(e.details),
      ].map(escape).join("\t"));
    return [headers.join("\t"), ...lines].join("\r\n");
  }

  function categoryLabel(category) {
    return CATEGORY_LABELS[category] || category || "기타";
  }

  // 즉시 로드
  load();

  return {
    record,
    list,
    clear,
    subscribe,
    exportJsonl,
    exportTsv,
    categoryLabel,
    get count() { return entries.length; },
    get all() { return entries.slice(); },
    CATEGORY_LABELS,
  };
})();

window.AuditLog = AuditLog;
