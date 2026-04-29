"use strict";

const Store = (() => {
  const STORAGE_KEY = "hk2401-cable-system:v1";
  const DEFAULT_VESSEL_ID = "HK2401";
  const DEFAULT_DATA_URL = "data/hk2401-cables.tsv";

  const CABLE_FIELDS = [
    "sys",
    "circuitNo",
    "cableType",
    "cableDia",
    "fromEquipment",
    "fromCode",
    "fmMargin",
    "toEquipment",
    "toCode",
    "toMargin",
    "total",
    "route",
    "node",
    "installDate",
    "conFromDate",
    "conToDate",
    "rev",
  ];

  const FIELD_LABELS = {
    sys: "SYS",
    circuitNo: "CIRCUIT NO.",
    cableType: "CABLE TYPE",
    cableDia: "CABLE DIA",
    fromEquipment: "FROM EQUIPMENT",
    fromCode: "FROM CODE",
    fmMargin: "FM MARJIN",
    toEquipment: "TO EQUIPMENT",
    toCode: "TO CODE",
    toMargin: "TO MARJIN",
    total: "TOTAL",
    route: "ROUTE",
    node: "NODE구분",
    installDate: "포설일자",
    conFromDate: "결선 FROM",
    conToDate: "결선 TO",
    rev: "REV",
  };

  const emptyState = () => ({
    activeVesselId: DEFAULT_VESSEL_ID,
    vessels: [{
      id: DEFAULT_VESSEL_ID,
      name: "HK2401",
      projectName: "HK2401 호선",
      owner: "YANASE DOCK",
      manager: "",
      phone: "",
      memo: "기본 Cable List 포함 호선",
      createdAt: new Date().toISOString(),
    }],
    cablesByVessel: { [DEFAULT_VESSEL_ID]: [] },
    phonesByVessel: { [DEFAULT_VESSEL_ID]: [] },
  });

  let state = emptyState();
  const listeners = new Set();
  const undoStack = [];
  const UNDO_LIMIT = 50;
  let undoSuspended = false;
  let lastSerialized = JSON.stringify(state);

  function today() {
    return new Date().toISOString().slice(0, 10);
  }

  function structuredCloneSafe(value) {
    return JSON.parse(JSON.stringify(value));
  }

  function notify(event = "change") {
    listeners.forEach((listener) => listener(event, structuredCloneSafe(state)));
  }

  function snapshot() {
    if (undoSuspended) return;
    try {
      undoStack.push(lastSerialized);
      if (undoStack.length > UNDO_LIMIT) undoStack.shift();
    } catch (e) {
      // ignore
    }
  }

  function save() {
    snapshot();
    const next = JSON.stringify(state);
    localStorage.setItem(STORAGE_KEY, next);
    lastSerialized = next;
    notify("save");
  }

  function undo() {
    if (!undoStack.length) return false;
    const previous = undoStack.pop();
    try {
      state = JSON.parse(previous);
      lastSerialized = previous;
      localStorage.setItem(STORAGE_KEY, previous);
      notify("undo");
      return true;
    } catch (e) {
      return false;
    }
  }

  function canUndo() {
    return undoStack.length > 0;
  }

  function load() {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) {
      state = emptyState();
      return;
    }

    try {
      const parsed = JSON.parse(raw);
      state = {
        ...emptyState(),
        ...parsed,
        vessels: Array.isArray(parsed.vessels) ? parsed.vessels : [],
        cablesByVessel: parsed.cablesByVessel || {},
        phonesByVessel: parsed.phonesByVessel || {},
      };
    } catch (error) {
      console.error("Cable List 저장 데이터를 복구하지 못했습니다.", error);
      state = emptyState();
    }

    ensureDefaultVessel();
    Object.keys(state.cablesByVessel).forEach((vesselId) => {
      state.cablesByVessel[vesselId] = state.cablesByVessel[vesselId].map(normalizeCable);
    });
    lastSerialized = JSON.stringify(state);
    undoStack.length = 0;
  }

  function ensureDefaultVessel() {
    if (!state.vessels.some((vessel) => vessel.id === DEFAULT_VESSEL_ID)) {
      state.vessels.unshift(emptyState().vessels[0]);
    }
    state.cablesByVessel[DEFAULT_VESSEL_ID] ||= [];
    state.phonesByVessel[DEFAULT_VESSEL_ID] ||= [];
    if (!state.activeVesselId) state.activeVesselId = DEFAULT_VESSEL_ID;
  }

  function normalizeId(value) {
    return String(value || "")
      .trim()
      .replace(/\s+/g, "")
      .toUpperCase();
  }

  function normalizeHeader(value) {
    if (window.ColumnMapper && typeof ColumnMapper.normalizeHeader === "function") {
      return ColumnMapper.normalizeHeader(value);
    }
    return String(value || "")
      .replace(/﻿/g, "")
      .replace(/[\r\n"'\s.]/g, "")
      .toUpperCase();
  }

  function parseDelimited(text, delimiter = "\t") {
    const rows = [];
    let row = [];
    let cell = "";
    let inQuotes = false;

    for (let index = 0; index < text.length; index += 1) {
      const char = text[index];
      const next = text[index + 1];

      if (char === "\"") {
        if (inQuotes && next === "\"") {
          cell += "\"";
          index += 1;
        } else {
          inQuotes = !inQuotes;
        }
        continue;
      }

      if (!inQuotes && char === delimiter) {
        row.push(cell);
        cell = "";
        continue;
      }

      if (!inQuotes && (char === "\n" || char === "\r")) {
        if (char === "\r" && next === "\n") index += 1;
        row.push(cell);
        if (row.some((value) => String(value).trim() !== "")) rows.push(row);
        row = [];
        cell = "";
        continue;
      }

      cell += char;
    }

    row.push(cell);
    if (row.some((value) => String(value).trim() !== "")) rows.push(row);
    return rows;
  }

  function parseNumber(value) {
    const normalized = String(value ?? "").replace(/,/g, "").trim();
    if (!normalized) return 0;
    const parsed = Number(normalized);
    return Number.isFinite(parsed) ? parsed : 0;
  }

  function normalizeDate(value) {
    const raw = String(value ?? "").trim();
    if (!raw) return "";
    if (raw instanceof Date && !Number.isNaN(raw.valueOf())) return raw.toISOString().slice(0, 10);
    const excelSerial = Number(raw);
    if (Number.isFinite(excelSerial) && excelSerial > 20000 && excelSerial < 60000) {
      const date = new Date(Math.round((excelSerial - 25569) * 86400 * 1000));
      return date.toISOString().slice(0, 10);
    }
    const match = raw.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$/);
    if (match) {
      const [, year, month, day] = match;
      return `${year}-${month.padStart(2, "0")}-${day.padStart(2, "0")}`;
    }
    return raw;
  }

  function cleanText(value) {
    return String(value ?? "").replace(/\r\n/g, "\n").trim();
  }

  function findHeaderIndex(rows) {
    return rows.findIndex((row) => row.some((cell) => normalizeHeader(cell).includes("CIRCUITNO")));
  }

  function rowsToCables(rows) {
    if (!rows.length) return [];

    let headerIndex;
    let columns;

    // ColumnMapper가 로드되어 있으면 강화된 헤더 탐지/매핑 사용
    if (window.ColumnMapper) {
      headerIndex = ColumnMapper.detectHeaderRow(rows);
      if (headerIndex < 0) headerIndex = findHeaderIndex(rows);
      if (headerIndex < 0) return [];
      columns = ColumnMapper.buildIndexMap(rows[headerIndex]);
    } else {
      headerIndex = findHeaderIndex(rows);
      if (headerIndex < 0) return [];
      const headers = rows[headerIndex].map(normalizeHeader);
      const at = (...keys) => {
        const normalizedKeys = keys.map(normalizeHeader);
        return headers.findIndex((header) => normalizedKeys.includes(header));
      };
      columns = {
        sys: at("SYS"),
        circuitNo: at("CIRCUITNO"),
        cableType: at("CABLETYPE"),
        cableDia: at("CABLEDIA"),
        fromEquipment: at("FROMEQUIPMENT"),
        fromCode: at("FROMCODE"),
        fmMargin: at("FMMARJIN", "FMMARGIN"),
        toEquipment: at("TOEQUIPMENT"),
        toCode: at("TOCODE"),
        toMargin: at("TOMARJIN", "TOMARGIN"),
        total: at("TOTAL"),
        route: at("ROUTE", "MERGEDROUTE"),
        node: at("NODE구분", "NODE"),
        installDate: at("포설일자", "포설DATE", "INSTALLDATE"),
        conFromDate: at("결선FROM", "FROM결선", "CONFROM", "CONNECTIONFROM"),
        conToDate: at("결선TO", "TO결선", "CONTO", "CONNECTIONTO"),
        inspection: at("검사", "INSPECTION"),
        rev: at("REV"),
      };
    }

    const get = (row, key) => {
      const index = columns[key];
      return index >= 0 ? cleanText(row[index]) : "";
    };

    return rows.slice(headerIndex + 1)
      .map((row, index) => normalizeCable({
        id: `cable-${Date.now()}-${index}-${Math.random().toString(16).slice(2)}`,
        sys: get(row, "sys"),
        circuitNo: get(row, "circuitNo"),
        cableType: get(row, "cableType"),
        cableDia: get(row, "cableDia"),
        fromEquipment: get(row, "fromEquipment"),
        fromCode: get(row, "fromCode"),
        fmMargin: get(row, "fmMargin"),
        toEquipment: get(row, "toEquipment"),
        toCode: get(row, "toCode"),
        toMargin: get(row, "toMargin"),
        total: get(row, "total"),
        route: get(row, "route") || "local",
        node: get(row, "node"),
        installDate: get(row, "installDate"),
        conFromDate: get(row, "conFromDate"),
        conToDate: get(row, "conToDate"),
        inspection: get(row, "inspection"),
        rev: get(row, "rev"),
        createdAt: new Date().toISOString(),
      }))
      .filter((row) => row.circuitNo || row.cableType);
  }

  function normalizeCable(input) {
    return {
      id: input.id || `cable-${Date.now()}-${Math.random().toString(16).slice(2)}`,
      sys: cleanText(input.sys).toUpperCase(),
      circuitNo: cleanText(input.circuitNo),
      cableType: cleanText(input.cableType),
      cableDia: parseNumber(input.cableDia),
      fromEquipment: cleanText(input.fromEquipment),
      fromCode: cleanText(input.fromCode).toUpperCase(),
      fmMargin: cleanText(input.fmMargin),
      toEquipment: cleanText(input.toEquipment),
      toCode: cleanText(input.toCode).toUpperCase(),
      toMargin: cleanText(input.toMargin),
      total: parseNumber(input.total),
      route: cleanText(input.route) || "local",
      node: cleanText(input.node),
      installDate: normalizeDate(input.installDate),
      conFromDate: normalizeDate(input.conFromDate),
      conToDate: normalizeDate(input.conToDate),
      inspection: cleanText(input.inspection),
      rev: cleanText(input.rev),
      deleted: Boolean(input.deleted) || /삭제|DELETE/i.test(cleanText(input.rev)),
      createdAt: input.createdAt || new Date().toISOString(),
      updatedAt: input.updatedAt || "",
    };
  }

  function patchDefaultProgressDates() {
    const cables = state.cablesByVessel[DEFAULT_VESSEL_ID] || [];
    const progressPatches = new Map(Object.entries({
      L0113: { conFromDate: "2026-04-27", conToDate: "2026-04-27" },
      L0114: { conFromDate: "2026-04-27", conToDate: "2026-04-27" },
      L0712: { conFromDate: "2026-04-27", conToDate: "2026-04-27" },
      L0722: { conFromDate: "2026-04-27" },
      L0723: { conFromDate: "2026-04-27", conToDate: "2026-04-27" },
      L2403: { conFromDate: "2026-04-27", conToDate: "2026-04-27" },
    }));

    let changed = false;
    cables.forEach((cable) => {
      const patch = progressPatches.get(cleanText(cable.circuitNo).toUpperCase());
      if (!patch) return;
      Object.entries(patch).forEach(([field, value]) => {
        if (!cable[field]) {
          cable[field] = value;
          changed = true;
        }
      });
    });

    return changed;
  }

  function getActiveVessel() {
    ensureDefaultVessel();
    return state.vessels.find((vessel) => vessel.id === state.activeVesselId) || state.vessels[0];
  }

  function getCables(vesselId = getActiveVessel().id) {
    return state.cablesByVessel[vesselId] || [];
  }

  function getPhones(vesselId = getActiveVessel().id) {
    return state.phonesByVessel[vesselId] || [];
  }

  function getKpi() {
    const allCables = getCables();
    const cables = allCables.filter((cable) => !cable.deleted);
    const totalLength = cables.reduce((sum, cable) => sum + (Number(cable.total) || 0), 0);
    const installedLength = cables.reduce((sum, cable) => sum + (cable.installDate ? Number(cable.total) || 0 : 0), 0);
    const conFrom = cables.filter((cable) => cable.conFromDate).length;
    const conTo = cables.filter((cable) => cable.conToDate).length;

    return {
      registeredCount: allCables.length,
      totalCount: cables.length,
      totalLength,
      installedLength,
      pendingLength: Math.max(0, totalLength - installedLength),
      installRate: totalLength > 0 ? installedLength / totalLength : 0,
      conFrom,
      conTo,
      conFromRate: cables.length > 0 ? conFrom / cables.length : 0,
      conToRate: cables.length > 0 ? conTo / cables.length : 0,
    };
  }

  function groupBy(field) {
    const map = new Map();
    getCables().filter((cable) => !cable.deleted).forEach((cable) => {
      const key = cable[field] || "미분류";
      const current = map.get(key) || { key, count: 0, total: 0, installed: 0, conFrom: 0, conTo: 0 };
      current.count += 1;
      current.total += Number(cable.total) || 0;
      if (cable.installDate) current.installed += Number(cable.total) || 0;
      if (cable.conFromDate) current.conFrom += 1;
      if (cable.conToDate) current.conTo += 1;
      map.set(key, current);
    });
    return [...map.values()].sort((a, b) => b.total - a.total || b.count - a.count);
  }

  function appendRev(existing, message) {
    const cleanMessage = cleanText(message);
    if (!cleanMessage) return cleanText(existing);
    const before = cleanText(existing);
    if (!before) return cleanMessage;
    if (before.includes(cleanMessage)) return before;
    return `${before}\n• ${cleanMessage}`;
  }

  function ensureRevPrefix(rev) {
    const clean = cleanText(rev);
    if (!clean) return "";
    if (clean.startsWith("•") || clean.includes("\n•")) return clean;
    // legacy "A | B | C" form -> convert to bulleted
    if (clean.includes(" | ")) {
      return clean.split(" | ").map((line) => `• ${line.trim()}`).join("\n");
    }
    return `• ${clean}`;
  }

  function formatChange(field, before, after) {
    const label = FIELD_LABELS[field] || field;
    const from = cleanText(before) || "-";
    const to = cleanText(after) || "-";
    return `${label}: ${from} -> ${to}`;
  }

  function sanitizePatch(patch) {
    const output = {};
    Object.entries(patch || {}).forEach(([field, value]) => {
      if (!CABLE_FIELDS.includes(field) && field !== "deleted") return;
      if (field === "cableDia" || field === "total") output[field] = parseNumber(value);
      else if (field === "installDate" || field === "conFromDate" || field === "conToDate") output[field] = normalizeDate(value);
      else if (field === "sys" || field === "fromCode" || field === "toCode") output[field] = cleanText(value).toUpperCase();
      else if (field === "deleted") output[field] = Boolean(value);
      else output[field] = cleanText(value);
    });
    return output;
  }

  function updateCable(id, patch, options = {}) {
    const vesselId = getActiveVessel().id;
    const cables = getCables(vesselId);
    const index = cables.findIndex((cable) => cable.id === id);
    if (index < 0) throw new Error("수정할 케이블을 찾을 수 없습니다.");

    const sanitized = sanitizePatch(patch);
    const target = cables[index];
    const changes = [];

    Object.entries(sanitized).forEach(([field, value]) => {
      const previous = target[field] ?? "";
      if (String(previous) !== String(value)) {
        changes.push(formatChange(field, previous, value));
        target[field] = value;
      }
    });

    if (!changes.length) return target;
    target.updatedAt = new Date().toISOString();

    if (options.recordRev !== false) {
      const reason = options.reason || "셀 수정";
      const message = `${reason} (${today()}) ${changes.join("; ")}`;
      target.rev = appendRev(target.rev, message);
    }

    if (window.AuditLog) {
      const cat = (options.reason || "").indexOf("일괄") >= 0 ? "bulk-edit" : "cell-edit";
      AuditLog.record(cat, `${target.circuitNo || target.id} ${options.reason || "셀 수정"}`, {
        id: target.id, circuitNo: target.circuitNo, changes,
      });
    }

    save();
    return target;
  }

  function importCableRows(rows, options = {}) {
    const incoming = rowsToCables(rows);
    if (!incoming.length) throw new Error("불러올 Cable List가 없습니다.");

    const vesselId = getActiveVessel().id;
    const sourceName = options.sourceName || "import";

    if (options.replace) {
      state.cablesByVessel[vesselId] = incoming.map((cable) => ({
        ...cable,
        rev: appendRev(cable.rev, `불러오기 교체 (${sourceName}, ${today()})`),
      }));
      save();
      return { added: incoming.length, updated: 0, deleted: 0, unchanged: 0, totalIncoming: incoming.length };
    }

    state.cablesByVessel[vesselId] ||= [];
    const current = state.cablesByVessel[vesselId];
    const byCircuit = new Map(current.map((cable, index) => [cleanText(cable.circuitNo).toUpperCase(), { cable, index }]));
    const summary = { added: 0, updated: 0, deleted: 0, restored: 0, missingDeleted: 0, unchanged: 0, totalIncoming: incoming.length };
    const seenKeys = new Set();
    const stamp = new Date().toISOString().slice(0, 16).replace("T", " ");

    incoming.forEach((rawCable) => {
      const incomingCable = normalizeCable(rawCable);
      const key = cleanText(incomingCable.circuitNo).toUpperCase();
      if (!key) return;
      seenKeys.add(key);
      const hit = byCircuit.get(key);
      const sourceRev = cleanText(incomingCable.rev);
      const incomingDeleted = incomingCable.deleted || incomingCable.total === 0 || /삭제|DELETE/i.test(sourceRev);

      if (!hit) {
        const action = incomingDeleted ? "DELETE" : "ADD";
        current.push({
          ...incomingCable,
          deleted: incomingDeleted,
          rev: appendRev("", `[${action}] ${stamp} (${sourceName}) 신규`),
          createdAt: new Date().toISOString(),
          updatedAt: new Date().toISOString(),
        });
        summary.added += 1;
        return;
      }

      const target = hit.cable;
      target.rev = ensureRevPrefix(target.rev);
      const patch = {};
      const changes = [];

      CABLE_FIELDS.forEach((field) => {
        if (field === "rev") return;
        const next = incomingCable[field] ?? "";
        const previous = target[field] ?? "";
        if (String(previous) !== String(next)) {
          patch[field] = next;
          changes.push(formatChange(field, previous, next));
        }
      });

      let actionTag = "UPDATE";
      if (incomingDeleted && !target.deleted) {
        patch.deleted = true;
        actionTag = "DELETE";
        changes.push("status: active -> deleted");
      } else if (!incomingDeleted && target.deleted) {
        patch.deleted = false;
        actionTag = "RESTORE";
        changes.push("status: deleted -> active");
      }

      if (!changes.length) {
        summary.unchanged += 1;
        return;
      }

      Object.assign(target, sanitizePatch(patch));
      target.updatedAt = new Date().toISOString();
      const message = `[${actionTag}] ${stamp} (${sourceName}) ${changes.join("; ")}`;
      target.rev = appendRev(target.rev, message);

      if (actionTag === "DELETE") summary.deleted += 1;
      else if (actionTag === "RESTORE") summary.restored += 1;
      else summary.updated += 1;
    });

    if (options.markMissingAsDeleted) {
      current.forEach((cable) => {
        const key = cleanText(cable.circuitNo).toUpperCase();
        if (!key || seenKeys.has(key) || cable.deleted) return;
        cable.deleted = true;
        cable.updatedAt = new Date().toISOString();
        cable.rev = appendRev(ensureRevPrefix(cable.rev), `[DELETE] ${stamp} (${sourceName}) 신규 파일에 누락 → 삭제 처리`);
        summary.missingDeleted += 1;
      });
    }

    save();

    if (window.AuditLog) {
      AuditLog.record("import", `Excel 가져오기: ${sourceName}`, {
        sourceName,
        ...summary,
      });
    }

    return summary;
  }

  async function loadDefaultCables({ replace = false } = {}) {
    const response = await fetch(DEFAULT_DATA_URL, { cache: "no-store" });
    if (!response.ok) {
      throw new Error(`HK2401 Cable List 파일을 읽지 못했습니다. (${response.status})`);
    }
    const text = await response.text();
    const rows = parseDelimited(text, "\t");
    const cables = rowsToCables(rows);
    if (!cables.length) {
      throw new Error("HK2401 Cable List에서 케이블 행을 찾지 못했습니다.");
    }

    const vesselId = getActiveVessel().id;
    if (replace || !state.cablesByVessel[vesselId]?.length) {
      state.cablesByVessel[vesselId] = cables;
    } else {
      state.cablesByVessel[vesselId] = [...state.cablesByVessel[vesselId], ...cables];
    }
    save();
    return cables.length;
  }

  function toTsv(cables = getCables()) {
    const headers = [
      "SYS",
      "CIRCUIT NO.",
      "CABLE TYPE",
      "CABLE DIA",
      "FROM EQUIPMENT",
      "FROM CODE",
      "FM MARJIN",
      "TO EQUIPMENT",
      "TO CODE",
      "TO MARJIN",
      "TOTAL",
      "ROUTE",
      "NODE구분",
      "포설일자",
      "결선 FROM",
      "결선 TO",
      "REV",
    ];
    const escape = (value) => {
      const text = String(value ?? "");
      return /[\t\r\n"]/.test(text) ? `"${text.replace(/"/g, "\"\"")}"` : text;
    };
    const lines = cables.map((cable) => [
      cable.sys,
      cable.circuitNo,
      cable.cableType,
      cable.cableDia || "",
      cable.fromEquipment,
      cable.fromCode,
      cable.fmMargin,
      cable.toEquipment,
      cable.toCode,
      cable.toMargin,
      cable.total || "",
      cable.route,
      cable.node,
      cable.installDate,
      cable.conFromDate,
      cable.conToDate,
      cable.rev,
    ].map(escape).join("\t"));
    return [headers.join("\t"), ...lines].join("\r\n");
  }

  return {
    DEFAULT_VESSEL_ID,
    FIELD_LABELS,

    async init() {
      load();
      ensureDefaultVessel();
      let changed = false;
      if (!getCables(DEFAULT_VESSEL_ID).length) {
        await loadDefaultCables({ replace: true });
        changed = true;
      }
      changed = patchDefaultProgressDates() || changed;
      if (changed) save();
      else save();
    },

    subscribe(listener) {
      listeners.add(listener);
      return () => listeners.delete(listener);
    },

    getState() {
      return structuredCloneSafe(state);
    },

    getActiveVessel,
    getCables,
    getPhones,
    getKpi,
    groupBy,
    toTsv,
    parseDelimited,
    rowsToCables,
    loadDefaultCables,
    importCableRows,
    updateCable,

    setActiveVessel(id) {
      const normalized = normalizeId(id);
      if (!state.vessels.some((vessel) => vessel.id === normalized)) {
        throw new Error("등록되지 않은 호선입니다.");
      }
      const previous = state.activeVesselId;
      state.activeVesselId = normalized;
      save();
      if (window.AuditLog && previous !== normalized) {
        AuditLog.record("vessel-switch", `호선 이동: ${previous} → ${normalized}`, { from: previous, to: normalized });
      }
    },

    addVessel(input) {
      const id = normalizeId(input.id || input.name);
      if (!id) throw new Error("호선번호를 입력하세요.");
      if (state.vessels.some((vessel) => vessel.id === id)) {
        throw new Error(`${id} 호선은 이미 등록되어 있습니다.`);
      }
      const vessel = {
        id,
        name: cleanText(input.name) || id,
        projectName: cleanText(input.projectName) || `${id} 호선`,
        owner: cleanText(input.owner),
        manager: cleanText(input.manager),
        phone: cleanText(input.phone),
        memo: cleanText(input.memo),
        createdAt: new Date().toISOString(),
      };
      state.vessels.push(vessel);
      state.cablesByVessel[id] = [];
      state.phonesByVessel[id] = [];
      save();
      if (window.AuditLog) {
        AuditLog.record("vessel-add", `${vessel.id} 호선 등록`, vessel);
      }
      return vessel;
    },

    addCable(input) {
      const vesselId = getActiveVessel().id;
      const cable = normalizeCable(input);
      if (!cable.circuitNo) throw new Error("CIRCUIT NO.를 입력하세요.");
      if (!cable.cableType) throw new Error("CABLE TYPE을 입력하세요.");
      cable.rev = appendRev(cable.rev, `신규 등록 (${today()})`);
      state.cablesByVessel[vesselId] ||= [];
      state.cablesByVessel[vesselId].unshift(cable);
      save();
      if (window.AuditLog) {
        AuditLog.record("cable-add", `${cable.circuitNo} 신규 케이블 등록`, {
          id: cable.id, circuitNo: cable.circuitNo, cableType: cable.cableType,
        });
      }
      return cable;
    },

    addPhone(input) {
      const vesselId = getActiveVessel().id;
      const label = cleanText(input.label);
      const phone = cleanText(input.phone);
      if (!label) throw new Error("전화명 또는 위치를 입력하세요.");
      if (!phone) throw new Error("전화번호를 입력하세요.");
      const entry = {
        id: `phone-${Date.now()}-${Math.random().toString(16).slice(2)}`,
        label,
        phone,
        location: cleanText(input.location),
        memo: cleanText(input.memo),
        createdAt: new Date().toISOString(),
      };
      state.phonesByVessel[vesselId] ||= [];
      state.phonesByVessel[vesselId].push(entry);
      save();
      return entry;
    },

    deletePhone(id) {
      const vesselId = getActiveVessel().id;
      state.phonesByVessel[vesselId] = getPhones(vesselId).filter((entry) => entry.id !== id);
      save();
    },

    importTsvText(text, { replace = false, sourceName = "TSV" } = {}) {
      const rows = parseDelimited(text, "\t");
      return importCableRows(rows, { replace, sourceName });
    },

    undo,
    canUndo,
  };
})();

window.Store = Store;
