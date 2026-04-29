"use strict";

const I18n = (() => {
  const STORAGE_KEY = "secms-cable-system:lang";
  const SUPPORTED = ["ko", "en"];
  const DEFAULT_LANG = "ko";

  const STRINGS = {
    ko: {
      "app.title": "Shipboard Electrical Cable Management System · 포설/결선 실적",
      "app.brandSub": "Shipboard Electrical cable Management System",
      "app.noVessel": "호선 미선택",
      "app.connectionLabel": (id, count) => `${id} Cable List · ${count}건`,

      "topbar.loadDefault": "HK2401 기본 불러오기",
      "topbar.import": "📂 파일 불러오기",
      "topbar.export": "📥 전체 내보내기",
      "topbar.backVessel": "호선 선택",
      "topbar.admin": "관리자 페이지",
      "topbar.logout": "Logout",
      "topbar.lang": "EN",

      "footer.contact": "프로그램 문의",

      "nav.vessel": "🚢 호선선택",
      "nav.cableList": "📋 케이블 리스트",
      "nav.dashboard": "📊 대시보드",
      "nav.report": "📋 실적보고",
      "nav.calendar": "🗓 달력 뷰",
      "nav.cableRegister": "➕ 호선케이블 등록",

      "vessel.pageTitle": "호선 관리",
      "vessel.pageSub": "작업할 호선을 선택하거나 새 호선을 등록합니다. 호선별 Cable List는 완전히 분리되어 저장됩니다.",
      "vessel.activeSection": "현재 접속 호선",
      "vessel.selectSection": "호선 선택",
      "vessel.registerSection": "새 호선 등록",
      "vessel.listSection": "등록된 호선",
      "vessel.field.id": "호선번호",
      "vessel.field.name": "호선명",
      "vessel.field.project": "프로젝트명",
      "vessel.field.owner": "선주/업체",
      "vessel.field.manager": "담당자",
      "vessel.field.phone": "대표전화",
      "vessel.field.memo": "메모",
      "vessel.btn.register": "호선 등록",
      "vessel.btn.connect": "접속",
      "vessel.btn.connected": "접속 중",
      "vessel.btn.call": "전화",
      "vessel.cardCables": "Cable List",
      "vessel.cardCablesUnit": "건",
      "vessel.cardLengthUnit": "m",
      "vessel.cardInstall": "포설",
      "vessel.cardOwner": "선주/업체",
      "vessel.cardManager": "담당",
      "vessel.cardPhone": "전화",
      "vessel.toastConnected": (id) => `${id} 호선에 접속했습니다.`,
      "vessel.toastRegistered": (id) => `${id} 호선을 등록했습니다.`,
      "vessel.empty": "등록된 호선이 없습니다.",
      "vessel.cardProject": "프로젝트",
      "vessel.cardPhoneLabel": "전화",
    },
    en: {
      "app.title": "Shipboard Electrical Cable Management System · Install / Connection",
      "app.brandSub": "Shipboard Electrical Cable Management System",
      "app.noVessel": "No Ship selected",
      "app.connectionLabel": (id, count) => `${id} Cable List · ${count} rows`,

      "topbar.loadDefault": "Load HK2401 default",
      "topbar.import": "📂 Import file",
      "topbar.export": "📥 Export all",
      "topbar.backVessel": "Select ship",
      "topbar.admin": "Admin",
      "topbar.logout": "Logout",
      "topbar.lang": "한",

      "footer.contact": "Contact",

      "nav.vessel": "🚢 Ships",
      "nav.cableList": "📋 Cable List",
      "nav.dashboard": "📊 Dashboard",
      "nav.report": "📋 Report",
      "nav.calendar": "🗓 Calendar",
      "nav.cableRegister": "➕ Register cable",

      "vessel.pageTitle": "Ship Management",
      "vessel.pageSub": "Select an active ship or register a new one. Each ship's Cable List is stored separately.",
      "vessel.activeSection": "Active ship",
      "vessel.selectSection": "Select ship",
      "vessel.registerSection": "Register new ship",
      "vessel.listSection": "Registered ships",
      "vessel.field.id": "Ship ID",
      "vessel.field.name": "Ship name",
      "vessel.field.project": "Project",
      "vessel.field.owner": "Owner / Yard",
      "vessel.field.manager": "Manager",
      "vessel.field.phone": "Phone",
      "vessel.field.memo": "Notes",
      "vessel.btn.register": "Register ship",
      "vessel.btn.connect": "Connect",
      "vessel.btn.connected": "Connected",
      "vessel.btn.call": "Call",
      "vessel.cardCables": "Cable List",
      "vessel.cardCablesUnit": "rows",
      "vessel.cardLengthUnit": "m",
      "vessel.cardInstall": "Installed",
      "vessel.cardOwner": "Owner",
      "vessel.cardManager": "Manager",
      "vessel.cardPhone": "Phone",
      "vessel.toastConnected": (id) => `Connected to ${id}.`,
      "vessel.toastRegistered": (id) => `Ship ${id} registered.`,
      "vessel.empty": "No ships registered yet.",
      "vessel.cardProject": "Project",
      "vessel.cardPhoneLabel": "Phone",
    },
  };

  let current = (() => {
    try {
      const stored = localStorage.getItem(STORAGE_KEY);
      if (stored && SUPPORTED.includes(stored)) return stored;
    } catch (_) { /* ignore */ }
    return DEFAULT_LANG;
  })();

  function get() { return current; }

  function set(lang) {
    if (!SUPPORTED.includes(lang)) return;
    current = lang;
    try { localStorage.setItem(STORAGE_KEY, lang); } catch (_) { /* ignore */ }
    document.documentElement.lang = lang;
  }

  function toggle() {
    set(current === "ko" ? "en" : "ko");
  }

  function t(key, ...args) {
    const value = (STRINGS[current] && STRINGS[current][key]) ?? (STRINGS[DEFAULT_LANG][key] ?? key);
    return typeof value === "function" ? value(...args) : value;
  }

  document.documentElement.lang = current;

  return { get, set, toggle, t, SUPPORTED };
})();
