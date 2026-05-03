"use strict";

const App = (() => {
  const PAGE_SIZE = 1000000;  // 엑셀처럼 한 페이지 무한 스크롤 (페이저 비활성)
  const XLSX_CDN = "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js";
  const USER_SESSION_KEY = "hk2401-cable-system:operator";
  let ALLOWED_OPERATOR_NAMES = ["신도식", "김용수"]; // CF 환경변수로 오버라이드
  let operatorPassword = ""; // CF USER_PASSWORD
  let adminId = "";          // CF ADMIN_id
  let adminPassword = "";    // CF ADMIN_PASSWORD

  let appMode = "name";
  let currentView = "cableList";
  let currentUser = null;
  let pendingView = "cableList";
  let adminTab = "overview";
  let calendarYear = 2026;
  let calendarMonth = 2;
  let calendarMode = "install";
  let reportPeriod = "month"; // "month" | "week"
  let reportWeek = isoWeekOfMonth(new Date(calendarYear, calendarMonth, 15));
  let tablePage = 1;
  let tableFilter = {
    query: "",
    sys: "",
    node: "",
    type: "",
    installStatus: "",
    conFromStatus: "",
    conToStatus: "",
  };
  let tableSelectedIds = new Set();
  let lastSelectedCableId = "";
  let calendarSelectedDate = "";
  let cableTab = "active"; // "active" | "deleted"
  let shipTab = "switch"; // "switch" | "register" | "cable"

  const views = {
    cableList: { label: "📋 케이블 리스트", render: renderCableList },
    dashboard: { label: "📊 대시보드", render: renderDashboard },
    report: { label: "📋 실적보고", render: renderReport },
    calendar: { label: "🗓 달력 뷰", render: renderCalendar },
    shipInfo: { label: "🚢 호선관리", render: renderShipManage },
  };

  const excelColumns = [
    { field: "sys", label: "SYS", className: "col-sys center" },
    { field: "circuitNo", label: "CIRCUIT<br>NO.", className: "col-circuit mono center" },
    { field: "cableType", label: "CABLE<br>TYPE", className: "col-type center" },
    { field: "cableDia", label: "CABLE<br>DIA", className: "col-dia num" },
    { field: "fromEquipment", label: "FROM EQUIPMENT", className: "col-equipment center" },
    { field: "fromCode", label: "FROM CODE", className: "col-code center" },
    { field: "fmMargin", label: "FM<br>MARJIN", className: "col-margin num" },
    { field: "toEquipment", label: "TO EQUIPMENT", className: "col-equipment center" },
    { field: "toCode", label: "TO CODE", className: "col-code center" },
    { field: "toMargin", label: "TO<br>MARJIN", className: "col-margin num" },
    { field: "total", label: "TOTAL", className: "col-total num" },
    { field: "route", label: "ROUTE", className: "col-route center excel-red" },
    { field: "node", label: "NODE구분", className: "col-node center excel-red" },
    { field: "installDate", label: "포설일자", className: "col-date center", kind: "date", headerClass: "yellow" },
    { field: "conFromDate", label: "결선<br>FROM", className: "col-date center", kind: "date", headerClass: "green red-text" },
    { field: "conToDate", label: "결선<br>TO", className: "col-date center", kind: "date", headerClass: "green red-text" },
    { field: "rev", label: "REV", className: "col-rev center" },
  ];

  function el(id) {
    return document.getElementById(id);
  }

  function excelColumnLetter(index) {
    let n = index;
    let letters = "";
    while (n > 0) {
      const rem = (n - 1) % 26;
      letters = String.fromCharCode(65 + rem) + letters;
      n = Math.floor((n - 1) / 26);
    }
    return letters;
  }

  function main() {
    return el("main-content");
  }

  function escapeHtml(value) {
    return String(value ?? "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  function escapeAttr(value) {
    return escapeHtml(value);
  }

  function pct(value) {
    return `${(Number(value || 0) * 100).toFixed(1)}%`;
  }

  function num(value, digits = 0) {
    return Number(value || 0).toLocaleString("ko-KR", {
      maximumFractionDigits: digits,
      minimumFractionDigits: digits,
    });
  }

  function today() {
    const date = new Date();
    date.setMinutes(date.getMinutes() - date.getTimezoneOffset());
    return date.toISOString().slice(0, 10);
  }

  function toast(message) {
    const target = el("toast");
    if (!target) return;
    target.textContent = message;
    target.classList.add("show");
    clearTimeout(toast.timer);
    toast.timer = setTimeout(() => target.classList.remove("show"), 3000);
  }

  function formData(form) {
    return Object.fromEntries(new FormData(form).entries());
  }

  function setMode(mode) {
    appMode = mode;
    document.body.dataset.mode = mode;
  }

  async function fetchConfig() {
    try {
      const resp = await fetch("/api/config");
      if (!resp.ok) return;
      const cfg = await resp.json();
      if (Array.isArray(cfg.users) && cfg.users.length) {
        ALLOWED_OPERATOR_NAMES = cfg.users.filter(Boolean);
      }
      if (cfg.userPassword) operatorPassword = cfg.userPassword;
      if (cfg.adminId) adminId = cfg.adminId;
      if (cfg.adminPassword) adminPassword = cfg.adminPassword;
    } catch (_) { /* 로컬 개발 환경: 기본값 유지 */ }
  }

  function normalizeOperatorName(value) {
    return String(value || "").trim().replace(/\s+/g, "").toLowerCase();
  }

  function isAllowedOperator(name) {
    const normalized = normalizeOperatorName(name);
    return ALLOWED_OPERATOR_NAMES.some((allowed) => normalizeOperatorName(allowed) === normalized);
  }

  function renderMode() {
    if (appMode === "name") {
      setMode("name");
      renderNameGate();
      return;
    }
    if (appMode === "select") {
      setMode("select");
      renderVesselSelect();
      return;
    }
    if (appMode === "login") {
      setMode("login");
      renderLoginScreen();
      return;
    }
    if (appMode === "admin") {
      setMode("admin");
      renderAdminPanel();
      return;
    }

    setMode("app");
    renderShell();
    views[currentView].render();
    updateConnectionBadge();
  }

  function enterApp(view = "cableList") {
    currentView = views[view] ? view : "cableList";
    currentUser ||= { name: "Installation Operator", role: "Standalone" };
    history.replaceState(null, "", `#${currentView}`);
    setMode("app");
    renderMode();
  }

  function renderShell() {
    const nav = el("nav-tabs");
    if (!nav) return;

    nav.innerHTML = Object.entries(views)
      .filter(([, view]) => !view.hidden)
      .map(([key, view]) => `
        <button class="nav-tab ${key === currentView ? "active" : ""}" data-view="${key}" type="button">
          ${view.label}
        </button>`)
      .join("");

    nav.querySelectorAll("[data-view]").forEach((button) => {
      button.addEventListener("click", () => navigate(button.dataset.view));
    });
  }

  function updateConnectionBadge() {
    const vessel = Store.getActiveVessel();
    const badge = el("connection-badge");
    if (badge) {
      badge.textContent = `${vessel.id} Cable List-포설실적용_260205.xlsb · ${Store.getCables().length.toLocaleString("ko-KR")}건`;
    }
    updateUndoButton();
  }

  function updateUndoButton() {
    const btn = el("btn-undo");
    if (!btn) return;
    const can = !!(Store.canUndo && Store.canUndo());
    btn.disabled = !can;
    btn.classList.toggle("is-disabled", !can);
  }

  function navigate(view) {
    if (!views[view]) return;
    currentView = view;
    tablePage = 1;
    history.replaceState(null, "", `#${view}`);
    renderMode();
  }

  function pageHead(title, sub = "", action = "") {
    return `
      <div class="page-head">
        <div>
          <h1 class="page-title">${escapeHtml(title)}</h1>
          ${sub ? `<p class="page-sub">${escapeHtml(sub)}</p>` : ""}
        </div>
        ${action ? `<div class="actions">${action}</div>` : ""}
      </div>`;
  }

  function renderVesselSelect() {
    const state = Store.getState();
    main().innerHTML = `
      <section class="secms-auth-shell">
        <div class="secms-auth-device wide">
          <div class="secms-login-logo-wrap">
            <img class="secms-login-logo" src="https://pub-1f16461a46af495aa0fb95334ed9207f.r2.dev/assets/secms_logo.png" alt="SEcMS">
          </div>
          <div class="secms-kicker">Shipboard Electrical Cable Management System</div>
          <h1 class="secms-login-title">호선 선택</h1>
          <p class="secms-login-sub">처음에는 반드시 작업할 호선을 선택합니다. HK2401 Cable List는 기본으로 준비되어 있습니다.</p>

          <div class="vessel-select-grid">
            ${state.vessels.map((vessel) => {
              const cables = Store.getCables(vessel.id);
              const totalLength = cables.reduce((sum, cable) => sum + (Number(cable.total) || 0), 0);
              return `
                <article class="vessel-card">
                  <div class="vessel-code">${escapeHtml(vessel.id)}</div>
                  <h2>${escapeHtml(vessel.projectName || vessel.name)}</h2>
                  <p>${escapeHtml(vessel.owner || "OWNER 미지정")}</p>
                  <div class="vessel-meta">
                    <span>Cable ${num(cables.length)}건</span>
                    <span>Total ${num(totalLength)}m</span>
                  </div>
                  <button class="secms-primary-btn" data-select-vessel="${escapeAttr(vessel.id)}" type="button">${escapeHtml(vessel.id)} 접속</button>
                </article>`;
            }).join("")}
          </div>

          <details class="secms-fold">
            <summary>새 호선 등록</summary>
            <form id="quick-vessel-form" class="grid three">
              <div class="field"><label>호선번호</label><input name="id" placeholder="HK2402" required></div>
              <div class="field"><label>호선명</label><input name="name" placeholder="HK2402"></div>
              <div class="field"><label>프로젝트명</label><input name="projectName" placeholder="35K FD"></div>
              <div class="field"><label>선주/업체</label><input name="owner" placeholder="YANASE DOCK"></div>
              <div class="field"><label>담당자</label><input name="manager" placeholder="담당자"></div>
              <div class="field"><label>대표전화</label><input name="phone" type="tel" placeholder="010-0000-0000"></div>
              <div class="actions" style="grid-column:1 / -1"><button class="btn green" type="submit">등록</button></div>
            </form>
          </details>

          <div class="secms-footer-note">SEcMS · by bijay kim Copyright © 2026</div>
        </div>
      </section>`;

    document.querySelectorAll("[data-select-vessel]").forEach((button) => {
      button.addEventListener("click", () => {
        Store.setActiveVessel(button.dataset.selectVessel);
        enterApp("cableList");
      });
    });

    el("quick-vessel-form").addEventListener("submit", (event) => {
      event.preventDefault();
      try {
        const vessel = Store.addVessel(formData(event.currentTarget));
        toast(`${vessel.id} 호선을 등록했습니다.`);
        Store.setActiveVessel(vessel.id);
        enterApp("cableList");
      } catch (error) {
        toast(error.message);
      }
    });
  }

  function renderLoginScreen() {
    renderNameGate();
  }

  function renderNameGate() {
    const vessel = Store.getActiveVessel();
    main().innerHTML = `
      <section class="secms-auth-shell">
        <form id="operator-name-form" class="secms-auth-device" autocomplete="off">
          <div class="secms-login-logo-wrap">
            <img class="secms-login-logo" src="https://pub-1f16461a46af495aa0fb95334ed9207f.r2.dev/assets/secms_logo.png" alt="SEcMS" onerror="this.style.display='none'">
          </div>
          <div class="secms-kicker">Shipboard Electrical Cable Management System</div>
          <h1 class="secms-login-title">${escapeHtml(vessel.id)} Cable List</h1>
          <p class="secms-login-sub">허용된 이름과 비밀번호로 접속합니다.</p>
          <label class="name-gate-field">
            <span>이름</span>
            <input name="operatorName" id="operator-name-input" placeholder="이름 입력" required autofocus>
          </label>
          <label class="name-gate-field">
            <span>비밀번호</span>
            <input name="operatorPassword" id="operator-pw-input" type="password" placeholder="비밀번호" required>
          </label>
          <div class="name-gate-actions">
            <button class="secms-primary-btn" type="submit">입장</button>
            <button class="name-gate-admin" id="btn-name-admin" type="button" title="관리자 설정" aria-label="관리자 설정">⚙</button>
          </div>
          <div class="secms-footer-note">SEcMS · by bijay kim Copyright © 2026</div>
        </form>
      </section>`;

    el("operator-name-form").addEventListener("submit", (event) => {
      event.preventDefault();
      const data = formData(event.currentTarget);
      const name = data.operatorName;
      const pw = data.operatorPassword;
      if (!isAllowedOperator(name)) {
        toast("등록된 작업자 이름만 접속할 수 있습니다.");
        el("operator-name-input").select();
        return;
      }
      if (operatorPassword && pw !== operatorPassword) {
        toast("비밀번호가 틀렸습니다.");
        el("operator-pw-input").value = "";
        el("operator-pw-input").focus();
        return;
      }
      currentUser = { name: String(name).trim(), role: "Operator" };
      sessionStorage.setItem(USER_SESSION_KEY, JSON.stringify(currentUser));
      Store.setActiveVessel(Store.DEFAULT_VESSEL_ID);
      enterApp(pendingView || "cableList");
    });
    el("btn-name-admin").addEventListener("click", () => {
      const id = prompt("관리자 ID를 입력하세요");
      if (id === null) return;
      if (adminId && id.trim() !== adminId) {
        toast("관리자 ID가 올바르지 않습니다.");
        return;
      }
      const pw = prompt("관리자 비밀번호를 입력하세요");
      if (pw === null) return;
      if (adminPassword && pw !== adminPassword) {
        toast("관리자 비밀번호가 틀렸습니다.");
        return;
      }
      currentUser = { name: "Admin", role: "Admin" };
      setMode("admin");
      renderMode();
    });
  }

  function renderAdminPanel() {
    const state = Store.getState();
    const active = Store.getActiveVessel();
    const kpi = Store.getKpi();
    main().innerHTML = `
      <section class="admin-shell">
        <header class="admin-header">
          <div class="admin-brand">
            <div class="admin-mark">AD</div>
            <div>
              <h1>Admin</h1>
              <p>SEcMS Installation standalone control</p>
            </div>
            <span class="admin-badge">ADMIN</span>
          </div>
          <div class="admin-actions">
            <button class="btn ghost" id="btn-admin-main" type="button">메인</button>
            <button class="btn ghost" id="btn-admin-select" type="button">호선 선택</button>
            <button class="btn red" id="btn-admin-logout" type="button">Logout</button>
          </div>
        </header>
        <nav class="admin-tabs">
          ${["overview", "projects", "users", "cable", "settings"].map((tab) => `
            <button class="${adminTab === tab ? "active" : ""}" data-admin-tab="${tab}" type="button">${adminLabel(tab)}</button>
          `).join("")}
        </nav>
        <div class="admin-content">${renderAdminContent(adminTab, state, active, kpi)}</div>
      </section>`;

    el("btn-admin-main").addEventListener("click", () => enterApp("cableList"));
    el("btn-admin-select").addEventListener("click", () => {
      setMode("select");
      renderMode();
    });
    el("btn-admin-logout").addEventListener("click", () => {
      currentUser = null;
      sessionStorage.removeItem(USER_SESSION_KEY);
      setMode("name");
      renderMode();
    });
    document.querySelectorAll("[data-admin-tab]").forEach((button) => {
      button.addEventListener("click", () => {
        adminTab = button.dataset.adminTab;
        renderAdminPanel();
      });
    });
  }

  function adminLabel(tab) {
    return ({ overview: "Overview", projects: "Projects", users: "Users", cable: "Cable List", settings: "Settings" })[tab] || tab;
  }

  function renderAdminContent(tab, state, active, kpi) {
    if (tab === "projects") {
      return `
        <section class="panel">
          <h2 class="panel-title">호선 관리</h2>
          <div class="table-wrap">
            <table class="data-table">
              <thead><tr><th>호선</th><th>프로젝트</th><th>선주/업체</th><th>케이블</th><th>담당</th><th>전화</th></tr></thead>
              <tbody>
                ${state.vessels.map((vessel) => `
                  <tr>
                    <td><strong>${escapeHtml(vessel.id)}</strong></td>
                    <td>${escapeHtml(vessel.projectName || vessel.name)}</td>
                    <td>${escapeHtml(vessel.owner || "-")}</td>
                    <td class="num">${num(Store.getCables(vessel.id).length)}</td>
                    <td>${escapeHtml(vessel.manager || "-")}</td>
                    <td>${escapeHtml(vessel.phone || "-")}</td>
                  </tr>
                `).join("")}
              </tbody>
            </table>
          </div>
        </section>`;
    }

    if (tab === "users") {
      return `
        <section class="panel">
          <h2 class="panel-title">사용자/권한</h2>
          <div class="admin-user-list">
            <div><strong>Installation Operator</strong><span>User · Cable List 수정</span></div>
            <div><strong>Admin</strong><span>Admin · 호선/불러오기/내보내기 관리</span></div>
            <div><strong>SEcMS 통합 예정</strong><span>Firebase/권한 체계 연결 대기</span></div>
          </div>
        </section>`;
    }

    if (tab === "cable") {
      return `
        <section class="panel">
          <h2 class="panel-title">Cable List 운영 상태</h2>
          <div class="stat-grid">
            <div class="stat"><div class="stat-label">Active Vessel</div><div class="stat-value">${escapeHtml(active.id)}</div></div>
            <div class="stat"><div class="stat-label">Rows</div><div class="stat-value">${num(Store.getCables().length)}</div></div>
            <div class="stat green"><div class="stat-label">Installed</div><div class="stat-value">${pct(kpi.installRate)}</div></div>
            <div class="stat orange"><div class="stat-label">REV Records</div><div class="stat-value">${num(Store.getCables().filter((cable) => cable.rev).length)}</div></div>
          </div>
        </section>`;
    }

    if (tab === "settings") {
      return `
        <section class="panel">
          <h2 class="panel-title">통합 준비 설정</h2>
          <div class="setting-grid">
            <div><strong>Standalone Build</strong><span>현재는 로컬 브라우저 저장소 기반으로 동작합니다.</span></div>
            <div><strong>SEcMS Installation</strong><span>나중에 E:\\code-project\\SEcMS의 Installation 메뉴로 이관할 수 있도록 화면과 용어를 맞췄습니다.</span></div>
            <div><strong>Dual Service</strong><span>Cable List 업데이트 로직은 독립 모듈화되어 이원화 서비스에 재사용 가능합니다.</span></div>
          </div>
        </section>`;
    }

    return `
      <div class="stat-grid">
        <div class="stat"><div class="stat-label">호선</div><div class="stat-value">${num(state.vessels.length)}</div></div>
        <div class="stat"><div class="stat-label">Cable Rows</div><div class="stat-value">${num(kpi.registeredCount || kpi.totalCount)}</div></div>
        <div class="stat green"><div class="stat-label">Total Length</div><div class="stat-value">${num(kpi.totalLength)}m</div></div>
        <div class="stat orange"><div class="stat-label">Phone</div><div class="stat-value">${num(Store.getPhones().length)}</div></div>
      </div>
      <section class="panel">
        <h2 class="panel-title">관리자 개요</h2>
        <p class="admin-copy">호선 선택, 로그인, Cable List 편집, Excel 업데이트, REV 이력 누적 기능을 SEcMS 스타일로 단독 빌드했습니다.</p>
      </section>`;
  }

  function renderShipManage() {
    const state = Store.getState();
    const active = Store.getActiveVessel();
    const tabs = [
      { id: "switch", label: "🔁 호선 이동" },
      { id: "register", label: "🚢 호선 등록" },
      { id: "cable", label: "➕ 호선케이블 등록" },
    ];
    if (!tabs.find((t) => t.id === shipTab)) shipTab = "switch";

    main().innerHTML = `
      ${pageHead("호선 관리", "현재 접속 호선 이동, 신규 호선 등록, 케이블 등록을 한 화면에서 관리합니다.")}

      <section class="panel">
        <h2 class="panel-title">현재 접속 호선</h2>
        <div class="stat-grid">
          <div class="stat"><div class="stat-label">호선번호</div><div class="stat-value">${escapeHtml(active.id)}</div></div>
          <div class="stat green"><div class="stat-label">Cable List</div><div class="stat-value">${num(Store.getCables().length)}</div></div>
          <div class="stat orange"><div class="stat-label">전화</div><div class="stat-value">${num(Store.getPhones().length)}</div></div>
          <div class="stat"><div class="stat-label">프로젝트</div><div class="stat-value">${escapeHtml(active.projectName || active.name)}</div></div>
        </div>
      </section>

      <div class="cable-tabs ship-tabs">
        ${tabs.map((tab) => `
          <button class="cable-tab ${shipTab === tab.id ? "active" : ""}" data-ship-tab="${tab.id}" type="button">${tab.label}</button>
        `).join("")}
      </div>

      <div id="ship-tab-body">${renderShipTabBody(state, active)}</div>`;

    document.querySelectorAll("[data-ship-tab]").forEach((button) => {
      button.addEventListener("click", () => {
        shipTab = button.dataset.shipTab;
        renderShipManage();
      });
    });

    bindShipTabHandlers();
  }

  function renderShipTabBody(state, active) {
    if (shipTab === "switch") {
      return `
        <section class="panel">
          <h2 class="panel-title">호선 이동 · 접속 호선 변경</h2>
          <p class="panel-desc">접속할 호선을 선택하면 Cable List, 실적, 달력 뷰가 해당 호선의 데이터로 즉시 전환됩니다.</p>
          <div class="list">
            ${state.vessels.map((vessel) => `
              <div class="item">
                <div>
                  <div class="item-title">${escapeHtml(vessel.id)} · ${escapeHtml(vessel.projectName || vessel.name)}</div>
                  <div class="item-meta">선주/업체 ${escapeHtml(vessel.owner || "-")} · 담당 ${escapeHtml(vessel.manager || "-")} · 전화 ${escapeHtml(vessel.phone || "-")}</div>
                </div>
                <div class="actions">
                  ${vessel.phone ? `<a class="btn ghost small" href="tel:${encodeURIComponent(vessel.phone)}">전화</a>` : ""}
                  <button class="btn small ${vessel.id === active.id ? "green" : ""}" data-connect="${escapeAttr(vessel.id)}" type="button">
                    ${vessel.id === active.id ? "접속 중" : "이 호선으로 이동"}
                  </button>
                </div>
              </div>`).join("")}
          </div>
        </section>`;
    }

    if (shipTab === "register") {
      return `
        <section class="panel">
          <h2 class="panel-title">새 호선 등록</h2>
          <form id="vessel-form" class="grid three">
            <div class="field"><label>호선번호</label><input name="id" placeholder="HK2402" required></div>
            <div class="field"><label>호선명</label><input name="name" placeholder="HK2402"></div>
            <div class="field"><label>프로젝트명</label><input name="projectName" placeholder="35K FD"></div>
            <div class="field"><label>선주/업체</label><input name="owner" placeholder="YANASE DOCK"></div>
            <div class="field"><label>담당자</label><input name="manager" placeholder="담당자"></div>
            <div class="field"><label>대표전화</label><input name="phone" type="tel" placeholder="010-0000-0000"></div>
            <div class="field" style="grid-column:1 / -1"><label>메모</label><textarea name="memo" placeholder="호선 관련 메모"></textarea></div>
            <div class="actions" style="grid-column:1 / -1"><button class="btn" type="submit">호선 등록</button></div>
          </form>
        </section>`;
    }

    // shipTab === "cable"
    return `
      <section class="panel">
        <div class="panel-head-row">
          <h2 class="panel-title">${escapeHtml(active.id)} 호선 케이블 등록</h2>
          <button class="btn ghost small" id="btn-reset-default" type="button">HK2401 Cable List 다시 불러오기</button>
        </div>
        <form id="cable-form" class="grid six">
          <div class="field"><label>SYS</label><input name="sys" maxlength="4" placeholder="L"></div>
          <div class="field"><label>CIRCUIT NO.</label><input name="circuitNo" placeholder="L0101" required></div>
          <div class="field"><label>CABLE TYPE</label><input name="cableType" placeholder="DE1" required></div>
          <div class="field"><label>CABLE DIA</label><input name="cableDia" inputmode="decimal" placeholder="13"></div>
          <div class="field"><label>TOTAL</label><input name="total" inputmode="decimal" placeholder="55"></div>
          <div class="field"><label>NODE구분</label><input name="node" placeholder="ACC"></div>

          <div class="field"><label>FROM EQUIPMENT</label><input name="fromEquipment" placeholder="W/H GROUP PANEL"></div>
          <div class="field"><label>FROM CODE</label><input name="fromCode" placeholder="ANP"></div>
          <div class="field"><label>FM MARJIN</label><input name="fmMargin" placeholder="5"></div>
          <div class="field"><label>TO EQUIPMENT</label><input name="toEquipment" placeholder="FWD MASTHEAD LIGHT"></div>
          <div class="field"><label>TO CODE</label><input name="toCode" placeholder="FM"></div>
          <div class="field"><label>TO MARJIN</label><input name="toMargin" placeholder="20"></div>

          <div class="field" style="grid-column:1 / span 3"><label>ROUTE</label><input name="route" placeholder="local 또는 AU05,AU18"></div>
          <div class="field"><label>포설일자</label><input name="installDate" type="date"></div>
          <div class="field"><label>결선 FROM</label><input name="conFromDate" type="date"></div>
          <div class="field"><label>결선 TO</label><input name="conToDate" type="date"></div>
          <div class="field" style="grid-column:1 / -1"><label>REV</label><input name="rev" placeholder="추가/수정 사유"></div>

          <div class="actions" style="grid-column:1 / -1">
            <button class="btn green" type="submit">케이블 등록</button>
            <button class="btn ghost" type="reset">초기화</button>
          </div>
        </form>
      </section>`;
  }

  function bindShipTabHandlers() {
    if (shipTab === "switch") {
      document.querySelectorAll("[data-connect]").forEach((button) => {
        button.addEventListener("click", () => {
          Store.setActiveVessel(button.dataset.connect);
          toast(`${button.dataset.connect} 호선으로 이동했습니다.`);
          renderMode();
        });
      });
      return;
    }

    if (shipTab === "register") {
      const form = el("vessel-form");
      if (form) {
        form.addEventListener("submit", (event) => {
          event.preventDefault();
          try {
            const vessel = Store.addVessel(formData(event.currentTarget));
            toast(`${vessel.id} 호선을 등록했습니다.`);
            event.currentTarget.reset();
            shipTab = "switch";
            renderShipManage();
          } catch (error) {
            toast(error.message);
          }
        });
      }
      return;
    }

    if (shipTab === "cable") {
      const cableForm = el("cable-form");
      if (cableForm) {
        cableForm.addEventListener("submit", (event) => {
          event.preventDefault();
          try {
            const cable = Store.addCable(formData(event.currentTarget));
            toast(`${cable.circuitNo} 케이블을 등록했습니다.`);
            event.currentTarget.reset();
            updateConnectionBadge();
          } catch (error) {
            toast(error.message);
          }
        });
      }
      const resetBtn = el("btn-reset-default");
      if (resetBtn) {
        resetBtn.addEventListener("click", async () => {
          if (!confirm("현재 호선의 Cable List를 HK2401 기본 Cable List로 교체할까요?")) return;
          try {
            const count = await Store.loadDefaultCables({ replace: true });
            toast(`HK2401 Cable List ${count.toLocaleString("ko-KR")}건을 불러왔습니다.`);
            renderMode();
          } catch (error) {
            toast(error.message);
          }
        });
      }
    }
  }

  function renderDashboard() {
    const vessel = Store.getActiveVessel();
    const cables = activeCables();
    const kpi = Store.getKpi();
    const byType = aggregateRows(cables, (cable) => cable.cableType).slice(0, 16);
    const byNode = aggregateRows(cables, (cable) => cable.node).slice(0, 10);
    const connectionDone = Math.min(kpi.conFrom, kpi.conTo);
    const connectionRate = kpi.totalCount ? connectionDone / kpi.totalCount : 0;

    main().innerHTML = `
      <section class="legacy-shell legacy-dashboard">
        <div class="legacy-page-head">
          <div>
            <h1>포설·결선 현황 대시보드</h1>
            <p>CABLE LIST(0424) · ${num(kpi.registeredCount || kpi.totalCount)}건 · ${escapeHtml(formatKoreanDateTime(new Date()))}</p>
          </div>
          <div class="legacy-actions">
            <button class="btn ghost" id="btn-dashboard-report" type="button">📄 보고서 내보내기</button>
            <button class="btn ghost" id="btn-dashboard-export" type="button">📦 전체 내보내기</button>
          </div>
        </div>

        <div class="legacy-kpi-grid">
          ${legacyKpiCard("📏", "포설 총량", `${num(kpi.totalLength)} m`, "전체 케이블")}
          ${legacyKpiCard("✅", "포설 완료", `${num(kpi.installedLength)} m`, `${pct(kpi.installRate)} 달성`, "green")}
          ${legacyKpiCard("⌛", "포설 미완료", `${num(kpi.pendingLength)} m`, "잔여 물량", "orange")}
          ${legacyKpiCard("🔌", "결선FROM 완료", `${num(kpi.conFrom)} EA`, `${pct(kpi.conFromRate)} 달성`, "blue")}
          ${legacyKpiCard("🔋", "결선TO 완료", `${num(kpi.conTo)} EA`, `${pct(kpi.conToRate)} 달성`, "purple")}
          ${legacyKpiCard("📋", "전체 케이블", `${num(kpi.registeredCount || kpi.totalCount)} EA`, "등록 수량")}
        </div>

        <section class="legacy-progress-card">
          <div class="progress-label">포설 진행률</div>
          <div class="legacy-progress-track"><span style="width:${Math.round(kpi.installRate * 100)}%"></span></div>
          <b>${pct(kpi.installRate)}</b>
          <div class="progress-label">결선FROM</div>
          <div class="legacy-progress-track blue"><span style="width:${Math.round(kpi.conFromRate * 100)}%"></span></div>
          <b>${pct(kpi.conFromRate)}</b>
          <div class="progress-label">결선TO</div>
          <div class="legacy-progress-track purple"><span style="width:${Math.round(kpi.conToRate * 100)}%"></span></div>
          <b>${pct(kpi.conToRate)}</b>
        </section>

        <div class="legacy-chart-grid">
          ${legacyGroupedBar("TYPE별 전체/포설 물량 (m)", byType)}
          ${legacyConnectionBar("NODE별 결선 진행률 (EA)", byNode)}
          ${legacyDonutCard("포설 진행률 도넛", kpi.installRate, "포설완료", "미포설", "green")}
          ${legacyDonutCard("결선 진행률 도넛", connectionRate, "결선FROM완료", "결선TO완료", "blue")}
        </div>

        <section class="legacy-section">
          <h2>TYPE별 집계</h2>
          ${legacySummaryTable("CABLE TYPE", byType)}
        </section>
      </section>`;

    el("btn-dashboard-report").addEventListener("click", () => navigate("report"));
    el("btn-dashboard-export").addEventListener("click", () => {
      download(`${vessel.id}-Cable-List.tsv`, Store.toTsv());
      toast("UTF-8 BOM 포함 TSV로 내보냈습니다.");
    });
  }

  function formatKoreanDateTime(date) {
    return date.toLocaleString("ko-KR", {
      year: "numeric",
      month: "numeric",
      day: "numeric",
      weekday: "short",
      hour: "numeric",
      minute: "2-digit",
    });
  }

  function legacyKpiCard(icon, label, value, caption, tone = "") {
    return `
      <section class="legacy-kpi ${tone}">
        <div class="legacy-kpi-icon">${icon}</div>
        <div>
          <span>${escapeHtml(label)}</span>
          <strong>${escapeHtml(value)}</strong>
          <small>${escapeHtml(caption)}</small>
        </div>
      </section>`;
  }

  function legacyGroupedBar(title, rows) {
    const max = Math.max(...rows.map((row) => row.totalLength), 1);
    const chartH = 180;
    return `
      <section class="legacy-chart-card">
        <h2>${escapeHtml(title)}</h2>
        <div class="legacy-bar-legend"><span class="green"></span>포설완료 <span class="orange"></span>미포설 <small style="margin-left:6px;color:#94a3b8;">(전체물량 = 포설 + 미포설)</small></div>
        <div class="legacy-stacked-chart" style="--chart-h:${chartH}px;">
          ${rows.map((row) => {
            const totalH = Math.round((row.totalLength / max) * chartH);
            const installedShare = row.totalLength ? row.installedLength / row.totalLength : 0;
            const installedH = Math.round(totalH * installedShare);
            const pendingH = Math.max(0, totalH - installedH);
            const rate = Math.round(installedShare * 100);
            return `
              <div class="legacy-stacked-column" title="${escapeAttr(row.key)} 전체 ${num(row.totalLength)}m / 포설 ${num(row.installedLength)}m (${rate}%)">
                <div class="stacked-bar" style="height:${totalH}px">
                  <span class="seg pending" style="height:${pendingH}px"></span>
                  <span class="seg installed" style="height:${installedH}px"></span>
                </div>
                <em>${rate}%</em>
                <b>${escapeHtml(row.key)}</b>
              </div>`;
          }).join("")}
        </div>
      </section>`;
  }

  function legacyConnectionBar(title, rows) {
    const max = Math.max(...rows.map((row) => row.count), 1);
    return `
      <section class="legacy-chart-card">
        <h2>${escapeHtml(title)}</h2>
        <div class="legacy-bar-legend"><span class="blue-leg"></span>결선FROM <span class="purple-leg"></span>결선TO <small style="margin-left:6px;color:#94a3b8;">(구역 케이블 수 대비)</small></div>
        <div class="legacy-conn-bars">
          ${rows.map((row) => {
            const fromRate = row.count ? row.conFromCount / row.count : 0;
            const toRate = row.count ? row.conToCount / row.count : 0;
            const fromW = Math.round((row.conFromCount / max) * 100);
            const toW = Math.round((row.conToCount / max) * 100);
            return `
              <div class="legacy-conn-row" title="${escapeAttr(row.key)} FROM ${num(row.conFromCount)}/${num(row.count)} TO ${num(row.conToCount)}/${num(row.count)}">
                <b>${escapeHtml(row.key)}</b>
                <div class="conn-stack">
                  <span class="conn-track"><i class="from-fill" style="width:${fromW}%"></i><em>${Math.round(fromRate * 100)}%</em></span>
                  <span class="conn-track"><i class="to-fill" style="width:${toW}%"></i><em>${Math.round(toRate * 100)}%</em></span>
                </div>
              </div>`;
          }).join("")}
        </div>
      </section>`;
  }

  function legacyHorizontalBar(title, rows) {
    const max = Math.max(...rows.map((row) => row.totalLength), 1);
    return `
      <section class="legacy-chart-card">
        <h2>${escapeHtml(title)}</h2>
        <div class="legacy-bar-legend"><span class="green"></span>포설완료 <span class="orange"></span>미포설</div>
        <div class="legacy-node-bars">
          ${rows.map((row) => {
            const installed = Math.round((row.installedLength / max) * 100);
            const pending = Math.round(((row.totalLength - row.installedLength) / max) * 100);
            return `
              <div class="legacy-node-row">
                <b>${escapeHtml(row.key)}</b>
                <span class="node-track">
                  <i class="installed" style="width:${installed}%"></i>
                  <i class="pending" style="width:${pending}%"></i>
                </span>
              </div>`;
          }).join("")}
        </div>
      </section>`;
  }

  function legacyDonutCard(title, value, completeLabel, pendingLabel, tone) {
    const safe = Math.max(0, Math.min(1, Number(value || 0)));
    return `
      <section class="legacy-chart-card legacy-donut-card">
        <h2>${escapeHtml(title)}</h2>
        <div class="legacy-donut ${tone}" style="--value:${Math.round(safe * 100)}"></div>
        <div class="legacy-donut-legend">
          <span class="${tone}"></span>${escapeHtml(completeLabel)}
          <span class="empty"></span>${escapeHtml(pendingLabel)}
        </div>
      </section>`;
  }

  function legacySummaryTable(firstLabel, rows) {
    const half = Math.ceil(rows.length / 2);
    const left = rows.slice(0, half);
    const right = rows.slice(half);
    function tableHtml(group) {
      return `
        <div class="legacy-table-wrap">
          <table class="legacy-table">
            <thead>
              <tr>
                <th>${escapeHtml(firstLabel)}</th>
                <th>포설완료(m)</th>
                <th>미포설(m)</th>
                <th>합계(m)</th>
                <th>진행률</th>
                <th>결선FROM</th>
                <th>결선TO</th>
                <th>총수량</th>
              </tr>
            </thead>
            <tbody>
              ${group.map((row) => {
                const rate = row.totalLength ? row.installedLength / row.totalLength : 0;
                return `
                  <tr>
                    <td>${escapeHtml(row.key)}</td>
                    <td class="num good">${num(row.installedLength)}</td>
                    <td class="num warn">${num(Math.max(0, row.totalLength - row.installedLength))}</td>
                    <td class="num">${num(row.totalLength)}</td>
                    <td class="rate">${pct(rate)}</td>
                    <td class="num blue-text">${num(row.conFromCount)}</td>
                    <td class="num purple-text">${num(row.conToCount)}</td>
                    <td class="num">${num(row.count)}</td>
                  </tr>`;
              }).join("")}
            </tbody>
          </table>
        </div>`;
    }
    return `<div class="legacy-table-two-col">${tableHtml(left)}${right.length ? tableHtml(right) : ""}</div>`;
  }

  function systemName(sys) {
    return ({
      C: "CONTROL",
      F: "FIRE",
      L: "LIGHT",
      N: "NAVI",
      P: "POWER",
    })[String(sys || "").trim().toUpperCase()] || (sys || "미분류");
  }

  function activeCables() {
    return Store.getCables().filter((cable) => !cable.deleted);
  }

  function aggregateRows(cables, keyFn) {
    const map = new Map();
    cables.forEach((cable) => {
      const key = keyFn(cable) || "미분류";
      const current = map.get(key) || {
        key,
        count: 0,
        totalLength: 0,
        installedCount: 0,
        installedLength: 0,
        conFromCount: 0,
        conToCount: 0,
        conFromLength: 0,
        conToLength: 0,
      };
      const length = Number(cable.total) || 0;
      current.count += 1;
      current.totalLength += length;
      if (cable.installDate) {
        current.installedCount += 1;
        current.installedLength += length;
      }
      if (cable.conFromDate) {
        current.conFromCount += 1;
        current.conFromLength += length;
      }
      if (cable.conToDate) {
        current.conToCount += 1;
        current.conToLength += length;
      }
      map.set(key, current);
    });
    return [...map.values()].sort((a, b) => b.totalLength - a.totalLength || a.key.localeCompare(b.key, "ko"));
  }

  function progressBarCell(rate) {
    const safe = Math.max(0, Math.min(1, Number(rate || 0)));
    return `<span class="excel-progress" style="--rate:${Math.round(safe * 100)}"><b>${pct(safe)}</b></span>`;
  }

  function compactDate(date = new Date()) {
    const year = String(date.getFullYear()).slice(2);
    const month = String(date.getMonth() + 1).padStart(2, "0");
    const day = String(date.getDate()).padStart(2, "0");
    return `${year}.${month}.${day}`;
  }

  function orderedSystemRows(cables) {
    const rows = aggregateRows(cables, (cable) => systemName(cable.sys));
    const order = ["POWER", "LIGHT", "CONTROL", "FIRE", "NAVI"];
    return [
      ...order.map((name) => rows.find((row) => row.key === name)).filter(Boolean),
      ...rows.filter((row) => !order.includes(row.key)),
    ];
  }

  function reportProgress(rate) {
    const safe = Math.max(0, Math.min(1, Number(rate || 0)));
    return `<span class="copy-progress" style="--rate:${Math.round(safe * 100)}"><b>${Math.round(safe * 100)}%</b></span>`;
  }

  function renderInstallCopyTable(vesselId, rows, kpi) {
    const totalRows = rows.slice(0, 10);
    return `
      <table class="copy-report-table install">
        <thead>
          <tr><th>호선</th><th>구역</th><th>TOTAL</th><th>포설</th><th>미포설</th><th>진도율</th></tr>
        </thead>
        <tbody>
          ${totalRows.map((row, index) => {
            const rate = row.totalLength ? row.installedLength / row.totalLength : 0;
            return `<tr>
              ${index === 0 ? `<td class="vessel-cell" rowspan="${totalRows.length}">${escapeHtml(vesselId)}</td>` : ""}
              <td>${escapeHtml(row.key)}</td>
              <td class="num">${num(Math.round(row.totalLength))}</td>
              <td class="num blue-text">${num(Math.round(row.installedLength))}</td>
              <td class="num red-text">${num(Math.round(Math.max(0, row.totalLength - row.installedLength)))}</td>
              <td>${reportProgress(rate)}</td>
            </tr>`;
          }).join("")}
          <tr class="sum-row">
            <td colspan="2">합계</td>
            <td class="num">${num(Math.round(kpi.totalLength))}</td>
            <td class="num blue-text">${num(Math.round(kpi.installedLength))}</td>
            <td class="num red-text">${num(Math.round(kpi.pendingLength))}</td>
            <td>${reportProgress(kpi.installRate)}</td>
          </tr>
        </tbody>
      </table>`;
  }

  function renderConnectionCopyTable(vesselId, rows) {
    const displayRows = rows.slice(0, 8);
    const totalCount = rows.reduce((sum, row) => sum + row.count, 0);
    const totalFrom = rows.reduce((sum, row) => sum + row.conFromCount, 0);
    const totalTo = rows.reduce((sum, row) => sum + row.conToCount, 0);
    return `
      <table class="copy-report-table connection">
        <thead>
          <tr>
            <th rowspan="2">호선</th><th rowspan="2">System</th>
            <th colspan="4">From</th><th colspan="4">To</th>
          </tr>
          <tr><th>TOTAL</th><th>결선</th><th>미결선</th><th>진도율</th><th>TOTAL</th><th>결선</th><th>미결선</th><th>진도율</th></tr>
        </thead>
        <tbody>
          ${displayRows.map((row, index) => {
            const fromRate = row.count ? row.conFromCount / row.count : 0;
            const toRate = row.count ? row.conToCount / row.count : 0;
            return `<tr>
              ${index === 0 ? `<td class="vessel-cell" rowspan="${displayRows.length}">${escapeHtml(vesselId)}</td>` : ""}
              <td class="system-cell">${escapeHtml(row.key)}</td>
              <td class="num">${num(row.count)}</td>
              <td class="num blue-text">${num(row.conFromCount)}</td>
              <td class="num">${num(Math.max(0, row.count - row.conFromCount))}</td>
              <td>${reportProgress(fromRate)}</td>
              <td class="num">${num(row.count)}</td>
              <td class="num purple-text">${num(row.conToCount)}</td>
              <td class="num">${num(Math.max(0, row.count - row.conToCount))}</td>
              <td>${reportProgress(toRate)}</td>
            </tr>`;
          }).join("")}
          <tr class="sum-row">
            <td colspan="2">합계</td>
            <td class="num">${num(totalCount)}</td>
            <td class="num blue-text">${num(totalFrom)}</td>
            <td class="num">${num(Math.max(0, totalCount - totalFrom))}</td>
            <td>${reportProgress(totalCount ? totalFrom / totalCount : 0)}</td>
            <td class="num">${num(totalCount)}</td>
            <td class="num purple-text">${num(totalTo)}</td>
            <td class="num">${num(Math.max(0, totalCount - totalTo))}</td>
            <td>${reportProgress(totalCount ? totalTo / totalCount : 0)}</td>
          </tr>
        </tbody>
      </table>`;
  }

  function collectCopyDaily(year, month) {
    const prefix = `${year}-${String(month + 1).padStart(2, "0")}`;
    const map = new Map();
    activeCables().forEach((cable) => {
      const length = Number(cable.total) || 0;
      const add = (date, key) => {
        if (!date || !date.startsWith(prefix)) return;
        const current = map.get(date) || { installCount: 0, installLength: 0, connectionCount: 0, connectionLength: 0 };
        if (key === "install") {
          current.installCount += 1;
          current.installLength += length;
        } else {
          current.connectionCount += 1;
          current.connectionLength += length;
        }
        map.set(date, current);
      };
      add(cable.installDate, "install");
      add(cable.conFromDate, "connection");
      add(cable.conToDate, "connection");
    });
    return map;
  }

  function renderCopyCalendar() {
    const daily = collectCopyDaily(calendarYear, calendarMonth);
    return `
      <div class="copy-calendar">
        <h3>${calendarYear}년 ${calendarMonth + 1}월 실적 달력</h3>
        <div class="copy-calendar-week"><span>일</span><span>월</span><span>화</span><span>수</span><span>목</span><span>금</span><span>토</span></div>
        <div class="copy-calendar-grid">
          ${monthCells(calendarYear, calendarMonth).map((cell) => {
            const record = cell.date ? daily.get(cell.date) : null;
            return `<div class="copy-calendar-cell ${cell.date ? "" : "empty"}">
              <b>${cell.day || ""}</b>
              ${record?.installCount ? `<em class="install">포설 ${num(record.installCount)}건 · ${num(Math.round(record.installLength))}m</em>` : ""}
              ${record?.connectionCount ? `<em class="connection">결선 ${num(record.connectionCount)}건 · ${num(Math.round(record.connectionLength))}m</em>` : ""}
            </div>`;
          }).join("")}
        </div>
      </div>`;
  }

  function renderCopyBarChart(rows) {
    const topRows = rows.slice(0, 8);
    const max = Math.max(1, ...topRows.map((row) => row.totalLength));
    return `
      <div class="copy-chart">
        <h3>구역별 포설 그래프</h3>
        ${topRows.map((row) => {
          const installed = Math.round((row.installedLength / max) * 100);
          const pending = Math.round((Math.max(0, row.totalLength - row.installedLength) / max) * 100);
          return `<p>
            <span>${escapeHtml(row.key)}</span>
            <i><em class="done" style="width:${installed}%"></em><em class="todo" style="width:${pending}%"></em></i>
            <b>${pct(row.totalLength ? row.installedLength / row.totalLength : 0)}</b>
          </p>`;
        }).join("")}
      </div>`;
  }

  function renderReport() {
    const vessel = Store.getActiveVessel();
    const cables = activeCables();
    const byNode = aggregateRows(cables, (cable) => cable.node).slice(0, 12);
    const bySystem = orderedSystemRows(cables);
    const kpi = Store.getKpi();
    const weeks = monthWeeks(calendarYear, calendarMonth);
    if (reportWeek > weeks.length) reportWeek = weeks.length;
    const activeWeek = weeks[reportWeek - 1] || weeks[0];

    main().innerHTML = `
      <section class="legacy-shell">
        <div class="report-toolbar">
          <div class="report-period-bar">
            <div class="report-period-tabs" role="tablist">
              <button class="report-period-tab ${reportPeriod === "month" ? "active" : ""}" data-period="month" type="button" role="tab" aria-selected="${reportPeriod === "month"}">월별 보고서</button>
              <button class="report-period-tab ${reportPeriod === "week" ? "active" : ""}" data-period="week" type="button" role="tab" aria-selected="${reportPeriod === "week"}">주차별 보고서</button>
            </div>
            <div class="report-period-controls">
              <button class="btn ghost small" id="btn-report-prev-month" type="button">◀</button>
              <select id="report-month-select" class="report-month-select">${monthOptions()}</select>
              <button class="btn ghost small" id="btn-report-next-month" type="button">▶</button>
              ${reportPeriod === "week" ? `
                <span class="report-period-divider"></span>
                <div class="report-week-chips">
                  ${weeks.map((week, index) => `
                    <button class="report-week-chip ${reportWeek === index + 1 ? "active" : ""}" data-week="${index + 1}" type="button">
                      ${index + 1}주차 <span>${week.startLabel}~${week.endLabel}</span>
                    </button>
                  `).join("")}
                </div>
              ` : ""}
            </div>
          </div>
          <div class="report-toolbar-actions">
            <button class="btn" id="btn-report-print" type="button">🖨 인쇄 / PDF</button>
            <button class="btn ghost" id="btn-report-export" type="button">📥 엑셀 저장</button>
          </div>
        </div>

        ${reportPeriod === "month"
          ? renderMonthlyReportSheet(vessel, byNode, bySystem, kpi)
          : renderWeeklyReportSheet(vessel, cables, byNode, bySystem, activeWeek)}
      </section>`;

    el("btn-report-print").addEventListener("click", () => window.print());
    el("btn-report-export").addEventListener("click", exportCalendarReport);
    document.querySelectorAll("[data-period]").forEach((button) => {
      button.addEventListener("click", () => {
        reportPeriod = button.dataset.period;
        renderReport();
      });
    });
    document.querySelectorAll("[data-week]").forEach((button) => {
      button.addEventListener("click", () => {
        reportWeek = Number(button.dataset.week);
        renderReport();
      });
    });
    const monthSelect = el("report-month-select");
    if (monthSelect) {
      monthSelect.value = `${calendarYear}-${String(calendarMonth + 1).padStart(2, "0")}`;
      monthSelect.addEventListener("change", (event) => {
        const [year, month] = event.target.value.split("-").map(Number);
        calendarYear = year;
        calendarMonth = month - 1;
        reportWeek = 1;
        renderReport();
      });
    }
    const prev = el("btn-report-prev-month");
    const next = el("btn-report-next-month");
    if (prev) prev.addEventListener("click", () => { stepReportMonth(-1); });
    if (next) next.addEventListener("click", () => { stepReportMonth(1); });
  }

  function stepReportMonth(delta) {
    const d = new Date(calendarYear, calendarMonth + delta, 1);
    calendarYear = d.getFullYear();
    calendarMonth = d.getMonth();
    reportWeek = 1;
    renderReport();
  }

  function renderMonthlyReportSheet(vessel, byNode, bySystem, kpi) {
    return `
      <article class="copy-report-sheet">
        <div class="copy-report-date">${compactDate()}</div>
        <section class="copy-report-main">
          <div class="copy-report-title">◎ ${escapeHtml(vessel.id)} CABLE 포설 현황</div>
          <div class="copy-report-owner">${calendarYear}년 ${calendarMonth + 1}월 누적</div>
          ${renderInstallCopyTable(vessel.id, byNode, kpi)}

          <div class="copy-report-title second">◎ ${escapeHtml(vessel.id)} CABLE 결선 현황</div>
          <div class="copy-report-owner">${calendarYear}년 ${calendarMonth + 1}월 누적</div>
          ${renderConnectionCopyTable(vessel.id, bySystem)}
        </section>

        <aside class="copy-report-side">
          <div class="copy-kpi-row">
            <div><span>포설 총량</span><b>${num(Math.round(kpi.totalLength))}m</b></div>
            <div><span>포설 완료</span><b>${num(Math.round(kpi.installedLength))}m</b></div>
            <div><span>미포설</span><b>${num(Math.round(kpi.pendingLength))}m</b></div>
          </div>
          ${renderCopyBarChart(byNode)}
          ${renderCopyCalendar()}
          <div class="copy-report-note">
            <b>REV 기준:</b> Cable List의 REV 열 변경 이력 기준<br>
            <b>기준 파일:</b> ${escapeHtml(vessel.id)} Cable List-포설실적용_260205.xlsb
          </div>
        </aside>
      </article>`;
  }

  function renderWeeklyReportSheet(vessel, cables, byNode, bySystem, week) {
    const weekStats = collectWeekStats(cables, week);
    const dailyMap = collectCopyDaily(calendarYear, calendarMonth);
    const dayCells = week.days.map((day) => {
      const key = day.toISOString().slice(0, 10);
      const record = dailyMap.get(key);
      const dow = ["일", "월", "화", "수", "목", "금", "토"][day.getDay()];
      return `<div class="weekly-day-cell ${day.getDay() === 0 ? "sun" : day.getDay() === 6 ? "sat" : ""}">
        <header><b>${day.getMonth() + 1}/${day.getDate()}</b><span>${dow}</span></header>
        <div class="weekly-day-body">
          ${record?.installCount ? `<em class="install">포설 ${num(record.installCount)}건<br><b>${num(Math.round(record.installLength))}m</b></em>` : `<em class="empty">—</em>`}
          ${record?.connectionCount ? `<em class="connection">결선 ${num(record.connectionCount)}건<br><b>${num(Math.round(record.connectionLength))}m</b></em>` : ""}
        </div>
      </div>`;
    }).join("");

    return `
      <article class="copy-report-sheet weekly-sheet">
        <div class="copy-report-date">${compactDate()}</div>
        <header class="weekly-header">
          <div>
            <div class="copy-report-title weekly-title">◎ ${escapeHtml(vessel.id)} ${calendarYear}년 ${calendarMonth + 1}월 ${week.index}주차 실적 보고</div>
            <div class="copy-report-owner">${week.startLabel} ~ ${week.endLabel} (${week.days.length}일)</div>
          </div>
          <div class="weekly-kpi-row">
            <div><span>주간 포설</span><b>${num(weekStats.installCount)}건</b><i>${num(Math.round(weekStats.installLength))}m</i></div>
            <div><span>주간 결선 FROM</span><b>${num(weekStats.conFromCount)}건</b><i>${num(Math.round(weekStats.conFromLength))}m</i></div>
            <div><span>주간 결선 TO</span><b>${num(weekStats.conToCount)}건</b><i>${num(Math.round(weekStats.conToLength))}m</i></div>
            <div><span>주간 진도 변화</span><b>+${pct(weekStats.installLength / Math.max(1, weekStats.totalLength))}</b><i>${num(Math.round(weekStats.totalLength))}m 기준</i></div>
          </div>
        </header>

        <section class="weekly-grid">
          <div class="weekly-card weekly-day-strip-card">
            <h3>일자별 실적 (${week.startLabel} ~ ${week.endLabel})</h3>
            <div class="weekly-day-strip">${dayCells}</div>
          </div>

          <div class="weekly-card">
            <h3>주간 구역별 포설 실적</h3>
            ${renderWeeklyNodeTable(weekStats.byNode, byNode)}
          </div>

          <div class="weekly-card">
            <h3>주간 SYSTEM별 결선 실적</h3>
            ${renderWeeklySystemTable(weekStats.bySystem, bySystem)}
          </div>

          <div class="weekly-card weekly-note">
            <h3>특이사항 / 다음 주 계획</h3>
            <ul>
              <li>${weekStats.installCount ? `금주 포설 ${num(weekStats.installCount)}건 / ${num(Math.round(weekStats.installLength))}m 완료.` : "금주 포설 실적 없음."}</li>
              <li>${weekStats.conFromCount + weekStats.conToCount ? `결선 진행 ${num(weekStats.conFromCount)}건(FROM) · ${num(weekStats.conToCount)}건(TO).` : "결선 실적 없음."}</li>
              <li>다음 주차 계획은 일일 작업 일지 기준으로 갱신 예정.</li>
            </ul>
            <div class="copy-report-note" style="margin-top:8px;">
              <b>REV 기준:</b> Cable List REV 열 변경 이력<br>
              <b>기준 파일:</b> ${escapeHtml(vessel.id)} Cable List-포설실적용_260205.xlsb
            </div>
          </div>
        </section>
      </article>`;
  }

  function renderWeeklyNodeTable(weekly, monthly) {
    if (!monthly.length) return `<p class="weekly-empty">데이터 없음</p>`;
    return `
      <table class="copy-report-table weekly-table">
        <thead><tr><th>구역</th><th>주간 포설(건)</th><th>주간 포설(m)</th><th>월누적 포설(m)</th><th>총 길이(m)</th><th>진도율</th></tr></thead>
        <tbody>
          ${monthly.map((row) => {
            const w = weekly.get(row.key) || { count: 0, length: 0 };
            const rate = row.totalLength ? row.installedLength / row.totalLength : 0;
            return `<tr>
              <td class="system-cell">${escapeHtml(row.key)}</td>
              <td class="num blue-text">${num(w.count)}</td>
              <td class="num blue-text">${num(Math.round(w.length))}</td>
              <td class="num">${num(Math.round(row.installedLength))}</td>
              <td class="num">${num(Math.round(row.totalLength))}</td>
              <td><span class="copy-progress" style="--rate:${Math.round(rate * 100)}"><b>${pct(rate)}</b></span></td>
            </tr>`;
          }).join("")}
        </tbody>
      </table>`;
  }

  function renderWeeklySystemTable(weekly, monthly) {
    if (!monthly.length) return `<p class="weekly-empty">데이터 없음</p>`;
    return `
      <table class="copy-report-table weekly-table">
        <thead><tr><th>SYSTEM</th><th>주 FROM(건)</th><th>주 TO(건)</th><th>월누적 FROM</th><th>월누적 TO</th><th>전체(EA)</th></tr></thead>
        <tbody>
          ${monthly.map((row) => {
            const w = weekly.get(row.key) || { conFromCount: 0, conToCount: 0 };
            return `<tr>
              <td class="system-cell">${escapeHtml(row.key)}</td>
              <td class="num blue-text">${num(w.conFromCount)}</td>
              <td class="num purple-text">${num(w.conToCount)}</td>
              <td class="num">${num(row.conFromCount)}</td>
              <td class="num">${num(row.conToCount)}</td>
              <td class="num">${num(row.count)}</td>
            </tr>`;
          }).join("")}
        </tbody>
      </table>`;
  }

  // Build week buckets for a month: each week starts on Sunday, capped to month bounds
  function monthWeeks(year, month) {
    const first = new Date(year, month, 1);
    const last = new Date(year, month + 1, 0);
    const weeks = [];
    let cursor = new Date(first);
    let index = 1;
    while (cursor <= last) {
      const dow = cursor.getDay();
      const weekStart = new Date(cursor);
      const weekEnd = new Date(cursor);
      weekEnd.setDate(weekEnd.getDate() + (6 - dow));
      if (weekEnd > last) weekEnd.setTime(last.getTime());
      const days = [];
      for (let d = new Date(weekStart); d <= weekEnd; d.setDate(d.getDate() + 1)) {
        days.push(new Date(d));
      }
      weeks.push({
        index,
        start: new Date(weekStart),
        end: new Date(weekEnd),
        startLabel: `${weekStart.getMonth() + 1}/${weekStart.getDate()}`,
        endLabel: `${weekEnd.getMonth() + 1}/${weekEnd.getDate()}`,
        days,
      });
      cursor = new Date(weekEnd);
      cursor.setDate(cursor.getDate() + 1);
      index += 1;
    }
    return weeks;
  }

  function isoWeekOfMonth(date) {
    const first = new Date(date.getFullYear(), date.getMonth(), 1);
    const offset = first.getDay();
    return Math.ceil((date.getDate() + offset) / 7);
  }

  function collectWeekStats(cables, week) {
    const byNode = new Map();
    const bySystem = new Map();
    let installCount = 0, installLength = 0;
    let conFromCount = 0, conFromLength = 0;
    let conToCount = 0, conToLength = 0;
    let totalLength = 0;
    const startKey = week.start.toISOString().slice(0, 10);
    const endKey = week.end.toISOString().slice(0, 10);
    cables.forEach((cable) => {
      const len = Number(cable.total) || 0;
      totalLength += len;
      const nodeKey = cable.node || "미분류";
      const sysKey = cable.sys || "미분류";
      if (cable.installDate && cable.installDate >= startKey && cable.installDate <= endKey) {
        installCount += 1;
        installLength += len;
        const cur = byNode.get(nodeKey) || { count: 0, length: 0 };
        cur.count += 1; cur.length += len;
        byNode.set(nodeKey, cur);
      }
      if (cable.conFromDate && cable.conFromDate >= startKey && cable.conFromDate <= endKey) {
        conFromCount += 1; conFromLength += len;
        const cur = bySystem.get(sysKey) || { conFromCount: 0, conToCount: 0 };
        cur.conFromCount += 1;
        bySystem.set(sysKey, cur);
      }
      if (cable.conToDate && cable.conToDate >= startKey && cable.conToDate <= endKey) {
        conToCount += 1; conToLength += len;
        const cur = bySystem.get(sysKey) || { conFromCount: 0, conToCount: 0 };
        cur.conToCount += 1;
        bySystem.set(sysKey, cur);
      }
    });
    return {
      installCount, installLength,
      conFromCount, conFromLength,
      conToCount, conToLength,
      totalLength,
      byNode, bySystem,
    };
  }

  function installationReportTable(vesselId, rows, kpi) {
    const totalInstalled = rows.reduce((sum, row) => sum + row.installedLength, 0);
    const totalLength = rows.reduce((sum, row) => sum + row.totalLength, 0);
    return `
      <div class="table-wrap">
        <table class="report-table">
          <thead>
            <tr><th>호선</th><th>구역</th><th>TOTAL(m)</th><th>포설(m)</th><th>미포설(m)</th><th>진도율</th><th>건수</th></tr>
          </thead>
          <tbody>
            ${rows.map((row, index) => {
              const rate = row.totalLength ? row.installedLength / row.totalLength : 0;
              return `<tr>
                ${index === 0 ? `<td rowspan="${rows.length + 1}">${escapeHtml(vesselId)}</td>` : ""}
                <td>${escapeHtml(row.key)}</td>
                <td class="num">${num(row.totalLength)}</td>
                <td class="num blue-text">${num(row.installedLength)}</td>
                <td class="num">${num(Math.max(0, row.totalLength - row.installedLength))}</td>
                <td>${progressBarCell(rate)}</td>
                <td class="num">${num(row.installedCount)} / ${num(row.count)}</td>
              </tr>`;
            }).join("")}
            <tr class="sum-row"><td>합계</td><td class="num">${num(totalLength)}</td><td class="num">${num(totalInstalled)}</td><td class="num">${num(Math.max(0, totalLength - totalInstalled))}</td><td>${progressBarCell(kpi.installRate)}</td><td class="num">${num(kpi.totalCount)}</td></tr>
          </tbody>
        </table>
      </div>`;
  }

  function connectionReportTable(vesselId, rows) {
    const total = rows.reduce((sum, row) => sum + row.count, 0);
    const totalFrom = rows.reduce((sum, row) => sum + row.conFromCount, 0);
    const totalTo = rows.reduce((sum, row) => sum + row.conToCount, 0);
    const progress = total ? Math.min(totalFrom, totalTo) / total : 0;
    return `
      <div class="table-wrap">
        <table class="report-table">
          <thead>
            <tr><th>호선</th><th>SYSTEM</th><th>TOTAL(EA)</th><th>결선 FROM</th><th>결선 TO</th><th>진도율</th><th>결선 길이(m)</th></tr>
          </thead>
          <tbody>
            ${rows.map((row, index) => {
              const rate = row.count ? Math.min(row.conFromCount, row.conToCount) / row.count : 0;
              return `<tr>
                ${index === 0 ? `<td rowspan="${rows.length + 1}">${escapeHtml(vesselId)}</td>` : ""}
                <td>${escapeHtml(row.key)}</td>
                <td class="num">${num(row.count)}</td>
                <td class="num blue-text">${num(row.conFromCount)}</td>
                <td class="num purple-text">${num(row.conToCount)}</td>
                <td>${progressBarCell(rate)}</td>
                <td class="num">${num(Math.min(row.conFromLength, row.conToLength))}</td>
              </tr>`;
            }).join("")}
            <tr class="sum-row"><td>합계</td><td class="num">${num(total)}</td><td class="num">${num(totalFrom)}</td><td class="num">${num(totalTo)}</td><td>${progressBarCell(progress)}</td><td class="num">-</td></tr>
          </tbody>
        </table>
      </div>`;
  }

  function renderCalendar() {
    const monthLabel = `${calendarYear}년 ${calendarMonth + 1}월`;
    const daily = collectDailyAll(calendarYear, calendarMonth);
    main().innerHTML = `
      <section class="legacy-shell calendar-page">
        <div class="calendar-toolbar">
          <div class="calendar-month-nav">
            <button class="btn ghost" id="btn-prev-month" type="button">◀ 이전</button>
            <select id="calendar-month-select">${monthOptions()}</select>
            <button class="btn ghost" id="btn-next-month" type="button">다음 ▶</button>
          </div>
          <div class="calendar-legend">
            <span class="calendar-legend-chip install">■ 포설</span>
            <span class="calendar-legend-chip connection">■ 결선</span>
          </div>
        </div>

        <section class="calendar-panel">
          <div class="calendar-title">${escapeHtml(monthLabel)}</div>
          <div class="calendar-week-head">
            <span class="sun">일</span><span>월</span><span>화</span><span>수</span><span>목</span><span>금</span><span class="sat">토</span>
          </div>
          <div class="calendar-grid">
            ${monthCells(calendarYear, calendarMonth).map((cell) => renderCalendarCell(cell, daily)).join("")}
          </div>
        </section>

        <section class="calendar-detail" id="calendar-detail">
          ${renderCalendarDetail(calendarSelectedDate)}
        </section>
      </section>`;

    el("btn-prev-month").addEventListener("click", () => changeMonth(-1));
    el("btn-next-month").addEventListener("click", () => changeMonth(1));
    el("calendar-month-select").addEventListener("change", (event) => {
      const [year, month] = event.target.value.split("-").map(Number);
      calendarYear = year;
      calendarMonth = month - 1;
      calendarSelectedDate = "";
      renderCalendar();
    });
    document.querySelectorAll(".calendar-cell[data-date]").forEach((cell) => {
      cell.addEventListener("click", () => {
        const date = cell.dataset.date;
        calendarSelectedDate = calendarSelectedDate === date ? "" : date;
        renderCalendar();
      });
    });
    bindCalendarDetailEvents();
  }

  function monthOptions() {
    const dates = activeCables().flatMap((cable) => [cable.installDate, cable.conFromDate, cable.conToDate]).filter(Boolean);
    const unique = [...new Set(dates.map((date) => date.slice(0, 7)))].sort();
    const values = unique.length ? unique : [`${calendarYear}-${String(calendarMonth + 1).padStart(2, "0")}`];
    return values.map((value) => {
      const [year, month] = value.split("-");
      return `<option value="${value}" ${Number(year) === calendarYear && Number(month) === calendarMonth + 1 ? "selected" : ""}>${Number(year)}년 ${Number(month)}월</option>`;
    }).join("");
  }

  function changeMonth(delta) {
    const next = new Date(calendarYear, calendarMonth + delta, 1);
    calendarYear = next.getFullYear();
    calendarMonth = next.getMonth();
    renderCalendar();
  }

  function monthCells(year, month) {
    const first = new Date(year, month, 1);
    const days = new Date(year, month + 1, 0).getDate();
    const cells = [];
    for (let i = 0; i < first.getDay(); i += 1) cells.push({ day: "", date: "" });
    for (let day = 1; day <= days; day += 1) {
      cells.push({ day, date: `${year}-${String(month + 1).padStart(2, "0")}-${String(day).padStart(2, "0")}` });
    }
    while (cells.length % 7 !== 0) cells.push({ day: "", date: "" });
    return cells;
  }

  function collectDaily(year, month, mode) {
    const prefix = `${year}-${String(month + 1).padStart(2, "0")}`;
    const map = new Map();
    activeCables().forEach((cable) => {
      const length = Number(cable.total) || 0;
      const add = (date, label) => {
        if (!date || !date.startsWith(prefix)) return;
        const current = map.get(date) || { count: 0, length: 0, labels: new Set() };
        current.count += 1;
        current.length += length;
        current.labels.add(label);
        map.set(date, current);
      };
      if (mode === "install") {
        add(cable.installDate, "포설");
      } else {
        add(cable.conFromDate, "FROM");
        add(cable.conToDate, "TO");
      }
    });
    return map;
  }

  function collectDailyAll(year, month) {
    const prefix = `${year}-${String(month + 1).padStart(2, "0")}`;
    const map = new Map();
    activeCables().forEach((cable) => {
      const length = Number(cable.total) || 0;
      const ensure = (date) => {
        if (!date || !date.startsWith(prefix)) return null;
        const cur = map.get(date) || {
          installCount: 0, installLength: 0,
          conFromCount: 0, conFromLength: 0,
          conToCount: 0, conToLength: 0,
        };
        map.set(date, cur);
        return cur;
      };
      const inst = ensure(cable.installDate);
      if (inst) { inst.installCount += 1; inst.installLength += length; }
      const cf = ensure(cable.conFromDate);
      if (cf) { cf.conFromCount += 1; cf.conFromLength += length; }
      const ct = ensure(cable.conToDate);
      if (ct) { ct.conToCount += 1; ct.conToLength += length; }
    });
    return map;
  }

  function renderCalendarCell(cell, daily) {
    const record = cell.date ? daily.get(cell.date) : null;
    const day = cell.day;
    const dayClass = cell.date ? new Date(cell.date).getDay() : -1;
    const events = [];
    if (record) {
      if (record.installCount) {
        events.push(`<div class="calendar-event install">포설 ${num(record.installCount)}건<br><b>${num(Math.round(record.installLength))}m</b></div>`);
      }
      const conCount = record.conFromCount + record.conToCount;
      const conLen = record.conFromLength + record.conToLength;
      if (conCount) {
        const detail = [
          record.conFromCount ? `FROM ${num(record.conFromCount)}` : "",
          record.conToCount ? `TO ${num(record.conToCount)}` : "",
        ].filter(Boolean).join(" · ");
        events.push(`<div class="calendar-event connection">결선 ${num(conCount)}건<br><span class="calendar-event-sub">${detail}</span><br><b>${num(Math.round(conLen))}m</b></div>`);
      }
    }
    const isSelected = cell.date && cell.date === calendarSelectedDate;
    const clickable = cell.date ? `data-date="${escapeAttr(cell.date)}"` : "";
    return `
      <div class="calendar-cell ${cell.date ? "" : "empty"} ${isSelected ? "selected" : ""}" ${clickable}>
        <div class="calendar-day ${dayClass === 0 ? "sun" : ""} ${dayClass === 6 ? "sat" : ""}">${day || ""}</div>
        ${events.join("")}
      </div>`;
  }

  function dailyCablesFor(date) {
    if (!date) return { install: [], conFrom: [], conTo: [] };
    return {
      install: activeCables().filter((cable) => cable.installDate === date),
      conFrom: activeCables().filter((cable) => cable.conFromDate === date),
      conTo: activeCables().filter((cable) => cable.conToDate === date),
    };
  }

  function renderCalendarDetail(date) {
    if (!date) {
      return `<div class="calendar-detail-empty">날짜를 클릭하면 해당일의 포설/결선 케이블 목록(SYS · CIRCUIT NO. · CABLE TYPE)이 여기에 표시됩니다.</div>`;
    }
    const { install, conFrom, conTo } = dailyCablesFor(date);
    const totalLen = install.reduce((sum, c) => sum + (Number(c.total) || 0), 0);
    const conLen = [...conFrom, ...conTo].reduce((sum, c) => sum + (Number(c.total) || 0), 0);

    const renderTable = (rows, kind) => {
      if (!rows.length) {
        return `<div class="calendar-detail-empty small">${kind === "install" ? "포설" : "결선"} 실적 없음</div>`;
      }
      return `
        <table class="calendar-detail-table">
          <thead>
            <tr><th>SYS</th><th>CIRCUIT NO.</th><th>CABLE TYPE</th>${kind === "connection" ? "<th>구분</th>" : ""}<th class="num">TOTAL(m)</th></tr>
          </thead>
          <tbody>
            ${rows.map((row) => `
              <tr>
                <td>${escapeHtml(row.cable.sys || "")}</td>
                <td class="mono">${escapeHtml(row.cable.circuitNo || "")}</td>
                <td>${escapeHtml(row.cable.cableType || "")}</td>
                ${kind === "connection" ? `<td>${escapeHtml(row.kind)}</td>` : ""}
                <td class="num">${num(Math.round(Number(row.cable.total) || 0))}</td>
              </tr>`).join("")}
          </tbody>
        </table>`;
    };

    const conRows = [
      ...conFrom.map((cable) => ({ cable, kind: "FROM" })),
      ...conTo.map((cable) => ({ cable, kind: "TO" })),
    ];

    return `
      <div class="calendar-detail-head">
        <div>
          <h3 class="calendar-detail-title">${escapeHtml(date)} 일자별 실적</h3>
          <p class="calendar-detail-sub">포설 ${num(install.length)}건 · ${num(Math.round(totalLen))}m &nbsp;|&nbsp; 결선 ${num(conRows.length)}건 · ${num(Math.round(conLen))}m</p>
        </div>
        <div class="actions">
          <button class="btn ghost small" id="btn-day-clear" type="button">닫기</button>
          <button class="btn green small" id="btn-day-export" type="button">📥 일자별 리스트 뽑기</button>
        </div>
      </div>

      <div class="calendar-detail-body">
        <section class="calendar-detail-section">
          <h4 class="calendar-detail-section-title install">포설 (${num(install.length)}건)</h4>
          ${renderTable(install.map((cable) => ({ cable })), "install")}
        </section>
        <section class="calendar-detail-section">
          <h4 class="calendar-detail-section-title connection">결선 (${num(conRows.length)}건)</h4>
          ${renderTable(conRows, "connection")}
        </section>
      </div>`;
  }

  function bindCalendarDetailEvents() {
    const clear = el("btn-day-clear");
    if (clear) {
      clear.addEventListener("click", () => {
        calendarSelectedDate = "";
        renderCalendar();
      });
    }
    const exp = el("btn-day-export");
    if (exp) exp.addEventListener("click", exportDailyList);
  }

  async function exportDailyList() {
    if (!calendarSelectedDate) return;
    const vessel = Store.getActiveVessel();
    const date = calendarSelectedDate;
    const { install, conFrom, conTo } = dailyCablesFor(date);

    const aoa = [];
    aoa.push([`◎ ${vessel.id} ${date} 일자별 실적`]);
    aoa.push([]);
    aoa.push([`■ 포설 (${install.length}건)`]);
    aoa.push(["SYS", "CIRCUIT NO.", "CABLE TYPE", "TOTAL(m)"]);
    install.forEach((cable) => {
      aoa.push([cable.sys || "", cable.circuitNo || "", cable.cableType || "", Math.round(Number(cable.total) || 0)]);
    });
    aoa.push([]);
    aoa.push([`■ 결선 (${conFrom.length + conTo.length}건)`]);
    aoa.push(["SYS", "CIRCUIT NO.", "CABLE TYPE", "구분", "TOTAL(m)"]);
    conFrom.forEach((cable) => {
      aoa.push([cable.sys || "", cable.circuitNo || "", cable.cableType || "", "FROM", Math.round(Number(cable.total) || 0)]);
    });
    conTo.forEach((cable) => {
      aoa.push([cable.sys || "", cable.circuitNo || "", cable.cableType || "", "TO", Math.round(Number(cable.total) || 0)]);
    });

    const filename = `${vessel.id}-daily-${date}.xlsx`;
    try {
      await ensureXlsx();
      const ws = XLSX.utils.aoa_to_sheet(aoa);
      ws["!cols"] = [{ wch: 14 }, { wch: 18 }, { wch: 22 }, { wch: 10 }, { wch: 12 }];
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Daily");
      XLSX.writeFile(wb, filename);
      toast(`${date} 일자별 실적을 Excel로 내보냈습니다.`);
    } catch (error) {
      download(filename.replace(/\.xlsx$/, ".tsv"), aoa.map((row) => row.join("\t")).join("\r\n"));
      toast("Excel 내보내기가 실패해 TSV로 저장했습니다.");
    }
  }

  async function exportCalendarReport() {
    const vessel = Store.getActiveVessel();
    const aoa = buildCalendarWorkbookRows();
    const filename = `${vessel.id}-calendar-report-${calendarYear}-${String(calendarMonth + 1).padStart(2, "0")}.xlsx`;
    try {
      await ensureXlsx();
      const worksheet = XLSX.utils.aoa_to_sheet(aoa);
      worksheet["!cols"] = [
        { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 14 },
        { wch: 14 }, { wch: 14 }, { wch: 18 }, { wch: 18 },
      ];
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, "Calendar Report");
      XLSX.writeFile(workbook, filename);
      toast("달력과 포설/결선 요약 보고서를 Excel로 내보냈습니다.");
    } catch (error) {
      download(filename.replace(/\.xlsx$/, ".tsv"), aoa.map((row) => row.join("\t")).join("\r\n"));
      toast("Excel 내보내기가 실패해 TSV로 저장했습니다.");
    }
  }

  function buildCalendarWorkbookRows() {
    const vessel = Store.getActiveVessel();
    const cables = activeCables();
    const byNode = aggregateRows(cables, (cable) => cable.node);
    const bySystem = aggregateRows(cables, (cable) => systemName(cable.sys));
    const kpi = Store.getKpi();
    const rows = [];

    rows.push([`◎ ${vessel.id} CABLE 포설 현황`]);
        rows.push(["호선", "구역", "TOTAL(m)", "포설(m)", "미포설(m)", "진도율", "포설건수"]);
    byNode.forEach((row, index) => {
      const rate = row.totalLength ? row.installedLength / row.totalLength : 0;
      rows.push([
        index === 0 ? vessel.id : "",
        row.key,
        Math.round(row.totalLength),
        Math.round(row.installedLength),
        Math.round(Math.max(0, row.totalLength - row.installedLength)),
        pct(rate),
        `${row.installedCount}/${row.count}`,
      ]);
    });
    rows.push(["합계", "", Math.round(kpi.totalLength), Math.round(kpi.installedLength), Math.round(kpi.pendingLength), pct(kpi.installRate), kpi.totalCount]);
    rows.push([]);
    rows.push([`◎ ${vessel.id} CABLE 결선 현황`]);
        rows.push(["호선", "SYSTEM", "TOTAL(EA)", "결선 FROM", "결선 TO", "진도율", "결선길이(m)"]);
    bySystem.forEach((row, index) => {
      const rate = row.count ? Math.min(row.conFromCount, row.conToCount) / row.count : 0;
      rows.push([
        index === 0 ? vessel.id : "",
        row.key,
        row.count,
        row.conFromCount,
        row.conToCount,
        pct(rate),
        Math.round(Math.min(row.conFromLength, row.conToLength)),
      ]);
    });
    rows.push([]);
    rows.push([`${calendarYear}년 ${calendarMonth + 1}월 달력`, calendarMode === "install" ? "포설 기준" : "결선 기준"]);
    rows.push(["일", "월", "화", "수", "목", "금", "토"]);
    const daily = collectDaily(calendarYear, calendarMonth, calendarMode);
    const cells = monthCells(calendarYear, calendarMonth);
    for (let i = 0; i < cells.length; i += 7) {
      const week = cells.slice(i, i + 7).map((cell) => {
        if (!cell.date) return "";
        const record = daily.get(cell.date);
        const text = record ? `${cell.day}\n${calendarMode === "install" ? "포설" : "결선"} ${record.count}건 / ${Math.round(record.length)}m` : String(cell.day);
        return text;
      });
      rows.push(week);
    }
    return rows;
  }

﻿  function renderCableList() {
    const all = Store.getCables();
    pruneTableSelection(all);
    const cables = sortCablesLikeExcel(filteredCables());
    const pages = Math.max(1, Math.ceil(cables.length / PAGE_SIZE));
    tablePage = Math.min(tablePage, pages);
    const start = (tablePage - 1) * PAGE_SIZE;
    const rows = cables.slice(start, start + PAGE_SIZE);
    const types = [...new Set(all.map((cable) => cable.cableType).filter(Boolean))].sort();
    const pageSelectedCount = rows.filter((cable) => tableSelectedIds.has(cable.id)).length;
    const filteredSelectedCount = cables.filter((cable) => tableSelectedIds.has(cable.id)).length;
    const hasPageRows = rows.length > 0;
    const pageAllSelected = hasPageRows && pageSelectedCount === rows.length;

    main().innerHTML = `
      <section class="legacy-shell data-page excel-data-page">
        <div class="data-head-row">
          <h1>케이블 리스트</h1>
          <div class="filtered-actions">
            <span>필터된 행 전체에 오늘 날짜 입력:</span>
            <button class="btn small green" id="btn-apply-install" type="button">포설일자</button>
            <button class="btn small" id="btn-apply-con-from" type="button">결선 FROM</button>
            <button class="btn small purple-btn" id="btn-apply-con-to" type="button">결선 TO</button>
            <button class="btn small ghost" id="btn-export-filtered" type="button">필터 결과 내보내기</button>
          </div>
        </div>

        <section class="legacy-filter-panel excel-control-deck">
          <div class="excel-slicer-stack">
            <div class="excel-slicer">
              <div class="excel-slicer-head">
                <strong>SYSTEM</strong>
                <span>${num(all.length)}건</span>
              </div>
              <div class="excel-slicer-body">
                ${renderSystemChips(all)}
              </div>
            </div>
            <div class="excel-slicer">
              <div class="excel-slicer-head">
                <strong>구역 / NODE</strong>
                <span>${num(all.filter((cable) => cable.node).length)}건</span>
              </div>
              <div class="excel-slicer-body">
                ${renderNodeChips(all)}
              </div>
            </div>
          </div>
          <div class="excel-ops-stack">
            <div class="legacy-filters excel-filter-grid">
              <div class="field"><label>CABLE TYPE</label><select id="flt-type"><option value="">전체</option>${types.map((type) => `<option ${type === tableFilter.type ? "selected" : ""}>${escapeHtml(type)}</option>`).join("")}</select></div>
              <div class="field"><label>포설일자</label><select id="flt-install-status">${statusOptions(tableFilter.installStatus)}</select></div>
              <div class="field"><label>결선 FROM</label><select id="flt-confrom-status">${statusOptions(tableFilter.conFromStatus)}</select></div>
              <div class="field"><label>결선 TO</label><select id="flt-conto-status">${statusOptions(tableFilter.conToStatus)}</select></div>
              <div class="field search-field"><label>검색 (CIRCUIT / FROM / TO)</label><input id="flt-query" value="${escapeAttr(tableFilter.query)}" placeholder="검색어 입력"></div>
              <button class="btn small ghost" id="btn-clear-filter" type="button">필터 초기화</button>
            </div>
            <div class="excel-ops-panel">
              <div class="selected-summary">
                <strong>${num(tableSelectedIds.size)}건 선택</strong>
                <span>현재 필터 ${num(filteredSelectedCount)}건 / 현재 페이지 ${num(pageSelectedCount)}건</span>
              </div>
              <div class="excel-bulk-row">
                <button class="btn small" id="btn-select-page" type="button">현재 페이지 선택</button>
                <button class="btn small" id="btn-select-filtered" type="button">필터 전체 선택</button>
                <button class="btn small ghost" id="btn-clear-selection" type="button">선택 해제</button>
              </div>
              <div class="excel-bulk-row">
                <label for="bulk-date">일괄 날짜</label>
                <input id="bulk-date" type="date" value="${today()}">
                <button class="btn small ghost" id="btn-bulk-today" type="button">오늘</button>
                <button class="btn small green" id="btn-bulk-install" type="button">선택행 포설일자</button>
                <button class="btn small" id="btn-bulk-confrom" type="button">선택행 결선 FROM</button>
                <button class="btn small purple-btn" id="btn-bulk-conto" type="button">선택행 결선 TO</button>
                <button class="btn small ghost" id="btn-bulk-clear-install" type="button">포설일자 비우기</button>
                <button class="btn small ghost" id="btn-bulk-clear-confrom" type="button">FROM 비우기</button>
                <button class="btn small ghost" id="btn-bulk-clear-conto" type="button">TO 비우기</button>
              </div>
            </div>
          </div>
        </section>

        <div class="data-count-row">
          <strong>${num(cables.length)}건 표시 중</strong>
          <span>${cables.length ? start + 1 : 0}-${Math.min(start + PAGE_SIZE, cables.length)} / ${num(cables.length)}</span>
        </div>

        <div class="excel-shell cable-data-wrap">
          <div class="excel-watermark">${tablePage} PAGE</div>
          <table class="excel-table cable-data-table ${cableTab === "deleted" ? "deleted-tab" : ""}" aria-label="Excel Cable List editable table">
            <thead>
              <tr class="excel-letters">
                <th class="excel-row-head"></th>
                <th class="col-select">SEL</th>
                ${excelColumns.map((_, index) => `<th>${excelColumnLetter(index + 1)}</th>`).join("")}
              </tr>
              <tr class="excel-headline">
                <th class="excel-row-head">12</th>
                <th class="col-select select-head">
                  <input id="chk-page-all" type="checkbox" ${pageAllSelected ? "checked" : ""} ${hasPageRows ? "" : "disabled"} aria-label="현재 페이지 전체 선택">
                </th>
                ${excelColumns.map((column) => `
                  <th class="${column.className || ""} ${column.headerClass || ""}">
                    <span class="excel-header-label">${column.label}</span>
                    <span class="excel-filter" aria-hidden="true"></span>
                  </th>
                `).join("")}
              </tr>
            </thead>
            <tbody>
              ${rows.map((cable, index) => renderDataRow(cable, start + index + 13)).join("")}
            </tbody>
          </table>
        </div>
        <div class="excel-help">셀은 바로 수정할 수 있고, 날짜 셀은 더블클릭으로 오늘 날짜 입력/해제가 됩니다. 왼쪽 선택란으로 여러 줄을 고른 뒤 일괄 날짜 입력도 가능합니다.</div>
        ${renderPager(pages)}
      </section>`;

    bindCableListEvents();
  }

  function statusOptions(selected) {
    return `
      <option value="" ${selected === "" ? "selected" : ""}>전체</option>
      <option value="filled" ${selected === "filled" ? "selected" : ""}>입력됨</option>
      <option value="empty" ${selected === "empty" ? "selected" : ""}>미입력</option>`;
  }

  function renderSystemChips(cables) {
    const counts = new Map();
    cables.forEach((cable) => {
      const key = cable.sys || "미분류";
      counts.set(key, (counts.get(key) || 0) + 1);
    });
    const chips = [...counts.entries()].sort((a, b) => a[0].localeCompare(b[0], "ko"));
    return [
      `<button class="slicer-chip ${tableFilter.sys === "" ? "active" : ""}" data-sys-filter="" type="button">전체 <span>${num(cables.length)}</span></button>`,
      ...chips.map(([sys, count]) => `
        <button class="slicer-chip ${tableFilter.sys === sys ? "active" : ""}" data-sys-filter="${escapeAttr(sys)}" type="button">
          ${escapeHtml(sys)} <span>${num(count)}</span>
        </button>
      `),
    ].join("");
  }

  function renderNodeChips(cables) {
    const counts = new Map();
    cables.forEach((cable) => {
      const key = cable.node || "미지정";
      counts.set(key, (counts.get(key) || 0) + 1);
    });
    const chips = [...counts.entries()].sort((a, b) => a[0].localeCompare(b[0], "ko"));
    return [
      `<button class="slicer-chip ${tableFilter.node === "" ? "active" : ""}" data-node-filter="" type="button">전체 <span>${num(cables.length)}</span></button>`,
      ...chips.map(([node, count]) => `
        <button class="slicer-chip ${tableFilter.node === node ? "active" : ""}" data-node-filter="${escapeAttr(node)}" type="button">
          ${escapeHtml(node)} <span>${num(count)}</span>
        </button>
      `),
    ].join("");
  }

  function pruneTableSelection(cables = Store.getCables()) {
    const liveIds = new Set(cables.map((cable) => cable.id));
    tableSelectedIds = new Set([...tableSelectedIds].filter((id) => liveIds.has(id)));
    if (lastSelectedCableId && !liveIds.has(lastSelectedCableId)) lastSelectedCableId = "";
  }

  function setTableSelection(ids, selected = true) {
    ids.forEach((id) => {
      if (!id) return;
      if (selected) tableSelectedIds.add(id);
      else tableSelectedIds.delete(id);
    });
  }

  function clearTableSelection() {
    tableSelectedIds = new Set();
    lastSelectedCableId = "";
  }

  function selectedCables() {
    const byId = new Map(Store.getCables().map((cable) => [cable.id, cable]));
    return [...tableSelectedIds].map((id) => byId.get(id)).filter(Boolean);
  }

  function renderDataRow(cable, rowNumber) {
    const deleted = cable.deleted || Number(cable.total) === 0 || /삭제/i.test(cable.rev || "") || /delete/i.test(cable.rev || "");
    const selected = tableSelectedIds.has(cable.id);
    return `<tr class="${deleted ? "deleted-row" : ""} ${selected ? "selected-row" : ""}" data-row-id="${escapeAttr(cable.id)}">
      <th class="excel-row-head">${rowNumber}</th>
      <td class="excel-cell excel-check">
        <input class="row-select" data-select-id="${escapeAttr(cable.id)}" type="checkbox" ${selected ? "checked" : ""} aria-label="${escapeAttr(cable.circuitNo || cable.id)} 선택">
      </td>
      ${excelColumns.map((column) => renderDataCell(cable, column.field, column.className || "")).join("")}
    </tr>`;
  }

  function renderDataCell(cable, field, className = "") {
    const value = formatCellValue(cable, field);
    const isDate = ["installDate", "conFromDate", "conToDate"].includes(field);
    const emptyDate = isDate && !value;
    const completeDateLine = cable.installDate && cable.conFromDate && cable.conToDate;
    const title = isDate ? "더블클릭: 오늘 날짜 입력/해제" : "클릭 후 바로 수정";
    return `
      <td
        class="excel-cell editable-cell ${className} ${emptyDate ? "empty-date" : ""} ${isDate && completeDateLine ? "excel-yellow" : ""}"
        contenteditable="true"
        spellcheck="false"
        title="${title}"
        data-cable-id="${escapeAttr(cable.id)}"
        data-field="${escapeAttr(field)}"
      >${escapeHtml(value || "")}</td>`;
  }

  function formatCellValue(cable, field) {
    if (field === "cableDia" || field === "total") {
      const raw = cable[field];
      if (raw === "" || raw === null || raw === undefined) return "";
      const numeric = Number(raw);
      if (Number.isFinite(numeric)) return num(numeric, 1).replace(/\.0$/, "");
      return raw;
    }
    return cable[field] ?? "";
  }

  function bindCableListEvents() {
    document.querySelectorAll("[data-cable-tab]").forEach((button) => {
      button.addEventListener("click", () => {
        cableTab = button.dataset.cableTab;
        tablePage = 1;
        clearTableSelection();
        renderCableList();
      });
    });

    ["flt-query", "flt-type", "flt-install-status", "flt-confrom-status", "flt-conto-status"].forEach((id) => {
      const target = el(id);
      if (!target) return;
      target.addEventListener(id === "flt-query" ? "input" : "change", () => {
        tableFilter = {
          ...tableFilter,
          query: el("flt-query").value,
          type: el("flt-type").value,
          installStatus: el("flt-install-status").value,
          conFromStatus: el("flt-confrom-status").value,
          conToStatus: el("flt-conto-status").value,
        };
        tablePage = 1;
        renderCableList();
      });
    });

    document.querySelectorAll("[data-sys-filter]").forEach((button) => {
      button.addEventListener("click", () => {
        tableFilter.sys = button.dataset.sysFilter || "";
        tablePage = 1;
        renderCableList();
      });
    });

    document.querySelectorAll("[data-node-filter]").forEach((button) => {
      button.addEventListener("click", () => {
        tableFilter.node = button.dataset.nodeFilter || "";
        tablePage = 1;
        renderCableList();
      });
    });

    el("btn-clear-filter").addEventListener("click", () => {
      tableFilter = { query: "", sys: "", node: "", type: "", installStatus: "", conFromStatus: "", conToStatus: "" };
      tablePage = 1;
      renderCableList();
    });

    el("btn-apply-install").addEventListener("click", () => applyTodayToFiltered("installDate"));
    el("btn-apply-con-from").addEventListener("click", () => applyTodayToFiltered("conFromDate"));
    el("btn-apply-con-to").addEventListener("click", () => applyTodayToFiltered("conToDate"));
    el("btn-export-filtered").addEventListener("click", () => {
      const vessel = Store.getActiveVessel();
      download(`${vessel.id}-Cable-List-filtered.tsv`, Store.toTsv(sortCablesLikeExcel(filteredCables())));
      toast("필터된 Cable List를 UTF-8 BOM 포함 TSV로 저장했습니다.");
    });

    const currentPageIds = sortCablesLikeExcel(filteredCables())
      .slice((tablePage - 1) * PAGE_SIZE, (tablePage - 1) * PAGE_SIZE + PAGE_SIZE)
      .map((cable) => cable.id);
    const headerCheck = el("chk-page-all");
    if (headerCheck) {
      const selectedOnPage = currentPageIds.filter((id) => tableSelectedIds.has(id)).length;
      headerCheck.indeterminate = selectedOnPage > 0 && selectedOnPage < currentPageIds.length;
      headerCheck.addEventListener("change", () => {
        setTableSelection(currentPageIds, headerCheck.checked);
        renderCableList();
      });
    }

    el("btn-select-page").addEventListener("click", () => {
      setTableSelection(currentPageIds, true);
      renderCableList();
    });
    el("btn-select-filtered").addEventListener("click", () => {
      setTableSelection(sortCablesLikeExcel(filteredCables()).map((cable) => cable.id), true);
      renderCableList();
    });
    el("btn-clear-selection").addEventListener("click", () => {
      clearTableSelection();
      renderCableList();
    });
    el("btn-bulk-today").addEventListener("click", () => {
      el("bulk-date").value = today();
    });
    el("btn-bulk-install").addEventListener("click", () => applySelectedDate("installDate", "set"));
    el("btn-bulk-confrom").addEventListener("click", () => applySelectedDate("conFromDate", "set"));
    el("btn-bulk-conto").addEventListener("click", () => applySelectedDate("conToDate", "set"));
    el("btn-bulk-clear-install").addEventListener("click", () => applySelectedDate("installDate", "clear"));
    el("btn-bulk-clear-confrom").addEventListener("click", () => applySelectedDate("conFromDate", "clear"));
    el("btn-bulk-clear-conto").addEventListener("click", () => applySelectedDate("conToDate", "clear"));

    const orderedIds = sortCablesLikeExcel(filteredCables()).map((cable) => cable.id);
    document.querySelectorAll(".row-select").forEach((input) => {
      input.addEventListener("click", (event) => {
        event.stopPropagation();
        const id = input.dataset.selectId;
        const checked = input.checked;
        if (event.shiftKey && lastSelectedCableId && lastSelectedCableId !== id) {
          const from = orderedIds.indexOf(lastSelectedCableId);
          const to = orderedIds.indexOf(id);
          if (from > -1 && to > -1) {
            const range = orderedIds.slice(Math.min(from, to), Math.max(from, to) + 1);
            setTableSelection(range, checked);
          } else {
            setTableSelection([id], checked);
          }
        } else {
          setTableSelection([id], checked);
        }
        lastSelectedCableId = id;
        renderCableList();
      });
    });

    document.querySelectorAll("[data-page]").forEach((button) => {
      button.addEventListener("click", () => {
        tablePage = Number(button.dataset.page);
        renderCableList();
      });
    });

    document.querySelectorAll(".editable-cell").forEach((cell) => {
      cell.addEventListener("focus", () => {
        document.querySelectorAll(".editable-cell.active-cell").forEach((active) => active.classList.remove("active-cell"));
        cell.classList.add("active-cell");
        cell.dataset.original = cell.textContent.trim();
      });
      cell.addEventListener("keydown", (event) => {
        if (event.key === "Enter") {
          event.preventDefault();
          cell.blur();
        }
        if (event.key === "Escape") {
          event.preventDefault();
          cell.textContent = cell.dataset.original || "";
          cell.blur();
        }
      });
      cell.addEventListener("dblclick", (event) => {
        if (["installDate", "conFromDate", "conToDate"].includes(cell.dataset.field)) {
          event.preventDefault();
          handleDateDoubleClick(cell);
        }
      });
      cell.addEventListener("blur", () => commitCell(cell));
    });
  }

  function applyTodayToFiltered(field) {
    const rows = filteredCables();
    const label = Store.FIELD_LABELS[field] || field;
    let updated = 0;
    let skipped = 0;
    rows.forEach((cable) => {
      if ((field === "conFromDate" || field === "conToDate") && !cable.installDate) {
        skipped += 1;
        return;
      }
      if (cable[field] === today()) return;
      try {
        Store.updateCable(cable.id, { [field]: today() }, { reason: `${label} 일괄 입력` });
        updated += 1;
      } catch (error) {
        skipped += 1;
      }
    });
    toast(`${label} ${updated.toLocaleString("ko-KR")}건 반영${skipped ? ` · 제외 ${skipped.toLocaleString("ko-KR")}건` : ""}`);
    renderCableList();
  }

  function applySelectedDate(field, mode = "set") {
    const rows = selectedCables();
    if (!rows.length) {
      toast("선택된 행이 없습니다.");
      return;
    }

    const label = Store.FIELD_LABELS[field] || field;
    const nextValue = mode === "set" ? (el("bulk-date").value || today()) : "";
    if (mode === "set" && !nextValue) {
      toast("일괄 입력할 날짜를 먼저 선택하세요.");
      return;
    }

    let updated = 0;
    let skipped = 0;
    rows.forEach((cable) => {
      if (mode === "set" && (field === "conFromDate" || field === "conToDate") && !cable.installDate) {
        skipped += 1;
        return;
      }
      if ((cable[field] || "") === nextValue) return;
      try {
        Store.updateCable(cable.id, { [field]: nextValue }, { reason: mode === "set" ? `${label} 선택행 일괄 입력` : `${label} 선택행 일괄 삭제` });
        updated += 1;
      } catch (error) {
        skipped += 1;
      }
    });

    const actionLabel = mode === "set" ? "반영" : "삭제";
    toast(`${label} ${updated.toLocaleString("ko-KR")}건 ${actionLabel}${skipped ? ` · 제외 ${skipped.toLocaleString("ko-KR")}건` : ""}`);
    renderCableList();
  }

  function commitCell(cell) {
    const field = cell.dataset.field;
    const id = cell.dataset.cableId;
    const original = cell.dataset.original || "";
    let next = cell.textContent.trim();
    if (["installDate", "conFromDate", "conToDate"].includes(field) && next === "-") next = "";
    if (original === next) return;

    const cable = Store.getCables().find((item) => item.id === id);
    if (!cable) {
      toast("수정할 케이블을 찾을 수 없습니다.");
      renderCableList();
      return;
    }

    if ((field === "conFromDate" || field === "conToDate") && next && !cable.installDate) {
      alert("포설실적부터 입력하세요.");
      renderCableList();
      return;
    }

    try {
      Store.updateCable(id, { [field]: next }, { reason: "셀 수정" });
      toast(`${cable.circuitNo} ${Store.FIELD_LABELS[field] || field} 수정 완료`);
      renderCableList();
    } catch (error) {
      toast(error.message);
      renderCableList();
    }
  }

  function handleDateDoubleClick(cell) {
    const field = cell.dataset.field;
    const id = cell.dataset.cableId;
    const cable = Store.getCables().find((item) => item.id === id);
    if (!cable) return;

    if ((field === "conFromDate" || field === "conToDate") && !cable.installDate) {
      alert("포설실적부터 입력하세요.");
      return;
    }

    const next = cable[field] ? "" : today();
    const label = Store.FIELD_LABELS[field] || field;
    try {
      Store.updateCable(id, { [field]: next }, { reason: next ? `${label} 입력` : `${label} 삭제` });
      toast(`${cable.circuitNo} ${label}: ${next || "삭제"}`);
      renderCableList();
    } catch (error) {
      toast(error.message);
    }
  }

  function filteredCables() {
    const query = tableFilter.query.trim().toLowerCase();
    return Store.getCables().filter((cable) => {
      const isDeleted = Boolean(cable.deleted);
      if (cableTab === "active" && isDeleted) return false;
      if (cableTab === "deleted" && !isDeleted) return false;
      if (tableFilter.sys && (cable.sys || "미분류") !== tableFilter.sys) return false;
      if (tableFilter.type && cable.cableType !== tableFilter.type) return false;
      if (tableFilter.node && cable.node !== tableFilter.node) return false;
      if (tableFilter.installStatus === "filled" && !cable.installDate) return false;
      if (tableFilter.installStatus === "empty" && cable.installDate) return false;
      if (tableFilter.conFromStatus === "filled" && !cable.conFromDate) return false;
      if (tableFilter.conFromStatus === "empty" && cable.conFromDate) return false;
      if (tableFilter.conToStatus === "filled" && !cable.conToDate) return false;
      if (tableFilter.conToStatus === "empty" && cable.conToDate) return false;
      if (query) {
        const haystack = [
          cable.sys, cable.circuitNo, cable.cableType, cable.fromEquipment, cable.toEquipment,
          cable.fromCode, cable.toCode, cable.route, cable.node, cable.rev,
        ].join(" ").toLowerCase();
        return haystack.includes(query);
      }
      return true;
    });
  }

  function sortCablesLikeExcel(cables) {
    return [...cables].sort((left, right) => {
      const leftTotal = Number(left.total) || 0;
      const rightTotal = Number(right.total) || 0;
      const leftKey = leftTotal > 0 ? leftTotal : Number.MAX_SAFE_INTEGER;
      const rightKey = rightTotal > 0 ? rightTotal : Number.MAX_SAFE_INTEGER;
      return leftKey - rightKey;
    });
  }

  function renderPager(pages) {
    if (pages <= 1) return "";
    const start = Math.max(1, tablePage - 2);
    const end = Math.min(pages, tablePage + 2);
    const buttons = [];
    buttons.push(`<button data-page="1" type="button">처음</button>`);
    for (let page = start; page <= end; page += 1) {
      buttons.push(`<button class="${page === tablePage ? "active" : ""}" data-page="${page}" type="button">${page}</button>`);
    }
    buttons.push(`<button data-page="${pages}" type="button">끝</button>`);
    return `<div class="pager">${buttons.join("")}</div>`;
  }

  function renderPhone() {
    const vessel = Store.getActiveVessel();
    const phones = Store.getPhones();
    main().innerHTML = `
      ${pageHead("Phone", `${vessel.id} 호선 연락처를 등록하고 바로 전화 연결합니다.`)}
      <section class="panel">
        <h2 class="panel-title">전화 등록</h2>
        <form id="phone-form" class="grid four">
          <div class="field"><label>전화명</label><input name="label" placeholder="현장소장" required></div>
          <div class="field"><label>전화번호</label><input name="phone" type="tel" placeholder="010-0000-0000" required></div>
          <div class="field"><label>위치</label><input name="location" placeholder="W/H, 기관실"></div>
          <div class="field"><label>메모</label><input name="memo" placeholder="통화 목적"></div>
          <div class="actions" style="grid-column:1 / -1"><button class="btn green" type="submit">전화 등록</button></div>
        </form>
      </section>

      <section class="panel">
        <h2 class="panel-title">전화 목록</h2>
        ${phones.length ? `<div class="list">
          ${phones.map((phone) => `
            <div class="item">
              <div>
                <div class="item-title">${escapeHtml(phone.label)} · <span class="mono">${escapeHtml(phone.phone)}</span></div>
                <div class="item-meta">${escapeHtml(phone.location || "-")} · ${escapeHtml(phone.memo || "-")}</div>
              </div>
              <div class="actions">
                <a class="btn small green" href="tel:${encodeURIComponent(phone.phone)}">전화걸기</a>
                <button class="btn small ghost" data-delete-phone="${escapeAttr(phone.id)}" type="button">삭제</button>
              </div>
            </div>`).join("")}
        </div>` : `<p class="empty">등록된 전화가 없습니다.</p>`}
      </section>`;

    el("phone-form").addEventListener("submit", (event) => {
      event.preventDefault();
      try {
        const phone = Store.addPhone(formData(event.currentTarget));
        toast(`${phone.label} 전화번호를 등록했습니다.`);
        renderPhone();
      } catch (error) {
        toast(error.message);
      }
    });

    document.querySelectorAll("[data-delete-phone]").forEach((button) => {
      button.addEventListener("click", () => {
        Store.deletePhone(button.dataset.deletePhone);
        toast("전화번호를 삭제했습니다.");
        renderPhone();
      });
    });
  }

  async function handleFileImport(event) {
    const file = event.target.files?.[0];
    event.target.value = "";
    if (!file) return;

    const confirmed = confirm(
      "업데이트 작업을 시작하시겠습니까?\n\n" +
      "[확인 사항]\n" +
      "불러오는 Excel/TSV 데이터는 현재 Cable List와 CIRCUIT NO. 기준으로 비교됩니다.\n" +
      "추가/수정/삭제/복원으로 판단된 내용은 REV 히스토리에 자동 기록됩니다."
    );
    if (!confirmed) return;

    const markMissingAsDeleted = confirm(
      "신규 파일에 없는 케이블을 삭제 처리할까요?\n\n" +
      "확인: 현재 List에는 있는데 신규 파일에는 없는 행을 [삭제]로 표시합니다.\n" +
      "취소: 누락된 행은 그대로 둡니다 (추가/수정만 반영)."
    );

    try {
      const rows = await readRowsFromFile(file);
      const summary = Store.importCableRows(rows, { sourceName: file.name, markMissingAsDeleted });
      const parts = [
        `추가 ${summary.added}`,
        `수정 ${summary.updated}`,
        `삭제 ${summary.deleted}`,
      ];
      if (summary.restored) parts.push(`복원 ${summary.restored}`);
      if (summary.missingDeleted) parts.push(`누락삭제 ${summary.missingDeleted}`);
      parts.push(`동일 ${summary.unchanged}`);
      toast(`업데이트 완료: ${parts.join(" · ")}`);
      renderMode();
    } catch (error) {
      console.error(error);
      toast(error.message || "불러오기 중 오류가 발생했습니다.");
    }
  }

  async function readRowsFromFile(file) {
    const lower = file.name.toLowerCase();
    if (/\.(xlsx|xls|xlsm|xlsb)$/.test(lower)) {
      await ensureXlsx();
      const workbook = XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: false, raw: false });
      const sheetName = chooseSheetName(workbook);
      const sheet = workbook.Sheets[sheetName];
      return XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: "" });
    }

    const text = await file.text();
    return Store.parseDelimited(text, lower.endsWith(".csv") ? "," : "\t");
  }

  function chooseSheetName(workbook) {
    const rev = workbook.SheetNames.find((name) => name.trim().toLowerCase() === "rev");
    if (rev) return rev;
    const withCircuit = workbook.SheetNames.find((name) => {
      const rows = XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1, raw: false, defval: "", range: 0 });
      return rows.slice(0, 30).some((row) => row.some((cell) => String(cell).replace(/[\s.\r\n"]/g, "").toUpperCase().includes("CIRCUITNO")));
    });
    return withCircuit || workbook.SheetNames[0];
  }

  function ensureXlsx() {
    if (window.XLSX) return Promise.resolve();
    return new Promise((resolve, reject) => {
      const existing = document.querySelector(`script[src="${XLSX_CDN}"]`);
      if (existing) {
        existing.addEventListener("load", resolve, { once: true });
        existing.addEventListener("error", reject, { once: true });
        return;
      }
      const script = document.createElement("script");
      script.src = XLSX_CDN;
      script.onload = resolve;
      script.onerror = () => reject(new Error("Excel 파서를 불러오지 못했습니다. 네트워크를 확인하세요."));
      document.head.appendChild(script);
    });
  }

  function download(filename, content, type = "text/tab-separated-values;charset=utf-8") {
    const blob = new Blob(["\uFEFF", content], { type });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = filename;
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    URL.revokeObjectURL(url);
  }

  function bindChrome() {
    const undoBtn = el("btn-undo");
    if (undoBtn) {
      undoBtn.addEventListener("click", () => {
        if (!Store.canUndo || !Store.canUndo()) {
          toast("되돌릴 작업이 없습니다.");
          return;
        }
        if (Store.undo()) {
          toast("이전 상태로 되돌렸습니다.");
          renderMode();
        }
      });
    }

    document.addEventListener("keydown", (event) => {
      const isUndoKey = (event.ctrlKey || event.metaKey) && !event.shiftKey && !event.altKey && (event.key === "z" || event.key === "Z");
      if (!isUndoKey) return;
      const target = event.target;
      const tag = (target?.tagName || "").toLowerCase();
      if (tag === "input" || tag === "textarea" || target?.isContentEditable) return;
      event.preventDefault();
      if (Store.canUndo && Store.canUndo() && Store.undo()) {
        toast("이전 상태로 되돌렸습니다.");
        renderMode();
      } else {
        toast("되돌릴 작업이 없습니다.");
      }
    });

    el("btn-load-default").addEventListener("click", async () => {
      if (!confirm("현재 호선의 Cable List를 HK2401 기본 Cable List로 교체할까요?")) return;
      try {
        const count = await Store.loadDefaultCables({ replace: true });
        toast(`HK2401 Cable List ${count.toLocaleString("ko-KR")}건을 불러왔습니다.`);
        renderMode();
      } catch (error) {
        toast(error.message);
      }
    });

    el("btn-import-file").addEventListener("click", () => el("file-input").click());
    el("file-input").addEventListener("change", handleFileImport);
    el("btn-export-all").addEventListener("click", () => {
      if (typeof window.exportToExcel === "function") {
        window.exportToExcel();
      } else {
        toast("Excel 라이브러리 로드 중... 잠시 후 다시 시도하세요.");
      }
    });
    el("btn-back-vessel").addEventListener("click", () => {
      setMode("select");
      renderMode();
    });
    el("btn-admin").addEventListener("click", () => {
      currentUser = { name: "Admin", role: "Admin" };
      setMode("admin");
      renderMode();
    });
    el("btn-logout").addEventListener("click", () => {
      currentUser = null;
      sessionStorage.removeItem(USER_SESSION_KEY);
      setMode("name");
      history.replaceState(null, "", location.pathname);
      renderMode();
    });
  }

  return {
    updateDateCells(cells, date) {
      const batches = {};
      cells.forEach((cell) => {
        const field = cell.dataset.field;
        const id = cell.dataset.cableId;
        if (field && id) {
          if (!batches[id]) batches[id] = {};
          batches[id][field] = date;
        }
      });
      let updated = 0;
      Object.entries(batches).forEach(([id, patch]) => {
        try {
          Store.updateCable(id, patch, { reason: date ? "다중셀 날짜 입력" : "다중셀 날짜 삭제" });
          updated++;
        } catch (_) {}
      });
      if (updated > 0) {
        toast(`${cells.length}개 셀에 날짜 적용 완료`);
        renderCableList();
      }
    },
    async init() {
      await fetchConfig(); // CF 환경변수 먼저 로드
      await Store.init();
      Store.subscribe(() => updateUndoButton());
      bindChrome();

      const hash = decodeURIComponent(location.hash.replace(/^#/, ""));
      Store.setActiveVessel(Store.DEFAULT_VESSEL_ID);
      const savedUser = sessionStorage.getItem(USER_SESSION_KEY);
      if (savedUser) {
        try {
          currentUser = JSON.parse(savedUser);
        } catch {
          sessionStorage.removeItem(USER_SESSION_KEY);
          currentUser = null;
        }
      }

      if (hash === "admin") {
        currentUser = { name: "Admin", role: "Admin" };
        setMode("admin");
      } else if (hash === "vessel" || hash === "select") {
        if (currentUser) {
          setMode("select");
        } else {
          pendingView = "cableList";
          setMode("name");
        }
      } else if (views[hash]) {
        pendingView = hash;
        if (currentUser) {
          currentView = hash;
          setMode("app");
        } else {
          setMode("name");
        }
      } else if (currentUser) {
        currentView = "cableList";
        setMode("app");
      } else {
        pendingView = "cableList";
        setMode("name");
      }
      renderMode();
    },
  };
})();

window.addEventListener("DOMContentLoaded", () => {
  App.init().catch((error) => {
    console.error(error);
    const main = document.getElementById("main-content");
    if (main) {
      main.innerHTML = `<section class="panel"><h1>초기화 오류</h1><p>${String(error.message || error)}</p></section>`;
    }
  });
});
