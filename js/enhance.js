"use strict";

/**
 * SEcMS · HK2401 Cable List · Enhance Layer (non-invasive)
 *
 * main.js를 수정하지 않고 다음을 추가한다:
 *  1. Excel-like 다중셀 선택 자동 attach (.excel-table.cable-data-table 감지)
 *  2. 헤더에 "📜 작업이력" 버튼 inject + 모달 viewer
 *  3. 다중셀 선택 카운트 floating badge
 *  4. SEcMS 통합 메뉴 톤(secms.tech 외부 링크) — main.js의 brand-link는 이미 secms.tech를 가리킴
 *
 * 의존: ExcelSelect, AuditLog, Store
 */
(function () {
  function el(id) { return document.getElementById(id); }

  // ============================================================
  // 1. Excel multi-cell selection: attach on every renderCableList
  // ============================================================
  let attachedTable = null;
  function tryAttachExcelSelect() {
    const table = document.querySelector(".excel-table.cable-data-table");
    if (!table) {
      if (attachedTable) {
        try { ExcelSelect.detach(attachedTable); } catch (e) { /* noop */ }
        attachedTable = null;
      }
      return;
    }
    if (table === attachedTable) {
      // 같은 테이블이지만 tbody가 main.js로 새로 그려졌을 수 있음 → 정렬 헤더 다시 attach
      enhanceSortableHeaders(table);
      return;
    }
    if (attachedTable) {
      try { ExcelSelect.detach(attachedTable); } catch (e) { /* noop */ }
    }
    attachedTable = table;
    ExcelSelect.attach(table, {
      columnsToSkip: ["col-select", "excel-row-head"],
      onCopy: (count, tsv) => {
        const lines = tsv.split(/\r?\n/).length;
        showToast(`${count}개 셀(${lines}행) Excel 형식으로 복사됨`);
        if (window.AuditLog) {
          AuditLog.record("export", `다중셀 복사: ${count}개`, { cellCount: count, rows: lines });
        }
      },
    });

    // 정렬 가능한 헤더 enhance
    enhanceSortableHeaders(table);
  }

  // ============================================================
  // 1-b. 헤더 클릭 정렬 (오름차순 ↔ 내림차순 토글)
  // ============================================================
  // 컬럼 인덱스 → 정렬 방향 ('asc' | 'desc' | null)
  const sortState = { colIndex: -1, dir: null };

  function enhanceSortableHeaders(table) {
    const headRow = table.querySelector("thead tr.excel-headline");
    if (!headRow) return;
    if (headRow.dataset.sortableEnhanced === "1") {
      // 매 렌더 후 화살표만 갱신
      updateSortIndicators(headRow);
      // 마지막 정렬 상태가 있으면 본문 다시 정렬
      if (sortState.colIndex >= 0 && sortState.dir) {
        applySortToBody(table, sortState.colIndex, sortState.dir);
      }
      return;
    }

    headRow.dataset.sortableEnhanced = "1";

    Array.from(headRow.children).forEach((th, idx) => {
      // 행번호/체크박스 컬럼 제외
      if (th.classList.contains("excel-row-head") || th.classList.contains("col-select")) return;
      th.classList.add("sortable-th");

      // 정렬 화살표 슬롯 추가
      if (!th.querySelector(".sort-indicator")) {
        const ind = document.createElement("span");
        ind.className = "sort-indicator";
        ind.setAttribute("aria-hidden", "true");
        th.appendChild(ind);
      }

      th.addEventListener("click", (e) => {
        // contenteditable 이벤트와 충돌 방지: 헤더는 contenteditable 아님
        e.preventDefault();
        if (sortState.colIndex === idx) {
          sortState.dir = sortState.dir === "asc" ? "desc"
                       : sortState.dir === "desc" ? null
                       : "asc";
          if (!sortState.dir) sortState.colIndex = -1;
        } else {
          sortState.colIndex = idx;
          sortState.dir = "asc";
        }
        updateSortIndicators(headRow);
        applySortToBody(table, sortState.colIndex, sortState.dir);
        if (window.AuditLog) {
          const label = (th.querySelector(".excel-header-label")?.textContent || th.textContent || "").replace(/\s+/g, " ").trim();
          AuditLog.record("system", `정렬: ${label} ${sortState.dir || "원본순"}`, { col: idx, dir: sortState.dir });
        }
      });
    });

    updateSortIndicators(headRow);
    if (sortState.colIndex >= 0 && sortState.dir) {
      applySortToBody(table, sortState.colIndex, sortState.dir);
    }
  }

  function updateSortIndicators(headRow) {
    Array.from(headRow.children).forEach((th, idx) => {
      const ind = th.querySelector(".sort-indicator");
      if (!ind) return;
      th.classList.remove("sort-asc", "sort-desc");
      if (idx === sortState.colIndex && sortState.dir) {
        ind.textContent = sortState.dir === "asc" ? "▲" : "▼";
        th.classList.add(sortState.dir === "asc" ? "sort-asc" : "sort-desc");
      } else {
        ind.textContent = "";
      }
    });
  }

  function applySortToBody(table, colIndex, dir) {
    const tbody = table.querySelector("tbody");
    if (!tbody || colIndex < 0 || !dir) return;
    const trs = Array.from(tbody.querySelectorAll("tr"));
    if (!trs.length) return;

    // 원본 순서 백업 (한 번만)
    if (!tbody.dataset.originalOrder) {
      trs.forEach((tr, i) => { tr.dataset.origIdx = String(i); });
      tbody.dataset.originalOrder = "1";
    }

    if (dir === null) {
      // 원본 순서 복원
      trs.sort((a, b) => Number(a.dataset.origIdx || 0) - Number(b.dataset.origIdx || 0));
    } else {
      const sign = dir === "asc" ? 1 : -1;
      trs.sort((a, b) => {
        const va = sortKey((a.children[colIndex] || {}).textContent || "");
        const vb = sortKey((b.children[colIndex] || {}).textContent || "");
        if (va.kind === "num" && vb.kind === "num") return (va.value - vb.value) * sign;
        if (va.kind === "date" && vb.kind === "date") return (va.value - vb.value) * sign;
        return va.text.localeCompare(vb.text, "ko", { numeric: true, sensitivity: "base" }) * sign;
      });
    }

    // DOM 재배치 (한 번에)
    const frag = document.createDocumentFragment();
    trs.forEach((tr) => frag.appendChild(tr));
    tbody.appendChild(frag);
  }

  function sortKey(raw) {
    const text = String(raw || "").trim();
    if (!text) return { kind: "empty", text: "", value: 0 };
    // 날짜 (YYYY-MM-DD)
    const dateMatch = text.match(/^(\d{4})[-./](\d{1,2})[-./](\d{1,2})$/);
    if (dateMatch) {
      const t = Date.parse(`${dateMatch[1]}-${dateMatch[2].padStart(2,"0")}-${dateMatch[3].padStart(2,"0")}`);
      if (!isNaN(t)) return { kind: "date", text, value: t };
    }
    // 숫자 (콤마 포함)
    const numMatch = text.replace(/,/g, "");
    if (/^-?\d+(\.\d+)?$/.test(numMatch)) {
      return { kind: "num", text, value: Number(numMatch) };
    }
    return { kind: "text", text, value: 0 };
  }

  // ============================================================
  // 2. Floating selection badge (.show 시 표시)
  // ============================================================
  function ensureSelectionBadge() {
    if (el("excel-select-badge")) return;
    const badge = document.createElement("div");
    badge.id = "excel-select-badge";
    badge.className = "excel-select-badge";
    document.body.appendChild(badge);
  }

  // ============================================================
  // 3. Audit Log 헤더 버튼 + 모달
  // ============================================================
  function ensureAuditButton() {
    if (el("btn-audit-log")) return;
    const topbar = document.querySelector(".topbar");
    if (!topbar) return;
    const adminBtn = el("btn-admin");
    if (!adminBtn) return;

    const btn = document.createElement("button");
    btn.id = "btn-audit-log";
    btn.className = "topbar-btn icon-only audit-symbol";
    btn.type = "button";
    btn.title = "작업 이력 (Audit Log)";
    btn.setAttribute("aria-label", "작업 이력");
    btn.textContent = "📜";
    btn.addEventListener("click", openAuditModal);
    adminBtn.parentNode.insertBefore(btn, adminBtn);
  }

  function openAuditModal() {
    let modal = el("audit-log-modal");
    if (!modal) {
      modal = document.createElement("div");
      modal.id = "audit-log-modal";
      modal.className = "audit-modal";
      modal.innerHTML = `
        <div class="audit-modal-backdrop" data-close></div>
        <div class="audit-modal-card" role="dialog" aria-modal="true" aria-labelledby="audit-modal-title">
          <header class="audit-modal-head">
            <h2 id="audit-modal-title">📜 작업 이력</h2>
            <div class="audit-modal-actions">
              <select id="audit-cat-filter" aria-label="카테고리 필터">
                <option value="">전체 카테고리</option>
              </select>
              <input id="audit-search" type="search" placeholder="검색어 입력..." aria-label="이력 검색">
              <button class="btn small ghost" id="audit-export-tsv" type="button">TSV 내보내기</button>
              <button class="btn small ghost" id="audit-export-jsonl" type="button">JSONL 내보내기</button>
              <button class="btn small red" id="audit-clear" type="button">전체 삭제</button>
              <button class="audit-modal-close" data-close type="button" aria-label="닫기">×</button>
            </div>
          </header>
          <div class="audit-modal-body">
            <div class="audit-summary" id="audit-summary"></div>
            <div class="audit-table-wrap">
              <table class="audit-table" id="audit-table">
                <thead>
                  <tr>
                    <th class="col-time">시각</th>
                    <th class="col-cat">카테고리</th>
                    <th class="col-actor">작업자</th>
                    <th class="col-vessel">호선</th>
                    <th class="col-action">내용</th>
                    <th class="col-detail">세부</th>
                  </tr>
                </thead>
                <tbody></tbody>
              </table>
            </div>
          </div>
        </div>`;
      document.body.appendChild(modal);

      // 카테고리 옵션 채우기
      const sel = modal.querySelector("#audit-cat-filter");
      if (sel && window.AuditLog && AuditLog.CATEGORY_LABELS) {
        Object.entries(AuditLog.CATEGORY_LABELS).forEach(([k, v]) => {
          const opt = document.createElement("option");
          opt.value = k;
          opt.textContent = v;
          sel.appendChild(opt);
        });
      }

      // 닫기 핸들러
      modal.querySelectorAll("[data-close]").forEach((node) => {
        node.addEventListener("click", () => {
          modal.classList.remove("show");
        });
      });
      // ESC
      document.addEventListener("keydown", (e) => {
        if (e.key === "Escape" && modal.classList.contains("show")) {
          modal.classList.remove("show");
        }
      });

      // 필터/검색
      modal.querySelector("#audit-cat-filter").addEventListener("change", renderAuditList);
      modal.querySelector("#audit-search").addEventListener("input", renderAuditList);
      modal.querySelector("#audit-export-tsv").addEventListener("click", () => {
        if (!window.AuditLog) return;
        downloadFile(`hk2401-audit-${new Date().toISOString().slice(0, 10)}.tsv`, "﻿" + AuditLog.exportTsv(), "text/tab-separated-values;charset=utf-8");
        showToast("이력을 TSV 파일로 저장했습니다.");
      });
      modal.querySelector("#audit-export-jsonl").addEventListener("click", () => {
        if (!window.AuditLog) return;
        downloadFile(`hk2401-audit-${new Date().toISOString().slice(0, 10)}.jsonl`, AuditLog.exportJsonl(), "application/x-ndjson;charset=utf-8");
        showToast("이력을 JSONL 파일로 저장했습니다.");
      });
      modal.querySelector("#audit-clear").addEventListener("click", () => {
        if (!window.AuditLog) return;
        if (!confirm("작업 이력을 모두 삭제할까요? 되돌릴 수 없습니다.")) return;
        AuditLog.clear();
        renderAuditList();
        showToast("작업 이력을 전체 삭제했습니다.");
      });

      // 변경 시 자동 갱신
      if (window.AuditLog && typeof AuditLog.subscribe === "function") {
        AuditLog.subscribe(() => {
          if (modal.classList.contains("show")) renderAuditList();
        });
      }
    }
    modal.classList.add("show");
    renderAuditList();
  }

  function renderAuditList() {
    if (!window.AuditLog) return;
    const modal = el("audit-log-modal");
    if (!modal) return;
    const cat = modal.querySelector("#audit-cat-filter").value;
    const search = modal.querySelector("#audit-search").value;
    const entries = AuditLog.list({ category: cat, search, limit: 500 });

    const summary = modal.querySelector("#audit-summary");
    summary.innerHTML = `<strong>${AuditLog.count.toLocaleString("ko-KR")}건</strong>의 이력 (현재 표시 ${entries.length.toLocaleString("ko-KR")}건)`;

    const tbody = modal.querySelector("#audit-table tbody");
    if (!entries.length) {
      tbody.innerHTML = `<tr><td colspan="6" class="audit-empty">이력이 없습니다.</td></tr>`;
      return;
    }

    tbody.innerHTML = entries.map((e) => `
      <tr>
        <td class="col-time mono">${formatTime(e.timestamp)}</td>
        <td class="col-cat">${escapeHtml(AuditLog.categoryLabel(e.category))}</td>
        <td class="col-actor">${escapeHtml(e.actor)}</td>
        <td class="col-vessel mono">${escapeHtml(e.vesselId)}</td>
        <td class="col-action">${escapeHtml(e.action)}</td>
        <td class="col-detail"><pre>${escapeHtml(JSON.stringify(e.details, null, 0))}</pre></td>
      </tr>
    `).join("");
  }

  // ============================================================
  // Helpers
  // ============================================================
  function formatTime(iso) {
    try {
      const d = new Date(iso);
      const Y = d.getFullYear();
      const M = String(d.getMonth() + 1).padStart(2, "0");
      const D = String(d.getDate()).padStart(2, "0");
      const h = String(d.getHours()).padStart(2, "0");
      const m = String(d.getMinutes()).padStart(2, "0");
      const s = String(d.getSeconds()).padStart(2, "0");
      return `${Y}-${M}-${D} ${h}:${m}:${s}`;
    } catch (e) {
      return iso;
    }
  }

  function escapeHtml(value) {
    return String(value ?? "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  function showToast(msg) {
    const t = el("toast");
    if (!t) { console.log(msg); return; }
    t.textContent = msg;
    t.classList.add("show");
    clearTimeout(showToast.timer);
    showToast.timer = setTimeout(() => t.classList.remove("show"), 2400);
  }

  function downloadFile(name, content, mime = "text/plain;charset=utf-8") {
    const blob = new Blob([content], { type: mime });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = name;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => {
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    }, 0);
  }

  // ============================================================
  // Boot
  // ============================================================
  function init() {
    ensureSelectionBadge();
    ensureAuditButton();
    tryAttachExcelSelect();

    // main()의 innerHTML이 매번 새로 그려지므로, MutationObserver로 감지
    const observer = new MutationObserver(() => {
      ensureAuditButton();
      tryAttachExcelSelect();
    });
    observer.observe(document.body, { childList: true, subtree: true });

    // 시스템 부팅 이력
    if (window.AuditLog) {
      AuditLog.record("system", "앱 부팅", {
        ua: navigator.userAgent,
        viewport: `${window.innerWidth}x${window.innerHeight}`,
      });
    }
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
})();
