"use strict";

/**
 * SEcMS · HK2401 Cable List · Excel-like Multi-Cell Selection
 *
 * Excel과 유사한 다중 셀 선택을 제공한다:
 *  - 마우스 드래그: 사각형 범위 선택
 *  - Shift+클릭: 마지막 anchor 셀로부터 범위 확장
 *  - Ctrl/Cmd+클릭: 개별 셀 토글 (다중 영역)
 *  - 행 헤더(SEL) 더블클릭/Shift: 해당 행 전체 선택
 *  - Ctrl+A: 화면의 모든 셀 선택
 *  - Ctrl+C: 선택 영역을 TSV로 클립보드에 복사 (Excel 호환)
 *  - Esc: 선택 해제
 *
 * 셀은 td.editable-cell 또는 .excel-cell 이어야 하며, contenteditable과 충돌하지 않도록
 * 셀이 편집 모드(focused contenteditable)일 때는 selection 무시.
 *
 * Public API:
 *   ExcelSelect.attach(table, { onCopy?, columnsToSkip? })
 *   ExcelSelect.detach(table)
 *   ExcelSelect.clear(table)
 */
const ExcelSelect = (() => {
  const ATTR_DATA_KEY = "__excelSelectState";

  // 셀의 (row, col) 인덱스를 계산
  function getCellPos(cell) {
    const tr = cell.closest("tr");
    if (!tr) return null;
    const tbody = tr.parentElement;
    if (!tbody) return null;
    const rowIndex = Array.prototype.indexOf.call(tbody.children, tr);
    // td 인덱스는 row 안에서의 cellIndex 기준
    let colIndex = 0;
    for (const c of tr.children) {
      if (c === cell) break;
      colIndex += 1;
    }
    return { rowIndex, colIndex, tr, tbody };
  }

  function getCellAt(tbody, rowIndex, colIndex) {
    const tr = tbody.children[rowIndex];
    if (!tr) return null;
    return tr.children[colIndex] || null;
  }

  function isSelectableCell(cell, state) {
    if (!cell || cell.tagName !== "TD") return false;
    if (state.columnsToSkip && state.columnsToSkip.length) {
      for (const skipClass of state.columnsToSkip) {
        if (cell.classList.contains(skipClass)) return false;
      }
    }
    return cell.classList.contains("excel-cell") || cell.classList.contains("editable-cell");
  }

  function clearSelection(state) {
    state.tbody.querySelectorAll(".excel-cell-selected").forEach((c) => c.classList.remove("excel-cell-selected"));
    state.tbody.querySelectorAll(".excel-cell-anchor").forEach((c) => c.classList.remove("excel-cell-anchor"));
    state.selectedCells.clear();
    state.anchor = null;
    state.dragging = false;
    updateBadge(state);
  }

  function setSelected(cell, state, on = true) {
    if (!cell) return;
    if (on) {
      cell.classList.add("excel-cell-selected");
      state.selectedCells.add(cell);
    } else {
      cell.classList.remove("excel-cell-selected");
      state.selectedCells.delete(cell);
    }
  }

  function selectRange(state, start, end) {
    const r1 = Math.min(start.rowIndex, end.rowIndex);
    const r2 = Math.max(start.rowIndex, end.rowIndex);
    const c1 = Math.min(start.colIndex, end.colIndex);
    const c2 = Math.max(start.colIndex, end.colIndex);

    // 선택 영역 채우기
    for (let r = r1; r <= r2; r += 1) {
      for (let c = c1; c <= c2; c += 1) {
        const cell = getCellAt(state.tbody, r, c);
        if (cell && isSelectableCell(cell, state)) setSelected(cell, state, true);
      }
    }
  }

  function rectFromSelection(state) {
    if (!state.selectedCells.size) return null;
    let r1 = Infinity, r2 = -1, c1 = Infinity, c2 = -1;
    state.selectedCells.forEach((cell) => {
      const pos = getCellPos(cell);
      if (!pos) return;
      r1 = Math.min(r1, pos.rowIndex);
      r2 = Math.max(r2, pos.rowIndex);
      c1 = Math.min(c1, pos.colIndex);
      c2 = Math.max(c2, pos.colIndex);
    });
    if (r2 < 0) return null;
    return { r1, r2, c1, c2 };
  }

  function buildTsvFromSelection(state) {
    const rect = rectFromSelection(state);
    if (!rect) return "";
    const { r1, r2, c1, c2 } = rect;
    const lines = [];
    for (let r = r1; r <= r2; r += 1) {
      const cells = [];
      for (let c = c1; c <= c2; c += 1) {
        const cell = getCellAt(state.tbody, r, c);
        // 선택 안 된 셀은 빈 문자열로 둔다 (Excel 패턴)
        if (cell && state.selectedCells.has(cell)) {
          // text-only로 추출 (contentEditable 화이트스페이스 보존)
          const text = (cell.innerText || cell.textContent || "").replace(/\r\n|\r/g, "\n");
          // TSV 이스케이프: \t, \n, " 포함되면 따옴표로 감싸고 " → ""
          if (/[\t\r\n"]/.test(text)) {
            cells.push(`"${text.replace(/"/g, '""')}"`);
          } else {
            cells.push(text);
          }
        } else {
          cells.push("");
        }
      }
      lines.push(cells.join("\t"));
    }
    return lines.join("\r\n");
  }

  function updateBadge(state) {
    const badge = document.getElementById("excel-select-badge");
    if (!badge) return;
    const count = state.selectedCells.size;
    if (count <= 1) {
      badge.classList.remove("show");
      badge.textContent = "";
      return;
    }
    const rect = rectFromSelection(state);
    if (!rect) return;
    const rows = rect.r2 - rect.r1 + 1;
    const cols = rect.c2 - rect.c1 + 1;
    badge.classList.add("show");
    badge.innerHTML = `<strong>${count}</strong>개 셀 · ${rows}행 × ${cols}열 선택 <span class="excel-select-hint">(Ctrl+C 복사)</span>`;
  }

  function isEditingActiveCell(cell) {
    // 현재 contenteditable로 포커스되어 있고 텍스트 selection이 있다면 편집 중
    if (!cell || cell.contentEditable !== "true") return false;
    if (document.activeElement !== cell) return false;
    const sel = document.getSelection();
    if (!sel) return false;
    if (sel.isCollapsed) return false;
    if (!cell.contains(sel.anchorNode)) return false;
    return true;
  }

  function attach(table, options = {}) {
    if (!table) return;
    const tbody = table.querySelector("tbody");
    if (!tbody) return;

    const state = {
      table,
      tbody,
      selectedCells: new Set(),
      anchor: null,            // {rowIndex, colIndex}
      anchorCell: null,
      dragging: false,
      dragStart: null,
      columnsToSkip: options.columnsToSkip || ["col-select", "excel-row-head"],
      onCopy: options.onCopy || null,
    };
    table[ATTR_DATA_KEY] = state;

    // 편집 진입 = active-cell 단일 (기존 로직과 동일)
    // 셀 클릭 시 selection 갱신
    function handleMouseDown(e) {
      if (e.button !== 0) return; // 좌클릭만
      const cell = e.target.closest("td");
      if (!cell || !isSelectableCell(cell, state)) return;

      const pos = getCellPos(cell);
      if (!pos) return;

      // contentEditable 셀에서 텍스트 편집 중이면 selection 무시 (드래그는 텍스트 선택)
      // 단 첫 mousedown은 single-cell anchor만 잡고, 드래그가 시작되면 다중 selection으로 전환
      if (e.shiftKey && state.anchor) {
        e.preventDefault();
        clearSelectionVisualOnly(state);
        selectRange(state, state.anchor, pos);
        updateBadge(state);
        return;
      }

      if (e.ctrlKey || e.metaKey) {
        e.preventDefault();
        if (state.selectedCells.has(cell)) {
          setSelected(cell, state, false);
        } else {
          setSelected(cell, state, true);
          state.anchor = pos;
          state.anchorCell = cell;
        }
        updateBadge(state);
        return;
      }

      // 단순 클릭 — anchor를 현재 셀로 설정
      // 이 시점에서 즉시 다른 셀 selection 클리어 (편집 가능하게)
      clearSelection(state);
      state.anchor = pos;
      state.anchorCell = cell;
      state.dragging = true;
      state.dragStart = pos;
      // 첫 셀은 selected 표시하지 않음 (편집 진입을 방해하지 않기 위해)
      // 드래그가 일어나면 setSelected 시작
    }

    function clearSelectionVisualOnly(state) {
      state.tbody.querySelectorAll(".excel-cell-selected").forEach((c) => c.classList.remove("excel-cell-selected"));
      state.selectedCells.clear();
    }

    function handleMouseMove(e) {
      if (!state.dragging || !state.dragStart) return;
      const cell = e.target.closest("td");
      if (!cell || !isSelectableCell(cell, state)) return;
      const pos = getCellPos(cell);
      if (!pos) return;

      // 드래그가 다른 셀로 이동했으면 selection 시작
      if (pos.rowIndex !== state.dragStart.rowIndex || pos.colIndex !== state.dragStart.colIndex) {
        e.preventDefault();
        clearSelectionVisualOnly(state);
        selectRange(state, state.dragStart, pos);
        updateBadge(state);
        // 텍스트 selection이 잡히지 않도록 contenteditable 일시 무력화
        try { document.getSelection().removeAllRanges(); } catch (e2) { /* ignore */ }
      }
    }

    function handleMouseUp(e) {
      state.dragging = false;
      // 단일 셀 클릭이면 selection 비움 (편집 모드로 진입)
      if (state.selectedCells.size <= 1) {
        clearSelectionVisualOnly(state);
      }
    }

    // Ctrl+C: 선택 영역을 TSV로 클립보드에 복사
    function handleCopy(e) {
      // 1셀 이상 selected가 있을 때만 가로채기
      if (state.selectedCells.size <= 1) return; // contenteditable 기본 동작
      const tsv = buildTsvFromSelection(state);
      if (!tsv) return;
      e.preventDefault();
      try {
        e.clipboardData.setData("text/plain", tsv);
        e.clipboardData.setData("text/html", buildHtmlFromSelection(state));
      } catch (err) {
        navigator.clipboard?.writeText(tsv);
      }
      if (state.onCopy) state.onCopy(state.selectedCells.size, tsv);
    }

    function buildHtmlFromSelection(state) {
      const rect = rectFromSelection(state);
      if (!rect) return "";
      const { r1, r2, c1, c2 } = rect;
      let html = "<table>";
      for (let r = r1; r <= r2; r += 1) {
        html += "<tr>";
        for (let c = c1; c <= c2; c += 1) {
          const cell = getCellAt(state.tbody, r, c);
          const text = cell && state.selectedCells.has(cell) ? (cell.innerText || "") : "";
          html += `<td>${text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")}</td>`;
        }
        html += "</tr>";
      }
      html += "</table>";
      return html;
    }

    function handleKeyDown(e) {
      if ((e.ctrlKey || e.metaKey) && (e.key === "a" || e.key === "A")) {
        // 편집 중이면 무시 (텍스트 select all)
        const active = document.activeElement;
        if (active && active.contentEditable === "true" && tbody.contains(active)) {
          // 셀 안에서 Ctrl+A = 셀 텍스트 select all (기본). 한 번 더 눌러야 전체 선택.
          // 단순화를 위해 그냥 기본 동작 둠.
          return;
        }
        e.preventDefault();
        clearSelectionVisualOnly(state);
        // 모든 selectable td에 selected 추가
        tbody.querySelectorAll("td").forEach((cell) => {
          if (isSelectableCell(cell, state)) setSelected(cell, state, true);
        });
        updateBadge(state);
      } else if (e.key === "Escape") {
        if (state.selectedCells.size > 0) {
          e.preventDefault();
          clearSelection(state);
        }
      }
    }

    table.addEventListener("mousedown", handleMouseDown);
    table.addEventListener("mousemove", handleMouseMove);
    document.addEventListener("mouseup", handleMouseUp);
    document.addEventListener("copy", handleCopy);
    document.addEventListener("keydown", handleKeyDown);

    state._cleanup = () => {
      table.removeEventListener("mousedown", handleMouseDown);
      table.removeEventListener("mousemove", handleMouseMove);
      document.removeEventListener("mouseup", handleMouseUp);
      document.removeEventListener("copy", handleCopy);
      document.removeEventListener("keydown", handleKeyDown);
    };

    return state;
  }

  function detach(table) {
    if (!table) return;
    const state = table[ATTR_DATA_KEY];
    if (state && state._cleanup) state._cleanup();
    delete table[ATTR_DATA_KEY];
  }

  function clear(table) {
    const state = table && table[ATTR_DATA_KEY];
    if (state) clearSelection(state);
  }

  return { attach, detach, clear };
})();

window.ExcelSelect = ExcelSelect;
