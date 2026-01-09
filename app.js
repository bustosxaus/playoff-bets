const SCRIPT_URL = "https://script.google.com/macros/s/AKfycbwG1aDwRpbHO-4zMnPuf1G_FnFC3oF9TcNPsQV0xkiyIZSy9tkOijY_2HCyabdB3EzAUQ/exec";

const statusEl = document.getElementById("status");
const tableWrap = document.getElementById("table-wrap");
const emptyEl = document.getElementById("empty");
const saveBtn = document.getElementById("save");
const reloadBtn = document.getElementById("reload");

const state = {
  columns: [],
  rows: [],
  dirty: false,
};

const lockedColumns = new Set(["Cumulative Balance", "Status Message"]);

function setStatus(message, tone = "muted") {
  statusEl.textContent = message;
  statusEl.dataset.tone = tone;
}

function loadJsonp(url, timeoutMs = 12000) {
  return new Promise((resolve, reject) => {
    const callbackName = `__sheetCallback_${Date.now()}_${Math.floor(Math.random() * 1000)}`;
    const script = document.createElement("script");
    let finished = false;

    function cleanup() {
      finished = true;
      delete window[callbackName];
      script.remove();
    }

    window[callbackName] = (payload) => {
      cleanup();
      resolve(payload);
    };

    script.onerror = () => {
      if (finished) {
        return;
      }
      cleanup();
      reject(new Error("Load failed"));
    };

    const joiner = url.includes("?") ? "&" : "?";
    script.src = `${url}${joiner}action=get&callback=${callbackName}`;
    document.body.appendChild(script);

    window.setTimeout(() => {
      if (finished) {
        return;
      }
      cleanup();
      reject(new Error("Load timed out"));
    }, timeoutMs);
  });
}

function normalizeValue(value) {
  if (value === null || value === undefined) {
    return "";
  }
  return String(value);
}

function coerceValue(value) {
  const trimmed = value.trim();
  if (!trimmed) {
    return "";
  }
  if (/^-?\d+(\.\d+)?$/.test(trimmed)) {
    return Number(trimmed);
  }
  return trimmed;
}

function markDirty(cell) {
  state.dirty = true;
  cell.dataset.dirty = "true";
  setStatus("Unsaved changes", "warn");
}

function buildTable() {
  tableWrap.innerHTML = "";

  if (!state.columns.length) {
    emptyEl.textContent = "No columns found in the sheet.";
    tableWrap.appendChild(emptyEl);
    return;
  }

  const table = document.createElement("table");
  table.className = "table";

  const thead = document.createElement("thead");
  const headRow = document.createElement("tr");
  state.columns.forEach((col) => {
    const th = document.createElement("th");
    th.textContent = col;
    headRow.appendChild(th);
  });
  thead.appendChild(headRow);
  table.appendChild(thead);

  const tbody = document.createElement("tbody");

  state.rows.forEach((row, rowIndex) => {
    const tr = document.createElement("tr");
    state.columns.forEach((_, colIndex) => {
      const columnName = state.columns[colIndex];
      const isLocked = lockedColumns.has(columnName);
      const td = document.createElement("td");
      const cell = document.createElement("div");
      cell.className = isLocked ? "cell locked" : "cell";
      cell.contentEditable = isLocked ? "false" : "true";
      cell.dataset.row = String(rowIndex);
      cell.dataset.col = String(colIndex);
      cell.textContent = normalizeValue(row[colIndex]);

      if (isLocked) {
        cell.title = "Formula-driven column";
      } else {
        cell.addEventListener("input", (event) => {
          const target = event.currentTarget;
          const r = Number(target.dataset.row);
          const c = Number(target.dataset.col);
          state.rows[r][c] = target.textContent;
          markDirty(target);
        });
      }

      td.appendChild(cell);
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });

  table.appendChild(tbody);
  tableWrap.appendChild(table);
}

async function loadData() {
  if (!SCRIPT_URL || SCRIPT_URL.includes("PASTE")) {
    emptyEl.textContent = "Paste your Apps Script web app URL into app.js to load data.";
    tableWrap.innerHTML = "";
    tableWrap.appendChild(emptyEl);
    setStatus("Missing Apps Script URL", "warn");
    return;
  }

  setStatus("Loading sheet...");

  try {
    const payload = await loadJsonp(SCRIPT_URL);

    if (payload.error) {
      throw new Error(payload.error);
    }

    state.columns = payload.columns || [];
    state.rows = (payload.rows || []).map((row) =>
      state.columns.map((_, i) => normalizeValue(row[i]))
    );
    state.dirty = false;
    setStatus("Loaded. Ready to edit.");
    buildTable();
  } catch (error) {
    setStatus("Failed to load data", "warn");
    emptyEl.textContent =
      `Could not load data: ${error.message}. ` +
      "Open the Apps Script URL in a new tab to authorize, then reload.";
    tableWrap.innerHTML = "";
    tableWrap.appendChild(emptyEl);
  }
}

async function saveData() {
  if (!state.columns.length) {
    return;
  }

  setStatus("Saving...");

  const cleanedRows = state.rows.map((row) => row.map((cell) => coerceValue(String(cell))));

  try {
    const response = await fetch(SCRIPT_URL, {
      method: "POST",
      mode: "no-cors",
      headers: {
        "Content-Type": "text/plain;charset=utf-8",
      },
      body: JSON.stringify({
        action: "update",
        columns: state.columns,
        rows: cleanedRows,
      }),
      redirect: "follow",
    });

    if (response.type !== "opaque") {
      if (!response.ok) {
        throw new Error(`Save failed: ${response.status}`);
      }
      const payload = await response.json();
      if (payload.error) {
        throw new Error(payload.error);
      }
    }

    state.dirty = false;
    document.querySelectorAll(".cell[data-dirty='true']").forEach((cell) => {
      cell.dataset.dirty = "false";
    });
    setStatus("Saved. Reload to confirm.");
  } catch (error) {
    setStatus("Save failed", "warn");
  }
}

saveBtn.addEventListener("click", saveData);
reloadBtn.addEventListener("click", loadData);

loadData();
