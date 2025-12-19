import { SheetEvaluator } from "./formula.js";

const $ = (sel) => document.querySelector(sel);
const el = (tag, attrs = {}, children = []) => {
  const n = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) {
    if (k === "class") n.className = v;
    else if (k.startsWith("on") && typeof v === "function") n.addEventListener(k.slice(2), v);
    else if (v === null || v === undefined) continue;
    else n.setAttribute(k, String(v));
  }
  for (const c of children) n.append(c);
  return n;
};

function formatValue(v) {
  if (v === null || v === undefined) return "";
  if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
  if (typeof v === "number") {
    if (Number.isInteger(v)) return String(v);
    // keep compact
    return String(Number(v.toFixed(3)));
  }
  return String(v);
}

function toInputValue(v, type) {
  if (v === null || v === undefined) return "";
  if (type === "number") return String(v);
  return String(v);
}

function parseUserInput(raw, type) {
  if (type === "number") {
    const s = (raw ?? "").trim();
    if (s === "") return "";
    const n = Number(s);
    return Number.isNaN(n) ? "" : n;
  }
  return (raw ?? "").toString();
}

function columnLabel(col) {
  // keep compact; Excel-style A/B/C… is familiar
  return col;
}

async function loadModel() {
  const res = await fetch("./model.json", { cache: "no-cache" });
  if (!res.ok) throw new Error(`Failed to load model.json (${res.status})`);
  return await res.json();
}

function buildTabs(sheetNames, active, onSelect) {
  const tabs = $("#tabs");
  tabs.innerHTML = "";
  for (const name of sheetNames) {
    tabs.append(
      el(
        "button",
        {
          class: "tab",
          type: "button",
          "aria-selected": name === active ? "true" : "false",
          onclick: () => onSelect(name),
        },
        [document.createTextNode(name.trim())]
      )
    );
  }
}

function renderInputs(sheetModel, overrides, onChange) {
  const container = $("#inputs");
  container.innerHTML = "";

  for (const inp of sheetModel.inputs) {
    const { label, cell, type } = inp;
    const key = cell.replace(/\$/g, "").toUpperCase();
    const base = sheetModel.cells[key]?.v ?? "";
    const cur = Object.prototype.hasOwnProperty.call(overrides, key) ? overrides[key] : base;

    const input = el("input", {
      inputmode: type === "number" ? "decimal" : "text",
      type: "text",
      value: toInputValue(cur, type),
      "data-cell": key,
      oninput: (e) => {
        const val = parseUserInput(e.target.value, type);
        onChange(key, val);
      },
    });

    container.append(
      el("div", { class: "field" }, [
        el("label", {}, [document.createTextNode(`${label} (${key})`)]),
        input,
      ])
    );
  }
}

function renderGrid(sheetName, sheetModel, overrides) {
  const grid = $("#grid");
  const errBox = $("#error");
  errBox.style.display = "none";
  errBox.textContent = "";

  const { startRow, endRow, columns } = sheetModel.table;
  const evaluator = new SheetEvaluator(sheetModel, overrides);

  // Build header
  const thead = el("thead");
  const trh = el("tr");
  for (const col of columns) {
    trh.append(el("th", {}, [document.createTextNode(columnLabel(col))]));
  }
  thead.append(trh);

  const tbody = el("tbody");
  for (let r = startRow; r <= endRow; r++) {
    const tr = el("tr");
    let rowHasAny = false;
    for (const col of columns) {
      const ref = `${col}${r}`;
      let v = "";
      try {
        v = evaluator.evalCell(ref);
      } catch (e) {
        v = "";
      }
      const text = formatValue(v);
      if (text !== "") rowHasAny = true;
      tr.append(el("td", {}, [document.createTextNode(text)]));
    }
    // Only show rows that have a part name in col A (or any visible content)
    if (rowHasAny) tbody.append(tr);
  }

  grid.innerHTML = "";
  grid.append(thead, tbody);
}

function setPwaStatus() {
  const txt = $("#pwaText");
  const parts = [];
  parts.push(navigator.onLine ? "online" : "offline");
  if ("serviceWorker" in navigator) parts.push("sw");
  txt.textContent = parts.join(" • ");
}

async function registerServiceWorker() {
  if (!("serviceWorker" in navigator)) return;
  try {
    await navigator.serviceWorker.register("./sw.js", { scope: "./" });
  } catch {
    // ignore (still works without SW)
  }
}

function loadOverrides(sheetName) {
  try {
    const raw = localStorage.getItem(`overrides:${sheetName}`);
    return raw ? JSON.parse(raw) : {};
  } catch {
    return {};
  }
}

function saveOverrides(sheetName, overrides) {
  try {
    localStorage.setItem(`overrides:${sheetName}`, JSON.stringify(overrides));
  } catch {
    // ignore
  }
}

function showError(msg) {
  const errBox = $("#error");
  errBox.textContent = msg;
  errBox.style.display = "block";
}

async function main() {
  setPwaStatus();
  window.addEventListener("online", setPwaStatus);
  window.addEventListener("offline", setPwaStatus);

  await registerServiceWorker();
  setPwaStatus();

  let model;
  try {
    model = await loadModel();
  } catch (e) {
    showError(String(e?.message || e));
    return;
  }

  const sheetNames = Object.keys(model.sheets);
  let active = sheetNames[0] || "";

  function render() {
    const sheetModel = model.sheets[active];
    if (!sheetModel) return;
    $("#sheetTitle").textContent = active.trim();

    const overrides = loadOverrides(active);
    renderInputs(sheetModel, overrides, (cell, value) => {
      const next = { ...loadOverrides(active), [cell]: value };
      // allow clearing back to template value
      if (value === "" || value === null || value === undefined) delete next[cell];
      saveOverrides(active, next);
      renderGrid(active, sheetModel, next);
    });
    renderGrid(active, sheetModel, overrides);
  }

  function select(name) {
    active = name;
    buildTabs(sheetNames, active, select);
    render();
  }
  buildTabs(sheetNames, active, select);
  render();
}

main();


