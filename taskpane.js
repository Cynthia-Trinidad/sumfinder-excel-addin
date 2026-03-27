/* ============================================================
   SumFinder — taskpane.js
   Lógica principal del add-in para Excel
   ============================================================ */

"use strict";

// ── Estado global ────────────────────────────────────────────
let capturedRange = null;   // Dirección del rango (ej: "A1:A50")
let capturedValues = [];    // Valores numéricos del rango
let capturedAddresses = []; // Dirección de cada celda individual
let currentTab = "manual";  // "manual" | "cell"
let lastResults = [];       // Últimos resultados encontrados

// ── Inicialización de Office.js ──────────────────────────────
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    console.log("SumFinder listo.");
  }
});

// ── Cambiar pestaña de entrada de monto ─────────────────────
function switchTab(tab) {
  currentTab = tab;
  document.getElementById("tabManual").classList.toggle("active", tab === "manual");
  document.getElementById("tabCell").classList.toggle("active", tab === "cell");
  document.getElementById("panelManual").style.display = tab === "manual" ? "flex" : "none";
  document.getElementById("panelCell").style.display  = tab === "cell"   ? "flex" : "none";
}

// ── Capturar rango seleccionado en Excel ─────────────────────
async function captureRange() {
  try {
    await Excel.run(async (ctx) => {
      const selection = ctx.workbook.getSelectedRange();
      selection.load(["address", "values", "cellCount"]);
      await ctx.sync();

      if (selection.cellCount > 1000) {
        showError("El rango es demasiado grande (máx. 1000 celdas).");
        return;
      }

      const flat = [];
      const addrs = [];

      // Obtener dirección de cada celda individual
      const rows = selection.values.length;
      const cols = selection.values[0].length;

      // Parsear dirección base para reconstruir celda a celda
      const baseAddr = selection.address; // ej: "Sheet1!A1:C5"
      const sheetName = baseAddr.includes("!") ? baseAddr.split("!")[0] : "";
      const rangeOnly = baseAddr.includes("!") ? baseAddr.split("!")[1] : baseAddr;

      const match = rangeOnly.match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/);
      if (!match) {
        showError("No se pudo interpretar el rango seleccionado.");
        return;
      }

      const startCol = colLetterToIndex(match[1]);
      const startRow = parseInt(match[2]);

      for (let r = 0; r < rows; r++) {
        for (let c = 0; c < cols; c++) {
          const val = selection.values[r][c];
          if (typeof val === "number" && !isNaN(val)) {
            flat.push(val);
            addrs.push(indexToColLetter(startCol + c) + (startRow + r));
          }
        }
      }

      if (flat.length === 0) {
        showError("El rango seleccionado no contiene valores numéricos.");
        return;
      }

      capturedRange = rangeOnly;
      capturedValues = flat;
      capturedAddresses = addrs;

      const display = document.getElementById("rangeDisplay");
      display.textContent = `${rangeOnly}  (${flat.length} valores)`;
      display.classList.remove("empty");
      hideError();
    });
  } catch (e) {
    showError("Error al capturar rango: " + e.message);
  }
}

// ── Leer valor desde una celda de referencia ─────────────────
async function readCellValue() {
  const ref = document.getElementById("cellRefInput").value.trim().toUpperCase();
  if (!ref) { showError("Ingresa una referencia de celda (ej: B15)."); return; }

  try {
    await Excel.run(async (ctx) => {
      const cell = ctx.workbook.activeWorksheet.getRange(ref);
      cell.load("values");
      await ctx.sync();
      const val = cell.values[0][0];
      if (typeof val !== "number" || isNaN(val)) {
        showError(`La celda ${ref} no contiene un número.`);
        return;
      }
      document.getElementById("amountInput").value = val;
      switchTab("manual");
      hideError();
    });
  } catch (e) {
    showError("No se pudo leer la celda: " + e.message);
  }
}

// ── Ejecutar búsqueda principal ──────────────────────────────
async function runSearch() {
  hideError();
  clearResults();

  if (!capturedValues.length) {
    showError("Primero captura un rango de celdas.");
    return;
  }

  const rawAmount = parseFloat(document.getElementById("amountInput").value);
  if (isNaN(rawAmount)) {
    showError("Ingresa un monto objetivo válido.");
    return;
  }

  showLoading(true);
  document.getElementById("btnSearch").disabled = true;

  // Pequeño timeout para que el spinner aparezca antes del cómputo
  setTimeout(() => {
    try {
      const results = findThreeCombinations(capturedValues, capturedAddresses, rawAmount);
      lastResults = results;
      renderResults(results, rawAmount);
    } catch (e) {
      showError("Error durante la búsqueda: " + e.message);
    } finally {
      showLoading(false);
      document.getElementById("btnSearch").disabled = false;
    }
  }, 50);
}

// ── Algoritmo: encontrar 3 combinaciones ─────────────────────
// Usa una heurística greedy + perturbaciones para rangos grandes.
// Retorna: [{ sum, diff, indices, addresses }] ordenadas por cercanía
function findThreeCombinations(values, addresses, target) {
  const n = values.length;
  const precision = 1e-9;

  // Intentamos múltiples semillas para diversidad
  const candidates = [];

  // --- Estrategia 1: greedy ordenado descendente ---
  const idxDesc = [...Array(n).keys()].sort((a,b) => Math.abs(values[b]) - Math.abs(values[a]));
  candidates.push(greedy(values, target, idxDesc));

  // --- Estrategia 2: greedy ordenado ascendente ---
  const idxAsc = [...idxDesc].reverse();
  candidates.push(greedy(values, target, idxAsc));

  // --- Estrategia 3..12: greedy con shuffle aleatorio semilla ---
  for (let s = 0; s < 40; s++) {
    const shuffled = shuffleSeeded([...Array(n).keys()], s * 7 + 13);
    candidates.push(greedy(values, target, shuffled));
  }

  // --- Estrategia 13+: perturbaciones sobre los mejores ---
  const sorted = [...candidates].sort((a,b) => Math.abs(a.diff) - Math.abs(b.diff));
  for (let i = 0; i < Math.min(3, sorted.length); i++) {
    const perturbed = perturbSolution(values, target, sorted[i].selected, 20);
    candidates.push(...perturbed);
  }

  // Deduplicar y separar en tres buckets: exacto/cercano, menor, mayor
  const unique = dedup(candidates, precision);

  const exact   = unique.filter(c => Math.abs(c.diff) < precision).sort((a,b) => a.selected.length - b.selected.length);
  const less    = unique.filter(c => c.diff < -precision).sort((a,b) => Math.abs(a.diff) - Math.abs(b.diff));
  const more    = unique.filter(c => c.diff >  precision).sort((a,b) => Math.abs(a.diff) - Math.abs(b.diff));

  const pick = (arr) => arr.length > 0 ? arr[0] : null;

  // Si hay exacto, lo ponemos en slot "exacto"; si no, el más cercano
  const allSorted = unique.sort((a,b) => Math.abs(a.diff) - Math.abs(b.diff));

  let slot1 = exact.length > 0 ? exact[0] : (allSorted[0] || null);
  let slot2 = less[0]  || null;
  let slot3 = more[0]  || null;

  // Evitar repetidos entre slots
  const usedKeys = new Set();
  const addIfNew = (c) => {
    if (!c) return null;
    const key = c.selected.slice().sort().join(",");
    if (usedKeys.has(key)) return null;
    usedKeys.add(key);
    return c;
  };

  slot1 = addIfNew(slot1);
  slot2 = addIfNew(slot2) || (less[1] ? addIfNew(less[1]) : null);
  slot3 = addIfNew(slot3) || (more[1] ? addIfNew(more[1]) : null);

  // Convertir índices → objetos con addresses
  const toResult = (c, label) => {
    if (!c) return null;
    return {
      label,
      sum: c.sum,
      diff: c.diff,
      cells: c.selected.map(i => addresses[i]),
      indices: c.selected,
    };
  };

  return [
    toResult(slot1, "exact"),
    toResult(slot2, "less"),
    toResult(slot3, "more"),
  ].filter(Boolean);
}

// Greedy: va sumando valores en el orden dado hasta alcanzar target
function greedy(values, target, order) {
  let sum = 0;
  const selected = [];
  for (const i of order) {
    const v = values[i];
    // Incluir si acerca al objetivo
    if (Math.abs(sum + v - target) <= Math.abs(sum - target)) {
      sum += v;
      selected.push(i);
    }
  }
  return { sum, diff: sum - target, selected };
}

// Perturbaciones: swap/add/remove elementos de una solución
function perturbSolution(values, target, selected, iterations) {
  const n = values.length;
  const results = [];
  const selSet = new Set(selected);

  for (let it = 0; it < iterations; it++) {
    const newSel = new Set(selSet);
    const mode = it % 3;

    if (mode === 0 && newSel.size > 0) {
      // Remover uno aleatorio
      const arr = [...newSel];
      newSel.delete(arr[it % arr.length]);
    } else if (mode === 1) {
      // Agregar uno no seleccionado
      for (let i = it % n; ; i = (i+1) % n) {
        if (!newSel.has(i)) { newSel.add(i); break; }
      }
    } else if (mode === 2 && newSel.size > 0) {
      // Swap
      const arr = [...newSel];
      newSel.delete(arr[it % arr.length]);
      for (let i = (it * 3) % n; ; i = (i+1) % n) {
        if (!newSel.has(i)) { newSel.add(i); break; }
      }
    }

    const newArr = [...newSel];
    const sum = newArr.reduce((s, i) => s + values[i], 0);
    results.push({ sum, diff: sum - target, selected: newArr });
  }
  return results;
}

function dedup(candidates, eps) {
  const seen = new Map();
  for (const c of candidates) {
    const key = c.selected.slice().sort((a,b)=>a-b).join(",");
    if (!seen.has(key)) seen.set(key, c);
  }
  return [...seen.values()];
}

function shuffleSeeded(arr, seed) {
  let s = seed;
  for (let i = arr.length - 1; i > 0; i--) {
    s = (s * 1664525 + 1013904223) & 0xffffffff;
    const j = Math.abs(s) % (i + 1);
    [arr[i], arr[j]] = [arr[j], arr[i]];
  }
  return arr;
}

// ── Renderizar resultados ────────────────────────────────────
function renderResults(results, target) {
  const container = document.getElementById("results");
  container.innerHTML = "";

  const headerDiv = document.createElement("div");
  headerDiv.className = "results-header";
  headerDiv.innerHTML = `
    <div class="section-label" style="margin:0">Resultados</div>
    <button class="btn-clear" onclick="clearResults()">✕ Limpiar</button>
  `;
  container.appendChild(headerDiv);

  const labels = {
    exact: { badge: "badge-exact", text: "Más cercano", icon: "🎯" },
    less:  { badge: "badge-less",  text: "Un poco menos", icon: "▼" },
    more:  { badge: "badge-more",  text: "Un poco más",  icon: "▲" },
  };

  for (const r of results) {
    const meta = labels[r.label] || labels.exact;
    const diffStr = Math.abs(r.diff) < 1e-9
      ? "exacto"
      : (r.diff > 0 ? "+" : "") + formatNum(r.diff);

    const card = document.createElement("div");
    card.className = "result-card";
    card.dataset.label = r.label;

    card.innerHTML = `
      <div class="card-header">
        <span class="card-badge ${meta.badge}">${meta.icon} ${meta.text}</span>
        <span class="card-sum">
          ${formatNum(r.sum)}
          <span class="diff">(${diffStr})</span>
        </span>
      </div>
      <div class="card-body">
        <div class="cell-count">${r.cells.length} celda${r.cells.length !== 1 ? "s" : ""}</div>
        <div class="cells-list">${r.cells.join("  ·  ")}</div>
        <div class="action-row">
          <button class="btn-action" onclick="doAction('select', '${r.label}')">☑ Seleccionar</button>
          <button class="btn-action" onclick="doAction('copy',   '${r.label}')">⎘ Copiar</button>
          <button class="btn-action cut" onclick="doAction('cut', '${r.label}')">✂ Cortar</button>
        </div>
      </div>
    `;
    container.appendChild(card);
  }

  if (results.length === 0) {
    container.innerHTML += `<div style="color:var(--muted);font-size:13px;padding:8px;">No se encontraron combinaciones.</div>`;
  }

  container.classList.add("visible");
}

// ── Acciones sobre resultados ────────────────────────────────
async function doAction(action, label) {
  const result = lastResults.find(r => r.label === label);
  if (!result || result.cells.length === 0) return;

  try {
    await Excel.run(async (ctx) => {
      const ws = ctx.workbook.activeWorksheet;

      // Construir rango multi-área (unión de celdas)
      // Excel acepta rangos como "A1,A3,B5" para selección múltiple
      const rangeAddress = result.cells.join(",");
      const range = ws.getRanges(rangeAddress);

      if (action === "select") {
        range.select();
        await ctx.sync();

      } else if (action === "copy") {
        range.select();
        await ctx.sync();
        // Ejecutar Copy via context
        ctx.workbook.application.calculate(Excel.CalculationType.full);
        range.copyFrom(range, Excel.RangeCopyType.all, false, false);
        range.select();
        await ctx.sync();
        // Nota: Office.js no expone Copy directamente en todas las versiones.
        // La selección permite que el usuario copie con Ctrl+C.
        // Se muestra un toast informativo.
        showToast("Celdas seleccionadas — presiona Ctrl+C para copiar");

      } else if (action === "cut") {
        range.select();
        await ctx.sync();
        showToast("Celdas seleccionadas — presiona Ctrl+X para cortar");
      }
    });
  } catch (e) {
    showError("Error al ejecutar acción: " + e.message);
  }
}

// ── Toast de feedback ────────────────────────────────────────
function showToast(msg) {
  let toast = document.getElementById("toast");
  if (!toast) {
    toast = document.createElement("div");
    toast.id = "toast";
    toast.style.cssText = `
      position:fixed; bottom:60px; left:50%; transform:translateX(-50%);
      background:#21212c; border:1px solid #7c6cfc; border-radius:8px;
      padding:10px 16px; font-size:12px; color:#e8e8f0;
      font-family:'Syne',sans-serif; z-index:9999;
      box-shadow:0 4px 20px rgba(0,0,0,0.4);
      transition:opacity 0.3s;
    `;
    document.body.appendChild(toast);
  }
  toast.textContent = msg;
  toast.style.opacity = "1";
  clearTimeout(toast._timer);
  toast._timer = setTimeout(() => { toast.style.opacity = "0"; }, 3000);
}

// ── Helpers UI ───────────────────────────────────────────────
function showLoading(v) {
  document.getElementById("loading").classList.toggle("visible", v);
}

function showError(msg) {
  const box = document.getElementById("errorBox");
  box.textContent = "⚠ " + msg;
  box.classList.add("visible");
}

function hideError() {
  document.getElementById("errorBox").classList.remove("visible");
}

function clearResults() {
  const r = document.getElementById("results");
  r.innerHTML = "";
  r.classList.remove("visible");
  lastResults = [];
}

function formatNum(n) {
  return new Intl.NumberFormat("es-CL", {
    minimumFractionDigits: 0,
    maximumFractionDigits: 2,
  }).format(n);
}

// ── Helpers columnas Excel (A→1, Z→26, AA→27…) ──────────────
function colLetterToIndex(letters) {
  let n = 0;
  for (const ch of letters) n = n * 26 + ch.charCodeAt(0) - 64;
  return n; // 1-based
}

function indexToColLetter(n) {
  let s = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

// ── Bind botones ──────────────────────────────────────────────
document.addEventListener("DOMContentLoaded", () => {
  document.getElementById("btnCapture").addEventListener("click", captureRange);
});
