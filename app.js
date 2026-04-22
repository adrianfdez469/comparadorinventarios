"use strict";

(function () {
  const input = document.getElementById("xlsxInput");
  const btnCompare = document.getElementById("btnCompare");
  const btnDownloadHtml = document.getElementById("btnDownloadHtml");
  const statusEl = document.getElementById("status");
  const resultWrap = document.getElementById("resultWrap");
  const resultBody = document.getElementById("resultBody");
  const summaryEl = document.getElementById("summary");
  const excelHelp = document.getElementById("excelHelp");

  const overlay = document.getElementById("reason-overlay");
  const reasonText = document.getElementById("reason-text");
  const reasonClose = document.getElementById("reason-close");

  let lastRows = [];
  let lastFileName = "data.xlsx";

  function setStatus(msg, isError) {
    statusEl.textContent = msg;
    statusEl.style.color = isError ? "#b42318" : "#444";
  }

  function toText(v) {
    if (v === null || v === undefined) return "";
    return String(v).trim();
  }

  function fmtQty(v) {
    if (v === null || v === undefined) return "—";
    if (Math.abs(v - Math.round(v)) < 1e-9) return String(Math.round(v));
    return String(v);
  }

  function isZero(v) {
    return Math.abs(v) <= 1e-6;
  }

  function qtyEqual(a, b) {
    return Math.abs(a - b) <= 1e-6;
  }

  function parseQuantityCell(value, options) {
    if (value === null || value === undefined || value === "") return null;
    if (typeof value === "number" && Number.isFinite(value)) return value;

    const text = String(value).replace(/\s+/g, "");
    const m = text.match(/[-+]?\d[\d,.]*/);
    if (!m) return null;
    const token = m[0];

    try {
      if (options && options.sheet2NumberFormat) {
        // Hoja 2: punto miles, coma decimal => 11.835,00 -> 11835.00
        return Number(token.replace(/\./g, "").replace(",", "."));
      }

      // Modo general (hoja 1)
      if (token.includes(",") && token.includes(".")) {
        if (token.lastIndexOf(".") > token.lastIndexOf(",")) {
          return Number(token.replace(/,/g, ""));
        }
        return Number(token.replace(/\./g, "").replace(",", "."));
      }
      if (token.includes(",")) {
        if (/^[-+]?\d{1,3}(,\d{3})+$/.test(token)) {
          return Number(token.replace(/,/g, ""));
        }
        return Number(token.replace(",", "."));
      }
      return Number(token);
    } catch (_err) {
      return null;
    }
  }

  function buildInventoryMap(sheetRows, options) {
    const map = new Map();
    for (let i = 1; i < sheetRows.length; i += 1) {
      const row = sheetRows[i] || [];
      const name = toText(row[0]);
      if (!name) continue;
      const qty = parseQuantityCell(row[1], options);
      map.set(name, qty);
    }
    return map;
  }

  function buildCorrelation(sheetRows) {
    const pairs = [];
    for (let i = 1; i < sheetRows.length; i += 1) {
      const row = sheetRows[i] || [];
      const a = toText(row[0]);
      const b = toText(row[1]);
      if (!a && !b) continue;
      if (!a) continue;
      pairs.push([a, b]);
    }
    return pairs;
  }

  function explainFail(row) {
    if (row.in1 && row.q1 === null) {
      return `En la hoja 1, la cantidad de «${row.name1}» no contiene un número reconocible.`;
    }
    if (row.in2 && row.q2 === null) {
      return `En la hoja 2, la cantidad de «${row.name2}» no contiene un número reconocible.`;
    }
    const diff = Math.abs(row.cmp1 - row.cmp2);
    if (!row.in1) {
      const n = row.name1 || "(sin nombre en hoja 1)";
      return `«${n}» no existe en hoja 1; se toma 0. Hoja 2 = ${fmtQty(row.cmp2)}. Diferencia = ${fmtQty(diff)}.`;
    }
    if (!row.in2) {
      const n = row.name2 || "(sin nombre en hoja 2)";
      return `«${n}» no existe en hoja 2; se toma 0. Hoja 1 = ${fmtQty(row.cmp1)}. Diferencia = ${fmtQty(diff)}.`;
    }
    return `Las cantidades no coinciden: hoja 1 = ${fmtQty(row.cmp1)}, hoja 2 = ${fmtQty(row.cmp2)}. Diferencia = ${fmtQty(diff)}.`;
  }

  function computeRows(inv1, inv2, correlation) {
    const rows = [];
    const mapped1 = new Set();
    const mapped2 = new Set();

    function appendRow(name1, name2) {
      const in1 = Boolean(name1) && inv1.has(name1);
      const in2 = Boolean(name2) && inv2.has(name2);
      const q1 = in1 ? inv1.get(name1) : null;
      const q2 = in2 ? inv2.get(name2) : null;

      let ok = false;
      let cmp1 = 0;
      let cmp2 = 0;
      let failReason = "";

      if (in1 && q1 === null) {
        ok = false;
      } else if (in2 && q2 === null) {
        ok = false;
      } else {
        cmp1 = in1 ? q1 : 0;
        cmp2 = in2 ? q2 : 0;
        ok = qtyEqual(cmp1, cmp2);
      }

      const row = {
        name1: name1 || "",
        name2: name2 || "",
        in1,
        in2,
        q1,
        q2,
        cmp1,
        cmp2,
        ok,
        failReason: "",
      };

      if (!ok) row.failReason = explainFail(row);
      rows.push(row);
    }

    correlation.forEach(([name1, name2]) => {
      mapped1.add(name1);
      if (name2) mapped2.add(name2);
      appendRow(name1, name2);
    });

    for (const name1 of inv1.keys()) {
      if (!mapped1.has(name1)) appendRow(name1, "");
    }
    for (const name2 of inv2.keys()) {
      if (!mapped2.has(name2)) appendRow("", name2);
    }

    // Reglas pedidas: si falta en una hoja y cantidad existente = 0 => OK
    rows.forEach((r) => {
      if (!r.in1 && r.in2 && r.q2 !== null && isZero(r.q2)) {
        r.ok = true;
        r.failReason = "";
      } else if (!r.in2 && r.in1 && r.q1 !== null && isZero(r.q1)) {
        r.ok = true;
        r.failReason = "";
      }
    });

    return rows;
  }

  function openReason(text) {
    reasonText.textContent = text || "";
    overlay.classList.add("is-open");
    overlay.setAttribute("aria-hidden", "false");
    reasonClose.focus();
  }

  function closeReason() {
    overlay.classList.remove("is-open");
    overlay.setAttribute("aria-hidden", "true");
    reasonText.textContent = "";
  }

  function render(rows) {
    resultBody.innerHTML = "";
    let okCount = 0;
    let failCount = 0;

    rows.forEach((r) => {
      const tr = document.createElement("tr");
      const td1 = document.createElement("td");
      const td2 = document.createElement("td");
      const td3 = document.createElement("td");
      const td4 = document.createElement("td");
      const td5 = document.createElement("td");

      td1.textContent = r.name1;
      td2.textContent = fmtQty(r.in1 ? r.q1 : 0);
      td2.className = "num";
      td3.className = "status";
      td4.textContent = fmtQty(r.in2 ? r.q2 : 0);
      td4.className = "num";
      td5.textContent = r.name2;

      if (r.ok) {
        const ok = document.createElement("span");
        ok.className = "ok";
        ok.textContent = "✓ OK";
        td3.appendChild(ok);
        okCount += 1;
      } else {
        const btn = document.createElement("button");
        btn.type = "button";
        btn.className = "fail-reason-btn";
        btn.textContent = "✗ Fail";
        btn.addEventListener("click", () => openReason(r.failReason));
        td3.appendChild(btn);
        failCount += 1;
      }

      tr.appendChild(td1);
      tr.appendChild(td2);
      tr.appendChild(td3);
      tr.appendChild(td4);
      tr.appendChild(td5);
      resultBody.appendChild(tr);
    });

    summaryEl.textContent = `Total: ${rows.length} | OK: ${okCount} | FAIL: ${failCount}`;
    resultWrap.style.display = "block";
    btnDownloadHtml.disabled = rows.length === 0;
  }

  function makeDownloadHtml(rows) {
    const bodyRows = rows.map((r) => {
      const status = r.ok
        ? '<span class="ok">✓ OK</span>'
        : `<span class="fail">✗ Fail</span><br><small>${escapeHtml(r.failReason)}</small>`;
      return (
        "<tr>" +
        `<td>${escapeHtml(r.name1)}</td>` +
        `<td class="num">${escapeHtml(fmtQty(r.in1 ? r.q1 : 0))}</td>` +
        `<td class="status">${status}</td>` +
        `<td class="num">${escapeHtml(fmtQty(r.in2 ? r.q2 : 0))}</td>` +
        `<td>${escapeHtml(r.name2)}</td>` +
        "</tr>"
      );
    });

    return `<!doctype html>
<html lang="es"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>comparacion_inventario</title>
<style>
body{font-family:system-ui,-apple-system,Segoe UI,Roboto,sans-serif;padding:1rem}
table{width:100%;border-collapse:collapse}th,td{padding:.55rem;border:1px solid #ddd}th{background:#2c3e50;color:#fff}
.ok{color:#0d7d4d;font-weight:600}.fail{color:#c0392b;font-weight:600}.num{text-align:right}.status{text-align:center}
</style></head><body>
<h1>Comparación de inventarios</h1>
<p>Archivo: ${escapeHtml(lastFileName)}</p>
<table><thead><tr><th>Producto (hoja 1)</th><th>Cantidad hoja 1</th><th>Resultado</th><th>Cantidad hoja 2</th><th>Producto (hoja 2)</th></tr></thead>
<tbody>${bodyRows.join("")}</tbody></table></body></html>`;
  }

  function escapeHtml(s) {
    return String(s)
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#39;");
  }

  async function runCompare() {
    if (!window.XLSX) {
      setStatus("No se pudo cargar la librería XLSX. Revisa conexión a internet o CDN.", true);
      return;
    }
    if (!input.files || !input.files[0]) {
      setStatus("Selecciona primero un archivo Excel.", true);
      return;
    }
    const file = input.files[0];
    lastFileName = file.name;
    setStatus("Leyendo Excel y comparando...", false);

    try {
      const ab = await file.arrayBuffer();
      const wb = XLSX.read(ab, { type: "array" });
      if (!wb.SheetNames || wb.SheetNames.length < 3) {
        throw new Error("El archivo debe tener al menos 3 hojas.");
      }

      const s1 = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: true });
      const s2 = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[1]], { header: 1, raw: true });
      const s3 = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[2]], { header: 1, raw: true });

      const inv1 = buildInventoryMap(s1, { sheet2NumberFormat: false });
      const inv2 = buildInventoryMap(s2, { sheet2NumberFormat: true });
      const corr = buildCorrelation(s3);
      if (corr.length === 0) {
        throw new Error("La hoja 3 no tiene filas de correlación.");
      }

      const rows = computeRows(inv1, inv2, corr);
      lastRows = rows;
      render(rows);
      excelHelp.style.display = "none";
      setStatus("Comparación completada.", false);
    } catch (err) {
      resultWrap.style.display = "none";
      btnDownloadHtml.disabled = true;
      excelHelp.style.display = "block";
      setStatus(`Error: ${err && err.message ? err.message : "no se pudo procesar el archivo."}`, true);
    }
  }

  btnCompare.addEventListener("click", runCompare);
  btnDownloadHtml.addEventListener("click", () => {
    if (!lastRows.length) return;
    const blob = new Blob([makeDownloadHtml(lastRows)], { type: "text/html;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "comparacion_inventario.html";
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  });

  reasonClose.addEventListener("click", closeReason);
  overlay.addEventListener("click", (e) => {
    if (e.target === overlay) closeReason();
  });
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape" && overlay.classList.contains("is-open")) closeReason();
  });
})();
