/*
 * Supplier price-file generator — shared logic (browser + node).
 *
 * Mirrors generate_price_files.py 1:1. Given the rows parsed from a source
 * price list it builds the three output workbooks:
 *   PRODUCT_WCEW_<stem>.xlsx, PRODUCT_LIGHTING_<stem>.xlsx,
 *   SUPPLIER_LIGHTING_<stem>.xlsx  (W/ codes, W001)
 *   SUPPLIER_WCEW_<stem>.xlsx      (ECI/Magnalux, optional 4th template)
 *
 * Requires SheetJS (XLSX) to be passed in, so this file works both in the
 * browser (window.XLSX) and under node (require('xlsx')).
 */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.PriceGen = factory();
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  // ---- constants (copied from the reference templates) ---------------------
  var TRADE_MULT = 1.25, RETAIL_MULT = 1.45, COST_UPLIFT = 1.10,
      PRICE_A_MULT = 1.35, VAT_MULT = 1.23, NEG = -100;

  var WCEW_HEADERS = ["Product Code", "Product Description", "Product Group Code",
    "Sales Analysis Code", "Sales VAT Code", "Purchase Analysis Code",
    "Purchase VAT Code", "Alternate Code", "Bin Location", "Unit",
    "Current Cost Price (ex VAT)", "Trade Markup %", "Retail Markup %",
    "CPOS Markup %", "Price A", "Price B", "Price C", "Price D", "Price E",
    "Price F", "Price G", "Price H", "Price I", "Price J", "Price K",
    "Price L", "Price M", "Price N"];
  var LIGHTING_HEADERS = ["Product Code", "Product Description", "Product Group Code",
    "Sales Analysis Code", "Sales VAT Code", "Purchase Analysis Code",
    "Purchase VAT Code", "Bin Location", "Unit", "Current Cost Price (ex VAT)",
    "Trade Price (ex VAT)", "Retail Price (ex VAT)", "CPOS Price (inc VAT)",
    "Price A", "Price B", "Price C", "Price D", "Price E", "Price F", "Price G",
    "Price H", "Price I", "Price J", "Price K", "Price L", "Price M", "Price N"];
  var SUPPLIER_HEADERS = ["Product Code", "Supplier", "Supplier Product Code",
    "Price", "Last Purchased Date"];

  var WCEW_WIDTHS = [31.9, 63, 19.1, 18.6, 14.7, 22.1, 18.3, 19.1, 11.7, 4.9,
    25.3, 15.4, 15.6, 15.1];
  var LIGHTING_WIDTHS = [34.7, 63, 19.1, 18.6, 14.7, 22.1, 18.3, 11.7, 23.6,
    25.3, 19.1, 19.3];
  var SUPPLIER_WIDTHS = [34.7, 14.7, 34.7, 10.3, 18.9];

  var PRICE_PRIORITY = ["our price", "trade price", "rrp", "cost", "price"];
  var CODE_ALIASES = ["our product code", "product code", "code", "sku"];
  var DESC_ALIASES = ["short description", "description", "desc"];
  var ORDER_ALIASES = ["order code", "supplier code", "alternate code"];

  // ---- helpers -------------------------------------------------------------
  function roundUpWhole(x) { return Math.ceil(x); }
  function roundUpHalf(x) { return Math.ceil(x * 2) / 2; }

  function norm(s) {
    return s == null ? "" : String(s).replace(/\s+/g, " ").trim().toLowerCase();
  }

  function teresaCct(code) {
    if (/Z1L(\b|-|$)/.test(code || "")) return " - 4K";
    if (/Z1R(\b|-|$)/.test(code || "")) return " - 3K";
    return "";
  }

  // find the header row within the first 15 rows of an array-of-arrays
  function findHeader(aoa) {
    var best = null;
    var scan = Math.min(aoa.length, 15);
    for (var r = 0; r < scan; r++) {
      var headers = {};
      var row = aoa[r] || [];
      for (var c = 0; c < row.length; c++) {
        var v = row[c];
        if (typeof v === "string" && v.trim()) headers[norm(v)] = c;
      }
      var keys = Object.keys(headers);
      var codeCol = matchCol(headers, CODE_ALIASES);
      var priceCol = findPriceCol(headers).col;
      // A real header row has the code and price in DIFFERENT columns and
      // several labelled columns — this rejects prose/notes rows that merely
      // happen to contain the words "code" and "price" in one cell.
      if (codeCol !== undefined && priceCol !== undefined &&
          codeCol !== priceCol && keys.length >= 3)
        return { row: r, headers: headers };
      if (!best || keys.length > Object.keys(best.headers).length)
        best = { row: r, headers: headers };
    }
    return best;
  }

  function matchCol(headers, aliases) {
    var h;
    for (var i = 0; i < aliases.length; i++)
      for (h in headers) if (h === aliases[i]) return headers[h];
    for (i = 0; i < aliases.length; i++)
      for (h in headers) if (h.indexOf(aliases[i]) >= 0) return headers[h];
    return undefined;
  }

  function findPriceCol(headers) {
    for (var i = 0; i < PRICE_PRIORITY.length; i++)
      for (var h in headers)
        if (h.indexOf(PRICE_PRIORITY[i]) >= 0) return { col: headers[h], name: h };
    return { col: undefined, name: undefined };
  }

  // parse an array-of-arrays into normalized product rows + meta
  function extractRows(aoa) {
    var hdr = findHeader(aoa);
    if (!hdr) throw new Error("No header row found.");
    var headers = hdr.headers;
    var codeC = matchCol(headers, CODE_ALIASES);
    var descC = matchCol(headers, DESC_ALIASES);
    var orderC = matchCol(headers, ORDER_ALIASES);
    var candC = matchCol(headers, ["all candidates", "candidates", "candidate", "arithmetic"]);
    var price = findPriceCol(headers);
    if (codeC === undefined || price.col === undefined)
      throw new Error("Could not identify code/price columns. Headers: " +
        Object.keys(headers).join(", "));

    var rows = [], skipped = [];
    for (var r = hdr.row + 1; r < aoa.length; r++) {
      var row = aoa[r] || [];
      var code = row[codeC];
      var p = row[price.col];
      if (code == null || String(code).trim() === "") continue;
      if (typeof p !== "number" || isNaN(p)) {
        skipped.push({ code: String(code), why: "non-numeric price: " + p });
        continue;
      }
      rows.push({
        code: String(code).trim(),
        desc: descC !== undefined ? (row[descC] == null ? "" : String(row[descC])) : "",
        order: orderC !== undefined ? (row[orderC] == null ? "" : String(row[orderC])) : "",
        candidates: candC !== undefined ? (row[candC] == null ? "" : String(row[candC])) : "",
        base: p
      });
    }
    return { rows: rows, priceName: price.name, skipped: skipped,
             hasCandidates: candC !== undefined };
  }

  // ---- 4th template: ECI / Magnalux supplier import ------------------------
  var SI_COLOUR = { BK: "BLACK", WH: "WHITE", GY: "GREY" };
  var SI_VARIANTS = ["SQUARE", "1L", "2L", "ROUND", "GX53", "E27", "G9",
    "GU10", "SEN", "OP"];

  function siToks(s) {
    var m = String(s).toUpperCase().replace(/\//g, " ").match(/[A-Z0-9]+/g);
    return m || [];
  }

  // Parse the "All candidates (with arithmetic)" free-text for the two prices
  // and the ECI description. e.g. "ECI TERESA 50 ROUND 10.50. Magnalux 7.68."
  function parseCandidates(txt) {
    var out = { eciPrice: null, magPrice: null, eciDesc: "" };
    if (!txt) return out;
    var t = String(txt);
    // ECI form A: "ECI <desc> <price> …"  (price is the last number in the
    // ECI segment, up to Magnalux/Stock/end). e.g. "ECI TERESA 50 ROUND 10.50."
    var m = t.match(/ECI\b([\s\S]*?)(?:Magnalux|Stock|$)/i);
    if (m) {
      var seg = m[1];
      var nums = seg.match(/\d+(?:\.\d+)?/g);
      if (nums) {
        out.eciPrice = parseFloat(nums[nums.length - 1]);
        out.eciDesc = seg.replace(/\d+(?:\.\d+)?[\s\S]*$/, "").trim();
      }
    }
    // ECI form B (fallback): "supplier <price> (ECI …)"  e.g.
    // "supplier 14.95 (ECI direct)". Deliberately NOT matched by "(Derived …",
    // "(Stock …" or "(… NO ECI)" — those aren't a direct ECI price.
    if (out.eciPrice == null) {
      var mb = t.match(/supplier\s+(\d+(?:\.\d+)?)\s*\(\s*ECI/i);
      if (mb) out.eciPrice = parseFloat(mb[1]);
    }
    var mm = t.match(/Magnalux[^\d]*(\d+(?:\.\d+)?)/i);
    if (mm) out.magPrice = parseFloat(mm[1]);
    return out;
  }

  // Parse a supplier price list (array-of-arrays) into [{desc,code,price}].
  function parsePriceList(aoa) {
    var CODE = ["product code", "code", "sku", "item"];
    var DESC = ["description", "desc"];
    var PRICE = ["list price", "nett", "net price", "price", "cost"];
    var best = null, scan = Math.min(aoa.length, 12);
    for (var r = 0; r < scan; r++) {
      var hdr = {}, rowc = aoa[r] || [];
      for (var c = 0; c < rowc.length; c++) {
        var v = rowc[c];
        if (typeof v === "string" && v.trim()) hdr[norm(v)] = c;
      }
      var keys = Object.keys(hdr);
      var hc = keys.some(function (h) { return CODE.some(function (a) { return h.indexOf(a) >= 0; }); });
      var hp = keys.some(function (h) { return PRICE.some(function (a) { return h.indexOf(a) >= 0; }); });
      if (hc && hp) { best = { row: r, headers: hdr }; break; }
      if (!best || keys.length > Object.keys(best.headers).length) best = { row: r, headers: hdr };
    }
    var H = best.headers;
    var cc = matchCol(H, CODE), dc = matchCol(H, DESC), pc = matchCol(H, PRICE);
    var out = [];
    for (var rr = best.row + 1; rr < aoa.length; rr++) {
      var row = aoa[rr] || [];
      var code = cc !== undefined ? row[cc] : null;
      var desc = dc !== undefined ? row[dc] : null;
      var price = pc !== undefined ? row[pc] : null;
      if (code == null || String(code).trim() === "") continue;
      out.push({
        code: String(code).trim(),
        desc: desc == null ? "" : String(desc).trim(),
        price: typeof price === "number" ? price : null
      });
    }
    return out;
  }

  // Find the supplier's product code by confident match; null if unsure.
  function matchSupplierCode(family, colour, variantToks, price, list) {
    var best = null, bestScore = -1, tol = 0.02;
    for (var i = 0; i < list.length; i++) {
      var rowToks = siToks(list[i].desc).concat(siToks(list[i].code));
      if (family && rowToks.indexOf(family) < 0) continue;
      if (colour && rowToks.indexOf(colour) < 0) continue;
      var vt = 0;
      for (var k = 0; k < variantToks.length; k++)
        if (rowToks.indexOf(variantToks[k]) >= 0) vt++;
      var priceClose = price != null && list[i].price != null &&
        Math.abs(list[i].price - price) <= tol;
      if (vt === 0 && !priceClose) continue;   // not confident enough
      var score = vt + (priceClose ? 3 : 0);
      if (score > bestScore) { bestScore = score; best = list[i]; }
    }
    return best ? best.code : null;
  }

  // Build the 4th template. suppliers pair a code (E18/M68) with its list and
  // which parsed price to use. Returns { data, meta }.
  function buildSupplierImport(XLSX, rows, opts) {
    var eci = opts.eciList || [], mag = opts.magList || [];
    var suppliers = [
      { code: opts.eciCode || "E18", list: eci, use: "eciPrice" },
      { code: opts.magCode || "M68", list: mag, use: "magPrice" }
    ];
    var data = [], matched = { E18: 0, M68: 0 }, serial = todaySerial();
    rows.forEach(function (p) {
      var cand = parseCandidates(p.candidates);
      var U = p.code.toUpperCase();
      var family = U.split("-")[0];
      var colour = null;
      Object.keys(SI_COLOUR).forEach(function (k) {
        if (U.slice(-(k.length + 1)) === "-" + k) colour = SI_COLOUR[k];
      });
      var qToks = siToks(p.code).concat(siToks(cand.eciDesc));
      var variantToks = SI_VARIANTS.filter(function (v) { return qToks.indexOf(v) >= 0; });
      suppliers.forEach(function (s) {
        var price = cand[s.use];
        var supCode = matchSupplierCode(family, colour, variantToks, price, s.list);
        if (supCode) matched[s.code] = (matched[s.code] || 0) + 1;
        data.push([
          str(p.code), txt(s.code),
          supCode ? str(supCode) : null,
          price != null ? num(price, "#,##0.00") : null,
          formula("TODAY()", "mm-dd-yy", serial)
        ]);
      });
    });
    var ws = makeSheet(XLSX, SUPPLIER_HEADERS, data, SUPPLIER_WIDTHS);
    return {
      data: bookBytes(XLSX, ws, "ProdSuppRecImport"),
      meta: { rows: data.length, matched: matched }
    };
  }

  // ---- cell helpers --------------------------------------------------------
  function txt(v) { return { t: "s", v: String(v), z: "@" }; }
  function str(v) { return { t: "s", v: String(v) }; }
  function num(v, z) { var c = { t: "n", v: v }; if (z) c.z = z; return c; }
  function formula(f, z, v) { return { t: "n", f: f, v: v, z: z }; }

  // Excel serial for today's local date (days since 1899-12-30).
  function todaySerial() {
    var n = new Date();
    return Math.floor((Date.UTC(n.getFullYear(), n.getMonth(), n.getDate()) -
      Date.UTC(1899, 11, 30)) / 86400000);
  }

  function makeSheet(XLSX, headers, dataRows, widths) {
    var ws = {};
    headers.forEach(function (h, c) {
      ws[XLSX.utils.encode_cell({ r: 0, c: c })] = { t: "s", v: h };
    });
    dataRows.forEach(function (row, ri) {
      row.forEach(function (cell, c) {
        if (cell != null) ws[XLSX.utils.encode_cell({ r: ri + 1, c: c })] = cell;
      });
    });
    ws["!ref"] = XLSX.utils.encode_range({
      s: { r: 0, c: 0 }, e: { r: dataRows.length, c: headers.length - 1 }
    });
    if (widths) ws["!cols"] = widths.map(function (w) { return { wch: w }; });
    return ws;
  }

  function bookBytes(XLSX, ws, sheetName) {
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
    return XLSX.write(wb, { bookType: "xlsx", type: "array" });
  }

  // ---- builders ------------------------------------------------------------
  function buildWcew(XLSX, rows, opts) {
    var data = rows.map(function (p) {
      var cost = opts.wcewRoundHalf ? roundUpHalf(p.base) : p.base;
      var desc = p.desc + (opts.cct ? teresaCct(p.code) : "");
      var row = [txt(p.code), str(desc), txt("077"), txt("001"), txt("3"),
        txt("001"), txt("3"), txt(p.order), txt(opts.brand), txt("EA"),
        num(cost), num(25), num(45), num(45), num(35), num(20)];
      for (var i = 0; i < 5; i++) row.push(num(NEG));       // C..G
      row.push(num(10));                                     // H
      for (i = 0; i < 6; i++) row.push(num(NEG));            // I..N
      return row;
    });
    return bookBytes(XLSX, makeSheet(XLSX, WCEW_HEADERS, data, WCEW_WIDTHS),
      "PRODUCT TEMPLATE");
  }

  function buildLighting(XLSX, rows, opts) {
    var data = rows.map(function (p) {
      var b = p.base;
      var desc = p.desc + (opts.cct ? teresaCct(p.code) : "");
      var row = [str(opts.prefix + p.code), str(desc), txt("009"), txt("001"),
        txt("3"), txt("001"), txt("3"), str(opts.brand), str("EA"),
        num(b * COST_UPLIFT, "0.00"), num(b * TRADE_MULT, "0.00"),
        num(b * RETAIL_MULT, "0.00"), num(roundUpWhole(b * RETAIL_MULT * VAT_MULT)),
        num(b * PRICE_A_MULT, "0.0")];
      for (var i = 0; i < 13; i++) row.push(num(NEG));       // B..N
      return row;
    });
    return bookBytes(XLSX, makeSheet(XLSX, LIGHTING_HEADERS, data, LIGHTING_WIDTHS),
      "PRODUCT TEMPLATE");
  }

  function buildSupplier(XLSX, rows, opts) {
    var data = rows.map(function (p) {
      return [str(opts.prefix + p.code), txt(opts.supplierCode), str(p.code),
        num(p.base * COST_UPLIFT, "#,##0.00"),
        formula("TODAY()", "mm-dd-yy", todaySerial())];
    });
    return bookBytes(XLSX, makeSheet(XLSX, SUPPLIER_HEADERS, data, SUPPLIER_WIDTHS),
      "ProdSuppRecImport");
  }

  // Build all three files. Returns { files:[{name,data}], meta:{...} }
  function buildFiles(XLSX, aoa, opts) {
    opts = opts || {};
    opts.brand = opts.brand || "FUMAGALLI";
    opts.supplierCode = opts.supplierCode || "W001";
    opts.prefix = opts.prefix || "W/";
    if (opts.cct === undefined) opts.cct = true;
    var stem = opts.stem || "OUTPUT";
    var parsed = extractRows(aoa);
    var rows = parsed.rows;
    var files = [
      { name: "PRODUCT_WCEW_" + stem + ".xlsx", data: buildWcew(XLSX, rows, opts) },
      { name: "PRODUCT_LIGHTING_" + stem + ".xlsx", data: buildLighting(XLSX, rows, opts) },
      { name: "SUPPLIER_LIGHTING_" + stem + ".xlsx", data: buildSupplier(XLSX, rows, opts) }
    ];
    var meta = { count: rows.length, priceName: parsed.priceName,
                 skipped: parsed.skipped, hasCandidates: parsed.hasCandidates };

    // 4th template (optional): ECI / Magnalux supplier import.
    if (opts.eciList && opts.magList) {
      if (!parsed.hasCandidates) {
        meta.supplierImportError =
          'Source has no "All candidates (with arithmetic)" column — the 4th ' +
          "template needs it to read the ECI/Magnalux prices.";
      } else {
        var si = buildSupplierImport(XLSX, rows, opts);
        files.push({ name: "SUPPLIER_WCEW_" + stem + ".xlsx", data: si.data });
        meta.supplierImport = si.meta;
      }
    }
    return { files: files, meta: meta };
  }

  return {
    buildFiles: buildFiles, extractRows: extractRows, findHeader: findHeader,
    parsePriceList: parsePriceList, parseCandidates: parseCandidates,
    buildSupplierImport: buildSupplierImport,
    teresaCct: teresaCct, roundUpHalf: roundUpHalf, roundUpWhole: roundUpWhole
  };
});
