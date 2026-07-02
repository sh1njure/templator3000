/*
 * Supplier price-file generator — shared logic (browser + node).
 *
 * Mirrors generate_price_files.py 1:1. Given the rows parsed from a source
 * price list it builds the three output workbooks:
 *   PRODUCT_WCEW_<stem>.xlsx, PRODUCT_LIGHTING_<stem>.xlsx,
 *   SUPPLIER_DETAILS_<stem>.xlsx
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
      var hasCode = keys.some(function (h) {
        return CODE_ALIASES.some(function (a) { return h.indexOf(a) >= 0; });
      });
      var hasPrice = keys.some(function (h) {
        return PRICE_PRIORITY.some(function (p) { return h.indexOf(p) >= 0; });
      });
      if (hasCode && hasPrice) return { row: r, headers: headers };
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
        base: p
      });
    }
    return { rows: rows, priceName: price.name, skipped: skipped };
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
    return {
      files: [
        { name: "PRODUCT_WCEW_" + stem + ".xlsx", data: buildWcew(XLSX, rows, opts) },
        { name: "PRODUCT_LIGHTING_" + stem + ".xlsx", data: buildLighting(XLSX, rows, opts) },
        { name: "SUPPLIER_DETAILS_" + stem + ".xlsx", data: buildSupplier(XLSX, rows, opts) }
      ],
      meta: { count: rows.length, priceName: parsed.priceName, skipped: parsed.skipped }
    };
  }

  return {
    buildFiles: buildFiles, extractRows: extractRows, findHeader: findHeader,
    teresaCct: teresaCct, roundUpHalf: roundUpHalf, roundUpWhole: roundUpWhole
  };
});
