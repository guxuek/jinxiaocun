/* ============================================================
 * Excel 导入/导出 v7 —— 四个导入源 + 共享列定义导出
 *   1) 三合一数据（入库NK / 出库XK / 在途表）
 *   2) C仓数据（主仓流水）
 *   3) 其他仓数据（自动拆分「序列号+说明」，识别仓位列）
 *   4) 系统出库数据
 * ============================================================ */
(function (global) {
"use strict";
var C = global.CORE;
var MAX_COL = 80;
var HEADER_SCAN = 12;

function requireXlsx() {
  if (typeof XLSX === "undefined") {
    throw new Error("缺少 Excel 组件（vendor/xlsx.full.min.js）。请确认 vendor 文件夹与本程序放在一起。");
  }
  return XLSX;
}

function readWorkbook(file) {
  var X = requireXlsx();
  return file.arrayBuffer().then(function (buf) {
    var wb;
    try { wb = X.read(buf, { type: "array", cellDates: true, dense: false }); }
    catch (e) { throw new Error("无法解析文件：" + (e && e.message ? e.message : "格式不支持")); }
    var out = {};
    wb.SheetNames.forEach(function (name) {
      var ws = wb.Sheets[name];
      var rows = X.utils.sheet_to_json(ws, { header: 1, raw: true, defval: "", blankrows: false });
      out[name] = rows.map(function (r) { return (r || []).slice(0, MAX_COL); });
    });
    return out;
  });
}
function findHeaderIndex(rows, kind) {
  var fields = C.SOURCE_FIELDS[kind] || [];
  var aliasSet = {};
  fields.forEach(function (f) { f.aliases.forEach(function (a) { aliasSet[C.headerKey(a)] = 1; }); });
  var best = 0, bestScore = -1;
  var limit = Math.min(rows.length, HEADER_SCAN);
  for (var i = 0; i < limit; i++) {
    var row = rows[i] || []; var score = 0;
    for (var j = 0; j < row.length; j++) {
      var k = C.headerKey(C.cellStr(row[j]));
      if (!k) continue;
      score += aliasSet[k] ? 2 : 0.1;
    }
    if (score > bestScore) { bestScore = score; best = i; }
  }
  return best;
}
function autoMap(headers, kind) {
  var mapping = {}, used = {};
  var fields = C.SOURCE_FIELDS[kind] || [];
  /* 在途表：date 字段强制优先绑定「采购月份」列，防止被其他列抢占 */
  if (kind === "transit") {
    for (var i = 0; i < headers.length; i++) {
      if (C.headerKey(headers[i]) === "采购月份") { mapping.date = headers[i]; used[headers[i]] = true; break; }
    }
  }
  fields.forEach(function (field) {
    if (mapping[field.key]) return;
    var aliases = field.aliases.map(C.headerKey);
    for (var i = 0; i < headers.length; i++) {
      var h = headers[i];
      if (used[h]) continue;
      if (aliases.indexOf(C.headerKey(h)) >= 0) { mapping[field.key] = h; used[h] = true; break; }
    }
  });
  return mapping;
}
function excelDateCell(v) {
  if (v instanceof Date) return C.isoFromMs(v.getTime()) || "";
  if (typeof v === "number") return C.isoFromMs(C.parseDateMs(v)) || String(v);
  return C.isoFromMs(C.parseDateMs(C.cellStr(v))) || C.cellStr(v);
}
function previewSheet(sheetName, rows, kind) {
  var hi = findHeaderIndex(rows, kind);
  var headers = (rows[hi] || []).map(function (h, i) { var s = C.cellStr(h); return s || ("列" + (i + 1)); });
  var mapping = autoMap(headers, kind);
  var missing = (C.SOURCE_FIELDS[kind] || []).filter(function (f) { return f.required && !mapping[f.key]; }).map(function (f) { return f.label; });
  var body = rows.slice(hi + 1);
  return { sheetName: sheetName, headers: headers, mapping: mapping, missingRequired: missing, preview: body.slice(0, 5).map(function (r) { return r.map(C.cellStr); }), totalRows: body.length, records: body, kind: kind };
}
function parseRecords(pv, kind) {
  var idx = {};
  pv.headers.forEach(function (h, i) { idx[h] = i; });
  function col(rec, name) { if (name == null) return ""; var i = idx[name]; return i == null || i < 0 ? "" : rec[i]; }
  var rows = [], errors = [];
  for (var i = 0; i < pv.records.length; i++) {
    var rec = pv.records[i], rowNo = i + 2;
    try {
      var serialRaw = C.cellStr(col(rec, pv.mapping.serial));
      var split = C.splitSerialNote(serialRaw);
      var serial = split.serial || serialRaw;
      if (!C.looksLikeSerial(serial)) continue;
      if (kind === "warehouse") {
        rows.push({ id: "wh-" + i, serial: serial, docType: C.cellStr(col(rec, pv.mapping.docType)), docStatus: C.cellStr(col(rec, pv.mapping.docStatus)), ioTime: excelDateCell(col(rec, pv.mapping.ioTime)), warehouse: "C仓", note: split.note });
      } else if (kind === "otherWarehouse") {
        rows.push({ id: "ow-" + i, serial: serial, docType: C.cellStr(col(rec, pv.mapping.docType)), docStatus: C.cellStr(col(rec, pv.mapping.docStatus)), ioTime: excelDateCell(col(rec, pv.mapping.ioTime)), warehouse: C.cellStr(col(rec, pv.mapping.warehouse)) || "其他仓", note: split.note });
      } else if (kind === "systemOutbound") {
        rows.push({ id: "so-" + i, serial: serial, orderNo: C.cellStr(col(rec, pv.mapping.orderNo) || rec[0]), downPayment: C.parseMoney(pv.mapping.downPayment ? col(rec, pv.mapping.downPayment) : rec[29]), downPaymentRaw: C.cellStr(pv.mapping.downPayment ? col(rec, pv.mapping.downPayment) : rec[29]), date: excelDateCell(col(rec, pv.mapping.date) || rec[15]), reason: C.cellStr(col(rec, pv.mapping.reason)) || "ERP出库", modelRaw: C.cellStr(col(rec, pv.mapping.modelRaw)), note: split.note });
      } else if (kind === "inbound") {
        rows.push({ id: "nk-" + i, date: excelDateCell(col(rec, pv.mapping.date)), modelRaw: C.cellStr(col(rec, pv.mapping.modelRaw)), color: C.cellStr(col(rec, pv.mapping.color)), memoryRaw: C.cellStr(col(rec, pv.mapping.memoryRaw)), conditionRaw: C.cellStr(col(rec, pv.mapping.conditionRaw)), serial: serial, inboundStatus: C.cellStr(col(rec, pv.mapping.inboundStatus)), cost: C.parseMoney(col(rec, pv.mapping.cost)), supplier: C.cellStr(col(rec, pv.mapping.supplier)), outboundReason: C.cellStr(col(rec, pv.mapping.outboundReason)) });
      } else if (kind === "outbound") {
        rows.push({ id: "xk-" + i, orderNo: C.cellStr(col(rec, pv.mapping.orderNo) || rec[0]), orderNoRaw: C.cellStr(col(rec, pv.mapping.orderNo) || rec[0]), cash: C.parseMoney(pv.mapping.cash ? col(rec, pv.mapping.cash) : rec[10]), cashRaw: C.cellStr(pv.mapping.cash ? col(rec, pv.mapping.cash) : rec[10]), orderDate: excelDateCell(col(rec, pv.mapping.orderDate)), serial: serial, modelRaw: C.cellStr(col(rec, pv.mapping.modelRaw)), conditionRaw: C.cellStr(col(rec, pv.mapping.conditionRaw)), cost: C.parseMoney(col(rec, pv.mapping.cost)), reason: C.cellStr(col(rec, pv.mapping.reason)) });
      } else if (kind === "transit") {
        rows.push({ id: "tr-" + i, date: excelDateCell(col(rec, pv.mapping.date)), modelRaw: C.cellStr(col(rec, pv.mapping.modelRaw)), memoryRaw: C.cellStr(col(rec, pv.mapping.memoryRaw)), color: C.cellStr(col(rec, pv.mapping.color)), serial: serial, cny: C.parseMoney(col(rec, pv.mapping.cny)), vnd: C.parseMoney(col(rec, pv.mapping.vnd)), status: C.cellStr(col(rec, pv.mapping.status)) || "在途", conditionRaw: C.cellStr(col(rec, pv.mapping.conditionRaw)), _sheet: pv.sheetName });
      }
    } catch (e) { errors.push({ row: rowNo, message: e && e.message ? e.message : "解析失败" }); }
  }
  return { rows: rows, errors: errors };
}
function importSingle(file, kind) {
  return readWorkbook(file).then(function (sheets) {
    var names = Object.keys(sheets);
    if (!names.length) throw new Error("文件里没有工作表");
    /* 多子表类型（其他仓 / 系统出库）：合并所有工作表，工作表名作为仓位名 */
    var mergeAll = (kind === "otherWarehouse" || kind === "systemOutbound") && names.length > 1;
    if (mergeAll) {
      var all = [], errors = [], sheetNames = [], missing = [];
      names.forEach(function (n) {
        var pv = previewSheet(n, sheets[n], kind);
        if (pv.missingRequired.length) { return; } /* 跳过无关键列的表 */
        var parsed = parseRecords(pv, kind);
        parsed.rows.forEach(function (r) {
          if (kind === "otherWarehouse" && (!r.warehouse || r.warehouse === "其他仓")) r.warehouse = n; /* 用工作表名当仓位 */
          all.push(r);
        });
        errors = errors.concat(parsed.errors);
        sheetNames.push(n);
        if (!missing.length) missing = pv.missingRequired;
      });
      if (!all.length) throw new Error("没有识别到有效数据行（需要『序列号/唯一码』列）。已扫描工作表：" + names.join("、"));
      return { sheetName: sheetNames.join("、"), mapping: {}, missingRequired: missing, rows: all, errors: errors };
    }
    /* 单表类型：选最佳工作表 */
    var best = null;
    names.forEach(function (n) {
      var pv = previewSheet(n, sheets[n], kind);
      var score = Object.keys(pv.mapping).length - pv.missingRequired.length * 10 + pv.totalRows / 1000;
      if (!best || score > best.score) best = { pv: pv, score: score };
    });
    var parsed = parseRecords(best.pv, kind);
    if (!parsed.rows.length) throw new Error("没有识别到有效数据行（需要『序列号/唯一码』列）。工作表：" + best.pv.sheetName);
    /* 其他仓单表时也用工作表名兜底仓位 */
    if (kind === "otherWarehouse") parsed.rows.forEach(function (r) { if (!r.warehouse || r.warehouse === "其他仓") r.warehouse = best.pv.sheetName; });
    return { sheetName: best.pv.sheetName, mapping: best.pv.mapping, missingRequired: best.pv.missingRequired, rows: parsed.rows, errors: parsed.errors };
  });
}
function importErp(file) {
  return readWorkbook(file).then(function (sheets) {
    var names = Object.keys(sheets);
    var pick = C.pickErpSheets(names);
    var result = { sheets: pick, inbound: null, outbound: null, transit: null, transitOptions: [], errors: [] };
    ["inbound", "outbound", "transit"].forEach(function (kind) {
      var name = pick[kind];
      if (!name) return;
      var pv = previewSheet(name, sheets[name], kind);
      var parsed = parseRecords(pv, kind);
      result[kind] = { sheetName: name, rows: parsed.rows, mapping: pv.mapping };
      result.errors = result.errors.concat(parsed.errors);
    });
    /* 所有「在途XX」子表作为可切换选项（在途表 / 在途8月…） */
    (pick.transitAll || []).forEach(function (n) {
      var pv = previewSheet(n, sheets[n], "transit");
      var parsed = parseRecords(pv, "transit");
      if (parsed.rows.length) {
        result.transitOptions.push({ sheetName: n, rows: parsed.rows, sheetDate: C.transitSheetDate(n) });
      }
    });
    if (!result.inbound && !result.outbound && !result.transit) {
      var pv = previewSheet(names[0], sheets[names[0]], "inbound");
      var parsed = parseRecords(pv, "inbound");
      if (parsed.rows.length) { result.inbound = { sheetName: names[0], rows: parsed.rows, mapping: pv.mapping }; result.errors = result.errors.concat(parsed.errors); }
      else throw new Error("未找到『入库NK / 出库XK / 在途表』工作表，也无法按入库表解析第一个工作表");
    }
    return result;
  });
}

/* ============================================================
 * 共享 Sheet 构建器（表头来自 CORE.COLUMN_DEFS）
 * ============================================================ */
function buildSheet(def, body) {
  var X = requireXlsx();
  var headers = def.headers;
  var aoa = [], merges = [];
  for (var i = 0; i < headers.length; i++) {
    var row = [], col = 0;
    for (var j = 0; j < headers[i].length; j++) {
      var h = headers[i][j];
      if (typeof h === "string") { row.push(h); col++; }
      else {
        row.push(h.text);
        var span = h.colspan || 1;
        if (span > 1) merges.push({ s: { r: i, c: col }, e: { r: i, c: col + span - 1 } });
        col += span;
      }
    }
    aoa.push(row);
  }
  body.forEach(function (r) { aoa.push(r.slice()); });
  var ws = X.utils.aoa_to_sheet(aoa);
  if (merges.length) ws["!merges"] = merges;
  if (def.widths) ws["!cols"] = def.widths.map(function (w) { return { wch: w }; });
  return ws;
}

function vnd(n) { return n == null ? "" : Math.round(n); }
function cny(n, fx) { fx = fx || 3600; return n == null ? "" : +(Math.round(n) / fx).toFixed(2); }
function pctN(n, d) { return d > 0 ? Math.round(n / d * 1000) / 10 + "%" : "—"; }

function sumVnd(list) { return list.reduce(function (a, u) { return a + (u.costVnd || 0); }, 0); }
function sumSales(list) { return list.reduce(function (a, r) { return a + (r.sales || 0); }, 0); }

/* 品类汇总（新公式：日销量=销量/窗口；可售天数=库存含在途÷日销量） */
function buildCategoryRows(stock, transit, turnover, s) {
  var win = s.salesWindowDays || 60;
  function byCond(cond) {
    return turnover.filter(function (r) { return r.condition === cond; });
  }
  function row(label, cond) {
    var rows = byCond(cond);
    var list = stock.filter(function (u) { return u.condition === cond; });
    var tr = transit.filter(function (u) { return u.condition === cond; });
    var problem = list.filter(function (u) { return u.isProblem; });
    var clean = list.filter(function (u) { return !u.isProblem; });
    var inclN = list.length + tr.length;
    var sales = sumSales(rows);
    var daily = sales > 0 ? sales / win : 0;
    var sellDays = daily > 0 ? Math.round(inclN / daily) : (inclN > 0 ? "—" : 0);
    var stockVnd = sumVnd(list), problemVnd = sumVnd(problem), cleanVnd = sumVnd(clean);
    return [label, inclN, problem.length, clean.length, sales, +daily.toFixed(2), sellDays,
      vnd(stockVnd), vnd(problemVnd), vnd(cleanVnd), vnd(Math.round(rows.reduce(function (a, r) { return a + (r.stockValue || 0); }, 0) * 0)),
      cny(stockVnd, s.fxRate), cny(problemVnd, s.fxRate), cny(cleanVnd, s.fxRate)];
  }
  return [row("新机", "新机"), row("二手机", "二手机"), row("未识别", "未知"), row("合计", null)].map(function (r, i) {
    if (i < 3) return r;
    /* 合计 */
    var rows = turnover;
    var inclN = stock.length + transit.length;
    var problem = stock.filter(function (u) { return u.isProblem; });
    var clean = stock.filter(function (u) { return !u.isProblem; });
    var sales = sumSales(rows);
    var daily = sales > 0 ? sales / win : 0;
    var sellDays = daily > 0 ? Math.round(inclN / daily) : "—";
    return ["合计", inclN, problem.length, clean.length, sales, +daily.toFixed(2), sellDays,
      vnd(sumVnd(stock)), vnd(sumVnd(problem)), vnd(sumVnd(clean)), "",
      cny(sumVnd(stock), s.fxRate), cny(sumVnd(problem), s.fxRate), cny(sumVnd(clean), s.fxRate)];
  });
}

/* 周转表数据行（含在途列、日销量、可售天数） */
function buildTurnoverRows(list, stockUnits, s) {
  var unitByKey = new Map();
  stockUnits.forEach(function (u) { var k = u.model + "|" + u.memory + "|" + u.condition; if (!unitByKey.has(k)) unitByKey.set(k, u); });
  var tot = { stock: 0, transit: 0, sales: 0, stockVnd: 0, salesVnd: 0 };
  var rows = list.map(function (x) {
    var k = x.model + "|" + x.memory + "|" + x.condition;
    var sample = unitByKey.get(k);
    var age = sample && sample.ageDays != null ? sample.ageDays : 0;
    var salesVnd = sample ? Math.round((sample.costVnd || 0) * x.sales) : 0;
    var sell = x.sellableDays == null ? "—" : x.sellableDays;
    tot.stock += x.stock; tot.transit += x.transit || 0; tot.sales += x.sales;
    tot.stockVnd += x.stockValue || 0; tot.salesVnd += salesVnd;
    var remark = x.sales > 0 ? "" : "近" + (s.salesWindowDays || 60) + "天无销量";
    return [x.model, x.memory, x.stockInclTransit != null ? x.stockInclTransit : x.stock, x.sales,
      +(x.dailySales || 0).toFixed(2), sell, vnd(x.stockValue), cny(x.stockValue, s.fxRate),
      vnd(salesVnd), remark, age];
  });
  var win = s.salesWindowDays || 60;
  var totDaily = tot.sales > 0 ? tot.sales / win : 0;
  var totSell = totDaily > 0 ? Math.round((tot.stock + tot.transit) / totDaily) : "—";
  rows.push(["合计", "", tot.stock + tot.transit, tot.sales, +totDaily.toFixed(2), totSell,
    vnd(tot.stockVnd), cny(tot.stockVnd, s.fxRate), vnd(tot.salesVnd), "", ""]);
  return rows;
}

function computeAge(stock) {
  var age = { "0-30天": { n: 0, v: 0 }, "30-60天": { n: 0, v: 0 }, "60-90天": { n: 0, v: 0 }, "90-180天": { n: 0, v: 0 }, ">=180天": { n: 0, v: 0 } };
  stock.forEach(function (u) {
    var a = u.ageDays || 0, k;
    if (a < 30) k = "0-30天"; else if (a < 60) k = "30-60天"; else if (a < 90) k = "60-90天"; else if (a < 180) k = "90-180天"; else k = ">=180天";
    age[k].n++; age[k].v += u.costVnd || 0;
  });
  return age;
}

/* 完整周转报告导出（6 Sheet，与工作台一致） */
function exportDashboardReport(snap, result) {
  var X = requireXlsx();
  var units = result.units;
  var stock = units.filter(function (u) { return u.status === "在库"; });
  var transit = units.filter(function (u) { return u.status === "在途"; });
  var s = snap.settings, kpi = result.kpi;
  var win = s.salesWindowDays || 60;
  var totalSales60 = result.turnover.reduce(function (a, r) { return a + (r.sales || 0); }, 0);
  var normalStock = stock.filter(function (u) { return !u.isProblem; }).length;
  /* 汇总可售天数 = 正常库存 × 60 ÷ 60天销量 */
  var sellableSummary = totalSales60 > 0 ? Math.round(normalStock * win / totalSales60) : "—";
  var wb = X.utils.book_new();

  var ws0 = X.utils.aoa_to_sheet([
    ["越南手机周转表"],
    ["库龄计算日 " + result.ageDate + "　库存/在途计算日 " + result.stockDate + "　汇率 " + s.fxRate + " VND/CNY　库存(含在途) " + (stock.length + transit.length) + " 台"],
    ["其他仓数据：" + (s.includeOtherWarehouses ? "已计入" : "未计入（设置中可勾选）")], [],
    ["关键数据"],
    ["库存+在途", (stock.length + transit.length) + " 台"],
    ["正常可售", normalStock + " 台"],
    ["近" + win + "天销量", totalSales60 + " 台"],
    ["汇总可售天数（正常库存×" + win + "÷销量）", sellableSummary],
    ["库存人民币", cny(sumVnd(stock), s.fxRate) + " 元"]
  ]);
  ws0["!cols"] = C.COLUMN_DEFS.summary.widths;
  X.utils.book_append_sheet(wb, ws0, "越南手机周转表");

  X.utils.book_append_sheet(wb, buildSheet(C.COLUMN_DEFS.category_summary, buildCategoryRows(stock, transit, result.turnover, s)), "1品类汇总");
  X.utils.book_append_sheet(wb, buildSheet(C.COLUMN_DEFS.turnover, buildTurnoverRows(result.turnover.filter(function (x) { return x.condition === "新机"; }), stock, s)), "2新机周转");
  X.utils.book_append_sheet(wb, buildSheet(C.COLUMN_DEFS.turnover, buildTurnoverRows(result.turnover.filter(function (x) { return x.condition === "二手机"; }), stock, s)), "3二手机周转");
  X.utils.book_append_sheet(wb, buildSheet(C.COLUMN_DEFS.turnover, buildTurnoverRows(result.turnover, stock, s)), "4全部周转");

  var age = computeAge(stock);
  X.utils.book_append_sheet(wb, buildSheet(C.COLUMN_DEFS.age_segment, [
    ["0-30天", age["0-30天"].n, cny(age["0-30天"].v, s.fxRate), pctN(age["0-30天"].n, stock.length)],
    ["30-60天", age["30-60天"].n, cny(age["30-60天"].v, s.fxRate), pctN(age["30-60天"].n, stock.length)],
    ["60-90天", age["60-90天"].n, cny(age["60-90天"].v, s.fxRate), pctN(age["60-90天"].n, stock.length)],
    ["90-180天", age["90-180天"].n, cny(age["90-180天"].v, s.fxRate), pctN(age["90-180天"].n, stock.length)],
    ["≥180天", age[">=180天"].n, cny(age[">=180天"].v, s.fxRate), pctN(age[">=180天"].n, stock.length)],
    ["在库合计", stock.length, cny(sumVnd(stock), s.fxRate), "100%"],
    ["近" + win + "天正常出货", totalSales60, "", ""]
  ]), "5库龄分段");

  X.utils.book_append_sheet(wb, buildSheet(C.COLUMN_DEFS.system_status, [
    ["📦 在库", "截至" + result.stockDate + " 入库>出库", kpi.inStock, ""],
    ["✅ 已出库", "入=出", kpi.sold, ""],
    ["🚚 在途", "截至" + result.stockDate + " 仍未入库", kpi.transit, ""],
    ["⏳ 待核", "出>入 / 仅仓库记录", kpi.pending, ""],
    ["⚠️ 异常机", "需人工处理", kpi.anomaly, ""],
    ["❗ 问题机", "入库原因≠出租", kpi.problem, ""]
  ]), "6系统状态");

  var detailRows = function (list) { return list.slice().sort(function (a,b) { return String(a.model||"").localeCompare(String(b.model||""), "en", {numeric:true}) || String(a.memory||"").localeCompare(String(b.memory||""), "en", {numeric:true}) || String(a.serial||"").localeCompare(String(b.serial||"")); }).map(function (u) { return [u.serial,u.model,u.memory,u.color||"",u.condition,u.status,u.isProblem?"是":"",u.costVnd==null?"":u.costVnd,u.erpInboundDate||"",u.erpAgeDays==null?"":u.erpAgeDays,u.whInboundDate||"",u.whAgeDays==null?"":u.whAgeDays,u.remark||""]; }); };
  var wsStock = X.utils.aoa_to_sheet([["序列号","机型","内存","颜色","新旧","状态","问题机","成本VND","入库表日期","入库表库龄","仓入库日期","仓库龄","备注"]].concat(detailRows(stock)));
  X.utils.book_append_sheet(wb, wsStock, "在库明细");
  var wsTransit = X.utils.aoa_to_sheet([["序列号","机型","内存","颜色","新旧","状态","成本VND","在途日","在途龄","备注"]].concat(detailRows(transit).map(function(r){return [r[0],r[1],r[2],r[3],r[4],r[5],r[7],r[8],r[9],r[12]];})));
  X.utils.book_append_sheet(wb, wsTransit, "在途明细");

  var buf = X.write(wb, { type: "array", bookType: "xlsx" });
  return new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
}

function exportTurnover(result) {
  var header = C.COLUMN_DEFS.turnover.headers[0];
  var rows = result.turnover.map(function (r) {
    return [r.model, r.memory, r.stockInclTransit != null ? r.stockInclTransit : r.stock, r.sales,
      +(r.dailySales || 0).toFixed(2), r.sellableDays == null ? "" : r.sellableDays,
      r.stockValue, "", "", r.sales > 0 ? "" : "近窗无销量", ""];
  });
  return blobOfSheet(header, rows, "周转表");
}
function exportUnits(units, name) {
  var header = C.COLUMN_DEFS.ledger.headers[0].slice(0, 16);
  var rows = (units || []).map(function (u) {
    var isTr = u.status === "在途";
    return [u.serial, u.model, u.memory, u.color || "", u.condition, u.status,
      u.inCount + "入/" + u.outCount + "出", u.whInCount + "入/" + u.whOutCount + "出",
      u.isProblem ? "是" : "", u.isAnomaly ? u.anomalyReason : (u.anomalyHandled ? "已处理" : ""),
      u.costVnd == null ? "" : u.costVnd,
      u.erpInboundDate || "", u.erpAgeDays == null ? "" : u.erpAgeDays,
      isTr ? (u.transitDate || "") : (u.whInboundDate || ""),
      isTr ? (u.transitAgeDays == null ? "" : u.transitAgeDays) : (u.whAgeDays == null ? "" : u.whAgeDays),
      u.remark || ""];
  });
  return blobOfSheet(header, rows, name || "台账");
}
function blobOfSheet(header, rows, sheetName) {
  var X = requireXlsx();
  var wb = X.utils.book_new();
  X.utils.book_append_sheet(wb, X.utils.aoa_to_sheet([header].concat(rows)), sheetName || "Sheet1");
  var buf = X.write(wb, { type: "array", bookType: "xlsx" });
  return new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
}
function exportBackup(snap) { return new Blob([JSON.stringify(snap, null, 2)], { type: "application/json" }); }
function importBackup(file) {
  return file.text().then(function (t) {
    var obj = JSON.parse(t);
    if (!obj || typeof obj !== "object" || !Array.isArray(obj.inbound)) throw new Error("备份文件格式不正确");
    var empty = C.emptySnapshot();
    return {
      warehouse: obj.warehouse || [], otherWarehouse: obj.otherWarehouse || [],
      systemOutbound: obj.systemOutbound || [],
      inbound: obj.inbound || [], outbound: obj.outbound || [], transit: obj.transit || [],
      overrides: obj.overrides || {}, settings: Object.assign(empty.settings, obj.settings || {}),
      importMeta: Object.assign(empty.importMeta || {}, obj.importMeta || {}),
      importHistory: obj.importHistory || {}, isSample: Boolean(obj.isSample)
    };
  });
}
function downloadBlob(blob, filename) {
  var url = URL.createObjectURL(blob);
  var a = document.createElement("a"); a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(function () { URL.revokeObjectURL(url); }, 3000);
}

global.EXCELIO = {
  readWorkbook: readWorkbook, previewSheet: previewSheet, parseRecords: parseRecords,
  importSingle: importSingle, importErp: importErp,
  exportTurnover: exportTurnover, exportDashboardReport: exportDashboardReport, exportUnits: exportUnits,
  exportBackup: exportBackup, importBackup: importBackup, downloadBlob: downloadBlob
};
})(typeof window !== "undefined" ? window : globalThis);
