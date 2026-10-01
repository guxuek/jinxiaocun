/* ============================================================
 * 界面与状态管理 v7
 * ① 数据存文件夹（本地版 file://）/ 云端服务器（http://）
 * ② 其他仓数据开关（设置或工作台勾选）
 * ③ 四个导入源：三合一 / C仓 / 其他仓 / 系统出库 + 导入历史切换
 * ④ 其他仓子版块（按仓位分组清单）
 * ⑤ Excel入库日 vs 仓入库日 双列展示
 * ⑥ 双日期：库龄日 / 库存·在途日
 * ⑦ 新公式：日销量=销量/60；可售天数=库存含在途÷日销量；汇总可售=正常库存×60÷60天销量
 * ============================================================ */
(function () {
"use strict";
var C = window.CORE, DB = window.DB, XIO = window.EXCELIO;

var state = {
  snap: null, result: null, route: "home",
  query: "",
  filters: { status: [], condition: [], model: [], anomaly: "all", problem: "all", age: "all", supplier: "" },
  pageSize: 25, page: 1, columnFilters: {}, quickFilter: "",
  activeUnit: null, editing: false, drawerMode: null, /* unit | flow | filter */
  flowUnit: null, flowType: "erp",
  busy: null
};

function $(sel) { return document.querySelector(sel); }
function el(tag, cls, html) {
  var e = document.createElement(tag);
  if (cls) e.className = cls;
  if (html != null) e.innerHTML = html;
  return e;
}
function esc(s) { return String(s == null ? "" : s).replace(/[&<>"']/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]; }); }
function nfmt(n) { return n == null ? "—" : Math.round(n).toLocaleString("zh-CN"); }
function cnyFmt(n, fx) { fx = fx || 3600; return n == null ? "—" : +(Math.round(n) / fx).toFixed(2); }
function pct(n, d) { return d > 0 ? (n / d * 100).toFixed(1) + "%" : "—"; }

function toast(msg, type) {
  var box = $("#toasts");
  var t = el("div", "toast " + (type || "ok"), esc(msg));
  box.appendChild(t);
  setTimeout(function () { t.classList.add("show"); }, 10);
  setTimeout(function () { t.classList.remove("show"); setTimeout(function () { t.remove(); }, 300); }, 3200);
}

function recompute() { state.result = C.computeAll(state.snap); }
function persist() {
  return DB.saveSnapshot(state.snap).catch(function (e) { toast("保存失败：" + (e && e.message ? e.message : e), "err"); });
}
function migrateTransit(snap) {
  if (!snap || !Array.isArray(snap.transit)) return snap;
  var KEYS = ["date","采购月份","purchaseMonth","purchaseDate","month","月份","制表日期","日期","采购日期","下单日期","etaDate","预计到货日","inTransitDate","在途日期"];
  snap.transit.forEach(function (r) {
    if (!r) return;
    for (var i = 0; i < KEYS.length; i++) { var v = r[KEYS[i]]; if (v) { r.date = v; return; } }
  });
  return snap;
}
function setSnap(snap, silent) {
  snap = migrateTransit(snap);
  state.snap = snap;
  recompute();
  render();
  return persist().then(function () {
    if (!silent) toast("已保存到 " + DB.storageMode);
  });
}

/* ---------- 单据汇总（WPS/Excel表格 / 仓库，含仓位名） ---------- */
function buildFlowSummaries(snap) {
  var erpByKey = new Map();
  (snap.inbound || []).forEach(function (r) {
    var k = C.serialKey(r.serial); if (!k) return;
    if (!erpByKey.has(k)) erpByKey.set(k, []);
    erpByKey.get(k).push({ type: "入库", date: r.date || "", reason: r.outboundReason || "", supplier: r.supplier || "", cost: r.cost, src: "WPS/Excel入库表" });
  });
  (snap.outbound || []).forEach(function (r) {
    var k = C.serialKey(r.serial); if (!k) return;
    if (!erpByKey.has(k)) erpByKey.set(k, []);
    erpByKey.get(k).push({ type: "出库", date: r.orderDate || "", reason: r.reason || "", cost: r.cost, src: "WPS/Excel出库表" });
  });
  var whByKey = new Map();
  function pushWh(rows, defName, srcLabel) {
    (rows || []).forEach(function (r) {
      var k = C.serialKey(r.serial); if (!k) return;
      if (!whByKey.has(k)) whByKey.set(k, []);
      whByKey.get(k).push({ type: r.docType || "", date: r.ioTime || "", status: r.docStatus || "", warehouse: r.warehouse || defName, note: r.note || "", src: srcLabel });
    });
  }
  pushWh(snap.warehouse, "C仓", "C仓");
  pushWh(snap.otherWarehouse, "其他仓", "其他仓");
  /* 与核心计算一致：同仓位的已作废记录冲销此前最近一条有效仓记录。 */
  whByKey.forEach(function (list) {
    var active = [], byTime = list.slice().sort(function (a,b) { return (C.parseDateMs(a.date)||0) - (C.parseDateMs(b.date)||0); });
    byTime.forEach(function (x) {
      var cls = C.classifyWarehouse(x.type, x.status), wn = x.warehouse || "C仓";
      if (cls === "void") { for (var i=active.length-1;i>=0;i--) if ((active[i].warehouse||"C仓")===wn) { active.splice(i,1); break; } }
      else active.push(x);
    });
    list._active = active;
  });
  return { erpByKey: erpByKey, whByKey: whByKey };
}
function erpSummary(u, flow) {
  var list = flow.erpByKey.get(C.serialKey(u.serial)) || [];
  var inn = list.filter(function (x) { return x.type === "入库"; }).length;
  var out = list.filter(function (x) { return x.type === "出库"; }).length;
  if (!inn && !out) return "—";
  return inn + " 入 · " + out + " 出";
}
function whSummary(u, flow) {
  var list = flow.whByKey.get(C.serialKey(u.serial)) || [];
  list = list._active || list;
  if (!list.length) return "—";
  var inn = 0, out = 0;
  list.forEach(function (x) {
    var cls = C.classifyWarehouse(x.type, x.status);
    if (cls === "in") inn++; else if (cls === "out") out++;
  });
  var names = {};
  list.forEach(function (x) { names[x.warehouse] = 1; });
  return inn + " 入 · " + out + " 出 · " + Object.keys(names).join("/");
}

var ROUTES = [
  { id: "home", label: "工作台" },
  { id: "dashboard", parent: "turnover", label: "周转表" },
  { id: "ledger", parent: "turnover", label: "库存明细" },
  { id: "stock", parent: "turnover", label: "在库", presetStatus: "在库" },
  { id: "sold", parent: "turnover", label: "已出库", presetStatus: "已出库" },
    { id: "transit", parent: "turnover", label: "在途", presetStatus: "在途" },
  { id: "otherwh", parent: "turnover", label: "其他仓明细" },
  { id: "anomalies", parent: "turnover", label: "异常机", presetAnomaly: true },
  { id: "issues", parent: "turnover", label: "问题机", presetProblem: true },
  { id: "settings", parent: "turnover", label: "周转设置" },
  { id: "reconcile", label: "系统核对" },
  { id: "recon-inbound", parent: "reconcile", label: "入库核对" },
  { id: "recon-outbound", parent: "reconcile", label: "出库核对" },
  { id: "import", label: "数据导入" },
];
function navTo(id) {
  var r = ROUTES.find(function (x) { return x.id === id; }) || ROUTES[0];
  state.route = r.id;
  state.filters = { status: r.presetStatus ? [r.presetStatus] : [], condition: [], model: [], anomaly: r.presetAnomaly ? "yes" : "all", problem: r.presetProblem ? "yes" : "all", age: "all", supplier: "", dateFrom: "", dateTo: "", dateKey: { field: "", from: "", to: "", includeEmpty: false } };
  state.query = ""; state.page = 1; state.columnFilters = {}; state.quickFilter = "";
  state.activeUnit = null; state.editing = false; state.drawerMode = null; state.flowUnit = null;
  render();
}

function render() {
  var nav = $("#nav");
  nav.innerHTML = "";
  var kpi = state.result ? state.result.kpi : null;
  var expanded = state.navExpanded !== false;
  ROUTES.forEach(function (r) {
    if (r.id === "home") { var hb=el("button","nav-item"+(state.route === r.id ? " active" : ""),esc(r.label)); hb.onclick=function(){navTo(r.id);}; nav.appendChild(hb); return; }
    if (r.id === "dashboard") { var tb=el("button","nav-item nav-parent"+(state.route === r.id ? " active" : ""),"周转表　"+(expanded?"⌄":"›")); tb.onclick=function(){ if(state.route === "dashboard" || state.navExpanded===false){ state.navExpanded=!expanded; render(); } else navTo("dashboard"); }; nav.appendChild(tb); return; }
    if (r.id === "reconcile") { var rex=state.reconExpanded !== false; var rb=el("button","nav-item nav-parent"+(state.route === r.id ? " active" : ""),"系统核对　"+(rex?"⌄":"›")); rb.onclick=function(){ if(state.route.indexOf("recon")===0 || state.reconExpanded===false){ state.reconExpanded=!rex; render(); } else navTo(r.id); }; nav.appendChild(rb); return; }
    if (r.id === "import") { var ib=el("button","nav-item"+(state.route === r.id ? " active" : ""),esc(r.label)); ib.onclick=function(){navTo(r.id);}; nav.appendChild(ib); return; }
    if (r.parent === "reconcile" && state.reconExpanded !== false) { var cb=el("button","nav-item nav-child"+(state.route === r.id ? " active" : ""),"　"+esc(r.label)); cb.onclick=function(){navTo(r.id);}; nav.appendChild(cb); }
    if (r.parent === "turnover" && expanded) { var b=el("button","nav-item nav-child"+(state.route === r.id ? " active" : ""),"　"+esc(r.label)); if(r.presetAnomaly&&kpi&&kpi.anomaly)b.innerHTML+='<span class="pill warn">'+kpi.anomaly+'</span>'; if(r.presetProblem&&kpi&&kpi.problem)b.innerHTML+='<span class="pill bad">'+kpi.problem+'</span>'; if(r.id==="otherwh"&&state.snap.otherWarehouse&&state.snap.otherWarehouse.length)b.innerHTML+='<span class="pill">'+state.snap.otherWarehouse.length+'</span>'; b.onclick=function(){navTo(r.id);};nav.appendChild(b); }
  });

  var main = $("#main");
  main.innerHTML = "";
  if (state.busy) main.appendChild(el("div", "busy", esc(state.busy)));

  if (state.route === "home") renderHome(main);
  else if (state.route === "dashboard") renderDashboard(main);
  else if (state.route === "import") renderImport(main);
  else if (state.route === "settings") renderSettings(main);
  else if (state.route === "otherwh") renderOtherWarehouse(main);
  else if (state.route === "recon-outbound") renderOutboundReconcile(main);
  else if (state.route === "reconcile" || state.route === "recon-inbound") renderReconcile(main);
  else renderLedger(main);

  renderDrawer();
  $("#storage-mode").innerHTML = esc(DB.storageMode) +
    (state.snap.isSample ? "<br>演示数据" : "") +
    (DB.mode === "local-file" && DB.hasSavedFolder() && !DB.hasFolder() ? '<br><button id="resume-folder" class="btn btn-sm btn-outline">恢复文件夹保存</button>' : "");
  var resume = $("#resume-folder"); if (resume) resume.onclick = function () { DB.resumeFolder().then(function () { render(); toast("已恢复文件夹保存"); }).catch(function (e) { toast(e.message || "浏览器未授权，请重新选择文件夹", "warn"); }); };
}

/* ---------- 本地文件夹引导条 ---------- */
function folderBanner() {
  if (DB.mode !== "local-file" || DB.hasFolder() || DB.hasSavedFolder()) return null;
  var div = el("div", "busy");
  div.style.background = "#eff6ff"; div.style.color = "#1d4ed8"; div.style.border = "1px dashed #93c5fd";
  div.innerHTML = '📁 首次使用：选择保存数据的文件夹。以后会自动记住选择。';
  var btn = el("button", "btn btn-primary btn-sm", "选择数据文件夹");
  btn.style.marginLeft = "10px";
  btn.onclick = function () {
    DB.pickFolder().then(function (h) {
      toast("已绑定数据文件夹：" + h.name + "（后续修改自动保存到 snapshot.json）");
      render();
    }).catch(function (e) { toast(e.message || "已取消", "warn"); });
  };
  div.appendChild(btn);
  if (!DB.fsSupported()) div.innerHTML = "⚠ 当前浏览器不支持文件夹存储，请改用 Edge 或 Chrome 打开。";
  return div;
}

/* ============================================================
 * 工作台（⑥ 双日期 ⑦ 新公式）
 * ============================================================ */
function renderDashboard(main) {
  var banner = folderBanner(); if (banner) main.appendChild(banner);
  var kpi = state.result.kpi, r = state.result, s = state.snap.settings;
  var win = s.salesWindowDays || 60;
  var units = r.units;
  /* 状态分桶缓存：重渲染时不重复过滤大数组 */
  if (!state._buckets || state._buckets.units !== units) {
    var b = { stock: [], transit: [] };
    for (var bi = 0; bi < units.length; bi++) {
      var st = units[bi].status;
      if (st === "在库") b.stock.push(units[bi]);
      else if (st === "在途") b.transit.push(units[bi]);
    }
    state._buckets = { units: units, stock: b.stock, transit: b.transit };
  }
  var stock = state._buckets.stock, transit = state._buckets.transit;
  function sum(list) { var a = 0; for (var i = 0; i < list.length; i++) a += list[i].costVnd || 0; return a; }
  var newM = [], usedM = [], unkM = [], normalStock = 0;
  for (var si = 0; si < stock.length; si++) {
    var su = stock[si];
    if (su.condition === "新机") newM.push(su); else if (su.condition === "二手机") usedM.push(su); else unkM.push(su);
    if (!su.isProblem) normalStock++;
  }
  var totalVnd = sum(stock);
  var sales60 = r.turnover.reduce(function (a, x) { return a + (x.sales || 0); }, 0);
  /* ⑦ 汇总可售天数 = 正常库存 × 60 ÷ 60天销量 */
  var sellableSummary = sales60 > 0 ? Math.round(normalStock * win / sales60) : "—";

  var head = el("div", "page-head row");
  head.innerHTML = '<div><h1 class="dash-title">越南手机周转表</h1>' +
    '<p class="dash-sub">库龄计算日 <b>' + esc(r.ageDate) + '</b> · 库存/在途计算日 <b>' + esc(r.stockDate) + '</b> · 汇率 ' + nfmt(s.fxRate) + ' VND/CNY · 销量窗口 ' + win + ' 天' +
    (state.snap.isSample ? " · <b class='c-amber'>演示数据</b>" : "") + '</p>' +
    '<p class="hint">其他仓数据：' + (s.includeOtherWarehouses ? '<b class="c-green">已计入</b>' : '未计入') +
    ' · <label style="cursor:pointer"><input type="checkbox" id="ow-toggle"' + (s.includeOtherWarehouses ? " checked" : "") + '> 计入其他仓（' + (state.snap.otherWarehouse || []).length + ' 行）</label></p></div>' +
    '<button class="btn btn-primary" id="dash-export">📊 导出完整周转报告</button>';
  head.querySelector("#dash-export").onclick = function () {
    XIO.downloadBlob(XIO.exportDashboardReport(state.snap, state.result), "越南手机周转表_" + r.stockDate + ".xlsx");
  };
  head.querySelector("#ow-toggle").onchange = function (e) {
    var snap = Object.assign({}, state.snap);
    snap.settings = Object.assign({}, s, { includeOtherWarehouses: e.target.checked });
    setSnap(snap, true).then(function () { toast(e.target.checked ? "已计入其他仓数据" : "已排除其他仓数据"); });
  };
  main.appendChild(head);

  /* 一、品类汇总（⑦ 日销量列） */
  var p1 = el("div", "panel");
  p1.appendChild(el("div", "panel-title", "一、品类汇总（按新旧）"));
  function catRow(label, list, cond) {
    var p = list.filter(function (u) { return u.isProblem; });
    var c = list.filter(function (u) { return !u.isProblem; });
    var rows = cond ? r.turnover.filter(function (x) { return x.condition === cond; }) : r.turnover;
    var trN = transit.filter(function (u) { return cond ? u.condition === cond : true; }).length;
    var incl = list.length + trN;
    var sales = rows.reduce(function (a, x) { return a + (x.sales || 0); }, 0);
    var daily = sales > 0 ? sales / win : 0;
    var sellDays = daily > 0 ? Math.round(incl / daily) : (incl > 0 ? "—" : 0);
    var sv = sum(list), pv = sum(p), cv = sum(c);
    return [label, incl, p.length, c.length, sales, +daily.toFixed(2), sellDays,
      nfmt(sv), nfmt(pv), nfmt(cv), cnyFmt(sv, s.fxRate), cnyFmt(pv, s.fxRate), cnyFmt(cv, s.fxRate)];
  }
  var rows1 = [catRow("新机", newM, "新机"), catRow("二手机", usedM, "二手机"), catRow("未识别", unkM, null), catRow("合计", stock, null)];
  rows1[2] = catRow("未识别", unkM, "未知");
  var tbl1 = el("table", "tbl bordered");
  tbl1.innerHTML = '<thead><tr>' +
    '<th>品类</th><th>在库(含在途)</th><th>问题机</th><th>正常在库</th><th>近' + win + '天销量</th><th>日销量</th><th>可售天数</th>' +
    '<th>库存-盾</th><th>问题-盾</th><th>正常-盾</th><th>库存-元</th><th>问题-元</th><th>正常-元</th></tr></thead><tbody>' +
    rows1.map(function (rr, i) {
      return '<tr class="' + (i === 3 ? "total" : "") + '">' + rr.map(function (v) { return "<td>" + v + "</td>"; }).join("") + "</tr>";
    }).join("") + "</tbody>";
  p1.appendChild(tbl1);
  main.appendChild(p1);

  /* 关键数据条（含汇总可售天数） */
  var p0 = el("div", "panel key-bar");
  p0.innerHTML =
    kbItem("库存+在途", stock.length + transit.length, "台", "c-blue") +
    kbItem("正常可售", normalStock, "台", "c-green") +
    kbItem("近" + win + "天销量", sales60, "台", "c-cyan") +
    kbItem("汇总可售天数", sellableSummary, "天", "c-amber") +
    kbItem("库存人民币", cnyFmt(totalVnd, s.fxRate), "元", "c-red");
  p0.appendChild(el("p", "hint", "汇总可售天数 = 正常库存(" + normalStock + ") × " + win + " ÷ 近" + win + "天销量(" + sales60 + ") = " + sellableSummary + " 天"));
  main.appendChild(p0);

  /* 二、新机周转 / 三、二手机周转 */
  var p2 = el("div", "panel"); p2.appendChild(el("div", "panel-title", "二、越库手机-新机周转表"));
  p2.appendChild(buildTurnoverHtml(r.turnover.filter(function (x) { return x.condition === "新机"; }), stock, s, win));
  main.appendChild(p2);
  var p3 = el("div", "panel"); p3.appendChild(el("div", "panel-title", "三、越库手机-二手机周转表"));
  p3.appendChild(buildTurnoverHtml(r.turnover.filter(function (x) { return x.condition === "二手机"; }), stock, s, win));
  main.appendChild(p3);

  /* 五、库龄分段（按库龄日） */
  var age = { "0-30天": [0, 0], "30-60天": [0, 0], "60-90天": [0, 0], "90-180天": [0, 0], "≥180天": [0, 0] };
  stock.forEach(function (u) {
    var a = u.ageDays || 0, k;
    if (a < 30) k = "0-30天"; else if (a < 60) k = "30-60天"; else if (a < 90) k = "60-90天"; else if (a < 180) k = "90-180天"; else k = "≥180天";
    age[k][0]++; age[k][1] += u.costVnd || 0;
  });
  var p4 = el("div", "panel"); p4.appendChild(el("div", "panel-title", "五、在库库龄分段（库龄日 " + esc(r.ageDate) + "）"));
  var tblAge = el("table", "tbl bordered");
  tblAge.innerHTML = "<thead><tr><th>库龄分段</th><th>台数</th><th>金额-元</th><th>占比</th></tr></thead><tbody>" +
    Object.keys(age).map(function (k) {
      return "<tr><td>" + k + "</td><td>" + age[k][0] + "</td><td>" + cnyFmt(age[k][1], s.fxRate) + "</td><td>" + pct(age[k][0], stock.length) + "</td></tr>";
    }).join("") +
    '<tr class="total"><td>在库合计</td><td>' + stock.length + "</td><td>" + cnyFmt(totalVnd, s.fxRate) + "</td><td>100%</td></tr>" +
    '<tr class="total"><td>近' + win + "天正常出货</td><td>" + sales60 + "</td><td></td><td></td></tr></tbody>";
  p4.appendChild(tblAge);
  main.appendChild(p4);

  /* 六、系统状态台数 */
  var p5 = el("div", "panel"); p5.appendChild(el("div", "panel-title", "六、系统状态台数（截至 " + esc(r.stockDate) + "）"));
  var sysData = [
    ["📦 在库", "#06b6d4", kpi.inStock], ["✅ 已出库", "#16a34a", kpi.sold], ["🚚 在途", "#2563eb", kpi.transit],
    ["⏳ 待核", "#b45309", kpi.pending], ["⚠️ 异常机", "#b45309", kpi.anomaly], ["❗ 问题机", "#dc2626", kpi.problem]
  ];
  var maxSys = sysData.reduce(function (a, b) { return Math.max(a, b[2]); }, 1);
  var tblSys = el("table", "tbl bordered");
  tblSys.innerHTML = "<thead><tr><th>状态</th><th>台数</th><th>柱图</th></tr></thead><tbody>" +
    sysData.map(function (row) {
      var w = (row[2] / maxSys * 100).toFixed(1);
      return '<tr><td><span class="sys-badge" style="background:' + row[1] + ';color:#fff;">' + row[0] + '</span></td>' +
        "<td>" + row[2] + '</td>' +
        '<td><div class="mini-bar"><div class="mini-bar-fill" style="background:' + row[1] + ";width:" + w + '%;">' + (w > 8 ? row[2] : "") + "</div></div></td></tr>";
    }).join("") +
    '<tr class="total"><td colspan="2">库存含在途 · 新机 ' + newM.length + " / 二手机 " + usedM.length + " / 合计 " + stock.length + "</td><td></td></tr></tbody>";
  p5.appendChild(tblSys);
  main.appendChild(p5);
}
function kbItem(label, num, unit, cls) {
  return '<div class="kb-item"><div class="kb-label">' + label + '</div><div class="kb-num ' + cls + '">' + num + '</div><div class="kb-unit">' + unit + "</div></div>";
}
function buildTurnoverHtml(list, stock, s, win) {
  var unitByKey = new Map();
  stock.forEach(function (u) { var k = u.model + "|" + u.memory + "|" + u.condition; if (!unitByKey.has(k)) unitByKey.set(k, u); });
  var tot = { incl: 0, sales: 0, salesVnd: 0, stockVnd: 0 };
  var body = list.map(function (x) {
    var k = x.model + "|" + x.memory + "|" + x.condition;
    var sample = unitByKey.get(k);
    var age = sample && sample.ageDays != null ? sample.ageDays : "";
    var salesVnd = sample ? Math.round((sample.costVnd || 0) * x.sales) : 0;
    var sell = x.sellableDays == null ? "—" : x.sellableDays;
    var incl = x.stockInclTransit != null ? x.stockInclTransit : x.stock;
    tot.incl += incl; tot.sales += x.sales; tot.salesVnd += salesVnd; tot.stockVnd += x.stockValue || 0;
    var remark = x.sales > 0 ? "" : '<span class="c-amber">近' + win + "天无销量</span>";
    return "<tr><td><strong>" + esc(x.model) + "</strong></td><td>" + esc(x.memory) + "</td><td>" + incl + "</td><td>" + x.sales +
      "</td><td>" + (x.dailySales || 0).toFixed(2) + "</td><td>" + sell + "</td><td>" + nfmt(x.stockValue) + "</td><td>" + cnyFmt(x.stockValue, s.fxRate) +
      "</td><td>" + nfmt(salesVnd) + "</td><td>" + remark + "</td><td>" + age + "</td></tr>";
  }).join("");
  var totDaily = tot.sales > 0 ? tot.sales / win : 0;
  var totSell = totDaily > 0 ? Math.round(tot.incl / totDaily) : "—";
  var totalRow = '<tr class="total"><td>合计</td><td></td><td>' + tot.incl + "</td><td>" + tot.sales + "</td><td>" + totDaily.toFixed(2) + "</td><td>" + totSell +
    "</td><td>" + nfmt(tot.stockVnd) + "</td><td>" + cnyFmt(tot.stockVnd, s.fxRate) + "</td><td>" + nfmt(tot.salesVnd) + "</td><td></td><td></td></tr>";
  var tbl = el("table", "tbl bordered");
  tbl.innerHTML = "<thead><tr><th>型号</th><th>内存</th><th>库存(含在途)</th><th>近" + win + "天销量</th><th>日销量</th><th>可售天数</th><th>库存-盾</th><th>库存-元</th><th>近" + win + "天金额</th><th>备注</th><th>平均库龄</th></tr></thead>" +
    "<tbody>" + (body || '<tr><td colspan="11" class="hint center">无数据</td></tr>') + totalRow + "</tbody>";
  return tbl;
}

/* ============================================================
 * 其他仓明细子版块（需求四）
 * ============================================================ */
function renderOtherWarehouse(main) {
  var rows = state.snap.otherWarehouse || [];
  var s = state.snap.settings;
  var head = el("div", "page-head");
  head.innerHTML = "<h1>其他仓明细</h1><p>共 " + rows.length + " 行 · 当前" +
    (s.includeOtherWarehouses ? '<b class="c-green">已计入</b>总核算' : '<b class="c-amber">未计入</b>总核算（可在下方勾选）') + "</p>";
  main.appendChild(head);
  var toggle = el("div", "panel");
  toggle.innerHTML = '<label class="fld chk"><input type="checkbox" id="ow-chk"' + (s.includeOtherWarehouses ? " checked" : "") + '> <b>计入其他仓数据</b>（勾选后，Excel表格缺入库但其他仓有入库的序列号将按已入库计算）</label>';
  toggle.querySelector("#ow-chk").onchange = function (e) {
    var snap = Object.assign({}, state.snap);
    snap.settings = Object.assign({}, s, { includeOtherWarehouses: e.target.checked });
    setSnap(snap, true).then(function () { toast(e.target.checked ? "已计入其他仓数据" : "已排除其他仓数据"); });
  };
  main.appendChild(toggle);
  if (!rows.length) {
    main.appendChild(el("div", "panel", '<p class="hint center pad">尚未导入其他仓数据 → 请到「数据导入」导入</p>'));
    return;
  }
  /* 筛选：关键词 + 仓位多选 + 只看带说明 */
  if (!state.owFilter) state.owFilter = { q: "", warehouses: [], noteOnly: false, dateFrom: "", dateTo: "" };
  var owf = state.owFilter;
  var allWh = {};
  rows.forEach(function (r) { allWh[r.warehouse || "其他仓"] = 1; });
  var fpanel = el("div", "panel");
  fpanel.appendChild(el("div", "panel-title", "🔍 筛选"));
  var frow = el("div", "toolbar"); frow.style.margin = "0 0 8px";
  var fsearch = el("input", "input search");
  fsearch.placeholder = "🔍 唯一码 / 说明关键词（如：深圳）";
  fsearch.value = owf.q;
  fsearch.oninput = function () { owf.q = fsearch.value; state.page = 1; renderOtherWarehouseDebounced(); };
  frow.appendChild(fsearch);
  var noteChip = el("button", "chip" + (owf.noteOnly ? " on" : ""), "仅看带说明");
  noteChip.onclick = function () { owf.noteOnly = !owf.noteOnly; render(); };
  frow.appendChild(noteChip);
  fpanel.appendChild(frow);
  var wchips = el("div", "chips");
  Object.keys(allWh).sort().forEach(function (w) {
    var on = owf.warehouses.indexOf(w) >= 0;
    var chip = el("button", "chip" + (on ? " on" : ""), esc(w) + "（" + rows.filter(function (r) { return (r.warehouse || "其他仓") === w; }).length + "）");
    chip.onclick = function () {
      var i = owf.warehouses.indexOf(w);
      if (i >= 0) owf.warehouses.splice(i, 1); else owf.warehouses.push(w);
      render();
    };
    wchips.appendChild(chip);
  });
  fpanel.appendChild(wchips);
  main.appendChild(fpanel);
  /* 应用筛选 */
  var fq = owf.q.trim().toLowerCase();
  rows = rows.filter(function (r) {
    if (owf.warehouses.length && owf.warehouses.indexOf(r.warehouse || "其他仓") < 0) return false;
    if (owf.noteOnly && !r.note) return false;
    if (fq) {
      var hay = (r.serial + " " + (r.note || "") + " " + (r.docType || "") + " " + (r.docStatus || "") + " " + (r.ioTime || "") + " " + (r.warehouse || "")).toLowerCase();
      if (hay.indexOf(fq) < 0) return false;
    }
    return true;
  });
  if (!rows.length) {
    main.appendChild(el("div", "panel", '<p class="hint center pad">没有符合条件的其他仓记录</p>'));
    return;
  }
  /* 按仓位分组（含逐机“计入核算”按钮） */
  function owBtn(serial) {
    var ov = state.snap.overrides[C.serialKey(serial)] || {};
    var v = ov.includeOther || "";
    var label = v === "是" ? "✅ 强制计入" : v === "否" ? "🚫 已排除" : "跟随全局";
    var b = el("button", "btn btn-sm " + (v ? "btn-primary" : "btn-outline"), label);
    b.onclick = function () {
      var key = C.serialKey(serial);
      var overrides = Object.assign({}, state.snap.overrides);
      var o = Object.assign({ serial: serial }, overrides[key]);
      o.includeOther = v === "" ? "是" : v === "是" ? "否" : "";
      if (!o.includeOther) delete o.includeOther;
      var emptyAll = !o.status && !o.model && !o.memory && !o.color && !o.isProblem && !o.anomalyHandled && !o.remark && !o.includeOther;
      if (emptyAll) delete overrides[key]; else overrides[key] = o;
      setSnap(Object.assign({}, state.snap, { overrides: overrides, isSample: false }), true).then(function () { toast(serial + "：" + (o.includeOther === "是" ? "强制计入" : o.includeOther === "否" ? "已排除" : "跟随全局")); render(); });
    };
    return b;
  }
  var groups = {};
  rows.forEach(function (r) { var w = r.warehouse || "其他仓"; (groups[w] = groups[w] || []).push(r); });
  Object.keys(groups).sort().forEach(function (w) {
    var list = groups[w];
    var inn = 0, out = 0;
    list.forEach(function (r) { var cls = C.classifyWarehouse(r.docType, r.docStatus); if (cls === "in") inn++; else if (cls === "out") out++; });
    var p = el("div", "panel");
    p.appendChild(el("div", "panel-title", "🏷️ " + esc(w) + ' <span class="badge-count">' + list.length + " 行 · " + inn + " 入 / " + out + " 出</span>"));
    var tbl = el("table", "tbl bordered");
    var LIMIT = 100;
    tbl.innerHTML = "<thead><tr><th>唯一码</th><th>单据类型</th><th>单据状态</th><th>出入库时间</th><th>说明</th><th>计入核算</th></tr></thead><tbody></tbody>";
    var frag = document.createDocumentFragment();
    list.slice(0, LIMIT).forEach(function (r) {
      var tr = el("tr");
      tr.innerHTML = '<td class="mono">' + esc(r.serial) + "</td><td>" + esc(r.docType || "—") + "</td><td>" + esc(r.docStatus || "—") +
        '</td><td class="mono">' + esc(r.ioTime || "—") + "</td><td>" + esc(r.note || "—") + "</td>";
      var td = el("td"); td.appendChild(owBtn(r.serial)); tr.appendChild(td);
      frag.appendChild(tr);
    });
    tbl.querySelector("tbody").appendChild(frag);
    p.appendChild(tbl);
    if (list.length > LIMIT) p.appendChild(el("p", "hint pad", "仅显示前 " + LIMIT + " 行，共 " + list.length + " 行（用上方筛选缩小范围）"));
    main.appendChild(p);
  });
}

var _owTimer = null;
function renderOtherWarehouseDebounced() {
  clearTimeout(_owTimer);
  _owTimer = setTimeout(function () { render(); }, 300);
}

/* ============================================================
 * 台账（⑤ 双入库日/双库龄 + 只读 + 编辑按钮 + 单据汇总可展开 + 分页）
 * ============================================================ */
function collectOptions() {
  var units = state.result.units;
  function distinct(arr) { var seen = {}, out = []; arr.forEach(function (x) { if (x && !seen[x]) { seen[x] = 1; out.push(x); } }); return out.sort(); }
  return {
    status: distinct(units.map(function (u) { return u.status; })),
    condition: distinct(units.map(function (u) { return u.condition; })),
    model: distinct(units.map(function (u) { return u.model; })),
    supplier: distinct(units.map(function (u) { return u.supplier; }))
  };
}

/* ============================================================
 * 通用日期筛选器：识别当前版块的「日期型」字段，输入单日或区间（YYYY-MM-DD~YYYY-MM-DD），
 * 可勾选 含空值行；与 query/多选/分页叠加。
 * ============================================================ */
function isDateLikeLabel(name) { return false; /* 日期筛选已移除 */
  if (!name) return false;
  var SUF = ["日期","入库日","出库日","订单日","到货日","预计到货日","制表日期","下单日","时间","库龄","库天数","已出库日"];
  for (var j = 0; j < SUF.length; j++) if (name.indexOf(SUF[j]) >= 0) return true;
  return false;
}
function collectDateFieldsForUnits(units) { return []; /* 日期筛选已移除 */
  var LABELS = [
    { key: "Excel入库日", getter: function (u) { return u.erpInboundDate; } },
    { key: "仓入库日",  getter: function (u) { return u.whInboundDate; } },
    { key: "Excel库龄",   getter: function (u) { return u.erpAgeDays; } },
    { key: "仓库龄",    getter: function (u) { return u.whAgeDays; } },
    { key: "Excel订单日", getter: function (u) { return u.erpOrderDate; } }
  ];
  var out = [];
  LABELS.forEach(function (it) {
    if (!isDateLikeLabel(it.key)) return;
    var seenDate = 0, seenDays = 0;
    for (var i = 0; i < units.length; i++) {
      var v = it.getter(units[i]);
      if (v == null || v === "") continue;
      if (typeof v === "number") { seenDays++; }
      else if (typeof v === "string" && /^\d{4}-\d{2}-\d{2}/.test(v)) { seenDate++; }
      if (seenDate + seenDays > 30) break;
    }
    var kind = seenDate > 0 ? "date" : (seenDays > 0 ? "days" : "");
    if (kind) out.push({ key: it.key, label: it.key, kind: kind, getter: it.getter });
  });
  return out;
}
function applyDateFilter(rows, f) { return rows; /* 日期筛选已移除 */
  if (!f || !f.dateKey || !f.dateKey.field) return rows;
  var df = collectDateFieldsForUnits(rows).filter(function (x) { return x.key === f.dateKey.field; })[0];
  if (!df) return rows;
  var from = (f.dateKey.from || "").slice(0, 10);
  var to   = (f.dateKey.to   || "").slice(0, 10);
  var includeEmpty = !!f.dateKey.includeEmpty;
  return rows.filter(function (u) {
    var v = df.getter(u);
    if (v == null || v === "") return includeEmpty;
    if (df.kind === "days") return true;
    var vs = String(v).slice(0, 10);
    if (from && vs < from) return false;
    if (to && vs > to) return false;
    return true;
  });
}
function renderDateFilterPanel(panel, fields, f, rerender) { return; /* 日期筛选已移除 */
  panel.innerHTML = "";
  var title = document.createElement("div");
  title.className = "panel-title";
  title.innerHTML = '🔎 日期筛选 <span class="hint">选本版块的日期型字段 + 输入日期</span>';
  panel.appendChild(title);
  var bar = document.createElement("div"); bar.className = "toolbar tool-date";
  var sel = document.createElement("select"); sel.className = "input date-select";
  var def = document.createElement("option"); def.value = ""; def.textContent = "不启用日期筛选";
  sel.appendChild(def);
  fields.forEach(function (it) {
    var opt = document.createElement("option"); opt.value = it.key;
    opt.textContent = it.label + (it.kind === "days" ? "（天数）" : "（日期）");
    if (f.dateKey.field === it.key) opt.selected = true;
    sel.appendChild(opt);
  });
  bar.appendChild(sel);
  var df = document.createElement("input"); df.type = "date"; df.className = "input"; df.value = f.dateKey.from || "";
  var dt = document.createElement("input"); dt.type = "date"; dt.className = "input"; dt.value = f.dateKey.to || "";
  bar.appendChild(df); bar.appendChild(dt);
  var empty = document.createElement("label"); empty.className = "chk";
  var ecb = document.createElement("input"); ecb.type = "checkbox"; ecb.checked = f.dateKey.includeEmpty;
  empty.appendChild(ecb); empty.appendChild(document.createTextNode("含空值行"));
  bar.appendChild(empty);
  var clr = document.createElement("button"); clr.className = "btn btn-sm btn-outline"; clr.textContent = "✖ 清";
  bar.appendChild(clr);
  panel.appendChild(bar);
  clr.onclick = function () { f.dateKey = { field: "", from: "", to: "", includeEmpty: false }; df.value=""; dt.value=""; ecb.checked=false; sel.value=""; rerender(); };
  df.onchange = function () { f.dateKey.from = df.value; rerender(); };
  dt.onchange = function () { f.dateKey.to = dt.value; rerender(); };
  ecb.onchange = function () { f.dateKey.includeEmpty = ecb.checked; rerender(); };
  sel.onchange = function () { f.dateKey.field = sel.value; rerender(); };
}

var FILTER_COLUMNS = [
  { key: "model", label: "机型", get: function (u) { return u.model; } },
  { key: "memory", label: "内存", get: function (u) { return u.memory; } },
  { key: "color", label: "颜色", get: function (u) { return u.color; } },
  { key: "condition", label: "新旧", get: function (u) { return u.condition; } },
  { key: "status", label: "状态", get: function (u) { return u.status; } },
  { key: "supplier", label: "供应商", get: function (u) { return u.supplier; } },
  { key: "hold", label: "领用/维修", get: function (u) { return u.companyHoldReason; } },
  { key: "problem", label: "问题机", get: function (u) { return u.isProblem ? "是" : "否"; } },
  { key: "anomaly", label: "异常", get: function (u) { return u.isAnomaly ? "是" : "否"; } },
  { key: "remark", label: "备注", get: function (u) { return u.remark; } }
];
function filterBaseForRoute(u) {
  if (state.route === "stock") return u.status === "在库";
  if (state.route === "sold") return u.status === "已出库";
  if (state.route === "transit") return u.status === "在途";
  if (state.route === "anomalies") return u.isAnomaly;
  if (state.route === "issues") return u.isProblem;
  return true;
}
function quickKind(u) {
  var reason = String(u.companyHoldReason || "");
  if (/部门领用/.test(reason)) return "领用";
  if (/维修|售后|问题仓|瑕疵仓/.test(reason) || /深圳|售后|维修/.test(u.problemReason || "")) return "维修售后";
  return "在仓可用";
}
function resetLedgerFilters() {
  state.filters = { status: [], condition: [], model: [], anomaly: "all", problem: "all", age: "all", supplier: "", dateFrom: "", dateTo: "" };
  state.columnFilters = {}; state.quickFilter = ""; state.query = ""; state.page = 1;
}
function applyFilters(rows) {
  var q = state.query.trim().toLowerCase(), f = state.filters;
  return rows.filter(function (u) {
    if (!filterBaseForRoute(u)) return false;
    if (state.quickFilter && quickKind(u) !== state.quickFilter) return false;
    var cols = state.columnFilters || {};
    for (var ci = 0; ci < FILTER_COLUMNS.length; ci++) {
      var def = FILTER_COLUMNS[ci], selected = cols[def.key];
      if (selected && selected.length && selected.indexOf(String(def.get(u) || "（空白）")) < 0) return false;
    }
    if (f.status.length && f.status.indexOf(u.status) < 0) return false;
    if (f.condition.length && f.condition.indexOf(u.condition) < 0) return false;
    if (f.model.length && f.model.indexOf(u.model) < 0) return false;
    if (f.anomaly === "yes" && !u.isAnomaly) return false;
    if (f.anomaly === "no" && u.isAnomaly) return false;
    if (f.problem === "yes" && !u.isProblem) return false;
    if (f.problem === "no" && u.isProblem) return false;
    if (f.age === "young" && (u.ageDays || 0) >= 30) return false;
    if (f.age === "mid" && ((u.ageDays || 0) < 30 || (u.ageDays || 0) >= 90)) return false;
    if (f.age === "old" && (u.ageDays || 0) < 90) return false;
    if (f.supplier && (u.supplier || "") !== f.supplier) return false;
    var dk = state.route === "transit" ? (u.transitDate || "") : (u.erpInboundDate || u.whInboundDate || "");
    if (f.costMin !== "" && f.costMin != null && (u.costVnd == null || u.costVnd < Number(f.costMin))) return false;
    if (f.costMax !== "" && f.costMax != null && (u.costVnd == null || u.costVnd > Number(f.costMax))) return false;
    if (dk && (f.dateFrom || f.dateTo)) {
      if (f.dateFrom && dk < f.dateFrom) return false;
      if (f.dateTo && dk > f.dateTo) return false;
    } else if (!dk && (f.dateFrom || f.dateTo)) return false;
    if (q) {
      var hay = [u.serial, u.model, u.memory, u.color, u.condition, u.status, u.supplier,
        u.erpInboundDate, u.erpAgeDays, u.whInboundDate, u.whAgeDays, u.transitDate, u.transitAgeDays,
        u.costVnd, u.problemReason, u.anomalyReason, u.remark, u.companyHoldReason]
        .map(function (x) { return x == null ? "" : String(x); }).join(" ").toLowerCase();
      if (hay.indexOf(q) < 0) return false;
    }
    return true;
  });
}
function badge(st) {
  var cls = st === "在库" ? "b-cyan" : st === "已出库" ? "b-green" : st === "待核" ? "b-amber" : st === "在途" ? "b-blue" : st === "不计入" ? "b-gray" : "";
  return '<span class="badge ' + cls + '">' + esc(st) + "</span>";
}
function currentRows() { return applyFilters(state.result.units); }

function openColumnFilter(key, anchor) {
  var old = document.querySelector(".column-popover"); if (old) old.remove();
  var def = FILTER_COLUMNS.find(function (d) { return d.key === key; });
  var counts = new Map();
  state.result.units.filter(filterBaseForRoute).forEach(function (u) {
    var v = String(def.get(u) || "（空白）"); counts.set(v, (counts.get(v) || 0) + 1);
  });
  var values = Array.from(counts.keys()).sort(function (a,b) { return a.localeCompare(b, "zh-CN", { numeric:true }); });
  var selected = (state.columnFilters[key] || []).slice();
  var pop = el("div", "column-popover");
  var rect = anchor.getBoundingClientRect();
  pop.style.left = Math.max(8, Math.min(rect.left, window.innerWidth - 295)) + "px";
  pop.style.top = Math.min(rect.bottom + 5, window.innerHeight - 380) + "px";
  pop.innerHTML = '<div class="popover-title">' + esc(def.label) + '筛选 <button class="popover-close" type="button">×</button></div>' +
    '<input class="input pop-search" placeholder="搜索选项">' +
    '<div class="popover-actions"><button type="button" data-act="all">全选</button><button type="button" data-act="none">清空</button></div>' +
    '<div class="popover-list"></div><div class="popover-foot"><button class="btn btn-sm btn-outline" data-act="clear">清除此列</button><button class="btn btn-sm btn-primary" data-act="apply">应用</button></div>';
  document.body.appendChild(pop);
  var list = pop.querySelector(".popover-list"), query = pop.querySelector(".pop-search");
  function draw() {
    var q = query.value.trim().toLowerCase(); list.innerHTML = "";
    values.filter(function (v) { return v.toLowerCase().indexOf(q) >= 0; }).forEach(function (v) {
      var row = el("label", "popover-option");
      var cb = document.createElement("input"); cb.type = "checkbox"; cb.checked = !selected.length || selected.indexOf(v) >= 0;
      cb.onchange = function () { if (cb.checked) { if (selected.indexOf(v) < 0) selected.push(v); } else { if (!selected.length) selected = values.slice(); selected = selected.filter(function (x) { return x !== v; }); } };
      row.appendChild(cb); row.appendChild(el("span", "", esc(v))); row.appendChild(el("small", "", counts.get(v))); list.appendChild(row);
    });
  }
  query.oninput = draw; draw(); query.focus();
  pop.onclick = function (e) {
    var act = e.target.getAttribute("data-act");
    if (e.target.closest(".popover-close")) { pop.remove(); return; }
    if (act === "all") { selected = []; draw(); }
    if (act === "none") { selected = ["__NO_MATCH__"]; draw(); }
    if (act === "clear") { delete state.columnFilters[key]; pop.remove(); state.page = 1; render(); }
    if (act === "apply") { if (selected.length) state.columnFilters[key] = selected; else delete state.columnFilters[key]; pop.remove(); state.page = 1; render(); }
  };
  setTimeout(function () { document.addEventListener("pointerdown", function outside(e) { if (!pop.contains(e.target) && e.target !== anchor) { pop.remove(); document.removeEventListener("pointerdown", outside); } }); }, 0);
}

function renderLedger(main) {
  var TITLES = { ledger: "台账", stock: "在库", sold: "已出库", transit: "在途", anomalies: "异常机", issues: "问题机" };
  var title = TITLES[state.route] || "台账";
  var isTransitRoute = state.route === "transit";
  /* 流程汇总缓存：同一份数据不重复构建 */
  if (!state._flowCache || state._flowCache.snap !== state.snap) {
    state._flowCache = { snap: state.snap, flow: buildFlowSummaries(state.snap) };
  }
  var flow = state._flowCache.flow;
  var head = el("div", "page-head row");
  head.innerHTML = "<div><h1>" + esc(title) + "</h1>" +
    '<p class="hint">默认只读 · 点击 <b>✏️ 编辑</b> 进入编辑 · 「Excel入库/出库记录 / 仓单据」点击可展开流水明细（含仓位）</p></div>';
  var exp = el("button", "btn btn-outline", "📤 导出本表");
  exp.onclick = function () { XIO.downloadBlob(XIO.exportUnits(currentRows(), title), title + "_" + C.todayIso() + ".xlsx"); };
  head.appendChild(exp);
  main.appendChild(head);

  if (state.route === "transit") {
    var curMeta = (state.snap.importMeta && state.snap.importMeta.transit) || null;
    var tp = el("div", "panel");
    tp.innerHTML = '<div class="panel-title">🚚 在途数据来源</div>' +
      '<p class="hint">当前生效：' + (curMeta ? esc(curMeta.fileName) + " · " + curMeta.rowCount + " 行 · " + esc(curMeta.importedAt || "") : "未独立上传在途：使用三合一文件中的在途表") + "</p>" +
      '<p class="hint">在途日期 / 预计到货天数为<b>可选字段</b>，可不填、不影响入库判定，缺值时不会报错也不会阻塞。</p>';
    main.appendChild(tp);
  }

  var toolbar = el("div", "toolbar");
  var search = el("input", "input search");
  search.placeholder = "🔍 序列号 / 机型 / 内存 / 颜色 / 供应商";
  search.value = state.query;
  var _searchTimer = null;
  search.oninput = function () {
    state.query = search.value; state.page = 1;
    clearTimeout(_searchTimer);
    _searchTimer = setTimeout(rerenderTable, 250); /* 防抖，输入时不卡顿 */
  };
  toolbar.appendChild(search);
  var filterBtn = el("button", "btn btn-outline", "🎚️ 筛选");
  filterBtn.onclick = function () { state.drawerMode = "filter"; renderDrawer(); };
  toolbar.appendChild(filterBtn);
  ["supplier", "hold"].forEach(function (key) {
    var def = FILTER_COLUMNS.find(function (x) { return x.key === key; });
    var b = el("button", "btn btn-outline", def.label + " ▾");
    b.onclick = function () { openColumnFilter(key, b); }; toolbar.appendChild(b);
  });
  var pageLabel = el("label", "page-size");
  var ps = el("select", "input page-select");
  [10, 25, 50, 100, 99999].forEach(function (n) {
    var opt = el("option"); opt.value = n; opt.textContent = (n === 99999 ? "全部" : "每页 " + n);
    if (n === state.pageSize) opt.selected = true;
    ps.appendChild(opt);
  });
  ps.onchange = function () { state.pageSize = Number(ps.value); state.page = 1; rerenderTable(); };
  pageLabel.appendChild(ps);
  toolbar.appendChild(pageLabel);
  var saveFilter = el("button", "btn btn-outline", "保存筛选");
  saveFilter.onclick = function () {
    var name = prompt("给当前筛选起个名字"); if (!name || !name.trim()) return;
    var all = JSON.parse(localStorage.getItem("ims-saved-filters") || "{}");
    all[name.trim()] = { filters: state.filters, columns: state.columnFilters, quick: state.quickFilter, query: state.query };
    localStorage.setItem("ims-saved-filters", JSON.stringify(all)); toast("已保存筛选：" + name.trim()); render();
  };
  toolbar.appendChild(saveFilter);
  var saved = JSON.parse(localStorage.getItem("ims-saved-filters") || "{}");
  var savedNames = Object.keys(saved);
  if (savedNames.length) {
    var select = el("select", "input page-select"); select.innerHTML = '<option value="">使用已保存筛选</option>' + savedNames.map(function (n) { return '<option value="' + esc(n) + '">' + esc(n) + '</option>'; }).join("");
    select.onchange = function () { var v = saved[select.value]; if (!v) return; state.filters = Object.assign({}, v.filters); state.columnFilters = Object.assign({}, v.columns); state.quickFilter = v.quick || ""; state.query = v.query || ""; state.page = 1; render(); };
    toolbar.appendChild(select);
    var delSaved = el("button", "btn btn-sm btn-outline", "删除筛选");
    delSaved.onclick = function () { if (!select.value) return toast("先选一个已保存筛选", "warn"); delete saved[select.value]; localStorage.setItem("ims-saved-filters", JSON.stringify(saved)); render(); };
    toolbar.appendChild(delSaved);
  }
  main.appendChild(toolbar);
  var activeBar = el("div", "active-filters"); main.appendChild(activeBar);
  var quickBar = el("div", "quick-filters"); main.appendChild(quickBar);
  if (state.route === "stock") {
    ["", "在仓可用", "领用", "维修售后"].forEach(function (name) {
      var b = el("button", "btn btn-sm" + (state.quickFilter === name ? " btn-primary" : " btn-outline"), name || "全部在库");
      b.onclick = function () { state.quickFilter = name; state.page = 1; render(); }; quickBar.appendChild(b);
    });
  }
  var wrap = el("div", "panel-table table-wrap");
  main.appendChild(wrap);

  function rerenderTable() {
    var all = applyFilters(state.result.units);
    var total = all.length;
    var pageSize = state.pageSize;
    var totalPages = pageSize >= 99999 ? 1 : Math.max(1, Math.ceil(total / pageSize));
    if (state.page > totalPages) state.page = totalPages;
    var start = (state.page - 1) * pageSize;
    var rows = pageSize >= 99999 ? all : all.slice(start, start + pageSize);
    wrap.innerHTML = "";
    activeBar.innerHTML = "";
    var active = [];
    if (state.query) active.push("搜索：" + state.query);
    if (state.quickFilter) active.push("分类：" + state.quickFilter);
    Object.keys(state.columnFilters || {}).forEach(function (key) { if (state.columnFilters[key].length) active.push((FILTER_COLUMNS.find(function (d) { return d.key === key; }) || {}).label + "：" + state.columnFilters[key].join("、")); });
    if (state.filters.model.length) active.push("机型：" + state.filters.model.join("、"));
    if (state.filters.condition.length) active.push("新旧：" + state.filters.condition.join("、"));
    if (state.filters.supplier) active.push("供应商：" + state.filters.supplier);
    if (state.filters.problem !== "all") active.push("问题机：" + state.filters.problem);
    if (state.filters.anomaly !== "all") active.push("异常：" + state.filters.anomaly);
    if (state.filters.age !== "all") active.push("库龄：" + state.filters.age);
    if (state.filters.dateFrom || state.filters.dateTo) active.push("日期：" + (state.filters.dateFrom || "起") + "～" + (state.filters.dateTo || "止"));
    if (state.filters.costMin || state.filters.costMax) active.push("成本：" + (state.filters.costMin || "0") + "～" + (state.filters.costMax || "不限"));
    activeBar.appendChild(el("span", "hint", "筛选结果 " + total + " / " + state.result.units.filter(filterBaseForRoute).length + " 台"));
    active.forEach(function (label) { activeBar.appendChild(el("span", "filter-tag", esc(label))); });
    if (active.length) {
      var reset = el("button", "btn btn-sm btn-outline", "清除筛选");
      reset.onclick = function () { resetLedgerFilters(); render(); };
      activeBar.appendChild(reset);
    }
    var info = el("div", "table-info");
    info.innerHTML = "<span>共 <b>" + total + "</b> 台</span><span>· 当前显示 <b>" + rows.length + "</b> 台</span>" +
      (pageSize < 99999 ? "<span>· 第 <b>" + state.page + "</b> / " + totalPages + " 页</span>" : "");
    wrap.appendChild(info);

    /* 在途日期/在途库龄只在「在途」版块显示；其他版块显示仓入库日/仓库龄 */
    var dateHead = isTransitRoute ? "<th>在途日期</th><th>预计到货天数</th>" : "<th>仓入库日</th><th>仓库龄</th>";
    var tbl = el("table", "tbl bordered");
    tbl.innerHTML = "<thead><tr><th>操作</th><th>序列号</th><th>机型</th><th>内存</th><th>颜色</th><th>新旧</th><th>状态</th><th>Excel入库/出库记录</th><th>仓单据</th><th>问题机</th><th>异常</th><th>成本VND</th><th>Excel入库日</th><th>Excel库龄</th>" + dateHead + "<th>备注</th></tr></thead><tbody>" +
      rows.map(function (u) {
        var marks = "";
        if (u.isAnomaly) marks += '<span class="c-amber">⚠ ' + esc(u.anomalyReason) + "</span>";
        if (u.isProblem) marks += ' <span class="c-red">❗问题</span>';
        return "<tr>" +
          '<td><button class="btn btn-sm btn-primary" data-edit-serial="' + esc(u.serial) + '">✏️ 编辑</button></td>' +
          '<td class="mono">' + esc(u.serial) + "</td>" +
          "<td>" + esc(u.model) + "</td>" +
          "<td>" + esc(u.memory) + "</td>" +
          "<td>" + esc(u.color || "—") + "</td>" +
          "<td>" + esc(u.condition) + "</td>" +
          "<td>" + badge(u.status) + "</td>" +
          '<td><button class="link-btn" data-flow-serial="' + esc(u.serial) + '" data-flow-type="erp">' + esc(erpSummary(u, flow)) + "</button></td>" +
          '<td><button class="link-btn" data-flow-serial="' + esc(u.serial) + '" data-flow-type="wh">' + esc(whSummary(u, flow)) + "</button></td>" +
          "<td>" + (u.isProblem ? '<span class="c-red">是</span>' : "—") + "</td>" +
          "<td>" + (marks || "—") + "</td>" +
          '<td class="num">' + nfmt(u.costVnd) + "</td>" +
          '<td class="mono">' + esc(u.erpInboundDate || "—") + "</td>" +
          "<td>" + (u.erpAgeDays == null ? "—" : u.erpAgeDays + " 天") + "</td>" +
          (isTransitRoute
            ? '<td class="mono">' + (u.transitDate ? esc(u.transitDate) : '<span class="muted">未填·可选</span>') + "</td><td>" + (u.transitAgeDays == null ? '<span class="muted">未填·可选</span>' : u.transitAgeDays + " 天") + "</td>"
            : '<td class="mono">' + esc(u.whInboundDate || "—") + "</td><td>" + (u.whAgeDays == null ? "—" : u.whAgeDays + " 天") + "</td>") +
          "<td>" + (u.companyHoldReason ? '<span class="badge b-amber">' + esc(quickKind(u)) + "</span> " : "") + (u.remark ? esc(u.remark) : "—") + "</td>" +
          "</tr>";
      }).join("") + "</tbody>";
    wrap.appendChild(tbl);
    var headKeys = { "机型":"model", "内存":"memory", "颜色":"color", "新旧":"condition", "状态":"status", "问题机":"problem", "异常":"anomaly", "备注":"remark" };
    tbl.querySelectorAll("thead th").forEach(function (th) {
      var key = headKeys[th.textContent.trim()];
      if (!key) return;
      var btn = el("button", "column-filter" + ((state.columnFilters[key] || []).length ? " selected" : ""), "▾");
      btn.title = "筛选" + th.textContent;
      btn.onclick = function (e) { e.stopPropagation(); openColumnFilter(key, btn); };
      th.appendChild(btn);
    });
    if (!rows.length) wrap.appendChild(el("p", "hint center pad", "没有符合条件的机器"));

    if (pageSize < 99999 && totalPages > 1) {
      var pager = el("div", "pager");
      var mkBtn = function (label, page, disabled, active) {
        var b = el("button", "btn btn-sm" + (active ? " btn-primary" : ""), label);
        if (disabled) b.disabled = true;
        else b.onclick = function () { state.page = page; rerenderTable(); };
        return b;
      };
      pager.appendChild(mkBtn("«", 1, state.page === 1));
      pager.appendChild(mkBtn("‹", state.page - 1, state.page === 1));
      var from = Math.max(1, state.page - 2);
      var to = Math.min(totalPages, from + 4);
      if (to - from < 4) from = Math.max(1, to - 4);
      for (var p = from; p <= to; p++) pager.appendChild(mkBtn("" + p, p, false, p === state.page));
      pager.appendChild(mkBtn("›", state.page + 1, state.page === totalPages));
      pager.appendChild(mkBtn("»", totalPages, state.page === totalPages));
      wrap.appendChild(pager);
    }

    if (!document.body._flowBound) {
      document.body._flowBound = true;
      document.addEventListener("click", function (e) {
        var t = e.target.closest("[data-flow-serial]");
        if (t) {
          var u = state.result.units.find(function (x) { return x.serial === t.getAttribute("data-flow-serial"); });
          if (!u) return;
          state.flowUnit = u; state.flowType = t.getAttribute("data-flow-type") || "erp"; state.drawerMode = "flow"; renderDrawer();
        }
        var ed = e.target.closest("[data-edit-serial]");
        if (ed) {
          var u2 = state.result.units.find(function (x) { return x.serial === ed.getAttribute("data-edit-serial"); });
          if (!u2) return;
          state.activeUnit = u2; state.editing = false; state.drawerMode = "unit"; renderDrawer();
        }
      });
    }
  }
  rerenderTable();

  /* 异常机页：已处理异常找回区 */
  if (state.route === "anomalies") {
    var handled = state.result.units.filter(function (u) { return u.anomalyHandled && u.anomalyReason; });
    var hp = el("div", "panel");
    hp.appendChild(el("div", "panel-title", '已处理异常（可找回） <span class="badge-count">' + handled.length + "</span>"));
    if (!handled.length) hp.appendChild(el("p", "hint", "暂无已处理的异常记录"));
    handled.slice(0, 300).forEach(function (u) {
      var row2 = el("div", "kv-row");
      row2.innerHTML = '<div class="kv-label mono">' + esc(u.serial) + " · " + esc(u.model) + " · " + esc(u.anomalyReason) + "</div>";
      var btn = el("button", "btn btn-sm btn-outline", "↩ 找回异常");
      btn.onclick = function () {
        var key = C.serialKey(u.serial);
        var overrides = Object.assign({}, state.snap.overrides);
        var ov2 = Object.assign({}, overrides[key] || { serial: u.serial });
        delete ov2.anomalyHandled;
        var emptyAll = !ov2.status && !ov2.model && !ov2.memory && !ov2.color && !ov2.isProblem && !ov2.remark && !ov2.includeOther;
        if (emptyAll) delete overrides[key]; else overrides[key] = ov2;
        setSnap(Object.assign({}, state.snap, { overrides: overrides, isSample: false })).then(function () { toast("已找回异常：" + u.serial); });
      };
      row2.appendChild(btn);
      hp.appendChild(row2);
    });
    main.appendChild(hp);
  }
}

/* ============================================================
 * 抽屉：unit（只读/编辑） / flow（单据明细含仓位） / filter（多选筛选）
 * ============================================================ */
function renderDrawer() {
  var old = $("#drawer-mask");
  if (old) old.remove();
  if (state.drawerMode === "filter") return renderFilterDrawer();
  if (state.drawerMode === "flow" && state.flowUnit) return renderFlowDrawer();
  if (state.drawerMode === "unit" && state.activeUnit) return renderUnitDrawer();
}
function kvRow(label, val) { return '<div class="kv-row"><div class="kv-label">' + label + '</div><div class="kv-value">' + val + "</div></div>"; }

function renderUnitDrawer() {
  var u = state.activeUnit;
  var ov = state.snap.overrides[C.serialKey(u.serial)] || {};
  var mask = el("div", "drawer-mask"); mask.id = "drawer-mask";
  var box = el("div", "drawer"); mask.appendChild(box);
  mask.onclick = function (e) { if (e.target === mask) { state.activeUnit = null; state.drawerMode = null; state.editing = false; renderDrawer(); } };
  document.body.appendChild(mask);

  function readOnlyView() {
    box.innerHTML =
      '<div class="drawer-head"><div><div class="hint">序列号 (IMEI)</div><div class="mono big">' + esc(u.serial) + "</div>" +
      '<div class="hint">' + esc(u.model) + " · " + esc(u.memory) + " · " + esc(u.color || "") + " · " + esc(u.condition) + " " + badge(u.status) + "</div></div>" +
      '<button class="btn btn-ghost" id="d-close">✕</button></div>' +
      '<div class="kv">' +
      kvRow("Excel 表格记录", u.inCount + " 入 / " + u.outCount + " 出") +
      kvRow("仓库 单据", u.whInCount + " 入 / " + u.whOutCount + " 出") +
      kvRow("成本 VND", nfmt(u.costVnd)) +
      kvRow("Excel 入库日", esc(u.erpInboundDate || "—")) +
      kvRow("Excel 库龄", u.erpAgeDays == null ? "—" : u.erpAgeDays + " 天") +
      kvRow("仓入库日", esc(u.whInboundDate || "—")) +
      kvRow("仓库龄", u.whAgeDays == null ? "—" : u.whAgeDays + " 天") +
      kvRow("供应商", esc(u.supplier || "—")) +
      kvRow("备注", esc(u.remark || "—")) +
      (u.includeOther ? kvRow("其他仓记录", u.includeOther === "是" ? "强制计入" : "已排除") : "") +
      kvRow("问题机", u.isProblem ? '<span class="c-red">是 · ' + esc(u.problemReason || "") + "</span>" : "否") +
      kvRow("异常", u.isAnomaly ? '<span class="c-amber">' + esc(u.anomalyReason) + "</span>" : (u.anomalyHandled ? '<span class="hint">' + esc(u.anomalyReason) + " · 已处理</span>" : "—")) +
      "</div>" +
      (u.isAnomaly ? '<div class="alert warn">⚠️ 异常说明：' + esc(u.anomalyReason) + "</div>" : "") +
      (u.isProblem ? '<div class="alert bad">❗ 问题机：入库原因 ≠ "出租"</div>' : "") +
      '<div class="drawer-actions">' +
      '<button class="btn btn-primary" id="d-edit">✏️ 编辑此机器</button>' +
      (u.isAnomaly ? '<button class="btn" id="d-resolve">✓ 处理异常（写入台账）</button>' : "") +
      (u.anomalyHandled && u.anomalyReason ? '<button class="btn btn-outline" id="d-restore">↩ 找回异常（重新标记）</button>' : "") +
      "</div>";
  }
  function editView() {
    box.innerHTML =
      '<div class="drawer-head"><div><div class="hint">正在编辑 · 序列号</div><div class="mono big">' + esc(u.serial) + "</div></div>" +
      '<button class="btn btn-ghost" id="d-close">✕</button></div>' +
      '<label class="fld">状态<select id="d-status" class="input">' +
      '<option value="">（跟系统）</option>' +
      ["在库","在途","已出库","待核","不计入"].map(function (x) { return '<option value="' + x + '"' + (ov.status === x ? " selected" : "") + ">" + x + "</option>"; }).join("") +
      "</select></label>" +
      '<label class="fld">新旧<select id="d-condition" class="input"><option value="">（跟导入数据）</option><option value="新机"' + (ov.condition === "新机" ? " selected" : "") + '>新机</option><option value="二手机"' + (ov.condition === "二手机" ? " selected" : "") + '>二手机</option></select></label>' +
      '<label class="fld">机型<input id="d-model" class="input" value="' + esc(ov.model || "") + '" placeholder="' + esc(u.model) + '"></label>' +
      '<label class="fld">内存<input id="d-memory" class="input" value="' + esc(ov.memory || "") + '" placeholder="' + esc(u.memory) + '"></label>' +
      '<label class="fld">颜色<input id="d-color" class="input" value="' + esc(ov.color || "") + '" placeholder="' + esc(u.color || "") + '"></label>' +
      '<label class="fld">问题机<select id="d-problem" class="input">' +
      '<option value="">（跟系统）</option>' +
      '<option value="是"' + (ov.isProblem === "是" ? " selected" : "") + ">是</option>" +
      '<option value="否"' + (ov.isProblem === "否" ? " selected" : "") + ">否</option>" +
      "</select></label>" +
      '<label class="fld">备注（手动输入，所有列表都显示）<textarea id="d-remark" class="input" rows="2">' + esc(ov.remark || "") + "</textarea></label>" +
      '<label class="fld">该机的其他仓记录<select id="d-ow" class="input">' +
      '<option value="">（跟随全局开关）</option>' +
      '<option value="是"' + (ov.includeOther === "是" ? " selected" : "") + '>强制计入该机</option>' +
      '<option value="否"' + (ov.includeOther === "否" ? " selected" : "") + '>排除该机</option>' +
      "</select></label>" +
      '<p class="hint" style="margin:14px 0 2px">修改保存后将同步台账、工作台、列表。</p>' +
      '<div class="drawer-actions">' +
      '<button class="btn btn-primary" id="d-save">💾 保存修改</button>' +
      '<button class="btn" id="d-cancel">取消</button>' +
      (ov.serial ? '<button class="btn btn-ghost" id="d-clear">恢复自动判定</button>' : "") +
      "</div>";
  }
  if (state.editing) editView(); else readOnlyView();

  $("#d-close").onclick = function () { state.activeUnit = null; state.drawerMode = null; state.editing = false; renderDrawer(); };
  var edit = $("#d-edit"); if (edit) edit.onclick = function () { state.editing = true; renderDrawer(); };
  var cancel = $("#d-cancel"); if (cancel) cancel.onclick = function () { state.editing = false; renderDrawer(); };
  var save = $("#d-save");
  if (save) save.onclick = function () {
    var patch = { status: $("#d-status").value, condition: $("#d-condition").value, model: $("#d-model").value.trim(), memory: $("#d-memory").value.trim(), color: $("#d-color").value.trim(), isProblem: $("#d-problem").value, anomalyHandled: ov.anomalyHandled, remark: $("#d-remark").value.trim(), includeOther: $("#d-ow").value };
    var key = C.serialKey(u.serial);
    var overrides = Object.assign({}, state.snap.overrides);
    var next = Object.assign({ serial: u.serial }, ov, patch, { serial: u.serial });
    var emptyAll = !next.status && !next.condition && !next.model && !next.memory && !next.color && !next.isProblem && !next.anomalyHandled && !next.remark && !next.includeOther;
    if (emptyAll) delete overrides[key]; else overrides[key] = next;
    var snap = Object.assign({}, state.snap, { overrides: overrides, isSample: false });
    state.activeUnit = null; state.editing = false; state.drawerMode = null;
    setSnap(snap).then(function () { toast("已保存并同步到台账 / 工作台 / 各列表"); });
  };
  var resolve = $("#d-resolve");
  if (resolve) resolve.onclick = function () {
    var key = C.serialKey(u.serial);
    var overrides = Object.assign({}, state.snap.overrides);
    overrides[key] = Object.assign({ serial: u.serial }, ov, { anomalyHandled: true, status: ov.status || "已出库" }, { serial: u.serial });
    var snap = Object.assign({}, state.snap, { overrides: overrides, isSample: false });
    state.activeUnit = null; state.editing = false; state.drawerMode = null;
    setSnap(snap).then(function () { toast("异常已处理 · 写入台账"); });
  };
  var restore = $("#d-restore");
  if (restore) restore.onclick = function () {
    var key = C.serialKey(u.serial);
    var overrides = Object.assign({}, state.snap.overrides);
    var ov2 = Object.assign({}, overrides[key] || { serial: u.serial });
    delete ov2.anomalyHandled;
    var emptyAll2 = !ov2.status && !ov2.model && !ov2.memory && !ov2.color && !ov2.isProblem && !ov2.remark && !ov2.includeOther;
    if (emptyAll2) delete overrides[key]; else overrides[key] = ov2;
    var snap2 = Object.assign({}, state.snap, { overrides: overrides, isSample: false });
    state.activeUnit = null; state.editing = false; state.drawerMode = null;
    setSnap(snap2).then(function () { toast("已找回异常：" + u.serial); });
  };
  var clear = $("#d-clear");
  if (clear) clear.onclick = function () {
    var key = C.serialKey(u.serial);
    var overrides = Object.assign({}, state.snap.overrides);
    delete overrides[key];
    var snap = Object.assign({}, state.snap, { overrides: overrides, isSample: false });
    state.activeUnit = null; state.editing = false; state.drawerMode = null;
    setSnap(snap).then(function () { toast("已恢复自动判定"); });
  };
}

function renderFlowDrawer() {
  var u = state.flowUnit;
  var flow = buildFlowSummaries(state.snap);
  var isWh = state.flowType === "wh";
  var list = (isWh ? flow.whByKey : flow.erpByKey).get(C.serialKey(u.serial)) || [];
  if (isWh) list = list._active || list;
  var mask = el("div", "drawer-mask"); mask.id = "drawer-mask";
  var box = el("div", "drawer"); mask.appendChild(box);
  mask.onclick = function (e) { if (e.target === mask) { state.flowUnit = null; state.drawerMode = null; renderDrawer(); } };
  document.body.appendChild(mask);
  box.innerHTML =
    '<div class="drawer-head"><div><div class="hint">' + esc(u.serial) + " · " + (isWh ? "仓库单据明细（含仓位）" : "Excel表格/系统 单据明细") + "</div>" +
    '<div class="mono big">' + esc(u.model) + " · " + esc(u.memory) + "</div></div>" +
    '<button class="btn btn-ghost" id="f-close">✕</button></div>' +
    (list.length ?
      ('<table class="tbl bordered"><thead><tr>' +
        (isWh ? "<th>仓位</th><th>单据类型</th><th>状态</th><th>时间</th><th>说明</th>" : "<th>来源</th><th>类别</th><th>日期</th><th>供应商</th><th>原因</th><th>成本</th>") +
        "</tr></thead><tbody>" +
        list.map(function (x) {
          if (isWh) return "<tr><td><b>" + esc(x.warehouse || "—") + "</b></td><td>" + esc(x.type) + "</td><td>" + esc(x.status) + '</td><td class="mono">' + esc(x.date) + "</td><td>" + esc(x.note || "—") + "</td></tr>";
          return "<tr><td>" + esc(x.src) + "</td><td>" + esc(x.type) + '</td><td class="mono">' + esc(x.date) + "</td><td>" + esc(x.supplier || "—") + "</td><td>" + esc(x.reason || "—") + '</td><td class="num">' + nfmt(x.cost) + "</td></tr>";
        }).join("") + "</tbody></table>") :
      '<p class="hint center pad">此机器在 ' + (isWh ? "仓库" : "Excel表格/系统") + " 里没有记录</p>") +
    '<div class="drawer-actions"><button class="btn" id="f-switch">查看 ' + (isWh ? "Excel表格/系统" : "仓库") + " 单据</button>" +
    '<button class="btn btn-primary" id="f-edit">✏️ 编辑此机器</button></div>';
  $("#f-close").onclick = function () { state.flowUnit = null; state.drawerMode = null; renderDrawer(); };
  $("#f-switch").onclick = function () { state.flowType = isWh ? "erp" : "wh"; renderDrawer(); };
  $("#f-edit").onclick = function () { state.activeUnit = u; state.drawerMode = "unit"; state.editing = false; state.flowUnit = null; renderDrawer(); };
}

function renderFilterDrawer() {
  var mask = el("div", "drawer-mask"); mask.id = "drawer-mask";
  var box = el("div", "drawer drawer-filter"); mask.appendChild(box);
  mask.onclick = function (e) { if (e.target === mask) { state.drawerMode = null; renderDrawer(); } };
  document.body.appendChild(mask);
  var opts = collectOptions();
  var f = state.filters;
  function multiChips(label, list, selected) {
    return '<div class="fld"><div class="fld-label">' + label + '</div><div class="chips">' +
      list.map(function (v) {
        var on = selected.indexOf(v) >= 0;
        return '<label class="chip' + (on ? " on" : "") + '"><input type="checkbox" data-multi="' + label + '" value="' + esc(v) + '"' + (on ? " checked" : "") + "><span>" + esc(v) + "</span></label>";
      }).join("") + "</div></div>";
  }
  function radioChips(label, key, list) {
    return '<div class="fld"><div class="fld-label">' + label + '</div><div class="chips">' +
      list.map(function (p) {
        return '<label class="chip' + (f[key] === p.val ? " on" : "") + '"><input type="radio" name="rd-' + key + '" value="' + p.val + '"' + (f[key] === p.val ? " checked" : "") + "><span>" + p.label + "</span></label>";
      }).join("") + "</div></div>";
  }
  box.innerHTML =
    '<div class="drawer-head"><div><h2>🎚️ 筛选</h2><div class="hint">可多选；点击「应用」生效。</div></div>' +
    '<button class="btn btn-ghost" id="flt-close">✕</button></div>' +
    '<div class="fld"><div class="fld-label">关键字</div><input id="flt-q" class="input" placeholder="序列号 / 机型 / 供应商" value="' + esc(state.query) + '"></div>' +
    multiChips("状态", opts.status, f.status) +
    multiChips("新旧", opts.condition, f.condition) +
    multiChips("机型", opts.model, f.model) +
    radioChips("异常", "anomaly", [{ label: "全部", val: "all" }, { label: "仅异常", val: "yes" }, { label: "排除异常", val: "no" }]) +
    radioChips("问题机", "problem", [{ label: "全部", val: "all" }, { label: "仅问题机", val: "yes" }, { label: "排除问题机", val: "no" }]) +
    radioChips("库龄", "age", [{ label: "全部", val: "all" }, { label: "0-30 天", val: "young" }, { label: "30-90 天", val: "mid" }, { label: "≥90 天", val: "old" }]) +
    '<div class="fld"><div class="fld-label">日期范围（在途页为在途日，其余为入库日）</div><div class="btn-row"><input id="flt-from" type="date" class="input" value="' + esc(f.dateFrom || "") + '"><input id="flt-to" type="date" class="input" value="' + esc(f.dateTo || "") + '"></div></div>' +
    '<div class="fld"><div class="fld-label">成本 VND 范围</div><div class="btn-row"><input id="flt-min" type="number" min="0" class="input" placeholder="最低" value="' + esc(f.costMin || "") + '"><input id="flt-max" type="number" min="0" class="input" placeholder="最高" value="' + esc(f.costMax || "") + '"></div></div>' +
    (opts.supplier.length ? '<div class="fld"><div class="fld-label">供应商</div><select id="flt-sup" class="input">' +
      ['<option value="">全部</option>'].concat(opts.supplier.map(function (x) { return '<option value="' + esc(x) + '"' + (f.supplier === x ? " selected" : "") + ">" + esc(x) + "</option>"; })).join("") +
      "</select></div>" : "") +
    '<div class="drawer-actions"><button class="btn btn-primary" id="flt-apply">✅ 应用筛选</button>' +
    '<button class="btn" id="flt-reset">重置</button></div>';
  $("#flt-close").onclick = function () { state.drawerMode = null; renderDrawer(); };
  $("#flt-reset").onclick = function () {
    resetLedgerFilters(); state.drawerMode = null;
    renderDrawer(); render();
  };
  $("#flt-apply").onclick = function () {
    var groups = {};
    box.querySelectorAll("[data-multi]").forEach(function (cb) {
      var g = cb.getAttribute("data-multi");
      (groups[g] = groups[g] || []).push({ v: cb.value, on: cb.checked });
    });
    state.filters.status = (groups["状态"] || []).filter(function (x) { return x.on; }).map(function (x) { return x.v; });
    state.filters.condition = (groups["新旧"] || []).filter(function (x) { return x.on; }).map(function (x) { return x.v; });
    state.filters.model = (groups["机型"] || []).filter(function (x) { return x.on; }).map(function (x) { return x.v; });
    function radio(n) { var x = box.querySelector('input[name="' + n + '"]:checked'); return x ? x.value : null; }
    state.filters.anomaly = radio("rd-anomaly") || "all";
    state.filters.problem = radio("rd-problem") || "all";
    state.filters.age = radio("rd-age") || "all";
    var sup = $("#flt-sup"); if (sup) state.filters.supplier = sup.value;
    state.filters.dateFrom = $("#flt-from").value; state.filters.dateTo = $("#flt-to").value;
    state.filters.costMin = $("#flt-min").value; state.filters.costMax = $("#flt-max").value;
    state.query = $("#flt-q").value;
    state.page = 1; state.drawerMode = null;
    renderDrawer(); render();
  };
}

/* ============================================================
 * 库存核对：按月份比较 WPS/Excel 入库表与仓库流水
 * ============================================================ */
function renderReconcile(main) {
  var month=state.reconcileMonth || C.todayIso().slice(0,7), snap=state.snap;
  var wh=(snap.warehouse||[]).map(function(r){return Object.assign({warehouse:"C仓"},r);}).concat((snap.otherWarehouse||[]).map(function(r){return Object.assign({warehouse:"其他仓"},r);}));
  function text(r){return JSON.stringify(r||{}).toLowerCase();} function isReturn(r){return /退机|退货|退回|返厂|换机/.test(text(r));}
  function serial(r){return C.serialKey(r.serial||r.序列号||r.唯一码||r.机器码||"");}
  function date(r){return String(r.date||r.入库日期||r.日期||r.ioTime||r.出入库时间||"").slice(0,19);}
  var allIn=(snap.inbound||[]).slice().sort(function(a,b){return date(a).localeCompare(date(b));}), whIn=wh.filter(function(r){return C.classifyWarehouse(r.docType,r.docStatus)==="in";}).sort(function(a,b){return date(a).localeCompare(date(b));});
  var inBy={}, whBy={}; allIn.forEach(function(r){var k=serial(r);(inBy[k]||(inBy[k]=[])).push(r);}); whIn.forEach(function(r){var k=serial(r);(whBy[k]||(whBy[k]=[])).push(r);});
  var whFilter=state.reconWarehouse||"all", resultFilter=state.reconResult||"all"; var filter={"recon-normal":"正常入库","recon-return":"退机入库","recon-cross":"跨月已入仓","recon-missing":"无仓库记录"}[state.route]||(state.reconFilter||"all"), query=(state.reconQuery||"").toLowerCase(), details=[];
  Object.keys(inBy).forEach(function(k){var rows=inBy[k].filter(function(r){return date(r).slice(0,7)===month;}); if(!rows.length)return; var whRows=whBy[k]||[];
    rows.forEach(function(row){var occurrence=inBy[k].indexOf(row)+1, matched=whRows[occurrence-1], type=isReturn(row)?"退机入库":"正常入库", status=matched?(date(matched).slice(0,7)===month?"同月已入仓":"跨月已入仓"):"无仓库记录"; if(filter!=="all"&&filter!==type&&filter!==status)return;if(resultFilter!=="all"&&resultFilter!==status)return;if(whFilter!=="all"&&history.filter(function(w){return (w.warehouse||"C仓")===whFilter;}).length===0)return;if(query&&k.toLowerCase().indexOf(query)<0)return; var history=wh.filter(function(w){return serial(w)===k;});details.push({serial:row.serial,type:type,status:status,row:row,occurrence:occurrence,matched:matched,history:history});});
  });
  /* 补入“只有仓库入库、入库表没有记录”的序列号，不能让它们混入正常入库匹配。 */
  Object.keys(whBy).forEach(function(k){ var wr=whBy[k].filter(function(w){return date(w).slice(0,7)===month;}); if(!wr.length || inBy[k]) return; if(filter!=="all"&&filter!=="仓库有入库，入库表缺记录")return; if(resultFilter!=="all"&&resultFilter!=="仓库有入库，入库表缺记录")return; if(whFilter!=="all"&&wr.every(function(w){return (w.warehouse||"C仓")!==whFilter; }))return; if(query&&k.toLowerCase().indexOf(query)<0)return; details.push({serial:(wr[0].serial||k),occurrence:"—",type:"仓库有入库，入库表缺记录",status:"仓库有入库，入库表缺记录",row:{date:"—"},matched:wr[0],history:wh.filter(function(w){return serial(w)===k;})}); });
  var counts={};details.forEach(function(x){counts[x.status]=(counts[x.status]||0)+1;});
  main.appendChild(el("div","page-head",'<h1>系统库存核对</h1><p>按序列号的第几次入库逐次匹配仓库的第几次入库。当前月份退机只匹配当前月份的退机入库记录，之前月份退机不会计入。</p>'));
  var bar=el("div","toolbar");bar.innerHTML='<label class="fld-label">月份 <input id="recon-month" class="input" type="month" value="'+esc(month)+'"></label><select id="recon-filter" class="input"><option value="all">全部入库类型</option><option value="正常入库">正常入库</option><option value="退机入库">退机入库</option></select><select id="recon-result" class="input"><option value="all">全部核对结果</option><option value="同月已入仓">同月已入仓</option><option value="跨月已入仓">跨月已入仓</option><option value="无仓库记录">无仓库记录</option><option value="仓库有入库，入库表缺记录">仓库有入库，入库表缺记录</option></select><select id="recon-warehouse" class="input"><option value="all">全部仓库</option><option value="C仓">C仓</option><option value="其他仓">其他仓</option></select><input id="recon-q" class="input" placeholder="搜索序列号" value="'+esc(state.reconQuery||"")+'"><button id="recon-run" class="btn btn-primary">查询</button>';main.appendChild(bar);$("#recon-filter").value=filter;$("#recon-result").value=state.reconResult||"all";$("#recon-warehouse").value=state.reconWarehouse||"all";$("#recon-run").onclick=function(){state.reconcileMonth=$("#recon-month").value;state.reconFilter=$("#recon-filter").value;state.reconResult=$("#recon-result").value;state.reconWarehouse=$("#recon-warehouse").value;state.reconQuery=$("#recon-q").value;render();};
  var cards=el("div","panel key-bar");cards.innerHTML=kbItem("本月匹配记录",details.length,"条","c-blue")+Object.keys(counts).map(function(k){return kbItem(k,counts[k],"条",k==="无仓库记录"?"c-red":k==="跨月已入仓"?"c-amber":"c-green");}).join("");var exp=el("button","btn btn-outline","导出当前核对结果");exp.onclick=function(){var csv="序列号,入库次数,入库类型,入库表日期,匹配仓库入库,核对结果\n"+details.map(function(x){return [x.serial,x.occurrence,x.type,date(x.row),x.matched?date(x.matched):"",x.status].map(function(v){return "\""+String(v).replace(/\"/g,"\"\"")+"\"";}).join(",");}).join("\n");XIO.downloadBlob(new Blob(["\ufeff"+csv],{type:"text/csv;charset=utf-8"}),"系统核对_"+month+".csv");};main.appendChild(cards);main.appendChild(exp);
  var panel=el("div","panel"),table=el("table","tbl bordered");table.innerHTML='<thead><tr><th>序列号</th><th>本次入库次数</th><th>入库类型</th><th>入库表日期</th><th>匹配仓库入库</th><th>核对结果</th><th>该序列号全部仓库流水</th></tr></thead><tbody>'+details.map(function(x){return '<tr><td class="mono">'+esc(x.serial)+'</td><td>第 '+x.occurrence+' 次</td><td>'+esc(x.type)+'</td><td>'+esc(date(x.row))+'</td><td>'+esc(x.matched?date(x.matched)+' / '+(x.matched.warehouse||'C仓'):'—')+'</td><td>'+esc(x.status)+'</td><td style="white-space:normal;text-align:left!important">'+(x.history.map(function(w){return esc([date(w),w.warehouse,w.docType,w.docStatus].join(' / '));}).join('<br>')||'无仓库流水')+'</td></tr>';}).join('')+'</tbody>';panel.appendChild(table);main.appendChild(panel);
}

function renderOutboundReconcile(main) {
  var month=state.outMonth||C.todayIso().slice(0,7), snap=state.snap;
  function norm(v){return String(v==null?"":v).replace(/[\u200b\uFEFF\s'"`‘’“”、，,：:;；|\\/]+/g,"").toUpperCase();}
  function serial(r){return C.serialKey(r.serial||r.手机序列号||r.序列号||"");} function date(r){return String(r.orderDate||r.date||r.订单完成时间||"").slice(0,19);} function order(r){if(!r)return "";var vals=[r.orderNo,r.orderNoRaw,r.订单编号,r.订单号,r.单号,r.单据号,r.交易单号,r.orderId];for(var oi=0;oi<vals.length;oi++){if(String(vals[oi]==null?"":vals[oi]).trim())return String(vals[oi]).trim();}for(var ok of Object.keys(r)){if(/订单|单据|交易|order.?id|order.?no/i.test(ok)&&String(r[ok]==null?"":r[ok]).trim())return String(r[ok]).trim();}return "";} function money(v){var n=C.parseMoney(v);return n==null?null:n;} function amount(row, kind){
    if(!row)return null;
    var vals=kind==="table"?[row.cash,row.cashRaw,row.cashAmount,row["现金"],row["现金金额"],row.cost,row["K"],row["K列现金"]]:[row.downPayment,row.downPaymentRaw,row["首付"],row["首付款"],row["首付金额"],row["首付款金额"],row["已付首付"],row.firstPayment,row.firstPay,row.deposit,row["AD"],row.cost];
    for(var vi=0;vi<vals.length;vi++){var n=money(vals[vi]);if(n!=null)return n;}
    var keys=Object.keys(row);
    for(var ki=0;ki<keys.length;ki++){var key=keys[ki];if(kind==="table"?/现金|cash|金额|^k$/i:/首付|down.?payment|first.?pay|deposit|^ad$/i.test(key)){var n2=money(row[key]);if(n2!=null)return n2;}}
    return null;
  }
  var tableRows=(snap.outbound||[]).filter(function(r){return date(r).slice(0,7)===month;}).map(function(r){if(r.cash==null&&r.cashRaw!=null)r.cash=money(r.cashRaw);return r;}), erpRows=(snap.systemOutbound||[]).filter(function(r){return date(r).slice(0,7)===month;}).map(function(r){if(r.downPayment==null&&r.downPaymentRaw!=null)r.downPayment=money(r.downPaymentRaw);return r;});
  var used={}, byOrder={}, bySerial={}; erpRows.forEach(function(r,i){var o=norm(order(r)), k=serial(r);if(o)(byOrder[o]||(byOrder[o]=[])).push({r:r,i:i});if(k)(bySerial[k]||(bySerial[k]=[])).push({r:r,i:i});});
  var resultFilter=state.outResult||"all", payFilter=state.outPay||"all", query=(state.outQuery||"").toLowerCase(), details=[];
  function take(list){if(!list)return null;for(var i=0;i<list.length;i++){if(!used[list[i].i]){used[list[i].i]=1;return list[i].r;}}return null;}
  function paymentStatus(t,e){var a=amount(t,"table"), b=amount(e,"erp");if(a==null&&b==null)return "无金额";if(a==null)return "表格缺现金";if(b==null)return "ERP缺首付";return Math.abs(a-b)<0.01?"首付一致":"首付差异";}
  tableRows.forEach(function(t){var e=take(byOrder[norm(order(t))])||take(bySerial[serial(t)]), status=e?(norm(order(t))&&norm(order(e))&&norm(order(t))!==norm(order(e))?"订单号不一致":serial(t)&&serial(e)&&serial(t)!==serial(e)?"序列号不一致":"一致"):"ERP缺少", pay=paymentStatus(t,e);if(resultFilter!=="all"&&resultFilter!==status)return;if(payFilter!=="all"&&payFilter!==pay)return;if(query&&norm(order(t)).toLowerCase().indexOf(query)<0&&serial(t).toLowerCase().indexOf(query)<0)return;details.push({t:t,e:e,status:status,pay:pay});});
  erpRows.forEach(function(e,i){if(used[i])return;var k=serial(e);if(resultFilter!=="all"&&resultFilter!=="表格缺少")return;if(payFilter!=="all"&&payFilter!==paymentStatus(null,e))return;if(query&&norm(order(e)).toLowerCase().indexOf(query)<0&&k.toLowerCase().indexOf(query)<0)return;details.push({t:null,e:e,status:"表格缺少",pay:paymentStatus(null,e)});});
  var counts={};details.forEach(function(x){counts[x.status]=(counts[x.status]||0)+1;});var tableTotal=tableRows.reduce(function(a,x){return a+(amount(x,"table")||0);},0),erpTotal=erpRows.reduce(function(a,x){return a+(amount(x,"erp")||0);},0);
  main.appendChild(el("div","page-head",'<h1>出库核对</h1><p>只比较当前月份“表格出库”与“ERP出库”。订单号会自动清理前导单引号、空格和不可见符号后匹配。</p>'));
  var bar=el("div","toolbar");bar.innerHTML='<label class="fld-label">月份 <input id="out-month" class="input" type="month" value="'+esc(month)+'"></label><select id="out-result" class="input"><option value="all">全部结果</option><option value="一致">一致</option><option value="ERP缺少">ERP缺少</option><option value="表格缺少">表格缺少</option><option value="订单号不一致">订单号不一致</option><option value="序列号不一致">序列号不一致</option></select><select id="out-pay" class="input"><option value="all">全部首付结果</option><option value="首付一致">首付一致</option><option value="首付差异">首付差异</option><option value="表格缺现金">表格缺现金</option><option value="ERP缺首付">ERP缺首付</option></select><input id="out-q" class="input" placeholder="订单号 / 序列号" value="'+esc(state.outQuery||"")+'"><button id="out-run" class="btn btn-primary">查询</button><button id="out-export" class="btn btn-outline">导出核对结果</button>';main.appendChild(bar);$("#out-result").value=resultFilter;$("#out-pay").value=payFilter;$("#out-run").onclick=function(){state.outMonth=$("#out-month").value;state.outResult=$("#out-result").value;state.outPay=$("#out-pay").value;state.outQuery=$("#out-q").value;render();};$("#out-export").onclick=function(){var csv="表格日期,表格订单号,表格序列号,表格现金K,ERP订单完成时间,ERP订单号,ERP序列号,ERP首付AD,出库比对,首付比对\\n"+details.map(function(x){return [x.t?date(x.t):"",x.t?order(x.t):"",x.t?x.t.serial:"",x.t?amount(x.t,"table"):"",x.e?date(x.e):"",x.e?order(x.e):"",x.e?x.e.serial:"",x.e?amount(x.e,"erp"):"",x.status,x.pay].map(function(v){return '"'+String(v).replace(/"/g,'""')+'"';}).join(',');}).join('\\n');XIO.downloadBlob(new Blob(["\\ufeff"+csv],{type:"text/csv;charset=utf-8"}),"出库及首付核对_"+month+".csv");};
  var cards=el("div","panel key-bar");cards.innerHTML=kbItem("表格出库",tableRows.length,"条","c-blue")+kbItem("ERP出库",erpRows.length,"条","c-violet")+kbItem("出库金额差异",tableRows.length-erpRows.length,"条","c-amber")+kbItem("首付差额",nfmt(tableTotal-erpTotal),"元","c-red");main.appendChild(cards);
  var panel=el("div","panel"),table=el("table","tbl bordered");table.innerHTML='<thead><tr><th>表格日期</th><th>表格订单号</th><th>表格序列号</th><th>表格现金(K)</th><th>ERP订单完成时间</th><th>ERP订单号</th><th>ERP序列号</th><th>ERP首付(AD)</th><th>出库比对</th><th>首付比对</th></tr></thead><tbody>'+details.map(function(x){var bad=x.status!=="一致"||x.pay==="首付差异";return '<tr class="'+(bad?'row-alert':'')+'"><td>'+esc(x.t?date(x.t):'—')+'</td><td class="mono">'+esc(x.t?order(x.t):'—')+'</td><td class="mono">'+esc(x.t?x.t.serial:'—')+'</td><td class="num">'+esc(x.t&&amount(x.t,"table")!=null?nfmt(amount(x.t,"table")):'—')+'</td><td>'+esc(x.e?date(x.e):'—')+'</td><td class="mono">'+esc(x.e?order(x.e):'—')+'</td><td class="mono">'+esc(x.e?x.e.serial:'—')+'</td><td class="num">'+esc(x.e&&amount(x.e,"erp")!=null?nfmt(amount(x.e,"erp")):'—')+'</td><td>'+esc(x.status)+'</td><td>'+esc(x.pay)+'</td></tr>';}).join('')+'</tbody>';panel.appendChild(table);main.appendChild(panel);
}

function renderHome(main){
 var k=state.result.kpi, r=state.result, stock=r.units.filter(function(u){return u.status==="在库";}), total=stock.length+k.transit;
 var head=el("div","home-head");head.innerHTML='<div><div class="home-date">星期三，'+esc(C.todayIso())+'</div><h1>早上好，开始管理今天的库存</h1><p>所有仓库数据已准备就绪，下面是最新概览。</p></div><div class="home-search">⌕　搜索序列号、机型或供应商　<span>林</span></div>';main.appendChild(head);
 var cards=el("div","home-kpis");cards.innerHTML='<div class="home-kpi mint"><small>总库存</small><b>'+nfmt(total)+'</b><span>台　↗ '+pct(k.inStock,total)+'</span></div><div class="home-kpi lavender"><small>在途库存</small><b>'+nfmt(k.transit)+'</b><span>台　↗ '+pct(k.transit,total)+'</span></div><div class="home-kpi peach"><small>本月入库</small><b>'+nfmt((snapMonthInbound(state.snap)||[]).length)+'</b><span>台　↗ 月度数据</span></div><div class="home-kpi rose"><small>异常待处理</small><b>'+nfmt(k.anomaly)+'</b><span>台　需要关注</span></div>';main.appendChild(cards);
 var charts=el("div","home-grid");var trend=el("div","home-chart panel");trend.innerHTML='<div class="panel-title">月度入库趋势</div><p class="hint">近 6 个月 · 台</p><svg class="line-chart" viewBox="0 0 620 170" preserveAspectRatio="none"><path d="M20 140 C90 126,120 132,170 100 S270 118,320 76 S420 98,470 42 S550 70,600 22" fill="none" stroke="#70c6ad" stroke-width="5" stroke-linecap="round"/><path d="M20 140 C90 126,120 132,170 100 S270 118,320 76 S420 98,470 42 S550 70,600 22 L600 160 L20 160 Z" fill="#cdeee3" opacity=".55"/></svg><div class="chart-labels">4月　　5月　　6月　　7月　　8月　　9月</div>';charts.appendChild(trend);var status=el("div","home-chart panel");status.innerHTML='<div class="panel-title">库存状态</div><p class="hint">按当前状态统计</p><div class="donut"><b>'+nfmt(total)+'</b><small>总库存</small></div><div class="legend"><span>● 在库 '+nfmt(k.inStock)+'</span><span>● 在途 '+nfmt(k.transit)+'</span><span>● 异常 '+nfmt(k.anomaly)+'</span></div>';charts.appendChild(status);main.appendChild(charts);
 var table=el("div","panel home-table");table.innerHTML='<div class="panel-title">最近库存记录 <span class="hint">共 '+nfmt(r.units.length)+' 条记录</span></div>';var t=el("table","tbl");t.innerHTML='<thead><tr><th>序列号</th><th>机型</th><th>内存</th><th>颜色</th><th>新旧</th><th>状态</th><th>成本VND</th><th>库龄</th></tr></thead><tbody>'+r.units.slice(0,5).map(function(u){return '<tr><td>'+esc(u.serial)+'</td><td>'+esc(u.model)+'</td><td>'+esc(u.memory)+'</td><td>'+esc(u.color)+'</td><td>'+esc(u.condition)+'</td><td>'+esc(u.status)+'</td><td>'+nfmt(u.costVnd)+'</td><td>'+esc(u.ageDays==null?'—':u.ageDays+'天')+'</td></tr>';}).join('')+'</tbody>';table.appendChild(t);main.appendChild(table);
}
function snapMonthInbound(s){var m=C.todayIso().slice(0,7);return (s.inbound||[]).filter(function(x){return String(x.date||'').slice(0,7)===m;});}

/* ============================================================
 * 数据导入：四个导入源 + 历史切换（需求三/五）
 * ============================================================ */
function renderImport(main) {
  var meta = state.snap.importMeta || {};
  var hist = state.snap.importHistory || {};
  main.appendChild(el("div", "page-head",
    "<h1>数据导入</h1><p>数据来源分为两类：WPS/Excel 周转表数据，以及仓库手动导入数据。每次导入<b>覆盖</b>对应源的最新数据，同时保留历史版本可切换回。</p>"));

  function metaLine(m) { return m ? esc(m.fileName) + " · " + m.rowCount + " 行 · " + esc(m.importedAt || "") : "尚未导入"; }

  function sourceCard(title, desc, source, applyImport) {
    var card = el("div", "panel");
    card.innerHTML = '<div class="panel-title">' + title + "</div><p class=\"hint\">" + desc + "</p>";
    var cur = source === "erp" ? metaLine(meta.inbound) : metaLine(meta[source]);
    card.appendChild(el("p", "hint", "当前生效：" + cur));
    var input = el("input"); input.type = "file"; input.accept = ".xlsx,.xls,.csv"; input.className = "file-input";
    var btn = el("button", "btn btn-primary", "选择文件并导入");
    btn.onclick = function () { input.click(); };
    input.onchange = function () {
      if (!input.files || !input.files[0]) return;
      var file = input.files[0];
      state.busy = "正在解析 " + file.name + " …"; render();
      applyImport(file).then(function (res) {
        state.busy = null;
        var snap = Object.assign({}, state.snap, { isSample: false, importMeta: Object.assign({}, state.snap.importMeta) });
        var now = new Date().toLocaleString("zh-CN");
        var holder = {};
        if (source === "erp") {
          if (res.inbound) { holder.inbound = res.inbound.rows; }
          if (res.outbound) { holder.outbound = res.outbound.rows; }
          if (res.transit) { holder.transit = res.transit.rows; }
          if (holder.inbound) { snap.inbound = holder.inbound; snap.importMeta.inbound = { fileName: file.name, importedAt: now, rowCount: holder.inbound.length, errorCount: res.errors.length }; }
          if (holder.outbound) { snap.outbound = holder.outbound; snap.importMeta.outbound = { fileName: file.name, importedAt: now, rowCount: holder.outbound.length, errorCount: res.errors.length }; }
          if (holder.transit) {
            if (snap.transitStandalone) {
              /* 已有独立上传的在途数据：三合一的在途表不覆盖，仅记录备查 */
              snap.importMeta.erpTransit = { fileName: file.name, importedAt: now, rowCount: holder.transit.length, errorCount: res.errors.length };
            } else {
              snap.transit = holder.transit; snap.importMeta.transit = { fileName: file.name, importedAt: now, rowCount: holder.transit.length, errorCount: res.errors.length };
            }
          }
          snap.transitOptions = res.transitOptions || [];
          snap.transitSheetName = res.transit ? res.transit.sheetName : "";
          if (holder.transit && res.transit) holder.transitOptions = res.transitOptions;
          toast("三合一导入完成：" + ["inbound", "outbound", "transit"].filter(function (k) { return holder[k]; }).length + " 个工作表");
        } else {
          holder[source] = res.rows;
          snap[source] = res.rows;
          if (source === "transit") snap.transitStandalone = true; /* 独立在途上传后优先于三合一 */
          snap.importMeta[source] = { fileName: file.name, importedAt: now, rowCount: res.rows.length, errorCount: res.errors.length };
          if (res.missingRequired && res.missingRequired.length) toast("注意：缺少推荐列 " + res.missingRequired.join("、"), "warn");
          toast(title + " 导入完成：" + res.rows.length + " 行" + (res.errors.length ? "，" + res.errors.length + " 行跳过" : ""));
        }
        snap.importHistory = C.pushImportHistory(state.snap, source, file.name, function (h) {
          Object.keys(holder).forEach(function (k) { h[k] = holder[k]; });
        });
        return setSnap(snap, true);
      }).catch(function (e) { state.busy = null; render(); toast((e && e.message) || "导入失败", "err"); });
    };
    card.appendChild(btn); card.appendChild(input);
    /* 历史版本切换 + 删除/清空 */
    var list = hist[source] || [];
    if (list.length) {
      var hbox = el("div"); hbox.style.marginTop = "10px";
      var hhead = el("div", "fld-label", "历史版本（切换数据源 / 删除旧版本）：");
      var clearAll = el("button", "btn btn-sm btn-ghost", "🗑 清空全部历史");
      clearAll.style.cssText = "color:var(--red);margin-left:10px;font-size:12px;padding:2px 8px";
      clearAll.onclick = function () {
        if (!confirm("确定删除「" + C.SOURCE_LABEL[source] + "」的全部历史版本吗？（当前生效数据不受影响）")) return;
        var h2 = Object.assign({}, state.snap.importHistory || {});
        delete h2[source];
        setSnap(Object.assign({}, state.snap, { importHistory: h2 }), true).then(function () { toast("已清空历史版本"); render(); });
      };
      hhead.appendChild(clearAll);
      hbox.appendChild(hhead);
      list.forEach(function (en, idx) {
        var row = el("div", "kv-row");
        var desc = esc(en.importedAt) + " · " + esc(en.fileName) + " · " + en.rowCount + " 行";
        if (source === "erp" && en.snapshot) {
          var parts = [];
          if (en.snapshot.inbound) parts.push("入库");
          if (en.snapshot.outbound) parts.push("出库");
          if (en.snapshot.transit) parts.push("在途");
          desc += "（含：" + parts.join("/") + "）";
        }
        row.innerHTML = '<div class="kv-label mono">' + desc + "</div>";
        var btns = el("div"); btns.style.cssText = "display:flex;gap:6px";
        var use = el("button", "btn btn-sm btn-outline", "切换到此版本");
        use.onclick = function () {
          var snap = C.applyImportEntry(state.snap, source, en);
          setSnap(snap).then(function () { toast("已切换到 " + en.importedAt + " 导入的「" + C.SOURCE_LABEL[source] + "」数据"); });
        };
        var del = el("button", "btn btn-sm btn-ghost", "✕");
        del.title = "删除此历史版本";
        del.style.color = "var(--red)";
        del.onclick = function () {
          var h2 = Object.assign({}, state.snap.importHistory || {});
          var l2 = (h2[source] || []).filter(function (x) { return x.id !== en.id; });
          if (l2.length) h2[source] = l2; else delete h2[source];
          setSnap(Object.assign({}, state.snap, { importHistory: h2 }), true).then(function () { toast("已删除该历史版本"); render(); });
        };
        btns.appendChild(use); btns.appendChild(del);
        row.appendChild(btns);
        hbox.appendChild(row);
      });
      card.appendChild(hbox);
    }
    return card;
  }

  var syncPanel=el("div","panel");
  syncPanel.innerHTML='<div class="panel-title">WPS 数据</div><p class="hint">本版不需要 BAT、Node.js 或同步助手。请在 WPS 下载 Excel 后，使用下方入口更新周转数据。私人 WPS 云表的自动同步需要官方授权接口或配套桌面应用，纯本地网页无法读取另一网页的登录数据。</p>';
  main.appendChild(syncPanel);
  main.appendChild(sourceCard("① WPS/Excel 周转数据（入库NK / 出库XK / 在途表）", "WPS表格内含『手机入库NK』『手机出库XK』『在途表』三个工作表。", "erp", function (f) { return XIO.importErp(f); }));
  main.appendChild(sourceCard("② 仓库数据（手动导入）", "仓库流水需要手动选择 Excel 导入；关键列：唯一码、单据类型、单据状态、出入库时间。", "warehouse", function (f) { return XIO.importSingle(f, "warehouse"); }));
  main.appendChild(sourceCard("③ 其他仓数据", "唯一码列若含说明（如「LYF4QQ9C7X 带回深圳售后」）将自动拆分。含仓位列会按仓位分组。", "otherWarehouse", function (f) { return XIO.importSingle(f, "otherWarehouse"); }));
  main.appendChild(sourceCard("④ ERP出库数据（按订单完成时间）", "按出库数据模板导入；使用“订单编号”、F列“手机序列号”、P列“订单完成时间”作为出库核对依据。", "systemOutbound", function (f) { return XIO.importSingle(f, "systemOutbound"); }));
  var transitCard = sourceCard("⑤ 在途数据", "列：序列号/机型/内存/颜色/金额。预期到货日期为<b>可选</b>，可不填。", "transit", function (f) { return XIO.importSingle(f, "transit"); });
  if (state.snap.transitStandalone) {
    var tbar = el("div", "btn-row"); tbar.style.marginTop = "8px";
    var tclear = el("button", "btn btn-sm btn-danger", "清除在途数据，回退到三合一在途表");
    tclear.onclick = function () {
      if (!confirm("清除独立在途数据？在途将回退使用三合一文件里的在途表。")) return;
      var snap2 = Object.assign({}, state.snap, { importMeta: Object.assign({}, state.snap.importMeta), isSample: false });
      snap2.transitStandalone = false;
      /* 从最近一次三合一历史里找回在途表 */
      var erpHist = (state.snap.importHistory && state.snap.importHistory.erp) || [];
      var found = null;
      for (var hi = 0; hi < erpHist.length; hi++) { if (erpHist[hi].snapshot && erpHist[hi].snapshot.transit && erpHist[hi].snapshot.transit.length) { found = erpHist[hi]; break; } }
      if (found) {
        snap2.transit = found.snapshot.transit;
        snap2.importMeta.transit = { fileName: found.fileName + "（三合一在途表）", importedAt: found.importedAt, rowCount: found.snapshot.transit.length, errorCount: 0 };
      } else if (snap2.importMeta.erpTransit && snap2.transitOptions && snap2.transitOptions.length) {
        /* 当前快照里 transitOptions 还留着三合一的在途 */
        var mainOp = snap2.transitOptions.filter(function (o) { return /在途表/.test(o.sheetName); })[0] || snap2.transitOptions[0];
        snap2.transit = mainOp.rows;
        snap2.importMeta.transit = { fileName: snap2.importMeta.erpTransit.fileName + "（三合一在途表）", importedAt: snap2.importMeta.erpTransit.importedAt, rowCount: mainOp.rows.length, errorCount: 0 };
      } else {
        snap2.transit = [];
        delete snap2.importMeta.transit;
      }
      setSnap(snap2).then(function () { toast("已回退到三合一的在途表（" + snap2.transit.length + " 行）"); render(); });
    };
    }
  main.appendChild(transitCard);

  var danger = el("div", "panel");
  danger.innerHTML = '<div class="panel-title">演示数据 / 清空</div><p class="hint">清空会删除全部数据（不可恢复），建议先备份。</p>';
  var row = el("div", "btn-row");
  var demo = el("button", "btn btn-outline", "载入演示数据");
  demo.onclick = function () { setSnap(C.generateSample(), true).then(function () { toast("已载入演示数据"); }); };
  var clear = el("button", "btn btn-danger", "清空全部数据");
  clear.onclick = function () {
    if (!confirm("确定要清空全部数据吗？本地数据文件夹里的 snapshot.json 也会被一并删除，此操作不可恢复！")) return;
    state.activeUnit = null; state.drawerMode = null; state.editing = false;
    setSnap(C.emptySnapshot(), true).then(function () { return DB.clearSnapshot(); }).then(function () { toast("已清空（含本地文件夹数据）"); });
  };
  row.appendChild(demo); row.appendChild(clear); danger.appendChild(row); main.appendChild(danger);
}

/* ============================================================
 * 设置（⑥ 双日期 + ② 其他仓开关）+ 备份
 * ============================================================ */
function renderSettings(main) {
  var s = state.snap.settings;
  main.appendChild(el("div", "page-head", "<h1>设置</h1><p>保存后立即重算全表。</p>"));
  var form = el("div", "panel form-grid");
  form.innerHTML =
    '<label class="fld">库龄计算日（只影响库龄/库龄分段）<input id="s-age" type="date" class="input" value="' + esc(s.ageDate || s.cutoffDate) + '"></label>' +
    '<label class="fld">库存 / 在途计算日（默认今天；回看某天的库存状态就改它）<input id="s-stock" type="date" class="input" value="' + esc(s.stockDate || C.todayIso()) + '"></label>' +
    '<label class="fld">汇率（1 人民币 = ? 越南盾）<input id="s-fx" type="number" min="1" step="1" class="input" value="' + esc(s.fxRate) + '"></label>' +
    '<label class="fld">销量窗口（天）<input id="s-window" type="number" min="1" max="365" class="input" value="' + esc(s.salesWindowDays) + '"></label>' +
    '<label class="fld">成本币种<select id="s-ccy" class="input">' +
    ["auto", "CNY", "VND"].map(function (v) { var label = v === "auto" ? "自动判断" : v; return '<option value="' + v + '"' + (s.costCurrency === v ? " selected" : "") + ">" + label + "</option>"; }).join("") + "</select></label>" +
    '<label class="fld chk"><input id="s-rental" type="checkbox"' + (s.includeRentalInSales ? " checked" : "") + "> 出租计入销量</label>" +
    '<label class="fld chk"><input id="s-unknown" type="checkbox"' + (s.unknownAsUsed ? " checked" : "") + "> 新旧未知按二手机</label>" +
    '<label class="fld chk"><input id="s-ow" type="checkbox"' + (s.includeOtherWarehouses ? " checked" : "") + "> <b>计入其他仓数据</b>（其他仓已导入 " + (state.snap.otherWarehouse || []).length + " 行）</label>" +
    '<label class="fld span2">正常出库原因（默认：出租，永不报问题机）<textarea id="s-reasons" class="input" rows="2">' + esc((s.normalOutReasons || []).join("\n")) + "</textarea></label>" +
    '<label class="fld span2"><b>问题机关键词</b>（出库原因含任一关键词才算问题机；默认只有“深圳”，即带回深圳售后等。想让“遗失/报废”等也算，在这里加一行一个关键词）<textarea id="s-pkw" class="input" rows="3">' + esc((s.problemKeywords || []).join("\n")) + "</textarea></label>" +
    '<label class="fld span2">机型别名（格式：关键词=标准名，每行一个）<textarea id="s-aliases" class="input mono" rows="6">' +
    esc((s.modelAliases || []).map(function (a) { return a.keyword + "=" + a.canonical; }).join("\n")) + "</textarea></label>";
  main.appendChild(form);
  var row = el("div", "btn-row");
  var save = el("button", "btn btn-primary", "保存并重算");
  save.onclick = function () {
    var ageD = $("#s-age").value, stockD = $("#s-stock").value;
    var fx = Number($("#s-fx").value), win = Number($("#s-window").value);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(ageD)) return toast("库龄日格式应为 YYYY-MM-DD", "err");
    if (!/^\d{4}-\d{2}-\d{2}$/.test(stockD)) return toast("库存日格式应为 YYYY-MM-DD", "err");
    if (!Number.isFinite(fx) || fx <= 0) return toast("汇率必须大于 0", "err");
    if (!Number.isFinite(win) || win < 1 || win > 365) return toast("窗口天数应在 1–365 之间", "err");
    var aliases = $("#s-aliases").value.split("\n").map(function (l) { var i = l.indexOf("="); if (i <= 0) return null; return { keyword: l.slice(0, i).trim(), canonical: l.slice(i + 1).trim() }; }).filter(function (x) { return x && x.keyword && x.canonical; });
    var snap = Object.assign({}, state.snap, { isSample: false });
    snap.settings = Object.assign({}, s, {
      ageDate: ageD, cutoffDate: ageD, stockDate: stockD,
      fxRate: fx, salesWindowDays: win, costCurrency: $("#s-ccy").value,
      includeRentalInSales: $("#s-rental").checked,
      unknownAsUsed: $("#s-unknown").checked,
      includeOtherWarehouses: $("#s-ow").checked,
      normalOutReasons: $("#s-reasons").value.split("\n").map(function (x) { return x.trim(); }).filter(Boolean),
      problemKeywords: $("#s-pkw").value.split("\n").map(function (x) { return x.trim(); }).filter(Boolean),
      modelAliases: aliases.length ? aliases : C.DEFAULT_MODEL_ALIASES.slice()
    });
    setSnap(snap);
  };
  row.appendChild(save); main.appendChild(row);

  var bak = el("div", "panel");
  bak.innerHTML = '<div class="panel-title">备份 / 恢复</div><p class="hint">本地版数据在你选择的文件夹 snapshot.json；云端版在服务器 data/snapshot.json。换设备前先导出备份。</p>';
  var brow = el("div", "btn-row");
  var out = el("button", "btn btn-outline", "导出备份（JSON）");
  out.onclick = function () { XIO.downloadBlob(XIO.exportBackup(state.snap), "进销存备份_" + C.todayIso() + ".json"); };
  var inp = el("input"); inp.type = "file"; inp.accept = ".json"; inp.style.display = "none";
  var inn = el("button", "btn btn-outline", "从备份恢复");
  inn.onclick = function () { inp.click(); };
  inp.onchange = function () {
    if (!inp.files || !inp.files[0]) return;
    XIO.importBackup(inp.files[0]).then(function (snap) { return setSnap(snap, true).then(function () { toast("已从备份恢复"); }); }).catch(function (e) { toast(e.message || "恢复失败", "err"); });
  };
  brow.appendChild(out); brow.appendChild(inn); bak.appendChild(brow); bak.appendChild(inp); main.appendChild(bak);
}

/* ---------- 启动 ---------- */
function boot() {
  DB.loadSnapshot().then(function (saved) {
    if (saved && !saved.isSample && C.hasUserRows(saved)) {
      var empty = C.emptySnapshot();
      var settings = Object.assign(empty.settings, saved.settings || {});
      if (!settings.modelAliases || !settings.modelAliases.length) settings.modelAliases = C.DEFAULT_MODEL_ALIASES.slice();
      if (!settings.stockDate) settings.stockDate = C.todayIso();
      if (!settings.ageDate) settings.ageDate = settings.cutoffDate || C.todayIso();
      state.snap = Object.assign(empty, saved, { settings: settings });
      state.snap.otherWarehouse = state.snap.otherWarehouse || [];
      state.snap.systemOutbound = state.snap.systemOutbound || [];
      state.snap.importHistory = state.snap.importHistory || {};
      state.snap.transitOptions = state.snap.transitOptions || [];
      if (!state.snap.settings.problemKeywords || !state.snap.settings.problemKeywords.length) state.snap.settings.problemKeywords = ["深圳"];
    } else {
      state.snap = C.generateSample();
    }
    recompute(); render();
  });
}
window.addEventListener("DOMContentLoaded", boot);
window.addEventListener("error", function (e) { var box = $("#toasts"); if (box) toast("运行错误：" + (e.message || "未知"), "err"); });
})();
