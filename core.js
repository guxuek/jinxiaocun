/* ============================================================
 * 进销存核心逻辑（纯计算，无 DOM 依赖）
 * v7：双日期（库龄日/库存日在途日）、其他仓合并开关、仓位标注、
 *     序列号+说明拆分、系统出库合并、新周转公式（日销量/可售天数）
 * ============================================================ */
(function (global) {
"use strict";

/* ---------- 文本规整 ---------- */
function collapse(s) { return String(s == null ? "" : s).replace(/\s+/g, " ").trim(); }
function looksLikeSerial(s) {
  var t = String(s == null ? "" : s).trim();
  return /^[A-Za-z0-9]{8,20}$/.test(t);
}
function serialKey(s) { return String(s == null ? "" : s).trim().toUpperCase(); }
function headerKey(s) {
  return String(s == null ? "" : s).toLowerCase().replace(/[\s_\-\/（）()：:．.·,，]+/g, "");
}
function cellStr(v) {
  if (v == null) return "";
  if (v instanceof Date) return isoFromMs(v.getTime()) || "";
  return collapse(v);
}

/* 序列号+说明 拆分：如 "LYF4QQ9C7X 带回深圳售后" → { serial, note } */
function splitSerialNote(s) {
  var raw = String(s == null ? "" : s).trim();
  if (!raw) return { serial: "", note: "" };
  var parts = raw.split(/\s+/);
  if (parts.length === 1) return { serial: raw, note: "" };
  for (var i = 0; i < parts.length; i++) {
    if (looksLikeSerial(parts[i])) {
      return { serial: parts[i], note: parts.slice(0, i).concat(parts.slice(i + 1)).join(" ") };
    }
  }
  return { serial: parts[0], note: parts.slice(1).join(" ") };
}

/* ---------- 机型 / 内存 / 新旧 ---------- */
var DEFAULT_MODEL_ALIASES = [
  { keyword: "iphone17promax", canonical: "iPhone 17 Pro Max" },
  { keyword: "iphone17pro", canonical: "iPhone 17 Pro" },
  { keyword: "iphone17", canonical: "iPhone 17" },
  { keyword: "iphone16promax", canonical: "iPhone 16 Pro Max" },
  { keyword: "iphone16pro", canonical: "iPhone 16 Pro" },
  { keyword: "iphone16plus", canonical: "iPhone 16 Plus" },
  { keyword: "iphone16", canonical: "iPhone 16" },
  { keyword: "iphone15promax", canonical: "iPhone 15 Pro Max" },
  { keyword: "iphone15pro", canonical: "iPhone 15 Pro" },
  { keyword: "iphone15plus", canonical: "iPhone 15 Plus" },
  { keyword: "iphone15", canonical: "iPhone 15" },
  { keyword: "iphone14promax", canonical: "iPhone 14 Pro Max" },
  { keyword: "iphone14pro", canonical: "iPhone 14 Pro" },
  { keyword: "iphone14plus", canonical: "iPhone 14 Plus" },
  { keyword: "iphone14", canonical: "iPhone 14" },
  { keyword: "iphone13promax", canonical: "iPhone 13 Pro Max" },
  { keyword: "iphone13pro", canonical: "iPhone 13 Pro" },
  { keyword: "iphone13", canonical: "iPhone 13" },
  { keyword: "iphone12promax", canonical: "iPhone 12 Pro Max" },
  { keyword: "iphone12pro", canonical: "iPhone 12 Pro" },
  { keyword: "iphone12", canonical: "iPhone 12" },
  { keyword: "iphone11promax", canonical: "iPhone 11 Pro Max" },
  { keyword: "iphone11pro", canonical: "iPhone 11 Pro" },
  { keyword: "iphone11", canonical: "iPhone 11" }
];
function normalizeModel(raw, aliases) {
  var k = String(raw == null ? "" : raw).toLowerCase().replace(/[^a-z0-9一-鿿]+/g, "");
  if (!k) return "";
  var list = aliases && aliases.length ? aliases : DEFAULT_MODEL_ALIASES;
  for (var i = 0; i < list.length; i++) {
    var kk = String(list[i].keyword).toLowerCase().replace(/[^a-z0-9一-鿿]+/g, "");
    if (kk && k.indexOf(kk) >= 0) return list[i].canonical;
  }
  return "";
}
function normalizeMemory(raw) {
  var m = String(raw == null ? "" : raw).toLowerCase().replace(/\s+/g, "");
  var mm = m.match(/(\d+)\s*(tb|g|gb)/);
  if (!mm) return "";
  return mm[1] + (mm[2] === "tb" ? "TB" : "GB");
}
function normalizeCondition(raw) {
  var s = String(raw == null ? "" : raw).toLowerCase();
  if (!s) return "未知";
  if (/mới|new|新机|全新/.test(s) && !/cũ|旧|二手|99|98/.test(s)) return "新机";
  if (/cũ|旧|二手|置换|置換|used|99|98|95/.test(s)) return "二手机";
  return "未知";
}
function classifyWarehouse(docType, docStatus) {
  var t = String(docType || ""), s = String(docStatus || "");
  if (/作废|作廢/.test(s)) return "void";
  if (/出仓|出庫|出库|销售出|其他出/.test(t) || /已出库|已出庫/.test(s)) return "out";
  if (/入仓|入库|入庫|导入|導入|采购入/.test(t) || /已入库|已入庫/.test(s)) return "in";
  return "other";
}
/* 问题机判定：关键词制。出库原因含任一问题关键词（默认“深圳”，如“带回深圳售后”）才算问题机；
   正常原因（出租等）永不报问题机；其余原因默认不算，可在设置里加关键词后再算。 */
function isProblemReason(reason, settings) {
  var r = String(reason == null ? "" : reason).trim();
  if (!r) return false;
  var normals = settings.normalOutReasons && settings.normalOutReasons.length ? settings.normalOutReasons : ["出租"];
  if (normals.indexOf(r) >= 0) return false;
  var kws = settings.problemKeywords && settings.problemKeywords.length ? settings.problemKeywords : ["深圳"];
  for (var i = 0; i < kws.length; i++) {
    var k = String(kws[i]).trim();
    if (k && r.indexOf(k) >= 0) return true;
  }
  return false;
}
/* 由在途表子表名推出“做表日期”：如「在途8月」「在途7月」→ 当年该月末；
   「在途表」本身不推（用行内日期列）。 */
function pickTransitDate(trs) {
  var KEYS = ["date","采购月份","purchaseMonth","purchaseDate","month","月份","制表日期","日期","采购日期","下单日期","etaDate","预计到到货日","预计到货日","inTransitDate","在途日期"];
  for (var i = 0; i < trs.length; i++) {
    var r = trs[i] || {};
    for (var j = 0; j < KEYS.length; j++) {
      var v = r[KEYS[j]];
      if (v === null || v === undefined || v === "") continue;
      if (v instanceof Date) return isoFromMs(v.getTime()) || "";
      return String(v);
    }
  }
  return null;
}
function transitSheetDate(sheetName, refDate) {
  var m = String(sheetName || "").match(/(\d{1,2})\s*月/);
  if (!m) return "";
  var ref = refDate ? new Date(refDate) : new Date();
  var y = ref.getFullYear();
  var month = +m[1];
  if (month < 1 || month > 12) return "";
  /* 月末日 */
  var d = new Date(y, month, 0);
  return isoFromMs(d.getTime());
}

/* ---------- 数字 / 日期 ---------- */
function parseMoney(v) {
  if (v == null || v === "") return null;
  if (typeof v === "number") return Number.isFinite(v) ? v : null;
  var s=String(v).trim().replace(/[，,\s￥¥₫]/g,"").replace(/[^\d.\-]/g,"");
  if (!s) return null;
  /* 兼容 1.200.000 / 1.200.000,50 等导出格式 */
  var dots=(s.match(/\./g)||[]).length;
  if (dots>1) { var parts=s.split("."); var tail=parts.pop(); s=parts.join("")+((tail.length===1||tail.length===2)?"."+tail:tail); }
  var n=Number(s); return Number.isFinite(n) ? n : null;
}
function parseDateMs(v) {
  if (v == null || v === "") return null;
  if (v instanceof Date) return v.getTime();
  if (typeof v === "number") {
    if (v > 1e12) return v;
    if (v > 20000 && v < 80000) return Math.round((v - 25569) * 86400000);
    return null;
  }
  var s = String(v).trim();
  var m = s.match(/(\d{4})[年\-\/\.](\d{1,2})[月\-\/\.](\d{1,2})/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]).getTime();
  m = s.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
  if (m) {
    /* 越南表格为 日/月/年；若第二位>12 则第一位是月份（美式），自动纠正 */
    var dd = +m[1], mm = +m[2], yy = +m[3];
    if (mm > 12 && dd <= 12) { var t0 = dd; dd = mm; mm = t0; }
    var dt = new Date(yy, mm - 1, dd);
    if (dt.getFullYear() !== yy || dt.getMonth() !== mm - 1 || dt.getDate() !== dd) return null;
    return dt.getTime();
  }
  var t = Date.parse(s);
  return Number.isFinite(t) ? t : null;
}
function pad2(n) { return n < 10 ? "0" + n : "" + n; }
function isoFromMs(ms) {
  if (ms == null) return "";
  var d = new Date(ms);
  return d.getFullYear() + "-" + pad2(d.getMonth() + 1) + "-" + pad2(d.getDate());
}
function dateOnly(s) {
  var ms = parseDateMs(s);
  return ms == null ? String(s || "").slice(0, 10) : isoFromMs(ms);
}
function daysBetween(fromIso, toIso) {
  var a = parseDateMs(fromIso), b = parseDateMs(toIso);
  if (a == null || b == null) return 0;
  return Math.max(0, Math.round((b - a) / 86400000));
}
function todayIso() { return isoFromMs(Date.now()); }
function fmtInt(n) { return n == null ? "—" : Math.round(n).toLocaleString("zh-CN"); }
function fmtMoney(n) { return n == null ? "—" : Math.round(n).toLocaleString("zh-CN"); }
function fmtDays(d) { return d == null ? "—" : d + " 天"; }

/* ---------- 列定义（四种导入源） ---------- */
var SOURCE_FIELDS = {
  warehouse: [
    { key: "serial", label: "唯一码", required: true, aliases: ["唯一码", "序列号", "序列號", "imei", "serial", "sn"] },
    { key: "docType", label: "单据类型", aliases: ["单据类型", "單據類型", "类型"] },
    { key: "docStatus", label: "单据状态", aliases: ["单据状态", "單據狀態", "状态"] },
    { key: "ioTime", label: "出入库时间", aliases: ["出入库时间", "出入庫時間", "时间", "日期"] }
  ],
  otherWarehouse: [
    { key: "serial", label: "唯一码", required: true, aliases: ["唯一码", "序列号", "序列號", "imei", "serial", "sn", "sn号", "编码", "唯一码/sn", "串号"] },
    { key: "docType", label: "单据类型", aliases: ["单据类型", "單據類型", "类型", "出入类型", "出入库类型", "业务类型", "单据", "类型名称", "出入方向"] },
    { key: "docStatus", label: "单据状态", aliases: ["单据状态", "單據狀態", "状态", "审核状态", "審核狀態"] },
    { key: "ioTime", label: "出入库时间", aliases: ["出入库时间", "出入庫時間", "时间", "日期", "创建时间", "創建時間", "操作时间", "出入时间", "业务日期"] },
    { key: "warehouse", label: "仓位", aliases: ["仓位", "倉位", "仓库", "仓别", "所属仓", "库房", "门店", "仓位名称", "所在仓位", "所在仓库", "所属仓位", "仓号", "仓库名称"] }
  ],
  systemOutbound: [
    { key: "serial", label: "手机序列号", required: true, aliases: ["手机序列号", "序列号", "序列號", "唯一码", "imei", "serial", "sn"] },
    { key: "date", label: "订单完成时间", aliases: ["订单完成时间", "訂單完成時間", "出库日期", "出庫日期", "日期", "时间", "出库时间"] },
    { key: "orderNo", label: "订单编号", aliases: ["订单编号", "訂單編號", "订单号", "訂單號", "单号", "单据号", "单据编号", "交易单号", "交易编号", "订单编码", "order id", "orderid", "order_no"] },
    { key: "downPayment", label: "首付", aliases: ["首付", "首付款", "首付金额", "首付款金额"] },
    { key: "reason", label: "出库原因", aliases: ["出库原因", "出庫原因", "原因", "备注"] },
    { key: "modelRaw", label: "机型", aliases: ["机型", "機型", "型号", "model"] }
  ],
  inbound: [
    { key: "date", label: "入库日期", aliases: ["入库日期", "入庫日期", "日期", "date"] },
    { key: "modelRaw", label: "机型原文", aliases: ["机型原文", "机型", "機型", "model", "型号"] },
    { key: "color", label: "颜色", aliases: ["颜色", "顏色", "color", "màu"] },
    { key: "memoryRaw", label: "内存原文", aliases: ["内存原文", "内存", "memory", "容量"] },
    { key: "conditionRaw", label: "新旧原文", aliases: ["新旧原文", "新旧", "成色", "condition"] },
    { key: "serial", label: "序列号", required: true, aliases: ["序列号", "序列號", "imei", "serial", "sn", "唯一码"] },
    { key: "inboundStatus", label: "入库状态", aliases: ["入库状态", "入庫狀態", "状态"] },
    { key: "cost", label: "成本", aliases: ["成本", "成本金额", "成本金額", "cost", "金额"] },
    { key: "supplier", label: "供应商", aliases: ["供应商", "供應商", "supplier", "采购来源"] },
    { key: "outboundReason", label: "已出库原因", aliases: ["已出库原因", "出庫原因", "出库原因", "原因"] }
  ],
  outbound: [
    { key: "orderDate", label: "表格日期", aliases: ["订单完成时间", "下单日期", "下單日期", "日期", "出库日期", "date"] },
    { key: "orderNo", label: "表格订单号", aliases: ["订单编号", "订单号", "訂單編號", "訂單號", "orderNo"] },
    { key: "serial", label: "表格序列号", required: true, aliases: ["序列号", "序列號", "imei", "serial", "sn", "唯一码"] },
    { key: "modelRaw", label: "机型原文", aliases: ["机型原文", "机型", "機型", "model", "型号"] },
    { key: "conditionRaw", label: "新旧原文", aliases: ["新旧原文", "新旧", "成色", "condition"] },
    { key: "cost", label: "成本", aliases: ["成本", "成本金额", "成本金額", "cost", "金额"] },
    { key: "reason", label: "出库原因", aliases: ["出库原因", "出庫原因", "原因"] },
    { key: "cash", label: "现金", aliases: ["现金", "現金", "现金K", "K列现金", "现金金额", "现金金额(K)", "实收现金"] }
  ],
  transit: [
    { key: "modelRaw", label: "机型原文", aliases: ["机型原文", "机型", "機型", "model", "型号"] },
    { key: "memoryRaw", label: "内存原文", aliases: ["内存原文", "内存", "memory", "容量"] },
    { key: "color", label: "颜色", aliases: ["颜色", "顏色", "color", "màu"] },
    { key: "serial", label: "序列号", required: true, aliases: ["序列号", "序列號", "imei", "serial", "sn", "唯一码"] },
    { key: "cny", label: "人民币", aliases: ["人民币", "人民幣", "人民币成本", "cny", "rmb"] },
    { key: "vnd", label: "越南盾", aliases: ["越南盾", "vnd", "盾", "采购成本(vnd)", "采购成本"] },
    { key: "date", label: "在途日期", aliases: ["在途日期", "日期", "月份", "制表日期", "date", "采购日期", "下单日期", "采购月份"] },
    { key: "status", label: "在途状态", aliases: ["状态", "狀態", "status", "在途状态"] },
    { key: "conditionRaw", label: "新旧原文", aliases: ["新旧原文", "新旧", "成色", "condition"] }
  ]
};

/* ============================================================
 * 共享列定义（工作台渲染 & Excel 导出唯一来源）
 * ============================================================ */
var COLUMN_DEFS = {
  category_summary: {
    headers: [
      [{ text: "品类", colspan: 1 }],
      ["数量", "在库(含在途)", "问题机", "正常在库", "近60天销量", "日销量", "可售天数"],
      ["金额-越南盾", "库存(含在途)", "问题机", "正常在库", "近60天金额"],
      ["金额-元", "库存(含在途)", "问题机", "正常在库"]
    ],
    widths: [12, 14, 10, 12, 14, 10, 10, 16, 16, 16, 16, 14, 14, 14]
  },
  turnover: {
    headers: [
      ["型号", "内存", "库存(含在途)", "近60天销量", "日销量", "可售天数", "库存-盾", "库存-元", "近60天金额", "备注", "平均库龄"]
    ],
    widths: [22, 10, 14, 14, 10, 10, 18, 14, 20, 16, 10]
  },
  age_segment: {
    headers: [["库龄分段", "台数", "金额-元", "占比"]],
    widths: [14, 12, 18, 10]
  },
  system_status: {
    headers: [["状态", "含义", "台数", "柱图"]],
    widths: [14, 36, 12, 80]
  },
  summary: {
    headers: [["指标", "数值"]],
    widths: [28, 24]
  },
  ledger: {
    headers: [["序列号", "机型", "内存", "颜色", "新旧", "状态", "ERP单据", "仓单据", "问题机", "异常", "成本VND", "ERP入库日", "ERP库龄", "仓入库日/在途日", "仓库龄/在途龄", "备注", "操作"]],
    widths: [18, 18, 8, 8, 8, 10, 16, 16, 12, 16, 14, 12, 9, 14, 12, 16, 12]
  }
};

/* ERP 三合一工作表识别 */
function pickErpSheets(names) {
  function key(n) { return n.toLowerCase().replace(/\s+/g, ""); }
  function find(needles, avoid) {
    avoid = avoid || [];
    for (var i = 0; i < needles.length; i++) {
      for (var j = 0; j < names.length; j++) {
        var k = key(names[j]);
        var bad = false;
        for (var a = 0; a < avoid.length; a++) if (k.indexOf(avoid[a]) >= 0) { bad = true; break; }
        if (!bad && k.indexOf(needles[i]) >= 0) return names[j];
      }
    }
    return null;
  }
  /* 在途表：收集所有“在途XX”子表（如 在途表 / 在途8月），优先名含当前月或“表”字的为主表 */
  var transitAll = names.filter(function (n) { return key(n).indexOf("在途") >= 0; });
  var main = null;
  for (var i = 0; i < transitAll.length; i++) {
    if (/在途表/.test(transitAll[i])) { main = transitAll[i]; break; }
  }
  if (!main && transitAll.length) main = transitAll[0];
  return {
    inbound: find(["手机入库", "手機入庫", "入库nk", "入库"], ["出库", "进出"]),
    outbound: find(["手机出库", "手機出庫", "出库xk", "出库"], ["入库", "进出"]),
    transit: main,
    transitAll: transitAll
  };
}

/* ---------- 设置 / 快照 ---------- */
function defaultSettings() {
  return {
    cutoffDate: todayIso(),          // 旧字段：现作为「库龄计算日」的别名
    ageDate: todayIso(),             // 库龄计算日
    stockDate: todayIso(),           // 库存 / 在途计算日（默认当前日期）
    fxRate: 3940,
    salesWindowDays: 60,
    costCurrency: "auto",
    includeRentalInSales: false,
    unknownAsUsed: true,
    includeOtherWarehouses: false,   // 默认不计算其他仓位数据（可勾选）
    normalOutReasons: ["出租"],
    problemKeywords: ["深圳"],        // 问题机关键词（原因含即算；默认只算“带回深圳售后”等）
    modelAliases: DEFAULT_MODEL_ALIASES.slice()
  };
}
function emptySnapshot() {
  return {
    warehouse: [],          // C仓（主仓）流水
    otherWarehouse: [],     // 其他仓流水（带 warehouse 仓位名 + note 说明）
    systemOutbound: [],     // 系统出库数据
    inbound: [], outbound: [], transit: [],
    overrides: {},
    settings: defaultSettings(),
    importMeta: {},
    importHistory: {},      // 每个数据源的导入历史，用于切换
    transitStandalone: false, // 是否使用独立上传的在途数据（true 时三合一的在途表不生效）
    transitOptions: [], transitSheetName: "",
    isSample: false
  };
}
function hasUserRows(s) {
  return !!(s && ((s.inbound && s.inbound.length) || (s.outbound && s.outbound.length) ||
    (s.warehouse && s.warehouse.length) || (s.transit && s.transit.length) ||
    (s.otherWarehouse && s.otherWarehouse.length) || (s.systemOutbound && s.systemOutbound.length)));
}

/* 导入历史：每次导入推入一条，可切换启用哪一次的数据 */
var SOURCE_LABEL = { erp: "三合一数据", warehouse: "C仓数据", otherWarehouse: "其他仓数据", systemOutbound: "ERP出库数据", transit: "在途数据" };
function pushImportHistory(snap, source, fileName, applyFn) {
  var hist = Object.assign({}, snap.importHistory || {});
  var list = (hist[source] || []).slice(0, 4); // 保留最近 5 次
  var entry = {
    id: "imp-" + Date.now(),
    fileName: fileName,
    importedAt: new Date().toLocaleString("zh-CN"),
    snapshot: {}
  };
  var holder = {};
  applyFn(holder);
  entry.snapshot = holder;
  entry.rowCount = Object.keys(holder).reduce(function (a, k) { return a + (holder[k] ? holder[k].length : 0); }, 0);
  list.unshift(entry);
  hist[source] = list;
  return hist;
}
function applyImportEntry(snap, source, entry) {
  var next = Object.assign({}, snap, { importMeta: Object.assign({}, snap.importMeta) });
  var s = entry.snapshot || {};
  if (source === "erp") {
    if (s.inbound) { next.inbound = s.inbound; next.importMeta.inbound = metaOf(entry, "入库"); }
    if (s.outbound) { next.outbound = s.outbound; next.importMeta.outbound = metaOf(entry, "出库"); }
    if (s.transit) { next.transit = s.transit; next.importMeta.transit = metaOf(entry, "在途"); }
  } else if (source === "warehouse") {
    next.warehouse = s.warehouse || []; next.importMeta.warehouse = metaOf(entry, "C仓");
  } else if (source === "otherWarehouse") {
    next.otherWarehouse = s.otherWarehouse || []; next.importMeta.otherWarehouse = metaOf(entry, "其他仓");
  } else if (source === "systemOutbound") {
    next.systemOutbound = s.systemOutbound || []; next.importMeta.systemOutbound = metaOf(entry, "系统出库");
  } else if (source === "transit") {
    next.transit = s.transit || []; next.importMeta.transit = metaOf(entry, "在途");
  }
  next.isSample = false;
  return next;
}
function metaOf(entry, label) {
  return { fileName: entry.fileName + "（" + label + "）", importedAt: entry.importedAt, rowCount: entry.rowCount, errorCount: 0 };
}

/* ---------- 演示数据 ---------- */
function generateSample() {
  var inbound = [
    { id: "nk-0", date: "2026-07-08", modelRaw: "iPhone14 Pro", color: "白色", memoryRaw: "128g", conditionRaw: "旧机", serial: "H34XV44T03", inboundStatus: "已入库", cost: 13320000, supplier: "中国采购调拨", outboundReason: "出租" },
    { id: "nk-1", date: "2026-08-01", modelRaw: "iPhone16 Pro Max", color: "原色", memoryRaw: "256g", conditionRaw: "新机", serial: "AAA111111111", inboundStatus: "已入库", cost: 28000000, supplier: "NXT", outboundReason: "出租" },
    { id: "nk-2", date: "2026-08-15", modelRaw: "iPhone15 Pro", color: "蓝", memoryRaw: "256g", conditionRaw: "新机", serial: "BBB222222222", inboundStatus: "已入库", cost: 22000000, supplier: "NXT", outboundReason: "出租" },
    { id: "nk-3", date: "2026-06-20", modelRaw: "iPhone13 Pro", color: "黑", memoryRaw: "128g", conditionRaw: "旧机", serial: "DDD444444444", inboundStatus: "已入库", cost: 13320000, supplier: "本地", outboundReason: "带回深圳售后" }
  ];
  var outbound = [
    { id: "xk-0", orderDate: "2026-08-20", serial: "BBB222222222", modelRaw: "iPhone15 Pro", conditionRaw: "新机", cost: 22000000, reason: "出租" },
    { id: "xk-1", orderDate: "2026-09-01", serial: "CCC333333333", modelRaw: "iPhone14 Pro Max", conditionRaw: "旧机", cost: 18000000, reason: "出售" },
    { id: "xk-2", orderDate: "2026-09-03", serial: "CCC333333333", modelRaw: "iPhone14 Pro Max", conditionRaw: "旧机", cost: 18000000, reason: "出售" }
  ];
  var transit = [
    { id: "tr-0", modelRaw: "iPhone16 Pro", memoryRaw: "256g", color: "白", serial: "GGG777777777", cny: 6200, vnd: null, status: "在途", conditionRaw: "新机" }
  ];
  var warehouse = [
    { id: "wh-0", serial: "AAA111111111", docType: "导入", docStatus: "已入库", ioTime: "2026-08-02 10:00" },
    { id: "wh-1", serial: "BBB222222222", docType: "销售出仓", docStatus: "已出库", ioTime: "2026-08-21 15:00" }
  ];
  var otherWarehouse = [
    { id: "ow-0", serial: "DDD444444444", docType: "其他入库", docStatus: "已入库", ioTime: "2026-07-01 09:00", warehouse: "D仓", note: "带回深圳售后" }
  ];
  var systemOutbound = [
    { id: "so-0", serial: "DDD444444444", date: "2026-09-02", reason: "系统出库", modelRaw: "iPhone13 Pro" }
  ];
  var s = defaultSettings();
  return {
    warehouse: warehouse, otherWarehouse: otherWarehouse, systemOutbound: systemOutbound,
    inbound: inbound, outbound: outbound, transit: transit, overrides: {}, settings: s,
    importMeta: {}, importHistory: {}, isSample: true
  };
}

/* ---------- 主核算 ---------- */
function firstText() {
  for (var i = 0; i < arguments.length; i++) {
    var v = arguments[i];
    if (v && String(v).trim()) return String(v).trim();
  }
  return "";
}
function toVnd(amount, settings, hint) {
  if (amount == null || !Number.isFinite(amount)) return null;
  var fx = settings.fxRate > 0 ? settings.fxRate : 3600;
  var mode = hint || settings.costCurrency;
  if (mode === "CNY" || (mode === "auto" && Math.abs(amount) > 0 && Math.abs(amount) < 100000)) {
    return Math.round(amount * fx);
  }
  return Math.round(amount);
}

function computeAll(input) {
  var settings = input.settings;
  var aliases = (settings.modelAliases && settings.modelAliases.length) ? settings.modelAliases : DEFAULT_MODEL_ALIASES;

  /* 双日期：ageDate 只算库龄；stockDate 决定「该日为止」的在库/在途/已出 */
  var ageDate = dateOnly(settings.ageDate || settings.cutoffDate);
  var stockDate = dateOnly(settings.stockDate || todayIso());
  var stockMs = parseDateMs(stockDate); if (stockMs == null) stockMs = Date.now();
  var windowDays = Math.max(1, settings.salesWindowDays || 60);
  var windowStartMs = stockMs - (windowDays - 1) * 86400000;

  function withinStock(d) { var ms = parseDateMs(d); if (ms == null) return true; return ms <= stockMs + 86399999; }

  function groupBy(rows, getKey, filter) {
    var map = new Map();
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      if (filter && !filter(r)) continue;
      var k = serialKey(getKey(r));
      if (!k) continue;
      if (map.has(k)) map.get(k).push(r); else map.set(k, [r]);
    }
    return map;
  }
  var inboundBy = groupBy(input.inbound, function (r) { return r.serial; });
  /* ERP 出库表只用于系统核对，不参与周转表库存计算。 */
  var allOutbound = (input.outbound || []);
  var outboundBy = groupBy(allOutbound, function (r) { return r.serial; });
  var transitBy = groupBy(input.transit, function (r) { return r.serial; }, function (r) {
    var st = r.status || "";
    return !(st && /已入库|已入庫/.test(st) && !/在途/.test(st));
  });
  /* 仓位合并：C仓恒计入；其他仓默认跟随全局开关，可被单机覆盖（ov.includeOther） */
  var cRows = (input.warehouse || []).map(function (r) { return Object.assign({ warehouse: r.warehouse || "C仓" }, r); });
  var oRowsAll = (input.otherWarehouse || []).map(function (r) { return Object.assign({ warehouse: r.warehouse || "其他仓" }, r); });
  var oRowsBy = groupBy(oRowsAll, function (r) { return r.serial; });
  var warehouseBy = groupBy(cRows, function (r) { return r.serial; });
  function whRowsFor(k, ov) {
    var base = warehouseBy.get(k) || [];
    var other = oRowsBy.get(k) || [];
    var inc = ov && ov.includeOther === "是" ? true : (ov && ov.includeOther === "否" ? false : !!settings.includeOtherWarehouses);
    return inc ? base.concat(other) : base;
  }

  var keys = new Set();
  [inboundBy, outboundBy, transitBy, warehouseBy, oRowsBy].forEach(function (m) { m.forEach(function (_, k) { keys.add(k); }); });

  var units = [];
  keys.forEach(function (k) {
    var insAll = inboundBy.get(k) || [];
    var outsAll = outboundBy.get(k) || [];
    var trs = transitBy.get(k) || [];
    var whs = warehouseBy.get(k) || [];
    /* 截至库存日的过滤 */
    var ov = input.overrides[k] || input.overrides[(insAll[0] && insAll[0].serial) || (outsAll[0] && outsAll[0].serial) || k];
    var whs = whRowsFor(k, ov);
    /* 仓单据按时间形成当前有效状态：同仓位的“已作废”冲销它之前最近一条有效记录。 */
    var effectiveWhs = [];
    whs.slice().sort(function (a, b) { return (parseDateMs(a.ioTime) || 0) - (parseDateMs(b.ioTime) || 0); }).forEach(function (w) {
      var cls0 = classifyWarehouse(w.docType, w.docStatus);
      var wn = w.warehouse || "C仓";
      if (cls0 === "void") {
        for (var vi = effectiveWhs.length - 1; vi >= 0; vi--) {
          if ((effectiveWhs[vi].warehouse || "C仓") === wn) { effectiveWhs.splice(vi, 1); break; }
        }
      } else effectiveWhs.push(w);
    });
    var ins = insAll.filter(function (r) { return withinStock(r.date); });
    var outs = outsAll.filter(function (r) { return withinStock(r.orderDate); });
    var inCount = ins.length, outCount = outs.length;
    var totalIn = insAll.length, totalOut = outsAll.length;
    /* 异常比对用总量：总数一致即不算异常（日期过滤只影响“截至某日”的在库状态） */
    var qtyMismatch = totalOut > totalIn;

    var whInCount = 0, whOutCount = 0;
    var latestWh = null, latestWhMs = -1, latestWhCls = "other";
    var whInDate = "";
    for (var i = 0; i < effectiveWhs.length; i++) {
      var w = effectiveWhs[i];
      if (!withinStock(w.ioTime)) continue;
      var cls = classifyWarehouse(w.docType, w.docStatus);
      if (cls === "void") continue;
      if (cls === "in") { whInCount++; if (!whInDate) whInDate = dateOnly(w.ioTime); }
      if (cls === "out") whOutCount++;
      var ms = parseDateMs(w.ioTime) || 0;
      if (ms >= latestWhMs) { latestWhMs = ms; latestWh = w; latestWhCls = cls; }
    }

    var currentWhIsOut = latestWhCls === "out";
    var currentWhIsIn = latestWhCls === "in";
    var hasErp = totalIn + totalOut > 0;
    /* 在途判定：在途表里有 + 截至库存日系统未入库（入库日在库存日之后也算未入库→在途）。
       ERP 与仓库冲突时一律以 ERP 数据/日期为准，不再报「仓出但ERP在库」。 */
    var hasTransit = trs.length > 0 && inCount === 0;
    var warehouseOnly = !hasErp && !hasTransit && whInCount + whOutCount > 0;
    var erpInStock = inCount > outCount;

    var anomalyReason = "";
    if (qtyMismatch) anomalyReason = "出多于入";
    else if (warehouseOnly) anomalyReason = "仅仓库有记录";
    var rawAnomaly = Boolean(anomalyReason);
    var anomalyHandled = Boolean(ov && ov.anomalyHandled) || Boolean(ov && ov.status);
    var isAnomaly = rawAnomaly && !anomalyHandled;

    var nkProblem = "";
    for (var j = 0; j < insAll.length; j++) if (isProblemReason(insAll[j].outboundReason, settings)) { nkProblem = insAll[j].outboundReason; break; }
    var xkProblem = "";
    for (var j2 = 0; j2 < outsAll.length; j2++) if (!outsAll[j2]._sys && isProblemReason(outsAll[j2].reason, settings)) { xkProblem = outsAll[j2].reason; break; }
    var owProblem = "";
    (oRowsBy.get(k) || []).forEach(function (r) { if (!owProblem && isProblemReason(r.note, settings)) owProblem = r.note; });
    var problemReason = nkProblem || xkProblem || owProblem || "";
    var isProblem = Boolean(problemReason);
    /* 公司仍持有但暂不在 C 仓的机器：部门领用、维修/售后、问题仓等。 */
    var companyHoldReason = "";
    effectiveWhs.forEach(function (r) {
      var text = String((r.warehouse || "") + " " + (r.note || ""));
      if (!companyHoldReason && /(部门领用|维修|售后|问题仓|瑕疵仓)/i.test(text)) companyHoldReason = text.trim();
    });
    if (ov && ov.isProblem === "是") { isProblem = true; if (!problemReason) problemReason = "手改"; }
    else if (ov && ov.isProblem === "否") { isProblem = false; problemReason = ""; }

    var status = "待核";
    if (ov && ov.status) status = ov.status;
    else if (qtyMismatch) status = "待核";
    else if (warehouseOnly) status = "待核";
    else if (erpInStock) status = "在库";
    else if (inCount > 0 && inCount === outCount) status = "已出库";
    else if (totalIn > 0 && totalIn === totalOut) status = "已出库";
    else if (totalIn > totalOut && latestWhCls !== "out") status = "在库";
    else if (hasTransit || totalIn > 0) status = "在途"; /* 在途日 字段可选，单元状态仍正常显示 */
    /* 领用/维修/深圳售后仍属于公司库存；出租及明确出售类出库保持原状态。 */
    if (!(ov && ov.status) && companyHoldReason && status === "已出库" && !currentWhIsOut) status = "在库";
    if (!(ov && ov.status) && isProblem && /深圳|售后|维修/.test(problemReason || "") && status === "已出库" && !currentWhIsOut) status = "在库";

    var modelRaw = firstText(ov && ov.model, insAll[0] && insAll[0].modelRaw, outsAll[0] && outsAll[0].modelRaw, trs[0] && trs[0].modelRaw);
    var memoryRaw = firstText(ov && ov.memory, insAll[0] && insAll[0].memoryRaw, trs[0] && trs[0].memoryRaw);
    var conditionRaw = firstText(insAll[0] && insAll[0].conditionRaw, outsAll[0] && outsAll[0].conditionRaw, trs[0] && trs[0].conditionRaw);
    var condition = (ov && ov.condition) || normalizeCondition(conditionRaw);
    if (condition === "未知" && settings.unknownAsUsed) condition = "二手机";

    /* 双入库日 / 双库龄：ERP 入库日 vs 仓库入库日（均按库龄日 ageDate 计算） */
    var erpInboundDate = firstText.apply(null, ins.map(function (r) { return r.date; }));
    if (!erpInboundDate) erpInboundDate = firstText.apply(null, insAll.map(function (r) { return r.date; }));
    var lastOutDate = firstText.apply(null, outsAll.slice().reverse().map(function (r) { return r.orderDate; }));
    var erpAgeDays = erpInboundDate ? Math.max(0, daysBetween(erpInboundDate, ageDate)) : null;
    var whAgeDays = whInDate ? Math.max(0, daysBetween(whInDate, ageDate)) : null;
    var ageDays = erpAgeDays != null ? erpAgeDays : (whAgeDays != null ? whAgeDays : 0);
    /* 在途日期（做在途表时的日期：优先行内日期列，否则用子表名推月末）与在途库龄 */
    var transitDate = pickTransitDate(trs);
    if (!transitDate && trs[0] && trs[0]._sheet) transitDate = transitSheetDate(trs[0]._sheet, ageDate);
    var transitAgeDays = transitDate ? Math.max(0, daysBetween(transitDate, ageDate)) : null;
    var owNote = firstText.apply(null, (oRowsBy.get(k) || []).map(function (r) { return r.note; }));

    var costVnd = toVnd(insAll[0] ? insAll[0].cost : null, settings);
    if (costVnd == null) costVnd = toVnd(outsAll[0] ? outsAll[0].cost : null, settings);
    if (costVnd == null) costVnd = toVnd(trs[0] ? trs[0].vnd : null, settings, "vnd");
    if (costVnd == null) costVnd = toVnd(trs[0] ? trs[0].cny : null, settings, "cny");

    units.push({
      serial: (insAll[0] && insAll[0].serial) || (outsAll[0] && outsAll[0].serial) || (trs[0] && trs[0].serial) || (latestWh && latestWh.serial) || k,
      model: normalizeModel(modelRaw, aliases) || modelRaw || "未识别",
      modelRaw: modelRaw,
      memory: normalizeMemory(memoryRaw) || memoryRaw,
      color: firstText(ov && ov.color, insAll[0] && insAll[0].color, trs[0] && trs[0].color),
      condition: condition, conditionRaw: conditionRaw,
      status: status,
      inCount: inCount, outCount: outCount,
      inCountAll: insAll.length, outCountAll: outsAll.length,
      whInCount: whInCount, whOutCount: whOutCount,
      costVnd: costVnd,
      inboundDate: erpInboundDate, erpInboundDate: erpInboundDate,
      whInboundDate: whInDate, lastOutDate: lastOutDate,
      erpAgeDays: erpAgeDays, whAgeDays: whAgeDays, ageDays: ageDays, erpOrderDate: lastOutDate,
      transitDate: transitDate || null, transitAgeDays: transitAgeDays,
      remark: firstText(ov && ov.remark, owNote, companyHoldReason ? "公司库存：" + companyHoldReason : ""),
      includeOther: (ov && ov.includeOther) || "",
      companyHoldReason: companyHoldReason,
      isProblem: isProblem, problemReason: problemReason,
      isAnomaly: isAnomaly,
      anomalyReason: isAnomaly ? anomalyReason : (anomalyHandled && rawAnomaly ? anomalyReason + "·已处理" : ""),
      anomalyHandled: anomalyHandled,
      supplier: firstText(insAll[0] && insAll[0].supplier),
      source: hasErp ? "ERP" : hasTransit ? "在途" : "仓库"
    });
  });

  units.sort(function (a, b) { return a.serial < b.serial ? -1 : a.serial > b.serial ? 1 : 0; });

  var kpi = { inStock: 0, sold: 0, transit: 0, pending: 0, anomaly: 0, problem: 0, stockValue: 0 };
  for (var u = 0; u < units.length; u++) {
    var un = units[u];
    if (un.status === "不计入") continue;
    if (un.status === "在库") { kpi.inStock++; kpi.stockValue += un.costVnd || 0; }
    else if (un.status === "已出库") kpi.sold++;
    else if (un.status === "在途") kpi.transit++;
    else if (un.status === "待核") kpi.pending++;
    if (un.isAnomaly) kpi.anomaly++;
    if (un.isProblem) kpi.problem++;
  }

  var unitByKey = new Map();
  units.forEach(function (x) { unitByKey.set(serialKey(x.serial), x); });

  /* 近窗销量（截至库存日 stockDate 往前 windowDays 天） */
  var salesCount = new Map();
  for (var o = 0; o < allOutbound.length; o++) {
    var ob = allOutbound[o];
    var oms = parseDateMs(ob.orderDate);
    if (oms == null || oms < windowStartMs || oms > stockMs + 86400000 - 1) continue;
    if (!settings.includeRentalInSales && /^出租$/.test(String(ob.reason || "").trim())) continue;
    var uu = unitByKey.get(serialKey(ob.serial));
    var skey = (uu ? uu.model : "未识别") + "|" + (uu ? uu.memory : "") + "|" + (uu ? uu.condition : "未知");
    salesCount.set(skey, (salesCount.get(skey) || 0) + 1);
  }

  /* 周转：库存含在途；日销量 = 近窗销量/窗口天数；可售天数 = 库存含在途 ÷ 日销量 */
  var stockMap = new Map();
  function ensureRow(key) {
    if (!stockMap.has(key)) {
      var parts = key.split("|");
      stockMap.set(key, { key: key, model: parts[0], memory: parts[1], condition: parts[2], stock: 0, transit: 0, sales: salesCount.get(key) || 0, dailySales: 0, sellableDays: null, stockValue: 0 });
    }
    return stockMap.get(key);
  }
  units.forEach(function (x) {
    if (x.status === "不计入") return;
    var row = ensureRow(x.model + "|" + x.memory + "|" + x.condition);
    if (x.status === "在库") { row.stock++; row.stockValue += x.costVnd || 0; }
    else if (x.status === "在途") row.transit++;
  });
  salesCount.forEach(function (_, key) { ensureRow(key); });

  var turnover = Array.from(stockMap.values()).map(function (row) {
    var inclTransit = row.stock + row.transit;
    var daily = row.sales > 0 ? row.sales / windowDays : 0;
    row.dailySales = daily;
    row.stockInclTransit = inclTransit;
    row.sellableDays = daily > 0 ? Math.round(inclTransit / daily) : (inclTransit > 0 ? null : 0);
    return row;
  });
  /* 周转表固定按机型、内存、新旧排序，保证同一机型集中显示。 */
  function naturalKey(v) { return String(v || "").toLowerCase().replace(/[^a-z0-9]+/g, ""); }
  turnover.sort(function (a, b) {
    return naturalKey(a.model).localeCompare(naturalKey(b.model), "en", { numeric: true }) ||
      naturalKey(a.memory).localeCompare(naturalKey(b.memory), "en", { numeric: true }) ||
      ({"新机":0,"二手机":1,"未知":2}[a.condition] || 9) - ({"新机":0,"二手机":1,"未知":2}[b.condition] || 9);
  });

  return { units: units, kpi: kpi, turnover: turnover, windowDays: windowDays, cutoffDate: ageDate, ageDate: ageDate, stockDate: stockDate };
}
function snapshotCompute(snap) { return computeAll(snap); }

/* ---------- 导出 ---------- */
var CORE = {
  collapse: collapse, looksLikeSerial: looksLikeSerial, serialKey: serialKey, headerKey: headerKey,
  cellStr: cellStr, splitSerialNote: splitSerialNote,
  normalizeModel: normalizeModel, normalizeMemory: normalizeMemory,
  normalizeCondition: normalizeCondition, classifyWarehouse: classifyWarehouse,
  isProblemReason: isProblemReason, parseMoney: parseMoney, parseDateMs: parseDateMs,
  isoFromMs: isoFromMs, dateOnly: dateOnly, daysBetween: daysBetween, todayIso: todayIso,
  fmtInt: fmtInt, fmtMoney: fmtMoney, fmtDays: fmtDays,
  DEFAULT_MODEL_ALIASES: DEFAULT_MODEL_ALIASES, SOURCE_FIELDS: SOURCE_FIELDS,
  pickErpSheets: pickErpSheets, defaultSettings: defaultSettings, emptySnapshot: emptySnapshot,
  hasUserRows: hasUserRows, generateSample: generateSample,
  pushImportHistory: pushImportHistory, applyImportEntry: applyImportEntry, SOURCE_LABEL: SOURCE_LABEL,
  computeAll: computeAll, snapshotCompute: snapshotCompute,
  transitSheetDate: transitSheetDate,
  pickTransitDate: pickTransitDate,
  COLUMN_DEFS: COLUMN_DEFS
};
if (typeof module !== "undefined" && module.exports) module.exports = CORE;
global.CORE = CORE;
})(typeof window !== "undefined" ? window : globalThis);
