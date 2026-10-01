/* 导入桥接：C仓/三合一优先用原生流式解析器（大文件不爆内存），
 * 其他仓/系统出库（多工作表合并）一律走 SheetJS 通道；
 * 原生解析返回 0 行或报格式错误时自动回退 SheetJS。 */
(function () {
"use strict";
if (typeof NATIVE_XLSX === "undefined" || !window.EXCELIO) return;
var sheetjsSingle = window.EXCELIO.importSingle;
var sheetjsErp = window.EXCELIO.importErp;
function normalizeNativeTransit(res) {
  if (!res || !Array.isArray(res.rows)) return;
  var KEYS = ["date","采购月份","purchaseMonth","purchaseDate","month","月份","制表日期","日期","采购日期","下单日期","etaDate","预计到货日","inTransitDate","在途日期"];
  var sheet = res.sheetName || "";
  res.rows.forEach(function (r) {
    if (!r) return;
    for (var i = 0; i < KEYS.length; i++) { var v = r[KEYS[i]]; if (v) { r.date = v; break; } }
    if (!r.date) {
      for (var k in r) { if (r[k] && typeof r[k] === "string" && /(采购)?(月份|日期)|date|month/i.test(k)) { r.date = r[k]; break; } }
      if (r instanceof Date) r.date = r.getTime ? new Date(r.getTime()).toISOString().slice(0,10) : r.date;
    }
    if (sheet && !r._sheet) r._sheet = sheet;
  });
}
var NATIVE_KINDS = { warehouse: 1, inbound: 1, outbound: 1, transit: 1 };
function isFormatErr(e) { return /不是有效的|不是 Excel|请上传/.test(e && e.message || ""); }
window.EXCELIO.importSingle = function (file, kind) {
  /* 其他仓 / 系统出库：必须走 SheetJS 的多子表合并逻辑 */
  if (!NATIVE_KINDS[kind]) return sheetjsSingle(file, kind);
  return NATIVE_XLSX.previewExcel(file, kind).then(function (res) {
    if (!res || !res.rows || !res.rows.length) return sheetjsSingle(file, kind);
    if (kind === "transit") normalizeNativeTransit(res);
    return res;
  }).catch(function (e) {
    if (isFormatErr(e)) return sheetjsSingle(file, kind);
    throw e;
  });
};
window.EXCELIO.importErp = function (file) {
  return NATIVE_XLSX.previewErp(file).then(function (res) {
    if (!res || (!res.inbound && !res.outbound && !res.transit)) return sheetjsErp(file);
    return res;
  }).catch(function (e) {
    if (isFormatErr(e)) return sheetjsErp(file);
    throw e;
  });
};
})();
