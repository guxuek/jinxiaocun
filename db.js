/* ============================================================
 * 数据存储层 v7 —— 双模式
 *   · 本地版（file:// 双击打开）：数据写入「用户选择的文件夹」里的
 *     snapshot.json（建议选在 index.html 旁边的“进销存数据”文件夹）。
 *     使用浏览器 File System Access API（Edge/Chrome 支持）。
 *     句柄持久在 IndexedDB，之后自动保存；首次需点一次“选择数据文件夹”。
 *   · 云端版（http(s):// 宝塔部署）：数据通过 /api/snapshot 存到
 *     服务器 data/snapshot.json，多人共用一份数据。
 *   任何情况下都在 IndexedDB 留一份兜底副本，防止误操作丢数据。
 * ============================================================ */
(function (global) {
"use strict";

var MODE = (typeof location !== "undefined" && location.protocol === "file:") ? "local-file" : "cloud";
var IDB_NAME = "ims-local-db", IDB_STORE = "kv";
var SNAP_FILE = "snapshot.json";

/* ---------- IndexedDB 极简 KV ---------- */
function idbOpen() {
  return new Promise(function (resolve, reject) {
    var req = indexedDB.open(IDB_NAME, 1);
    req.onupgradeneeded = function () { req.result.createObjectStore(IDB_STORE); };
    req.onsuccess = function () { resolve(req.result); };
    req.onerror = function () { reject(req.error); };
  });
}
function idbGet(key) {
  return idbOpen().then(function (db) {
    return new Promise(function (resolve, reject) {
      var tx = db.transaction(IDB_STORE, "readonly");
      var rq = tx.objectStore(IDB_STORE).get(key);
      rq.onsuccess = function () { resolve(rq.result == null ? null : rq.result); };
      rq.onerror = function () { reject(rq.error); };
    });
  });
}
function idbSet(key, val) {
  return idbOpen().then(function (db) {
    return new Promise(function (resolve, reject) {
      var tx = db.transaction(IDB_STORE, "readwrite");
      tx.objectStore(IDB_STORE).put(val, key);
      tx.oncomplete = function () { resolve(); };
      tx.onerror = function () { reject(tx.error); };
    });
  });
}
function idbDel(key) {
  return idbOpen().then(function (db) {
    return new Promise(function (resolve, reject) {
      var tx = db.transaction(IDB_STORE, "readwrite");
      tx.objectStore(IDB_STORE).delete(key);
      tx.oncomplete = function () { resolve(); };
      tx.onerror = function () { reject(tx.error); };
    });
  });
}

/* ---------- 本地文件（File System Access） ---------- */
var dirHandle = null, savedHandle = null, writePermission = false;
function fsSupported() { return typeof window !== "undefined" && !!window.showDirectoryPicker; }
function loadHandle() {
  return idbGet("dir-handle").then(function (h) {
    if (!h) return null;
    savedHandle = h;
    return h.queryPermission({ mode: "readwrite" }).then(function (p) {
      if (p === "granted") { dirHandle = h; writePermission = true; return h; }
      return null;
    });
  }).catch(function () { return null; });
}
function pickFolder() {
  if (!fsSupported()) return Promise.reject(new Error("当前浏览器不支持文件夹存储，请使用 Edge 或 Chrome"));
  return window.showDirectoryPicker({ mode: "readwrite" }).then(function (h) {
    dirHandle = h; savedHandle = h; writePermission = true;
    return idbSet("dir-handle", h).then(function () { return h; });
  });
}
function writeSnapshotFile(obj) {
  if (!dirHandle) return Promise.resolve(false);
  return dirHandle.getFileHandle(SNAP_FILE, { create: true }).then(function (fh) {
    return fh.createWritable().then(function (w) {
      return w.write(JSON.stringify(obj)).then(function () { return w.close(); }).then(function () { return true; });
    });
  }).catch(function () { return false; });
}
function readSnapshotFile() {
  if (!dirHandle) return Promise.resolve(null);
  return dirHandle.getFileHandle(SNAP_FILE).then(function (fh) {
    return fh.getFile().then(function (f) { return f.text(); });
  }).then(function (t) {
    try { return JSON.parse(t); } catch (e) { return null; }
  }).catch(function () { return null; });
}

/* ---------- 云端 API ---------- */
function cloudLoad() {
  return fetch("/api/snapshot", { cache: "no-store" }).then(function (r) {
    if (!r.ok) return null;
    return r.json().then(function (j) { return (j && Object.keys(j).length) ? j : null; });
  }).catch(function () { return null; });
}
function cloudSave(obj) {
  return fetch("/api/snapshot", {
    method: "PUT", headers: { "Content-Type": "application/json" }, body: JSON.stringify(obj)
  }).then(function (r) { return r.ok; }).catch(function () { return false; });
}

/* ---------- 对外接口 ---------- */
var DB = {
  mode: MODE,
  storageMode: MODE === "cloud" ? "云端服务器 data/snapshot.json" : "本地文件夹 snapshot.json",
  fsSupported: fsSupported,
  folderName: function () { return dirHandle ? dirHandle.name : ""; },
  hasFolder: function () { return !!dirHandle && writePermission; },
  hasSavedFolder: function () { return !!savedHandle; },
  resumeFolder: function () {
    var self = this;
    if (!savedHandle) return pickFolder();
    return savedHandle.requestPermission({ mode: "readwrite" }).then(function (p) {
      if (p !== "granted") throw new Error("浏览器未授权，请重新选择数据文件夹");
      dirHandle = savedHandle; writePermission = true;
      self.storageMode = "本地文件夹 " + savedHandle.name + "\\" + SNAP_FILE;
      return savedHandle;
    });
  },
  pickFolder: function () {
    var self = this;
    return pickFolder().then(function (h) {
      /* 选完文件夹立刻把当前数据落盘 */
      return idbGet("snapshot").then(function (snap) {
        if (snap) return writeSnapshotFile(snap);
        return false;
      }).then(function () {
        self.storageMode = "本地文件夹 " + h.name + "\\" + SNAP_FILE;
        return h;
      });
    });
  },
  loadSnapshot: function () {
    var self = this;
    if (MODE === "cloud") {
      return cloudLoad().then(function (j) {
        if (j) return j;
        return idbGet("snapshot");
      });
    }
    return loadHandle().then(function (h) {
      if (h) {
        self.storageMode = "本地文件夹 " + h.name + "\\" + SNAP_FILE;
        return readSnapshotFile().then(function (j) {
          if (j) return j;
          return idbGet("snapshot");
        });
      }
      return idbGet("snapshot");
    });
  },
  saveSnapshot: function (snap) {
    if (MODE === "cloud") {
      return cloudSave(snap).then(function (ok) {
        return idbSet("snapshot", snap).then(function () { return ok; });
      });
    }
    return writeSnapshotFile(snap).then(function () {
      return idbSet("snapshot", snap);
    });
  },
  clearSnapshot: function () {
    /* 清空 = 连本地文件夹里的 snapshot.json 一起删除（不留空壳文件） */
    if (MODE === "cloud") return cloudSave({}).then(function () { return idbDel("snapshot"); });
    var delFile = dirHandle
      ? dirHandle.removeEntry(SNAP_FILE).catch(function () { return writeSnapshotFile({}); })
      : Promise.resolve();
    return delFile.then(function () { return idbDel("snapshot"); });
  }
};

global.DB = DB;
})(typeof window !== "undefined" ? window : globalThis);
