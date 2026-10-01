/* ============================================================
 * 进销存 · 云端版服务器（零依赖 Node.js，宝塔面板可直接部署）
 * 用法：
 *   1) 宝塔 → 文件 → 上传本文件夹到 /www/wwwroot/jinxiaocun （任意目录）
 *   2) 宝塔 → 网站 → Node项目 → 添加项目：启动文件 server.js，端口 8321
 *      （或终端执行：cd 目录 && node server.js ，生产建议 pm2 start server.js）
 *   3) 浏览器访问 http://服务器IP:8321 或绑定域名
 * 数据保存在服务器 data/snapshot.json，记得在宝塔计划任务里定期备份该文件。
 * ============================================================ */
const http = require("http");
const fs = require("fs");
const path = require("path");

const ROOT = __dirname;
const DATA_DIR = path.join(ROOT, "data");
const SNAP_FILE = path.join(DATA_DIR, "snapshot.json");
const PORT = Number(process.env.PORT || 8321);

if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });

const MIME = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".txt": "text/plain; charset=utf-8",
  ".png": "image/png", ".jpg": "image/jpeg", ".svg": "image/svg+xml",
  ".ico": "image/x-icon"
};

http.createServer(function (req, res) {
  const url = decodeURIComponent((req.url || "/").split("?")[0]);

  /* 数据 API */
  if (url === "/api/snapshot") {
    if (req.method === "GET") {
      fs.readFile(SNAP_FILE, function (e, d) {
        if (e) { res.writeHead(200, { "Content-Type": "application/json" }); res.end("{}"); }
        else { res.writeHead(200, { "Content-Type": "application/json" }); res.end(d); }
      });
      return;
    }
    if (req.method === "PUT" || req.method === "POST") {
      let body = "";
      req.on("data", function (c) { body += c; if (body.length > 200 * 1024 * 1024) req.destroy(); });
      req.on("end", function () {
        try {
          /* 写入前自动备份上一份 */
          if (fs.existsSync(SNAP_FILE)) {
            fs.copyFileSync(SNAP_FILE, path.join(DATA_DIR, "snapshot.bak.json"));
          }
          fs.writeFileSync(SNAP_FILE, body || "{}");
          res.writeHead(200, { "Content-Type": "application/json" });
          res.end('{"ok":true}');
        } catch (err) {
          res.writeHead(500, { "Content-Type": "application/json" });
          res.end(JSON.stringify({ ok: false, error: String(err) }));
        }
      });
      return;
    }
    res.writeHead(405); res.end(); return;
  }

  /* 静态文件 */
  let p = path.normalize(path.join(ROOT, url === "/" ? "index.html" : url));
  if (!p.startsWith(ROOT)) { res.writeHead(403); res.end(); return; }
  fs.readFile(p, function (e, d) {
    if (e) { res.writeHead(404); res.end("Not found"); return; }
    res.writeHead(200, { "Content-Type": MIME[path.extname(p).toLowerCase()] || "application/octet-stream" });
    res.end(d);
  });
}).listen(PORT, function () {
  console.log("进销存云端版已启动：http://0.0.0.0:" + PORT);
  console.log("数据文件：" + SNAP_FILE);
});
