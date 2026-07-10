const http = require("http");
const fs = require("fs");
const path = require("path");
const https = require("https");

const host = "127.0.0.1";
const port = 4173;
const root = __dirname;
const API_FOOTBALL_KEY = process.env.API_FOOTBALL_KEY || "24ad8ed6257c8768b381665d7ccc318f";
const API_FOOTBALL_HOST = "v3.football.api-sports.io";

const mimeTypes = {
  ".html": "text/html; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".js": "application/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".jpeg": "image/jpeg",
  ".gif": "image/gif",
  ".svg": "image/svg+xml",
  ".ico": "image/x-icon",
  ".webp": "image/webp",
};

function send(res, status, body, type = "text/plain; charset=utf-8") {
  res.writeHead(status, { "Content-Type": type });
  res.end(body);
}

function sendJson(res, status, payload) {
  send(res, status, JSON.stringify(payload), "application/json; charset=utf-8");
}

function proxyApiFootball(req, res) {
  if (!API_FOOTBALL_KEY) {
    sendJson(res, 500, { error: "API_FOOTBALL_KEY no configurada" });
    return;
  }

  const incomingUrl = new URL(req.url || "/", `http://${host}:${port}`);
  const targetPath = incomingUrl.pathname.replace(/^\/api\/football/, "") || "/";
  const targetUrl = new URL(`https://${API_FOOTBALL_HOST}${targetPath}${incomingUrl.search}`);

  const proxyReq = https.request(
    targetUrl,
    {
      method: "GET",
      headers: {
        "x-apisports-key": API_FOOTBALL_KEY,
        Accept: "application/json",
      },
    },
    (proxyRes) => {
      const chunks = [];
      proxyRes.on("data", (chunk) => chunks.push(chunk));
      proxyRes.on("end", () => {
        const body = Buffer.concat(chunks);
        res.writeHead(proxyRes.statusCode || 502, {
          "Content-Type": proxyRes.headers["content-type"] || "application/json; charset=utf-8",
          "Access-Control-Allow-Origin": "*",
        });
        res.end(body);
      });
    }
  );

  proxyReq.on("error", (error) => {
    sendJson(res, 502, {
      error: "Error al conectar con API-Football",
      detail: error.message,
    });
  });

  proxyReq.end();
}

const server = http.createServer((req, res) => {
  if ((req.url || "").startsWith("/api/football/")) {
    proxyApiFootball(req, res);
    return;
  }

  const urlPath = decodeURIComponent((req.url || "/").split("?")[0]);
  const requested = urlPath === "/" ? "/index.html" : urlPath;
  const safePath = path.normalize(requested).replace(/^(\.\.[/\\])+/, "");
  const filePath = path.join(root, safePath);

  if (!filePath.startsWith(root)) {
    send(res, 403, "Forbidden");
    return;
  }

  fs.readFile(filePath, (err, data) => {
    if (err) {
      if (err.code === "ENOENT") {
        send(res, 404, "Not found");
        return;
      }
      send(res, 500, "Server error");
      return;
    }

    const ext = path.extname(filePath).toLowerCase();
    send(res, 200, data, mimeTypes[ext] || "application/octet-stream");
  });
});

server.listen(port, host, () => {
  console.log(`FutBrain disponible en http://${host}:${port}`);
});
