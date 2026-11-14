// Servidor local HTTPS para el complemento "Plumia".
// Sirve /public y reenvía /api/chat a DeepSeek (evita problemas de CORS).
// Sin dependencias: solo Node.js.
const https = require("https");
const fs = require("fs");
const path = require("path");

const PORT = 3443;
const LOG = path.join(__dirname, "server.log");
const MAX_LOG = 1024 * 1024; // el registro se rota al pasar de 1 MB
const MAX_BODY = 20 * 1024 * 1024; // tamaño máximo de una petición
const UPSTREAM_TIMEOUT = 120000; // tiempo máximo esperando al servicio de IA
function log(msg) {
  try {
    if (fs.existsSync(LOG) && fs.statSync(LOG).size > MAX_LOG) fs.renameSync(LOG, LOG + ".old");
    fs.appendFileSync(LOG, new Date().toISOString() + " " + msg + "\n");
  } catch (e) { /* ignore */ }
}
// Un error inesperado no debe tumbar el servidor.
process.on("uncaughtException", (e) => log("ERROR no capturado: " + (e && e.stack ? e.stack.split("\n")[0] : e)));
process.on("unhandledRejection", (e) => log("PROMESA rechazada: " + (e && e.message ? e.message : e)));
const PUBLIC = path.join(__dirname, "public");
const CERT_DIR = path.join(__dirname, "certs");
const PROVIDERS = {
  nvidia: "https://integrate.api.nvidia.com/v1/chat/completions",
  deepseek: "https://api.deepseek.com/chat/completions",
  openrouter: "https://openrouter.ai/api/v1/chat/completions",
};

const MIME = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".png": "image/png",
  ".svg": "image/svg+xml",
  ".json": "application/json",
};

const tls = {
  pfx: fs.readFileSync(path.join(CERT_DIR, "localhost.pfx")),
  passphrase: "deepseek-word",
};

async function readBody(req) {
  const chunks = [];
  let size = 0;
  for await (const c of req) {
    size += c.length;
    if (size > MAX_BODY) throw Object.assign(new Error("petición demasiado grande"), { status: 413 });
    chunks.push(c);
  }
  return Buffer.concat(chunks);
}

// Seguridad: solo se aceptan peticiones dirigidas a este equipo y hechas desde el propio panel.
// Evita que otra página web abierta en el navegador use el servidor (CSRF / DNS rebinding).
const ALLOWED_HOSTS = new Set([`localhost:${PORT}`, `127.0.0.1:${PORT}`, `[::1]:${PORT}`]);
const ALLOWED_ORIGIN = `https://localhost:${PORT}`;
function isTrusted(req, isApi) {
  if (!ALLOWED_HOSTS.has(String(req.headers.host || "").toLowerCase())) return false;
  const origin = req.headers.origin;
  if (origin && origin !== ALLOWED_ORIGIN && origin !== "null") return false;
  if (isApi && req.method === "POST" && !/^application\/json/i.test(req.headers["content-type"] || "")) return false;
  return true;
}

// Clave guardada en un archivo de texto junto al servidor (p. ej. clave-nvidia.txt).
function fileKey(provider) {
  try {
    return fs.readFileSync(path.join(__dirname, `clave-${provider}.txt`), "utf8").replace(/^﻿/, "").trim();
  } catch (e) {
    return "";
  }
}

// Catálogo en vivo de modelos de cada servicio (se guarda 1 hora en memoria).
const CATALOGS = {
  nvidia: "https://integrate.api.nvidia.com/v1/models",
  openrouter: "https://openrouter.ai/api/v1/models",
  deepseek: "https://api.deepseek.com/models",
};
// Modelos que no son de chat/texto (embeddings, voz, imagen, seguridad...).
const NOT_CHAT = /(embed|rerank|guard|safety|reward|vision-?only|vlm-embed|parse|ocr|retriev|pii|audio|asr|tts|clip|chatqa|diffusion|sdxl|flux|stable-?diffusion|whisper|riva|fastpitch|video-detector|calibration|nv-embed|nemoretriever|deplot|kosmos|paligemma|cosmos|bge|arctic-embed)/i;
const modelCache = {};

async function listModels(provider, key) {
  const hit = modelCache[provider];
  if (hit && Date.now() - hit.at < 3600e3) return hit.models;
  const headers = provider === "deepseek" && key ? { Authorization: "Bearer " + key } : {};
  const r = await fetch(CATALOGS[provider], { headers });
  if (!r.ok) throw new Error("catálogo " + r.status);
  const data = (await r.json()).data || [];
  let models = data
    .filter((m) => !NOT_CHAT.test(m.id))
    // Solo modelos que responden con texto (fuera música, imagen, vídeo...).
    .filter((m) => !m.architecture || (m.architecture.output_modalities || ["text"]).every((x) => x === "text"))
    .map((m) => ({
      id: m.id,
      name: m.name || m.id,
      // En OpenRouter solo es gratis lo que termina en ":free" (el precio 0 de texto no basta).
      free: provider === "nvidia" || /:free$/.test(m.id),
    }))
    .sort((a, b) => a.id.localeCompare(b.id));
  modelCache[provider] = { at: Date.now(), models };
  log(`catálogo ${provider}: ${models.length} modelos`);
  return models;
}

async function proxyChat(req, res) {
  const provider = req.headers["x-provider"] || "deepseek";
  const key = (req.headers["x-api-key"] || "").trim() || fileKey(provider);
  const target = PROVIDERS[provider];
  if (!key || !target) {
    log(`chat provider=${provider} sin clave (ni en el panel ni en clave-${provider}.txt)`);
    res.writeHead(401, { "Content-Type": "application/json" });
    return res.end(JSON.stringify({ error: { message: `Falta la clave. Pégala en Ajustes o en el archivo clave-${provider}.txt de la carpeta DeepSeekWordAddin.` } }));
  }
  // Si el panel se cierra o se pulsa Detener, se cancela también la petición a la IA.
  const abort = new AbortController();
  const timer = setTimeout(() => abort.abort(), UPSTREAM_TIMEOUT);
  res.on("close", () => { if (!res.writableEnded) abort.abort(); });
  try {
    const body = await readBody(req);
    const upstream = await fetch(target, {
      method: "POST",
      signal: abort.signal,
      headers: {
        "Content-Type": "application/json",
        Authorization: "Bearer " + key,
        "X-Title": "Plumia",
      },
      body,
    });
    let model = "?";
    try { model = JSON.parse(body.toString()).model; } catch (e) { /* ignore */ }
    log(`chat provider=${provider} model=${model} status=${upstream.status}`);
    if (!upstream.ok) {
      const text = await upstream.text();
      log("  error body: " + text.slice(0, 500));
      let original = text.slice(0, 300);
      try { const j = JSON.parse(text); original = j.error?.message || j.detail || j.title || original; } catch (e) { /* not json */ }
      // NVIDIA responde 403 "Authorization failed" cuando la clave es mala: se trata como 401.
      const status = upstream.status === 403 && /authoriz|unauthori|invalid.*(key|token)|api key/i.test(String(original)) ? 401 : upstream.status;
      res.writeHead(status, { "Content-Type": "application/json" });
      const keyHelp = { nvidia: "build.nvidia.com → API Keys (nvapi-...)", openrouter: "openrouter.ai → Keys (sk-or-...)", deepseek: "platform.deepseek.com" }[provider];
      let message = original;
      // 401 = la clave es mala. 403 suele ser "este modelo no está permitido", no un problema de la clave.
      if (status === 401) message = `la clave de ${provider} no es válida o está incompleta. Genera una nueva en ${keyHelp}.`;
      else if (status === 403) message = `acceso denegado a este modelo (${original}).`;
      else if (status === 402) message = "este modelo es de pago y no hay saldo. Se usará otro gratis.";
      else if (status === 404) message = "ese modelo ya no está disponible.";
      else if (status === 429) message = "modelo saturado o límite gratis alcanzado; se probará otro.";
      return res.end(JSON.stringify({ error: { message: `Error ${status} del servicio: ${message}`, code: status } }));
    }
    res.writeHead(upstream.status, {
      "Content-Type": upstream.headers.get("content-type") || "application/json",
      "Cache-Control": "no-cache",
    });
    if (!upstream.body) return res.end();
    for await (const chunk of upstream.body) res.write(chunk);
    res.end();
  } catch (e) {
    const status = e.status || (e.name === "AbortError" ? 504 : 502);
    log(`  exception (${status}): ${e.message}`);
    if (res.destroyed) return;
    if (!res.headersSent) res.writeHead(status, { "Content-Type": "application/json" });
    const msg = status === 504 ? "el servicio de IA tardó demasiado." : status === 413 ? "la petición es demasiado grande." : "no se pudo contactar con el servicio de IA: " + e.message;
    res.end(JSON.stringify({ error: { message: msg, code: status } }));
  } finally {
    clearTimeout(timer);
  }
}

function handler(req, res) {
  try {
    handle(req, res);
  } catch (e) {
    log("error en petición: " + e.message);
    if (!res.headersSent) res.writeHead(400);
    res.end();
  }
}

function handle(req, res) {
    const url = new URL(req.url, "https://localhost:" + PORT);
    const isApi = url.pathname.startsWith("/api/");
    log(`${req.method} ${url.pathname}`);
    if (!isTrusted(req, isApi)) {
      log(`  bloqueada (host=${req.headers.host}, origin=${req.headers.origin || "-"})`);
      res.writeHead(403);
      return res.end();
    }
    if (req.method === "POST" && url.pathname === "/api/chat") return proxyChat(req, res);
    if (url.pathname === "/api/models") {
      const p = url.searchParams.get("provider") || "nvidia";
      if (!CATALOGS[p]) { res.writeHead(400); return res.end("[]"); }
      const key = (req.headers["x-api-key"] || "").trim() || fileKey(p);
      return listModels(p, key)
        .then((models) => {
          res.writeHead(200, { "Content-Type": "application/json", "Cache-Control": "no-store" });
          res.end(JSON.stringify(models));
        })
        .catch((e) => {
          log(`catálogo ${p} error: ${e.message}`);
          res.writeHead(502, { "Content-Type": "application/json" });
          res.end(JSON.stringify({ error: e.message }));
        });
    }
    if (url.pathname === "/api/keys") {
      const status = {};
      for (const p of Object.keys(PROVIDERS)) status[p] = !!fileKey(p);
      res.writeHead(200, { "Content-Type": "application/json", "Cache-Control": "no-store" });
      return res.end(JSON.stringify(status));
    }

    if (req.method !== "GET" && req.method !== "HEAD") {
      res.writeHead(405);
      return res.end();
    }
    let rel;
    try { rel = decodeURIComponent(url.pathname); } catch (e) { res.writeHead(400); return res.end(); }
    if (rel === "/") rel = "/taskpane.html";
    const file = path.normalize(path.join(PUBLIC, rel));
    if (!file.startsWith(PUBLIC + path.sep)) {
      res.writeHead(403);
      return res.end();
    }
    fs.readFile(file, (err, data) => {
      if (err) {
        res.writeHead(404);
        return res.end("No encontrado");
      }
      res.writeHead(200, {
        "Content-Type": MIME[path.extname(file)] || "application/octet-stream",
        "Cache-Control": "no-store",
      });
      res.end(data);
    });
}

// Escucha en IPv4 e IPv6 de localhost (solo este equipo), para que Word llegue siempre aquí.
for (const host of ["127.0.0.1", "::1"]) {
  https
    .createServer(tls, handler)
    .on("error", (e) => log(`listen ${host}:${PORT} error: ${e.message}`))
    .listen(PORT, host, () => log(`escuchando https://${host.includes(":") ? "[" + host + "]" : host}:${PORT}`));
}
