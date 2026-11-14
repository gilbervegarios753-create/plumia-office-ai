// POST /api/chat — reenvía la conversación al servicio de IA elegido (versión en la nube).
const { PROVIDERS, UPSTREAM_TIMEOUT, originAllowed, resolveKey, friendlyError, effectiveStatus, sendJson } = require("./_lib");

const MAX_BODY = 4 * 1024 * 1024; // 4 MB por petición

async function readBody(req) {
  if (req.body !== undefined && req.body !== null) {
    return typeof req.body === "string" ? req.body : Buffer.isBuffer(req.body) ? req.body.toString("utf8") : JSON.stringify(req.body);
  }
  const chunks = [];
  let size = 0;
  for await (const c of req) {
    size += c.length;
    if (size > MAX_BODY) throw Object.assign(new Error("petición demasiado grande"), { status: 413 });
    chunks.push(c);
  }
  return Buffer.concat(chunks).toString("utf8");
}

module.exports = async (req, res) => {
  if (req.method !== "POST") return sendJson(res, 405, { error: { message: "Usa POST." } });
  if (!originAllowed(req) || !/^application\/json/i.test(req.headers["content-type"] || "")) {
    return sendJson(res, 403, { error: { message: "Petición no permitida." } });
  }
  const provider = String(req.headers["x-provider"] || "nvidia");
  const target = PROVIDERS[provider];
  const key = resolveKey(req, provider);
  if (!target) return sendJson(res, 400, { error: { message: "Servicio desconocido." } });
  if (!key) return sendJson(res, 401, { error: { message: `Falta la clave de ${provider}. Pégala en Ajustes (es gratis).`, code: 401 } });

  const abort = new AbortController();
  const timer = setTimeout(() => abort.abort(), UPSTREAM_TIMEOUT);
  res.on("close", () => { if (!res.writableEnded) abort.abort(); });
  try {
    const body = await readBody(req);
    if (body.length > MAX_BODY) return sendJson(res, 413, { error: { message: "La petición es demasiado grande.", code: 413 } });
    const upstream = await fetch(target, {
      method: "POST",
      signal: abort.signal,
      headers: { "Content-Type": "application/json", Authorization: "Bearer " + key, "X-Title": "Plumia" },
      body,
    });
    const text = await upstream.text();
    if (!upstream.ok) {
      let original = text.slice(0, 300);
      try { const j = JSON.parse(text); original = (j.error && j.error.message) || j.detail || j.title || original; } catch (e) { /* no es JSON */ }
      const status = effectiveStatus(upstream.status, original);
      return sendJson(res, status, { error: { message: `Error ${status} del servicio: ${friendlyError(status, provider, original)}`, code: status } });
    }
    res.statusCode = 200;
    res.setHeader("Content-Type", upstream.headers.get("content-type") || "application/json");
    res.setHeader("Cache-Control", "no-store");
    res.end(text);
  } catch (e) {
    const status = e.status || (e.name === "AbortError" ? 504 : 502);
    if (res.writableEnded || res.destroyed) return;
    const msg = status === 504 ? "el servicio de IA tardó demasiado." : status === 413 ? "la petición es demasiado grande." : "no se pudo contactar con el servicio de IA.";
    sendJson(res, status, { error: { message: msg, code: status } });
  } finally {
    clearTimeout(timer);
  }
};
