// GET /api/models?provider=nvidia|openrouter|deepseek — catálogo en vivo de modelos de texto.
const { CATALOGS, NOT_CHAT, originAllowed, resolveKey, sendJson } = require("./_lib");

const cache = {}; // se reutiliza mientras la función siga activa (máx. 1 hora)

module.exports = async (req, res) => {
  if (req.method !== "GET") return sendJson(res, 405, { error: "Usa GET." });
  if (!originAllowed(req)) return sendJson(res, 403, { error: "Petición no permitida." });
  const url = new URL(req.url, "https://localhost");
  const provider = url.searchParams.get("provider") || "nvidia";
  if (!CATALOGS[provider]) return sendJson(res, 400, []);
  const hit = cache[provider];
  if (hit && Date.now() - hit.at < 3600e3) return sendJson(res, 200, hit.models);
  try {
    const key = resolveKey(req, provider);
    const headers = provider === "deepseek" && key ? { Authorization: "Bearer " + key } : {};
    const r = await fetch(CATALOGS[provider], { headers });
    if (!r.ok) throw new Error("catálogo " + r.status);
    const data = (await r.json()).data || [];
    const models = data
      .filter((m) => !NOT_CHAT.test(m.id))
      .filter((m) => !m.architecture || (m.architecture.output_modalities || ["text"]).every((x) => x === "text"))
      .map((m) => ({ id: m.id, name: m.name || m.id, free: provider === "nvidia" || /:free$/.test(m.id) }))
      .sort((a, b) => a.id.localeCompare(b.id));
    cache[provider] = { at: Date.now(), models };
    sendJson(res, 200, models);
  } catch (e) {
    sendJson(res, 502, { error: e.message });
  }
};
