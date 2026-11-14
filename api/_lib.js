// Código compartido por las funciones en la nube de Plumia (Vercel).
// Los archivos que empiezan por "_" no se publican como direcciones.

const PROVIDERS = {
  nvidia: "https://integrate.api.nvidia.com/v1/chat/completions",
  openrouter: "https://openrouter.ai/api/v1/chat/completions",
  deepseek: "https://api.deepseek.com/chat/completions",
};

const CATALOGS = {
  nvidia: "https://integrate.api.nvidia.com/v1/models",
  openrouter: "https://openrouter.ai/api/v1/models",
  deepseek: "https://api.deepseek.com/models",
};

// Modelos que no son de chat/texto (embeddings, voz, imagen, seguridad...).
const NOT_CHAT = /(embed|rerank|guard|safety|reward|vision-?only|vlm-embed|parse|ocr|retriev|pii|audio|asr|tts|clip|chatqa|diffusion|sdxl|flux|stable-?diffusion|whisper|riva|fastpitch|video-detector|calibration|nv-embed|nemoretriever|deplot|kosmos|paligemma|cosmos|bge|arctic-embed)/i;

const UPSTREAM_TIMEOUT = 120000;

// Solo se aceptan peticiones hechas desde el propio panel (mismo dominio) o sin Origin.
function originAllowed(req) {
  const origin = req.headers.origin;
  if (!origin || origin === "null") return true;
  try {
    return new URL(origin).host === req.headers.host;
  } catch (e) {
    return false;
  }
}

// Clave del usuario (cabecera). Solo si se activa PLUMIA_SERVER_KEYS=1 se usan las claves
// del servidor (variables de entorno); por defecto está desactivado para que nadie gaste
// tu cuota sin pagar.
function resolveKey(req, provider) {
  const k = String(req.headers["x-api-key"] || "").replace(/[^\x21-\x7E]/g, "");
  if (k) return k;
  if (process.env.PLUMIA_SERVER_KEYS === "1") return process.env[`${provider.toUpperCase()}_API_KEY`] || "";
  return "";
}

function friendlyError(status, provider, original) {
  const keyHelp = { nvidia: "build.nvidia.com → API Keys (nvapi-...)", openrouter: "openrouter.ai → Keys (sk-or-...)", deepseek: "platform.deepseek.com" }[provider];
  if (status === 401) return `la clave de ${provider} no es válida o está incompleta. Genera una nueva en ${keyHelp}.`;
  if (status === 403) return `acceso denegado a este modelo (${original}).`;
  if (status === 402) return "este modelo es de pago y no hay saldo. Se usará otro gratis.";
  if (status === 404) return "ese modelo ya no está disponible.";
  if (status === 429) return "modelo saturado o límite gratis alcanzado; se probará otro.";
  return original;
}

// Algunos servicios (NVIDIA) responden 403 cuando la clave es mala: se trata como 401
// para que el panel no pruebe todos los modelos de ese servicio inútilmente.
function effectiveStatus(status, original) {
  if (status === 403 && /authoriz|unauthori|invalid.*(key|token)|api key/i.test(String(original))) return 401;
  return status;
}

// Límite DIARIO de la cuenta (p. ej. OpenRouter: 50 peticiones gratis al día). Afecta a todos los
// modelos gratis del servicio, así que el panel no debe seguir probando otros: se marca con quota="daily".
function quotaInfo(status, bodyText) {
  if (status !== 429 || !/free-models-per-day|per[- ]day|daily/i.test(String(bodyText))) return {};
  let resetAt;
  try { const r = JSON.parse(bodyText).error?.metadata?.headers?.["X-RateLimit-Reset"]; if (r) resetAt = Number(r); } catch (e) { /* sin hora de reinicio */ }
  return { quota: "daily", resetAt };
}

function sendJson(res, status, obj) {
  res.statusCode = status;
  res.setHeader("Content-Type", "application/json; charset=utf-8");
  res.setHeader("Cache-Control", "no-store");
  res.end(JSON.stringify(obj));
}

module.exports = { PROVIDERS, CATALOGS, NOT_CHAT, UPSTREAM_TIMEOUT, originAllowed, resolveKey, friendlyError, effectiveStatus, quotaInfo, sendJson };
