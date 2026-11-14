// GET /api/keys — en la nube no hay archivos de clave: cada usuario pone la suya en Ajustes.
// "cloud: true" le dice al panel que muestre las instrucciones de la versión en la nube.
const { sendJson } = require("./_lib");

module.exports = (req, res) => {
  const server = process.env.PLUMIA_SERVER_KEYS === "1";
  sendJson(res, 200, {
    cloud: true,
    nvidia: server && !!process.env.NVIDIA_API_KEY,
    openrouter: server && !!process.env.OPENROUTER_API_KEY,
    deepseek: server && !!process.env.DEEPSEEK_API_KEY,
  });
};
