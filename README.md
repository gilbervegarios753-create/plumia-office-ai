# ✒️ Plumia — tu equipo de IA dentro de Office

**Escribe lo que quieres y Plumia lo hace en tu documento.** Plumia es un complemento para **Word, Excel, PowerPoint y Outlook** con un chat de IA que **trabaja directamente en tu archivo**: redacta, corrige, da formato, crea tablas y diapositivas. No hay que copiar ni pegar nada.

> Funciona con modelos **gratuitos** (DeepSeek vía NVIDIA, y más de 400 modelos en OpenRouter). Si un modelo va lento o falla, **cambia solo a otro** y sigue trabajando.

---

## ✨ Qué hace

| Programa | Lo que Plumia puede hacer |
|---|---|
| **Word** | Escribir textos con títulos, listas y citas · corregir ortografía y gramática · mejorar, resumir, ampliar o traducir · buscar y reemplazar · dar formato (negrita, color, tamaño, alineación, estilos) · crear tablas · reorganizar párrafos |
| **Excel** | Leer la hoja · crear tablas con datos y fórmulas · dar formato y ajustar columnas |
| **PowerPoint** | Crear diapositivas con título y contenido · leer y mejorar el texto |
| **Outlook** | Resumir correos · redactar respuestas en el correo |

**Lo que lo hace diferente:**
- 🤖 **Agente autónomo:** lee el documento, decide qué hacer y aplica los cambios él mismo. **Ctrl+Z** lo deshace todo.
- 👥 **Multiagente:** las tareas grandes, como un informe con varias secciones, se reparten entre hasta 4 agentes que escriben **en paralelo**.
- ⚡ **Modo automático:** elige el mejor modelo gratuito disponible y cambia de modelo si uno va lento, se satura o falla.
- 🔒 **Privado:** todo pasa por un servidor local en tu PC. Tus claves se guardan solo en tu equipo.

---

## 🧱 Cómo funciona

```
Word / Excel / PowerPoint / Outlook
        │  (panel lateral: chat + herramientas)
        ▼
Servidor local HTTPS  →  https://localhost:3443   (server.js, sin dependencias)
        │  (reenvía las peticiones y protege tu clave)
        ▼
Servicio de IA: NVIDIA · OpenRouter · DeepSeek
```

| Archivo | Para qué sirve |
|---|---|
| `manifest.xml` | Registra el complemento en Word, Excel y PowerPoint (botón en la pestaña Inicio) |
| `manifest-outlook.xml` | Registra el complemento en Outlook (leer y redactar correos) |
| `server.js` | Servidor local HTTPS: sirve el panel y reenvía las peticiones a la IA |
| `public/taskpane.*` | Panel con chat, ajustes, cambio automático de modelo y equipo de agentes |
| `public/tools.js` | Herramientas que la IA usa en cada programa (leer, escribir, formato…) |
| `setup.ps1` | Instalador para Windows |

---

## ☁️ Instalación rápida (recomendada)

Todo está en **https://plumia-office-ai.vercel.app** — la página tiene el botón de descarga y los pasos para cada caso:

| Dónde | Cómo |
|---|---|
| **Windows** (Word, Excel, PowerPoint) | Descarga [`instalar-plumia.ps1`](public/descargas/instalar-plumia.ps1) → clic derecho → *Ejecutar con PowerShell* → reabre Word → **Inicio → Complementos → Más complementos → CARPETA COMPARTIDA → Plumia → Agregar** |
| **Word en la web / Mac** | Descarga [`plumia.xml`](public/descargas/plumia.xml) → *Insertar → Complementos → Mis complementos → Cargar mi complemento* |
| **Outlook** | Descarga [`plumia-outlook.xml`](public/descargas/plumia-outlook.xml) → *Obtener complementos → Mis complementos → Agregar desde archivo* |
| **Empresas y colegios** | Centro de administración de Microsoft 365 → *Aplicaciones integradas → Cargar aplicaciones personalizadas* → vínculo `https://plumia-office-ai.vercel.app/descargas/plumia.xml` |

Después, abre **Plumia → Ajustes** y pega tu clave **gratis** (NVIDIA u OpenRouter, ver abajo). No hace falta Node.js ni certificados.

---

## 📦 Instalación local (Windows, sin depender de internet para el panel)

**Requisitos:** Windows 10/11, Office de escritorio y [Node.js](https://nodejs.org) 18 o superior.

1. Descarga este repositorio (**Code → Download ZIP**) y descomprímelo.
2. Abre **PowerShell** en esa carpeta y ejecuta:
   ```powershell
   powershell -ExecutionPolicy Bypass -File setup.ps1
   ```
   Cuando Windows pregunte si confías en el certificado de `localhost`, pulsa **Sí**.
3. Abre Word y ve a **Inicio → Complementos → Más complementos → CARPETA COMPARTIDA → Plumia → Agregar**.
4. Consigue una clave **gratis** y pégala en el panel (**Ajustes**) o en el archivo correspondiente de `%USERPROFILE%\DeepSeekWordAddin\`:
   - **NVIDIA** (incluye DeepSeek): [build.nvidia.com](https://build.nvidia.com) → *Get API Key* → `clave-nvidia.txt`
   - **OpenRouter**: [openrouter.ai](https://openrouter.ai) → *Keys* → `clave-openrouter.txt`
5. Pulsa **Plumia** en la pestaña **Inicio** y escribe lo que quieres.

**Outlook:** Inicio → Obtener complementos → Mis complementos → *Agregar un complemento personalizado → Desde archivo* → elige `manifest-outlook.xml`.

El servidor local arranca solo al iniciar Windows.

---

## 🛡️ Seguridad

- El servidor solo escucha en tu propio equipo (`127.0.0.1` / `::1`).
- Rechaza peticiones de otras webs (protección contra CSRF y DNS rebinding) y bloquea el acceso a archivos fuera del panel.
- Las claves nunca se suben al repositorio (ver `.gitignore`).
- La IA trata el contenido de tus documentos como **datos, no como órdenes** (protección contra inyección de instrucciones).
- En modo automático solo usa modelos **gratuitos**. Los de pago solo se usan si los eliges tú.

---

## 🩺 Problemas frecuentes

| Problema | Solución |
|---|---|
| No veo el botón | Inicio → Complementos → Más complementos → CARPETA COMPARTIDA → Agregar |
| "Error del complemento" | El servidor no está en marcha: reinicia el PC o ejecuta `node server.js` en `%USERPROFILE%\DeepSeekWordAddin` |
| "No hay clave" / error 401 | Revisa la clave en Ajustes o en `clave-*.txt` (debe estar completa, sin asteriscos) |
| Va lento o "saturado" | Normal en modelos gratis: Plumia cambia de modelo sola. Si todos fallan, espera unos minutos |
| El registro de actividad | `%USERPROFILE%\DeepSeekWordAddin\server.log` |

---

## 🗺️ Hoja de ruta

- [x] Renombrar la interfaz a "Plumia"
- [x] Versión alojada en la nube (sin servidor local): https://plumia-office-ai.vercel.app
- [ ] Publicarla en Microsoft Marketplace (AppSource) — ficha y pasos listos en [docs/marketplace/FICHA.md](docs/marketplace/FICHA.md)

## ⚖️ Legal

[Política de privacidad](https://plumia-office-ai.vercel.app/privacidad.html) · [Términos de uso](https://plumia-office-ai.vercel.app/terminos.html) · [Soporte](https://plumia-office-ai.vercel.app/soporte.html)
- [ ] Cuentas de usuario y suscripción
- [ ] Plantillas: trabajo escolar, informe, carta formal, CV

---

## 📄 Licencia

© 2026 Plumia. **Todos los derechos reservados.** El código es visible para consulta y uso personal. Para cualquier uso comercial, redistribución o licencia, contacta con el autor. Consulta [LICENSE](LICENSE).

*Microsoft, Word, Excel, PowerPoint y Outlook son marcas de Microsoft. DeepSeek, NVIDIA y OpenRouter son marcas de sus respectivos dueños. Plumia no está afiliado a ninguno de ellos.*
