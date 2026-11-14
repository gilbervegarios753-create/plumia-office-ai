/* global Office, Word, Excel, PowerPoint */
// Herramientas que la IA puede usar para trabajar directamente en el documento.
// Cada herramienta tiene: definición (para el modelo) y ejecución (en Office).

let HOST = null;
const isHost = (name) => HOST === Office.HostType[name];
const hostName = () => Object.keys(Office.HostType).find((k) => Office.HostType[k] === HOST) || "Office";

function officeAsync(fn) {
  return new Promise((resolve, reject) => {
    fn((r) => (r.status === Office.AsyncResultStatus.Succeeded ? resolve(r.value) : reject(new Error(r.error.message))));
  });
}
function isCompose() {
  const item = Office.context.mailbox && Office.context.mailbox.item;
  return !!(item && typeof item.setSelectedDataAsync === "function");
}
const clip = (s, n = 40000) => (s.length > n ? s.slice(0, n) + `\n…[recortado, ${s.length - n} caracteres más]` : s);

// Nombres de estilo aceptados (español o inglés) → estilo integrado de Word.
const STYLE_MAP = {
  normal: "Normal", titulo: "Title", "título": "Title", title: "Title", subtitulo: "Subtitle", "subtítulo": "Subtitle", subtitle: "Subtitle",
  titulo1: "Heading1", titulo2: "Heading2", titulo3: "Heading3", titulo4: "Heading4",
  heading1: "Heading1", heading2: "Heading2", heading3: "Heading3", heading4: "Heading4",
  lista: "ListBullet", vinetas: "ListBullet", "viñetas": "ListBullet", listbullet: "ListBullet",
  numerada: "ListNumber", listanumerada: "ListNumber", listnumber: "ListNumber", cita: "Quote", quote: "Quote",
};
function builtInStyle(name) {
  if (!name) return null;
  const k = String(name).toLowerCase().replace(/[\s_-]/g, "");
  return STYLE_MAP[k] || null;
}
function applyStyle(par, estilo) {
  const s = builtInStyle(estilo);
  if (s) par.styleBuiltIn = s;
  else if (estilo) par.style = estilo; // estilo personalizado por nombre
}
function applyFont(target, a) {
  if (a.negrita !== undefined) target.font.bold = !!a.negrita;
  if (a.cursiva !== undefined) target.font.italic = !!a.cursiva;
  if (a.subrayado !== undefined) target.font.underline = a.subrayado ? "Single" : "None";
  if (a.tamano) target.font.size = Number(a.tamano);
  if (a.color) target.font.color = a.color;
  if (a.fuente) target.font.name = a.fuente;
  if (a.resaltado !== undefined) target.font.highlightColor = a.resaltado || null;
}
const ALIGN = { izquierda: "Left", centro: "Centered", centrado: "Centered", derecha: "Right", justificado: "Justified" };

// ---------------------------- WORD ----------------------------
const WORD_TOOLS = {
  leer_documento: {
    description: "Lee el documento completo de Word. Devuelve cada párrafo con su número [índice], su estilo y su texto. Úsalo antes de editar para saber qué hay.",
    parameters: { type: "object", properties: {} },
    run: () =>
      Word.run(async (ctx) => {
        const ps = ctx.document.body.paragraphs;
        ps.load("items/text,items/style,items/tableNestingLevel");
        const sel = ctx.document.getSelection();
        sel.load("text");
        await ctx.sync();
        const lines = ps.items.map((p, i) => `[${i}] (${p.style}${p.tableNestingLevel ? ", en tabla" : ""}) ${p.text}`);
        return clip(`Párrafos: ${ps.items.length}\nSelección actual: ${sel.text ? JSON.stringify(sel.text.slice(0, 500)) : "(nada seleccionado)"}\n\n${lines.join("\n")}`);
      }),
  },
  leer_seleccion: {
    description: "Devuelve el texto que el usuario tiene seleccionado en Word.",
    parameters: { type: "object", properties: {} },
    run: () =>
      Word.run(async (ctx) => {
        const sel = ctx.document.getSelection();
        sel.load("text");
        await ctx.sync();
        return sel.text || "(no hay nada seleccionado)";
      }),
  },
  insertar_parrafos: {
    description: "Inserta uno o varios párrafos nuevos, cada uno con su estilo (normal, titulo, titulo1, titulo2, titulo3, lista, numerada, cita). Úsalo para escribir contenido nuevo.",
    parameters: {
      type: "object",
      properties: {
        parrafos: {
          type: "array",
          items: { type: "object", properties: { texto: { type: "string" }, estilo: { type: "string" } }, required: ["texto"] },
        },
        posicion: { type: "string", enum: ["final", "inicio", "despues_seleccion", "antes_seleccion", "despues_parrafo", "antes_parrafo"], description: "Dónde insertar. Por defecto: final." },
        indice: { type: "integer", description: "Número de párrafo de referencia (solo para despues_parrafo/antes_parrafo)." },
      },
      required: ["parrafos"],
    },
    run: (a) =>
      Word.run(async (ctx) => {
        const pos = a.posicion || "final";
        const list = a.parrafos || [];
        let anchor = null;
        if (pos === "despues_parrafo" || pos === "antes_parrafo") {
          const ps = ctx.document.body.paragraphs;
          ps.load("items");
          await ctx.sync();
          anchor = ps.items[a.indice];
          if (!anchor) throw new Error(`No existe el párrafo ${a.indice} (hay ${ps.items.length}).`);
        } else if (pos === "despues_seleccion" || pos === "antes_seleccion") {
          const sps = ctx.document.getSelection().paragraphs;
          sps.load("items");
          await ctx.sync();
          anchor = pos === "despues_seleccion" ? sps.items[sps.items.length - 1] : sps.items[0];
        }
        const ordered = pos === "inicio" ? [...list].reverse() : list;
        let last = anchor;
        for (const item of ordered) {
          let p;
          if (pos === "final") p = ctx.document.body.insertParagraph(item.texto, "End");
          else if (pos === "inicio") p = ctx.document.body.insertParagraph(item.texto, "Start");
          else if (pos.startsWith("antes")) p = anchor.insertParagraph(item.texto, "Before");
          else { p = last.insertParagraph(item.texto, "After"); last = p; }
          applyStyle(p, item.estilo || "normal");
        }
        await ctx.sync();
        return `Insertados ${list.length} párrafo(s) en posición "${pos}".`;
      }),
  },
  reemplazar_seleccion: {
    description: "Reemplaza el texto seleccionado por el usuario con un texto nuevo (por ejemplo, la versión corregida o mejorada).",
    parameters: { type: "object", properties: { texto: { type: "string" } }, required: ["texto"] },
    run: (a) =>
      Word.run(async (ctx) => {
        const sel = ctx.document.getSelection();
        sel.load("text");
        await ctx.sync();
        if (!sel.text) throw new Error("No hay nada seleccionado. Usa reemplazar_parrafo o insertar_parrafos.");
        sel.insertText(a.texto, "Replace");
        await ctx.sync();
        return "Selección reemplazada.";
      }),
  },
  reemplazar_parrafo: {
    description: "Reemplaza el texto de un párrafo concreto (por su índice de leer_documento). Opcionalmente cambia su estilo.",
    parameters: {
      type: "object",
      properties: { indice: { type: "integer" }, texto: { type: "string" }, estilo: { type: "string" } },
      required: ["indice", "texto"],
    },
    run: (a) =>
      Word.run(async (ctx) => {
        const ps = ctx.document.body.paragraphs;
        ps.load("items");
        await ctx.sync();
        const p = ps.items[a.indice];
        if (!p) throw new Error(`No existe el párrafo ${a.indice} (hay ${ps.items.length}).`);
        p.insertText(a.texto, "Replace");
        if (a.estilo) applyStyle(p, a.estilo);
        await ctx.sync();
        return `Párrafo ${a.indice} reemplazado.`;
      }),
  },
  eliminar_parrafos: {
    description: "Elimina uno o varios párrafos por su índice.",
    parameters: { type: "object", properties: { indices: { type: "array", items: { type: "integer" } } }, required: ["indices"] },
    run: (a) =>
      Word.run(async (ctx) => {
        const ps = ctx.document.body.paragraphs;
        ps.load("items");
        await ctx.sync();
        const idx = [...new Set(a.indices)].sort((x, y) => y - x);
        for (const i of idx) if (ps.items[i]) ps.items[i].delete();
        await ctx.sync();
        return `Eliminados ${idx.length} párrafo(s).`;
      }),
  },
  buscar_y_reemplazar: {
    description: "Busca un texto en todo el documento y lo reemplaza en todas sus apariciones.",
    parameters: {
      type: "object",
      properties: { buscar: { type: "string" }, reemplazar: { type: "string" }, coincidir_mayusculas: { type: "boolean" }, palabra_completa: { type: "boolean" } },
      required: ["buscar", "reemplazar"],
    },
    run: (a) =>
      Word.run(async (ctx) => {
        const res = ctx.document.body.search(String(a.buscar).slice(0, 255), { matchCase: !!a.coincidir_mayusculas, matchWholeWord: !!a.palabra_completa });
        res.load("items");
        await ctx.sync();
        res.items.forEach((r) => r.insertText(a.reemplazar, "Replace"));
        await ctx.sync();
        return `${res.items.length} reemplazo(s) hechos.`;
      }),
  },
  dar_formato: {
    description: "Aplica formato a la selección, a uno o varios párrafos (por índice) o a todas las apariciones de un texto: negrita, cursiva, subrayado, tamaño, color (#RRGGBB o nombre), fuente, resaltado, alineación y estilo.",
    parameters: {
      type: "object",
      properties: {
        objetivo: { type: "string", enum: ["seleccion", "parrafos", "texto", "documento"] },
        indices: { type: "array", items: { type: "integer" }, description: "Para objetivo=parrafos." },
        texto: { type: "string", description: "Para objetivo=texto: el texto a buscar y formatear." },
        negrita: { type: "boolean" }, cursiva: { type: "boolean" }, subrayado: { type: "boolean" },
        tamano: { type: "number" }, color: { type: "string" }, fuente: { type: "string" }, resaltado: { type: "string" },
        alineacion: { type: "string", enum: ["izquierda", "centro", "derecha", "justificado"] },
        estilo: { type: "string" },
      },
      required: ["objetivo"],
    },
    run: (a) =>
      Word.run(async (ctx) => {
        let targets = [];
        let paragraphs = [];
        if (a.objetivo === "seleccion") {
          const sel = ctx.document.getSelection();
          targets = [sel];
          const sps = sel.paragraphs;
          sps.load("items");
          await ctx.sync();
          paragraphs = sps.items;
        } else if (a.objetivo === "parrafos") {
          const ps = ctx.document.body.paragraphs;
          ps.load("items");
          await ctx.sync();
          paragraphs = (a.indices || []).map((i) => ps.items[i]).filter(Boolean);
          targets = paragraphs;
        } else if (a.objetivo === "texto") {
          const res = ctx.document.body.search(String(a.texto || "").slice(0, 255), { matchCase: false });
          res.load("items");
          await ctx.sync();
          targets = res.items;
        } else {
          const ps = ctx.document.body.paragraphs;
          ps.load("items");
          await ctx.sync();
          paragraphs = ps.items;
          targets = [ctx.document.body];
        }
        targets.forEach((t) => applyFont(t, a));
        paragraphs.forEach((p) => {
          if (a.alineacion) p.alignment = ALIGN[a.alineacion] || "Left";
          if (a.estilo) applyStyle(p, a.estilo);
        });
        await ctx.sync();
        return `Formato aplicado a ${targets.length} elemento(s).`;
      }),
  },
  insertar_tabla: {
    description: "Inserta una tabla. 'filas' es una lista de filas; la primera fila se usa como encabezado.",
    parameters: {
      type: "object",
      properties: {
        filas: { type: "array", items: { type: "array", items: { type: "string" } } },
        posicion: { type: "string", enum: ["final", "despues_seleccion", "despues_parrafo"] },
        indice: { type: "integer" },
      },
      required: ["filas"],
    },
    run: (a) =>
      Word.run(async (ctx) => {
        const rows = a.filas || [];
        const cols = Math.max(...rows.map((r) => r.length));
        const values = rows.map((r) => Array.from({ length: cols }, (_, i) => String(r[i] ?? "")));
        const pos = a.posicion || "final";
        let table;
        if (pos === "final") {
          table = ctx.document.body.insertTable(values.length, cols, "End", values);
        } else {
          let anchor;
          if (pos === "despues_parrafo") {
            const ps = ctx.document.body.paragraphs;
            ps.load("items");
            await ctx.sync();
            anchor = ps.items[a.indice];
            if (!anchor) throw new Error(`No existe el párrafo ${a.indice}.`);
          } else {
            const sps = ctx.document.getSelection().paragraphs;
            sps.load("items");
            await ctx.sync();
            anchor = sps.items[sps.items.length - 1];
          }
          table = anchor.insertTable(values.length, cols, "After", values);
        }
        table.headerRowCount = 1;
        table.styleBuiltIn = "GridTable4_Accent1";
        await ctx.sync();
        return `Tabla de ${values.length}x${cols} insertada.`;
      }),
  },
};

// ---------------------------- EXCEL ----------------------------
const EXCEL_TOOLS = {
  leer_hoja: {
    description: "Lee el rango usado de la hoja activa (dirección y valores).",
    parameters: { type: "object", properties: {} },
    run: () =>
      Excel.run(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getActiveWorksheet();
        const r = sheet.getUsedRangeOrNullObject(true);
        r.load("address,text");
        sheet.load("name");
        await ctx.sync();
        if (r.isNullObject) return `Hoja "${sheet.name}" vacía.`;
        return clip(`Hoja "${sheet.name}", rango ${r.address}:\n` + r.text.map((row) => row.join("\t")).join("\n"));
      }),
  },
  leer_seleccion: {
    description: "Lee las celdas seleccionadas (dirección y valores).",
    parameters: { type: "object", properties: {} },
    run: () =>
      Excel.run(async (ctx) => {
        const r = ctx.workbook.getSelectedRange();
        r.load("address,text");
        await ctx.sync();
        return clip(`${r.address}:\n` + r.text.map((row) => row.join("\t")).join("\n"));
      }),
  },
  escribir_celdas: {
    description: "Escribe valores o fórmulas (empiezan por =) en un rango a partir de una celda inicial, p. ej. inicio='A1' y valores=[[\"Nombre\",\"Total\"],[\"Ana\",\"=B2*2\"]].",
    parameters: {
      type: "object",
      properties: { inicio: { type: "string" }, valores: { type: "array", items: { type: "array", items: {} } } },
      required: ["inicio", "valores"],
    },
    run: (a) =>
      Excel.run(async (ctx) => {
        const vals = a.valores;
        const cols = Math.max(...vals.map((r) => r.length));
        const norm = vals.map((r) => Array.from({ length: cols }, (_, i) => (r[i] === undefined || r[i] === null ? "" : r[i])));
        const sheet = ctx.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getRange(a.inicio).getResizedRange(norm.length - 1, cols - 1);
        range.formulas = norm;
        range.load("address");
        await ctx.sync();
        return `Escrito en ${range.address}.`;
      }),
  },
  dar_formato_celdas: {
    description: "Formatea un rango (p. ej. 'A1:D1'): negrita, color de fondo, color de letra, formato numérico, autoajustar columnas.",
    parameters: {
      type: "object",
      properties: {
        rango: { type: "string" }, negrita: { type: "boolean" }, fondo: { type: "string" }, color: { type: "string" },
        formato_numero: { type: "string" }, autoajustar: { type: "boolean" },
      },
      required: ["rango"],
    },
    run: (a) =>
      Excel.run(async (ctx) => {
        const r = ctx.workbook.worksheets.getActiveWorksheet().getRange(a.rango);
        if (a.negrita !== undefined) r.format.font.bold = !!a.negrita;
        if (a.fondo) r.format.fill.color = a.fondo;
        if (a.color) r.format.font.color = a.color;
        if (a.formato_numero) {
          r.load("rowCount,columnCount");
          await ctx.sync();
          r.numberFormat = Array.from({ length: r.rowCount }, () => Array(r.columnCount).fill(a.formato_numero));
        }
        if (a.autoajustar) r.format.autofitColumns();
        await ctx.sync();
        return `Formato aplicado a ${a.rango}.`;
      }),
  },
};

// ---------------------------- POWERPOINT ----------------------------
const PPT_TOOLS = {
  leer_seleccion: {
    description: "Lee el texto seleccionado en PowerPoint.",
    parameters: { type: "object", properties: {} },
    run: () => officeAsync((cb) => Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, cb)).then((t) => t || "(nada seleccionado)"),
  },
  leer_presentacion: {
    description: "Lee el texto de todas las diapositivas.",
    parameters: { type: "object", properties: {} },
    run: () =>
      PowerPoint.run(async (ctx) => {
        const slides = ctx.presentation.slides;
        slides.load("items");
        await ctx.sync();
        const out = [];
        for (let i = 0; i < slides.items.length; i++) {
          const shapes = slides.items[i].shapes;
          shapes.load("items/name,items/textFrame/textRange/text");
          await ctx.sync();
          const texts = shapes.items.map((s) => { try { return s.textFrame.textRange.text; } catch (e) { return ""; } }).filter(Boolean);
          out.push(`Diapositiva ${i + 1}: ${texts.join(" | ")}`);
        }
        return clip(out.join("\n") || "Presentación vacía.");
      }),
  },
  insertar_texto: {
    description: "Escribe texto en el cuadro de texto o la selección actual de la diapositiva.",
    parameters: { type: "object", properties: { texto: { type: "string" } }, required: ["texto"] },
    run: (a) => officeAsync((cb) => Office.context.document.setSelectedDataAsync(a.texto, { coercionType: Office.CoercionType.Text }, cb)).then(() => "Texto insertado."),
  },
  agregar_diapositiva: {
    description: "Añade una diapositiva nueva al final con un título y un cuerpo de texto.",
    parameters: { type: "object", properties: { titulo: { type: "string" }, texto: { type: "string" } }, required: ["titulo"] },
    run: (a) =>
      PowerPoint.run(async (ctx) => {
        ctx.presentation.slides.add();
        await ctx.sync();
        const slides = ctx.presentation.slides;
        slides.load("items");
        await ctx.sync();
        const slide = slides.items[slides.items.length - 1];
        const t = slide.shapes.addTextBox(a.titulo, { left: 40, top: 30, width: 640, height: 60 });
        t.textFrame.textRange.font.size = 32;
        t.textFrame.textRange.font.bold = true;
        if (a.texto) {
          const b = slide.shapes.addTextBox(a.texto, { left: 40, top: 110, width: 640, height: 360 });
          b.textFrame.textRange.font.size = 18;
        }
        await ctx.sync();
        return `Diapositiva ${slides.items.length} añadida.`;
      }),
  },
};

// ---------------------------- OUTLOOK ----------------------------
const OUTLOOK_TOOLS = {
  leer_correo: {
    description: "Lee el asunto y el cuerpo del correo abierto.",
    parameters: { type: "object", properties: {} },
    run: async () => {
      const item = Office.context.mailbox.item;
      const body = await officeAsync((cb) => item.body.getAsync(Office.CoercionType.Text, cb));
      let subject = item.subject;
      if (subject && typeof subject.getAsync === "function") subject = await officeAsync((cb) => subject.getAsync(cb));
      return clip(`Asunto: ${subject || ""}\n\n${body}`);
    },
  },
  escribir_en_correo: {
    description: "Escribe texto en el correo que se está redactando, en la posición del cursor (o reemplazando la selección).",
    parameters: { type: "object", properties: { texto: { type: "string" } }, required: ["texto"] },
    run: async (a) => {
      if (!isCompose()) throw new Error("Solo se puede escribir al redactar o responder un correo.");
      await officeAsync((cb) => Office.context.mailbox.item.body.setSelectedDataAsync(a.texto, { coercionType: Office.CoercionType.Text }, cb));
      return "Texto escrito en el correo.";
    },
  },
};

function currentTools() {
  let own = {};
  if (isHost("Word")) own = WORD_TOOLS;
  else if (isHost("Excel")) own = EXCEL_TOOLS;
  else if (isHost("PowerPoint")) own = PPT_TOOLS;
  else if (isHost("Outlook")) own = OUTLOOK_TOOLS;
  // Herramientas comunes a todos los programas (p. ej. el equipo de agentes, definido en taskpane.js).
  const shared = typeof SHARED_TOOLS !== "undefined" ? SHARED_TOOLS : {};
  return { ...own, ...shared };
}

// Definiciones en el formato de "tools" compatible con OpenAI.
function toolDefinitions() {
  return Object.entries(currentTools()).map(([name, t]) => ({
    type: "function",
    function: { name, description: t.description, parameters: t.parameters },
  }));
}

async function runTool(name, args) {
  const tool = currentTools()[name];
  if (!tool) throw new Error(`Herramienta desconocida: ${name}`);
  return String(await tool.run(args || {}));
}
