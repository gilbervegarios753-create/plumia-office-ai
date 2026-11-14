# Ficha de Plumia para Microsoft Marketplace (AppSource)

Textos y archivos listos para copiar en **Microsoft Partner Center** al crear la oferta *Office add-in*.

## Archivos

| Qué pide Partner Center | Archivo |
|---|---|
| Manifiesto Word, Excel y PowerPoint | [`manifest-cloud.xml`](../../manifest-cloud.xml) |
| Manifiesto Outlook (oferta aparte) | [`manifest-outlook-cloud.xml`](../../manifest-outlook-cloud.xml) |
| Logo de la tienda (PNG, 300×300) | [`public/assets/icon-300.png`](../../public/assets/icon-300.png) |
| Captura (PNG, 1366×768) | [`captura-1.png`](captura-1.png) |

Los dos manifiestos pasan el validador oficial (`npx office-addin-manifest validate`).

## Direcciones

| Campo | Valor |
|---|---|
| Política de privacidad | https://plumia-office-ai.vercel.app/privacidad.html |
| Términos de uso (EULA) | https://plumia-office-ai.vercel.app/terminos.html |
| Soporte | https://plumia-office-ai.vercel.app/soporte.html |
| Correo de soporte | gilbervegarios753@gmail.com |

## Textos de la ficha (español)

**Nombre:** Plumia

**Resumen (máx. 100 caracteres):**
> Tu equipo de IA en Word, Excel, PowerPoint y Outlook: escribe, corrige y da formato por ti.

**Descripción:**
> Plumia es un asistente de inteligencia artificial que trabaja directamente en tus documentos. Escribe lo que necesitas en el chat y Plumia lo hace en el archivo: sin copiar ni pegar.
>
> **Qué puede hacer**
> • Word: redactar textos con títulos, listas y citas; corregir ortografía y gramática; mejorar, resumir, ampliar o traducir; buscar y reemplazar; dar formato y crear tablas.
> • Excel: crear tablas con datos y fórmulas, y dar formato a las celdas.
> • PowerPoint: crear diapositivas con título y contenido, y mejorar sus textos.
> • Outlook: resumir correos y redactar respuestas.
>
> **Por qué Plumia**
> • Equipo de agentes: las tareas grandes, como un informe con varias secciones, se reparten entre varios agentes que escriben a la vez.
> • Modo automático: si un modelo va lento o está saturado, Plumia cambia sola a otro y sigue trabajando.
> • Más de 400 modelos de IA, incluidos DeepSeek, Llama, Mistral, Gemma y Nemotron.
> • Gratis: usa tu propia clave gratuita de NVIDIA u OpenRouter. Plumia no guarda tus documentos.
> • Puedes deshacer cualquier cambio con Ctrl+Z.
>
> **Para empezar** necesitas una clave gratuita de NVIDIA (build.nvidia.com) u OpenRouter (openrouter.ai). En el panel, abre Ajustes y pégala.

**Palabras clave:** IA, inteligencia artificial, redacción, corrector, resumen, traductor, DeepSeek, asistente

**Categorías sugeridas:** Productividad; Contenido y redacción (o la más cercana que ofrezca Partner Center)

**Idioma de la ficha:** Español. Se puede añadir una versión en inglés más adelante.

## Notas para los revisores de Microsoft

Partner Center tiene un campo de **notas para la certificación**. Plumia necesita una clave de un servicio de IA para responder, así que hay que darles una clave de prueba:

> Plumia requires the user's own AI provider key. For testing, open Plumia → Ajustes (Settings), choose "DeepSeek vía NVIDIA (GRATIS)" and paste this test key: **[PEGA AQUÍ UNA CLAVE DE PRUEBA DE NVIDIA]**. Then type in the chat, for example: "Escribe un párrafo sobre el agua con un título". Support: gilbervegarios753@gmail.com

⚠️ **Pega la clave de prueba solo en Partner Center, nunca en este repositorio** (es público). Crea una clave nueva solo para la revisión y bórrala en build.nvidia.com cuando aprueben Plumia.

## Pasos en Partner Center

1. Crea la cuenta de desarrollador en https://partner.microsoft.com (registro en el **Microsoft 365 and Copilot program**, cuenta *Individual*). La cuenta individual es gratuita.
2. **Marketplace offers → New offer → Office add-in**. Nombre: **Plumia**.
3. **Packages:** sube `manifest-cloud.xml`.
4. **Properties:** categorías, URLs de privacidad, términos y soporte de la tabla de arriba.
5. **Marketplace listings → Español:** resumen, descripción, palabras clave, logo `icon-300.png` y captura `captura-1.png`.
6. **Availability:** gratuita, todos los mercados (o los que prefieras).
7. **Notes for certification:** el texto de arriba con la clave de prueba.
8. **Review and publish.** La revisión de Microsoft suele tardar unos días; si algo falla, te envían un informe con los cambios necesarios.
9. Repite para Outlook con `manifest-outlook-cloud.xml` como una oferta aparte (opcional).

## Antes de cobrar

- El alojamiento usa el plan **Hobby** de Vercel, que solo permite uso **no comercial**. Mientras Plumia sea gratuita y con la clave de cada usuario, puede quedarse así. Para cobrar, pasa a **Vercel Pro** o a otro alojamiento que permita uso comercial.
- Al vender en Perú necesitarás **RUC**, **Libro de Reclamaciones virtual** y una política de **reembolsos**. Revisa los textos legales con un abogado de derecho digital antes del lanzamiento comercial.
