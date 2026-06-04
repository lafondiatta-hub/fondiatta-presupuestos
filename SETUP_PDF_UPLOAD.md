# Setup · Subir PDF → Seguimiento de Presupuestos

> **Estado 2026-05-12:** ✅ Setup completado.
> Web App standalone deployado en script `1zNrmzpbAzY5F1AMwbiP2EmXWo5EQv5_DUFjULxf5KsQCNlexN3FoC35Y` (NO el bound v4.0 que ya tenía doPost ocupado).
> Código tracked vía clasp en `./apps-script-pdf/`.
> API key Gemini: `fondiatta-presupuestos-pdf` en aistudio (proyecto `lafondiatta-planner-494114`).
> Si necesitás re-deployar, ver sección "Re-deploy" al final.

Pasos para activar la feature "📥 Subir PDF" del cotizador.
Sólo se hace una vez. Tiempo estimado: 10–15 minutos.

---

## 1 · Crear API Key de Gemini (gratis, sin tarjeta)

1. Andá a **https://aistudio.google.com/app/apikey** (loggeate con tu Google).
2. Click en **"Create API key"**.
3. Si te pide proyecto, elegí cualquiera (o "Create API key in new project").
4. Copiá la key (empieza con `AIza...`). Guardala momentáneamente.

> **Cuota gratis:** ~1500 PDFs por día con `gemini-2.5-flash`. Más que suficiente.

---

## 2 · Crear Apps Script bound a la Sheet de Seguimiento

1. Abrí la Sheet:
   https://docs.google.com/spreadsheets/d/1MrczSE46x3ecpUCuHJsTYI5Gm3VEvp33GtTPyBzsg98/edit

2. Menú: **Extensiones → Apps Script**.

3. Se abre un script nuevo. Borrá el `function myFunction()` que viene por default.

4. Pegá el contenido completo del archivo `apps_script_pdf_to_sheet.gs`
   (en `~/Documents/LaFondiatta/Proyectos/fondiatta-presupuestos/`).

5. Renombrá el proyecto (arriba a la izq): **"PDF Upload — Seguimiento Presupuestos"**.

6. **Guardá** (Ctrl+S o el ícono del disquete).

---

## 3 · Cargar la Gemini API Key como Script Property

1. En el editor de Apps Script, ícono de **engranaje ⚙ (Project Settings)** a la izq.
2. Bajá a **"Script Properties"** → **"Add script property"**.
3. **Property:** `GEMINI_API_KEY`
4. **Value:** (pegá la key del paso 1)
5. **Save script properties**.

---

## 4 · Probar que la key funciona

1. Volvé al editor (icono `<>` a la izq).
2. Arriba donde dice "Select function", elegí **`test_geminiKey`** y dale **Run**.
3. La primera vez te pide permisos: **Review permissions → tu cuenta → Advanced → Go to (unsafe) → Allow**.
4. Abrí los logs (View → Logs o Ctrl+Enter). Tiene que decir algo como:
   ```
   OK key length=39
   ```
   Si dice `FALTA GEMINI_API_KEY` revisá el paso 3.

---

## 5 · Deploy del Web App

1. Botón **Deploy → New deployment** (arriba a la derecha).
2. Click en el ⚙ junto a "Select type" → **Web app**.
3. Configurá:
   - **Description:** `PDF Upload v1`
   - **Execute as:** `Me (tu@gmail)`
   - **Who has access:** `Anyone` ← importante para que el HTML llegue sin auth
4. **Deploy**.
5. Te muestra una URL del tipo:
   ```
   https://script.google.com/macros/s/AKfycb...XYZ/exec
   ```
   **Copiala**.

---

## 6 · Pegar la URL en el cotizador

1. Abrí el archivo `index.html` del cotizador.
2. Buscá la línea (cerca de la línea 1015):
   ```js
   const APPS_SCRIPT_PDF_URL = '';
   ```
3. Reemplazá por:
   ```js
   const APPS_SCRIPT_PDF_URL = 'https://script.google.com/macros/s/AKfycb...XYZ/exec';
   ```
4. Guardá. Si tenés deploy automático en Vercel, esperá el redeploy.

---

## 7 · Probar

1. Abrí el cotizador en el browser.
2. Header (arriba a la izq) → botón **📥 Subir PDF**.
3. Arrastrá un PDF de propuesta.
4. Esperá 10–30 segundos (Gemini procesa).
5. Revisá los datos extraídos. **Editá lo que esté mal**.
6. Click **Registrar en Sheet**.
7. Verificá en la Sheet que apareció una fila nueva con ID auto-asignado `P-2026-XXX`.

---

## Limitaciones conocidas

- **No detecta duplicados.** Si subís dos veces el mismo PDF, se inserta dos veces. Si lo necesitás, después agregamos un check por contacto+fecha.
- **No actualiza filas existentes.** Siempre crea fila nueva. Para editar usá la Sheet directamente.
- **PDFs muy largos (>50 págs):** Gemini puede tardar más o cortar. Los presupuestos LF normales (5–20 págs) andan bien.
- **Datos no presentes en el PDF:** quedan vacíos. Los completás vos antes de confirmar.

---

## Si algo falla

| Problema | Causa probable | Fix |
|---|---|---|
| "Falta GEMINI_API_KEY" | No guardaste la property | Repetí paso 3 |
| "Gemini 401 / 403" | Key inválida o sin permisos | Regenerá la key en aistudio |
| "Gemini 429" | Pasaste la cuota del día | Esperá hasta mañana o agregá billing |
| "HTTP 401 / 403" | El deploy no es "Anyone" | Repetí paso 5 con permisos correctos |
| "Respuesta no-JSON" | Apps Script cacheó versión vieja | Deploy → New deployment (no "manage") |
| Datos extraídos mal | PDF muy raro o escaneado | Editalos a mano antes de confirmar |

---

## Si necesitás re-deploy (cambios al `.gs`)

Cada vez que modifiques el `.gs`:
1. Pegá el código nuevo en Apps Script editor → guardar.
2. **Deploy → New deployment** (no "Manage deployments" — eso reusa la URL vieja con código viejo).
3. Copiá la URL nueva (cambia cada deploy).
4. Pegala en `APPS_SCRIPT_PDF_URL` del `index.html`.

> Alternativa más cómoda: Deploy → Manage deployments → ícono lápiz → Version: New version → Deploy.
> Esto mantiene la URL pero actualiza el código.
