"use strict";
// Anexos ("¿Tienes en tu poder…? Anexar en imagen"): fotos o archivos que se
// integran al final del protocolo. Se guardan SOLO en este dispositivo
// (IndexedDB del navegador); nunca se envían a ningún lado.

const ADJ_DB = "protocolo_transporte_asmac";
const ADJ_STORE = "adjuntos";
const ADJ_LADO_MAX = 1800;     // px del lado mayor: legible y sin hacer pesado el Word
const ADJ_CALIDAD = 0.85;
const ADJ_PDF_PAGINAS_MAX = 6;

let _adjDb = null;
function adjDb() {
  if (_adjDb) return _adjDb;
  _adjDb = new Promise((ok, mal) => {
    const req = indexedDB.open(ADJ_DB, 1);
    req.onupgradeneeded = () => {
      const st = req.result.createObjectStore(ADJ_STORE, { keyPath: "id", autoIncrement: true });
      st.createIndex("clave", "clave");
    };
    req.onsuccess = () => ok(req.result);
    req.onerror = () => mal(req.error);
  });
  return _adjDb;
}

function adjTx(modo, fn) {
  return adjDb().then((db) => new Promise((ok, mal) => {
    const tx = db.transaction(ADJ_STORE, modo);
    const res = fn(tx.objectStore(ADJ_STORE));
    tx.oncomplete = () => ok(res && "result" in res ? res.result : res);
    tx.onerror = () => mal(tx.error);
    tx.onabort = () => mal(tx.error || new Error("No se pudo guardar el anexo (¿espacio lleno?)."));
  }));
}

/** Memoria: clave → [{id, clave, nombre, blob, w, h, url}] */
const ADJUNTOS = new Map();

async function cargarAdjuntos() {
  ADJUNTOS.clear();
  let todos = [];
  try { todos = await adjTx("readonly", (st) => st.getAll()); }
  catch { return; } // navegador sin IndexedDB (p. ej. modo privado muy restrictivo)
  for (const a of todos || []) {
    a.url = URL.createObjectURL(a.blob);
    if (!ADJUNTOS.has(a.clave)) ADJUNTOS.set(a.clave, []);
    ADJUNTOS.get(a.clave).push(a);
  }
}

const adjuntosDe = (clave) => ADJUNTOS.get(clave) || [];

async function guardarAdjunto(clave, nombre, { blob, w, h }) {
  const reg = { clave, nombre, blob, w, h };
  let id;
  try { id = await adjTx("readwrite", (st) => st.add(reg)); }
  catch { id = `mem-${Date.now()}-${Math.random()}`; } // sin IndexedDB: solo en memoria
  const a = { ...reg, id, url: URL.createObjectURL(blob) };
  if (!ADJUNTOS.has(clave)) ADJUNTOS.set(clave, []);
  ADJUNTOS.get(clave).push(a);
  return a;
}

async function borrarAdjunto(clave, id) {
  const lista = adjuntosDe(clave);
  const i = lista.findIndex((a) => a.id === id);
  if (i >= 0) { URL.revokeObjectURL(lista[i].url); lista.splice(i, 1); }
  if (typeof id === "number") { try { await adjTx("readwrite", (st) => st.delete(id)); } catch { /* nada */ } }
}

async function borrarAdjuntosConPrefijo(prefijo) {
  for (const clave of [...ADJUNTOS.keys()]) {
    if (!clave.startsWith(prefijo)) continue;
    for (const a of [...adjuntosDe(clave)]) await borrarAdjunto(clave, a.id);
    ADJUNTOS.delete(clave);
  }
}

async function borrarTodosLosAdjuntos() {
  for (const lista of ADJUNTOS.values()) lista.forEach((a) => URL.revokeObjectURL(a.url));
  ADJUNTOS.clear();
  try { await adjTx("readwrite", (st) => st.clear()); } catch { /* nada */ }
}

// ───────── Conversión a JPEG ─────────
function lienzoAJpeg(canvas) {
  return new Promise((ok, mal) => canvas.toBlob(
    (b) => (b ? ok({ blob: b, w: canvas.width, h: canvas.height }) : mal(new Error("No se pudo convertir la imagen."))),
    "image/jpeg", ADJ_CALIDAD));
}

function lienzoDesde(fuente, w, h) {
  const k = Math.min(1, ADJ_LADO_MAX / Math.max(w, h));
  const c = document.createElement("canvas");
  c.width = Math.max(1, Math.round(w * k));
  c.height = Math.max(1, Math.round(h * k));
  const ctx = c.getContext("2d");
  ctx.fillStyle = "#fff";
  ctx.fillRect(0, 0, c.width, c.height);
  ctx.drawImage(fuente, 0, 0, c.width, c.height);
  return c;
}

async function imagenAJpeg(file) {
  let fuente;
  try {
    fuente = await createImageBitmap(file, { imageOrientation: "from-image" });
  } catch {
    fuente = await new Promise((ok, mal) => {
      const img = new Image();
      const url = URL.createObjectURL(file);
      img.onload = () => { URL.revokeObjectURL(url); ok(img); };
      img.onerror = () => { URL.revokeObjectURL(url); mal(new Error(
        `No se pudo abrir "${file.name}". Usa una foto JPG o PNG (en iPhone: Ajustes → Cámara → Formatos → Más compatible).`)); };
      img.src = url;
    });
  }
  const w = fuente.width || fuente.naturalWidth, h = fuente.height || fuente.naturalHeight;
  return [await lienzoAJpeg(lienzoDesde(fuente, w, h))];
}

let _pdfjs = null;
async function pdfAJpegs(file) {
  if (!_pdfjs) {
    _pdfjs = await import("./vendor/pdfjs/pdf.min.mjs");
    _pdfjs.GlobalWorkerOptions.workerSrc = new URL("./vendor/pdfjs/pdf.worker.min.mjs", location.href).href;
  }
  const pdf = await _pdfjs.getDocument({ data: new Uint8Array(await file.arrayBuffer()), isEvalSupported: false }).promise;
  const paginas = Math.min(pdf.numPages, ADJ_PDF_PAGINAS_MAX);
  const salida = [];
  for (let n = 1; n <= paginas; n++) {
    const pag = await pdf.getPage(n);
    const base = pag.getViewport({ scale: 1 });
    const vp = pag.getViewport({ scale: Math.min(3, ADJ_LADO_MAX / Math.max(base.width, base.height)) });
    const c = document.createElement("canvas");
    c.width = Math.round(vp.width); c.height = Math.round(vp.height);
    const ctx = c.getContext("2d");
    ctx.fillStyle = "#fff"; ctx.fillRect(0, 0, c.width, c.height);
    await pag.render({ canvasContext: ctx, viewport: vp }).promise;
    salida.push(await lienzoAJpeg(c));
  }
  await pdf.destroy();
  return { imagenes: salida, omitidas: pdf.numPages - paginas };
}

/** Procesa un archivo elegido (foto, imagen o PDF) → lista de JPEG listos para el Word. */
async function procesarArchivoAnexo(file) {
  const esPdf = file.type === "application/pdf" || /\.pdf$/i.test(file.name);
  if (esPdf) return pdfAJpegs(file);
  return { imagenes: await imagenAJpeg(file), omitidas: 0 };
}
