"use strict";
// Llena la plantilla OFICIAL de ASMAC/CNAR ("Protocolo de transporte eventos
// 2026 CNAR.docx") en el navegador. Es el mismo llenado que hace scoutdb en el
// servidor (python-docx), portado a JavaScript: se abre el .docx (un zip), se
// edita word/document.xml y se vuelve a empaquetar. La plantilla nunca se
// modifica; el resultado es un archivo nuevo.

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const XML_NS = "http://www.w3.org/XML/1998/namespace";

const CATEGORIAS = ["ML", "TS", "CC", "CR", "S", "D", "OTRO"];
const CAT_JOVENES = ["ML", "TS", "CC", "CR"];
const CAT_ADULTOS = ["S", "D", "OTRO"];
const COL_ROSTER = { ML: 3, TS: 4, CC: 5, CR: 6, S: 7, D: 8, OTRO: 9 };

const hijos = (el, nombre) => Array.from(el.childNodes).filter((n) => n.nodeType === 1 && n.namespaceURI === W && n.localName === nombre);
const hijo = (el, nombre) => hijos(el, nombre)[0] || null;

function conteos(pasajeros) {
  const c = Object.fromEntries(CATEGORIAS.map((k) => [k, 0]));
  for (const p of pasajeros) c[CATEGORIAS.includes(p.cat) ? p.cat : "OTRO"]++;
  c.jovenes = CAT_JOVENES.reduce((s, k) => s + c[k], 0);
  c.adultos = CAT_ADULTOS.reduce((s, k) => s + c[k], 0);
  c.total = c.jovenes + c.adultos;
  return c;
}

class DocxLlenador {
  constructor(xmlTexto) {
    this.doc = new DOMParser().parseFromString(xmlTexto, "application/xml");
    if (this.doc.getElementsByTagName("parsererror").length) throw new Error("La plantilla está dañada (XML inválido).");
    this.body = this.doc.getElementsByTagNameNS(W, "body")[0];
  }

  tablas() { return hijos(this.body, "tbl"); }
  filas(tbl) { return hijos(tbl, "tr"); }

  /** Igual que python-docx `row.cells[i]`: índice por columna de la cuadrícula (gridSpan). */
  celda(tr, idx) {
    let col = 0;
    const trPr = hijo(tr, "trPr");
    const gb = trPr && hijo(trPr, "gridBefore");
    if (gb) col += Number(gb.getAttributeNS(W, "val") || 0);
    for (const tc of hijos(tr, "tc")) {
      const tcPr = hijo(tc, "tcPr");
      const gs = tcPr && hijo(tcPr, "gridSpan");
      const span = gs ? Number(gs.getAttributeNS(W, "val") || 1) : 1;
      if (idx >= col && idx < col + span) return tc;
      col += span;
    }
    throw new Error(`La plantilla cambió de estructura: no existe la columna ${idx} en una fila.`);
  }

  el(nombre) { return this.doc.createElementNS(W, "w:" + nombre); }

  /** Copia sólo el formato que copia scoutdb: negrita, cursiva, subrayado, tamaño y fuente. */
  formatoDe(run) {
    const rPr = run && hijo(run, "rPr");
    if (!rPr) return null;
    const nuevo = this.el("rPr");
    for (const n of ["rFonts", "b", "bCs", "i", "iCs", "u", "sz", "szCs"]) {
      const x = hijo(rPr, n);
      if (x) nuevo.appendChild(x.cloneNode(true));
    }
    return nuevo.childNodes.length ? nuevo : null;
  }

  nuevoRun(texto, plantillaRun) {
    const r = this.el("r");
    const rPr = this.formatoDe(plantillaRun);
    if (rPr) r.appendChild(rPr);
    String(texto).split("\n").forEach((linea, i) => {
      if (i) r.appendChild(this.el("br"));
      linea.split("\t").forEach((trozo, j) => {
        if (j) r.appendChild(this.el("tab"));
        if (trozo) {
          const t = this.el("t");
          t.setAttributeNS(XML_NS, "xml:space", "preserve");
          t.textContent = trozo;
          r.appendChild(t);
        }
      });
    });
    return r;
  }

  textoParrafo(p) {
    let s = "";
    for (const r of hijos(p, "r")) {
      for (const n of r.childNodes) {
        if (n.namespaceURI !== W) continue;
        if (n.localName === "t") s += n.textContent;
        else if (n.localName === "tab") s += "\t";
        else if (n.localName === "br" || n.localName === "cr") s += "\n";
      }
    }
    return s;
  }

  parrafoLimpio(p, texto, plantillaRun) {
    const runs = hijos(p, "r");
    if (plantillaRun === undefined) plantillaRun = runs[0] || null;
    for (const r of runs) p.removeChild(r);
    p.appendChild(this.nuevoRun(texto, plantillaRun));
  }

  /** Deja la celda con un único párrafo con `texto` (como _set_cell_clean de scoutdb). */
  celdaLimpia(tc, texto) {
    const ps = hijos(tc, "p");
    const plantillaRun = ps.map((p) => hijo(p, "r")).find(Boolean) || null;
    ps.slice(1).forEach((p) => tc.removeChild(p));
    let p0 = ps[0];
    if (!p0) { p0 = this.el("p"); tc.appendChild(p0); }
    this.parrafoLimpio(p0, texto, plantillaRun);
  }

  set(tr, idx, texto) { this.celdaLimpia(this.celda(tr, idx), texto ?? ""); }

  reemplazar(marcador, valor) {
    for (const p of Array.from(this.body.getElementsByTagNameNS(W, "p"))) {
      const t = this.textoParrafo(p);
      if (t.includes(marcador)) this.parrafoLimpio(p, t.split(marcador).join(valor));
    }
  }

  xml() { return new XMLSerializer().serializeToString(this.doc); }
}

const v = (x) => (x ?? "").toString().trim();

// Orden de los hijos que exige el esquema de Word (si se rompe, Word dice "contenido ilegible").
const ORDEN_RPR = ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline",
  "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing", "w", "kern",
  "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang"];
const ORDEN_PPR = ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr",
  "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct",
  "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind",
  "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc", "textDirection", "textAlignment",
  "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange"];
const ORDEN_TRPR = ["cnfStyle", "divId", "gridBefore", "gridAfter", "wBefore", "wAfter", "cantSplit", "trHeight",
  "tblHeader", "tblCellSpacing", "jc", "hidden", "ins", "del", "trPrChange"];
const ORDEN_TCPR = ["cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd", "noWrap", "tcMar",
  "textDirection", "tcFitText", "vAlign", "hideMark"];

/** Devuelve (creándolo si falta) el hijo `nombre` de `padre`, respetando el orden del esquema. */
function hijoOrdenado(d, padre, nombre, orden) {
  const existente = hijo(padre, nombre);
  if (existente) return existente;
  const nuevo = d.el(nombre);
  const pos = orden.indexOf(nombre);
  const siguiente = Array.from(padre.childNodes).find((n) =>
    n.nodeType === 1 && n.namespaceURI === W && orden.indexOf(n.localName) > pos);
  padre.insertBefore(nuevo, siguiente || null);
  return nuevo;
}

/** Contenedor de propiedades (rPr/pPr/trPr/tcPr): siempre es el primer hijo. */
function props(d, el, nombre) {
  let p = hijo(el, nombre);
  if (!p) { p = d.el(nombre); el.insertBefore(p, el.firstChild); }
  return p;
}

/** Marca una respuesta como lo hacen a mano en el formato: texto en negritas y resaltado amarillo. */
function resaltar(d, tc) {
  for (const r of Array.from(tc.getElementsByTagNameNS(W, "r"))) {
    const rPr = props(d, r, "rPr");
    hijoOrdenado(d, rPr, "b", ORDEN_RPR);
    hijoOrdenado(d, rPr, "highlight", ORDEN_RPR).setAttributeNS(W, "w:val", "yellow");
  }
}

const siNo = (tr, d, valor) => {
  if (valor === "si") resaltar(d, d.celda(tr, 9));
  else if (valor === "no") resaltar(d, d.celda(tr, 10));
};

/** Evita que un bloque de filas se parta entre dos páginas. */
function mantenerJuntas(d, filas) {
  filas.forEach((tr, i) => {
    hijoOrdenado(d, props(d, tr, "trPr"), "cantSplit", ORDEN_TRPR);
    if (i === filas.length - 1) return;
    for (const p of Array.from(tr.getElementsByTagNameNS(W, "p"))) hijoOrdenado(d, props(d, p, "pPr"), "keepNext", ORDEN_PPR);
  });
}

/** Fila que ocupa todo el ancho de la tabla de pasajeros, con el nombre del vehículo. */
function filaEncabezado(d, filaModelo, texto) {
  const tr = filaModelo.cloneNode(true);
  const tcs = hijos(tr, "tc");
  const ancho = tcs.reduce((s, tc) => {
    const w = hijo(hijo(tc, "tcPr") || tc, "tcW");
    return s + (Number(w?.getAttributeNS(W, "w")) || 0);
  }, 0);
  tcs.slice(1).forEach((tc) => tr.removeChild(tc));
  const tc = tcs[0];
  const tcPr = props(d, tc, "tcPr");
  const tcW = hijoOrdenado(d, tcPr, "tcW", ORDEN_TCPR);
  if (ancho) { tcW.setAttributeNS(W, "w:w", String(ancho)); tcW.setAttributeNS(W, "w:type", "dxa"); }
  hijoOrdenado(d, tcPr, "gridSpan", ORDEN_TCPR).setAttributeNS(W, "w:val", String(tcs.length));
  const shd = hijoOrdenado(d, tcPr, "shd", ORDEN_TCPR);
  shd.setAttributeNS(W, "w:val", "clear");
  shd.setAttributeNS(W, "w:color", "auto");
  shd.setAttributeNS(W, "w:fill", "E7E0F0");
  d.celdaLimpia(tc, texto);
  for (const r of Array.from(tc.getElementsByTagNameNS(W, "r"))) hijoOrdenado(d, props(d, r, "rPr"), "b", ORDEN_RPR);
  for (const p of Array.from(tc.getElementsByTagNameNS(W, "p"))) {
    const pPr = props(d, p, "pPr");
    hijoOrdenado(d, pPr, "keepNext", ORDEN_PPR);
    hijoOrdenado(d, pPr, "jc", ORDEN_PPR).setAttributeNS(W, "w:val", "left");
  }
  return tr;
}

function llenarVehiculo(d, filas, veh) {
  d.set(filas[1], 3, `Marca: ${v(veh.marca)}`);
  d.set(filas[1], 7, `Modelo: ${v(veh.modelo)}`);
  d.set(filas[2], 3, `Año: ${v(veh.anio)}`);
  d.set(filas[2], 5, `Capacidad (pasajeros): ${v(veh.capacidad)}`);
  d.set(filas[3], 3, `Placas: ${v(veh.placas)}`);
  d.set(filas[3], 6, `Color: ${v(veh.color)}`);
  d.set(filas[4], 0, `Nombre del conductor: ${v(veh.conductor)}`);
  d.set(filas[5], 0, `Número de licencia del conductor: ${v(veh.licencia)}`);
  d.set(filas[5], 4, `Expedida por: ${v(veh.expedida)}`);
  if (v(veh.poliza)) d.set(filas[6], 2, `Número de póliza del seguro del auto: ${v(veh.poliza)}`);
  siNo(filas[6], d, veh.chkPoliza);
  siNo(filas[7], d, veh.chkLicencia);
  siNo(filas[8], d, veh.chkTarjeta);
}

/**
 * @param {ArrayBuffer} plantilla  bytes del .docx oficial
 * @param {object} datos           estado del formulario
 * @param {Array} pasajeros        [{nombre, cum, cat, grupo?}] (grupo = encabezado del vehículo)
 * @param {Array} anexos           [{titulo, imagenes:[{bytes: Uint8Array (JPEG), w, h}]}]
 * @returns {Promise<Blob>}
 */
async function generarProtocoloDocx(plantilla, datos, pasajeros, anexos = []) {
  const zip = await JSZip.loadAsync(plantilla);
  const docFile = zip.file("word/document.xml");
  if (!docFile) throw new Error("La plantilla no es un .docx válido.");
  const d = new DocxLlenador(await docFile.async("string"));
  const tablas = d.tablas();
  if (tablas.length < 6) throw new Error("La plantilla ASMAC no tiene la estructura esperada (se esperaban 6 tablas).");
  const [t0, t1, t2, t3, t4, t5] = tablas.map((t) => d.filas(t));
  const e = datos.evento;
  const medios = datos.medios || [];

  // Título y lugar del evento (párrafos sueltos y celdas con el marcador)
  d.reemplazar("(AQUÍ EL NOMBRE DEL EVENTO)", v(e.nombre));
  d.reemplazar("(AQUÍ EL LUGAR DONDE SERÁ EL EVENTO)", v(e.lugar));

  // Tabla 0: cabecera
  d.set(t0[0], 2, v(e.provincia));
  d.set(t0[0], 9, v(e.grupos));
  d.set(t0[1], 3, v(e.coordNombre));
  d.set(t0[1], 10, `Cargo: ${v(e.coordCargo)}`);
  d.set(t0[2], 3, `Local: ${v(e.coordTelLocal)}`);
  d.set(t0[3], 3, `Celular: ${v(e.coordTelCelular)}`);
  d.set(t0[5], 3, v(e.enlaceNombre));
  d.set(t0[5], 11, `Cargo: ${v(e.enlaceCargo)}`);
  d.set(t0[6], 3, `Celular 1: ${v(e.enlaceCel1)}`);
  d.set(t0[7], 3, `Celular 2: ${v(e.enlaceCel2)}`);

  const c = conteos(pasajeros);
  d.set(t0[10], 1, String(c.ML));
  d.set(t0[10], 3, String(c.TS));
  d.set(t0[10], 4, String(c.CC));
  d.set(t0[10], 5, String(c.CR));
  d.set(t0[10], 7, String(c.S));
  d.set(t0[10], 8, String(c.D));
  d.set(t0[10], 12, String(c.OTRO));
  d.set(t0[11], 5, String(c.jovenes));
  d.set(t0[12], 5, String(c.adultos));

  // Tabla 1: salida y regreso
  d.set(t1[1], 0, `Fecha y hora ${v(e.salidaFechaHora)}`);
  d.set(t1[1], 1, `Fecha y hora ${v(e.regresoFechaHora)}`);
  d.set(t1[2], 0, `Fecha y hora estimada de llegada a ${v(e.lugar)}: ${v(e.salidaLlegada)}`);
  d.set(t1[2], 1, `Fecha y hora estimada de llegada a su ciudad de origen: ${v(e.regresoLlegada)}`);
  d.set(t1[3], 0, `Punto de salida: ${v(e.salidaPunto)}`);
  d.set(t1[3], 1, `Punto de llegada: ${v(e.regresoPunto)}`);

  // Tabla 2: medios (X) + transporte contratado + vehículos particulares
  const filaMedio = { contratado: 0, particular: 1, publico: 2, aereo: 3 };
  for (const [m, fila] of Object.entries(filaMedio)) if (medios.includes(m)) d.set(t2[fila], 8, "X");

  if (medios.includes("contratado")) {
    const k = datos.contratado || {};
    d.set(t2[6], 3, v(k.razonSocial));
    d.set(t2[7], 3, v(k.contactos));
    d.set(t2[8], 3, v(k.marcaModelo));
    d.set(t2[9], 3, v(k.placas));
    d.set(t2[10], 3, v(k.operador));
    siNo(t2[11], d, k.chkContrato);
    siNo(t2[12], d, k.chkFeMecanica);
    siNo(t2[13], d, k.chkSeguroViajero);
    siNo(t2[14], d, k.chkLicencia);
    siNo(t2[15], d, k.chkTarjeta);
  }

  const VB_INICIO = 17, VB_LARGO = 10;
  const tbl2 = tablas[2];
  const txtBloque = Array.from(t2[VB_INICIO].getElementsByTagNameNS(W, "t")).map((n) => n.textContent).join("").toUpperCase();
  if (!txtBloque.includes("VEHICULO PARTICULAR")) {
    throw new Error("La plantilla ASMAC cambió de estructura: no se encontró el bloque de vehículo particular.");
  }
  const vehiculos = medios.includes("particular") ? (datos.vehiculos || []) : [];
  const pristinos = t2.slice(VB_INICIO, VB_INICIO + VB_LARGO).map((tr) => tr.cloneNode(true));
  if (vehiculos.length) {
    const primero = t2.slice(VB_INICIO, VB_INICIO + VB_LARGO);
    llenarVehiculo(d, primero, vehiculos[0]);
    mantenerJuntas(d, primero.slice(0, VB_LARGO - 1)); // la última fila es el separador en blanco
    for (const veh of vehiculos.slice(1)) {
      const nuevos = pristinos.map((tr) => tr.cloneNode(true));
      nuevos.forEach((tr) => tbl2.appendChild(tr));
      llenarVehiculo(d, nuevos, veh);
      mantenerJuntas(d, nuevos.slice(0, VB_LARGO - 1));
    }
  }

  // Tabla 3: transporte público (autobús)
  if (medios.includes("publico")) {
    const p = datos.publico || {};
    d.set(t3[3], 1, v(p.lineaSalida));
    d.set(t3[3], 2, v(p.lineaRegreso));
    d.set(t3[4], 1, v(p.fechaSalida));
    d.set(t3[4], 2, v(p.fechaRegreso));
    d.set(t3[6], 0, v(p.descripcion));
  }

  // Tabla 4: transporte aéreo
  if (medios.includes("aereo")) {
    const a = datos.aereo || {};
    const tramo = (base, x = {}) => {
      d.set(t4[base], 1, v(x.aerolinea));
      if (v(x.aeropuerto)) {
        d.set(t4[base + 4], 1, "X");
        d.set(t4[base + 4], 2, `Otro (especifica): ${v(x.aeropuerto)}`);
      }
      d.set(t4[base + 5], 1, v(x.vuelo));
      d.set(t4[base + 6], 1, v(x.fecha));
      d.set(t4[base + 8], 0, v(x.descripcion));
    };
    tramo(2, a.ida);
    tramo(12, a.regreso);
  }

  // Tabla 5: lista de pasajeros (se clona la fila plantilla una vez por pasajero)
  if (t5.length < 3) throw new Error("La plantilla ASMAC cambió de estructura: no se encontró la tabla de pasajeros.");
  const tbl5 = tablas[5];
  const filaModelo = t5[2].cloneNode(true);
  t5.slice(2).forEach((tr) => tbl5.removeChild(tr));
  // Si los pasajeros traen `grupo` (vehículo asignado), se agrupan con una fila de encabezado por unidad.
  let grupoActual = null;
  pasajeros.forEach((pas, i) => {
    if (pas.grupo && pas.grupo !== grupoActual) {
      grupoActual = pas.grupo;
      tbl5.appendChild(filaEncabezado(d, filaModelo, pas.grupo));
    }
    const tr = filaModelo.cloneNode(true);
    tbl5.appendChild(tr);
    d.set(tr, 0, String(i + 1));
    d.set(tr, 1, v(pas.nombre));
    d.set(tr, 2, v(pas.cum));
    d.set(tr, COL_ROSTER[pas.cat] ?? 9, "X");
  });

  await agregarAnexos(zip, d, anexos);

  zip.file("word/document.xml", d.xml());
  return zip.generateAsync({
    type: "blob",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    compression: "DEFLATE",
  });
}

// ───────── Anexos en imagen (al final del documento) ─────────
const NS_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships";
const NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types";
const REL_IMAGEN = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const EMU_POR_TWIP = 635;

const escXml = (s) => String(s).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));

function parrafosDesdeXml(d, xmlInterno) {
  const frag = new DOMParser().parseFromString(
    `<raiz xmlns:w="${W}" xmlns:wp="${NS_WP}" xmlns:a="${NS_A}" xmlns:pic="${NS_PIC}" xmlns:r="${NS_R}">${xmlInterno}</raiz>`,
    "application/xml");
  if (frag.getElementsByTagName("parsererror").length) throw new Error("No se pudo armar el anexo.");
  return Array.from(frag.documentElement.childNodes).filter((n) => n.nodeType === 1).map((n) => d.doc.importNode(n, true));
}

async function agregarAnexos(zip, d, anexos) {
  const total = anexos.reduce((s, a) => s + a.imagenes.length, 0);
  if (!total) return;

  // Relaciones del documento
  const relsRuta = "word/_rels/document.xml.rels";
  const rels = new DOMParser().parseFromString(await zip.file(relsRuta).async("string"), "application/xml");
  const ids = new Set(Array.from(rels.getElementsByTagNameNS(NS_PKG_REL, "Relationship")).map((r) => r.getAttribute("Id")));

  // Tipo de contenido JPEG
  const ctRuta = "[Content_Types].xml";
  const ct = new DOMParser().parseFromString(await zip.file(ctRuta).async("string"), "application/xml");
  const tieneJpeg = Array.from(ct.getElementsByTagNameNS(NS_CT, "Default"))
    .some((x) => (x.getAttribute("Extension") || "").toLowerCase() === "jpeg");
  if (!tieneJpeg) {
    const def = ct.createElementNS(NS_CT, "Default");
    def.setAttribute("Extension", "jpeg");
    def.setAttribute("ContentType", "image/jpeg");
    ct.documentElement.insertBefore(def, ct.documentElement.firstChild);
  }

  // Espacio útil de la página (según la sección final de la plantilla)
  const sectPr = hijo(d.body, "sectPr");
  const pgSz = sectPr && hijo(sectPr, "pgSz"), pgMar = sectPr && hijo(sectPr, "pgMar");
  const num = (el, a, def) => Number(el?.getAttributeNS(W, a)) || def;
  const anchoTw = num(pgSz, "w", 12240) - num(pgMar, "left", 1418) - num(pgMar, "right", 1418);
  const altoTw = num(pgSz, "h", 15840) - num(pgMar, "top", 1701) - num(pgMar, "bottom", 1417);
  const maxW = anchoTw * EMU_POR_TWIP;
  const maxH = Math.round(altoTw * EMU_POR_TWIP * 0.72); // deja lugar al título del anexo

  const insertar = (nodos) => nodos.forEach((n) => d.body.insertBefore(n, sectPr));
  insertar(parrafosDesdeXml(d,
    `<w:p><w:r><w:br w:type="page"/></w:r></w:p>` +
    `<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="28"/></w:rPr><w:t>ANEXOS</w:t></w:r></w:p>` +
    `<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:sz w:val="20"/></w:rPr>` +
    `<w:t xml:space="preserve">Documentos en imagen que acompañan este protocolo (${total} ${total === 1 ? "imagen" : "imágenes"}).</w:t></w:r></w:p>`));

  let n = 0, docPrId = 900000001, primero = true;
  for (const anexo of anexos) {
    for (let i = 0; i < anexo.imagenes.length; i++) {
      const img = anexo.imagenes[i];
      n++;
      let rid = `rIdAnexo${n}`;
      while (ids.has(rid)) rid += "x";
      ids.add(rid);
      const archivo = `media/anexo${n}.jpeg`;
      zip.file(`word/${archivo}`, img.bytes);
      const rel = rels.createElementNS(NS_PKG_REL, "Relationship");
      rel.setAttribute("Id", rid);
      rel.setAttribute("Type", REL_IMAGEN);
      rel.setAttribute("Target", archivo);
      rels.documentElement.appendChild(rel);

      // Tamaño: llenar el ancho sin pasar del alto disponible, conservando proporción
      const k = Math.min(maxW / img.w, maxH / img.h);
      const cx = Math.round(img.w * k), cy = Math.round(img.h * k);
      const sufijo = anexo.imagenes.length > 1 ? ` (${i + 1} de ${anexo.imagenes.length})` : "";
      const salto = primero ? "" : `<w:p><w:r><w:br w:type="page"/></w:r></w:p>`;
      primero = false;
      insertar(parrafosDesdeXml(d, salto +
        `<w:p><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120"/></w:pPr>` +
        `<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">Anexo ${n}. ${escXml(anexo.titulo)}${sufijo}</w:t></w:r></w:p>` +
        `<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:drawing>` +
        `<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/>` +
        `<wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="${docPrId++}" name="Anexo ${n}"/>` +
        `<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>` +
        `<a:graphic><a:graphicData uri="${NS_PIC}"><pic:pic>` +
        `<pic:nvPicPr><pic:cNvPr id="${n}" name="anexo${n}.jpeg"/><pic:cNvPicPr/></pic:nvPicPr>` +
        `<pic:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
        `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
        `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
        `<a:ln w="6350"><a:solidFill><a:srgbClr val="8C8C8C"/></a:solidFill></a:ln></pic:spPr></pic:pic></a:graphicData></a:graphic>` +
        `</wp:inline></w:drawing></w:r></w:p>`));
    }
  }

  zip.file(relsRuta, new XMLSerializer().serializeToString(rels));
  zip.file(ctRuta, new XMLSerializer().serializeToString(ct));
}
