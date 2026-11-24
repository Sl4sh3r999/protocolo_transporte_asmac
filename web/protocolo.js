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
const siNo = (tr, d, valor) => {
  if (valor === "si") d.set(tr, 9, "X");
  else if (valor === "no") d.set(tr, 10, "X");
};

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
 * @param {Array} pasajeros        [{nombre, cum, cat}]
 * @returns {Promise<Blob>}
 */
async function generarProtocoloDocx(plantilla, datos, pasajeros) {
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
    llenarVehiculo(d, t2.slice(VB_INICIO, VB_INICIO + VB_LARGO), vehiculos[0]);
    for (const veh of vehiculos.slice(1)) {
      const nuevos = pristinos.map((tr) => tr.cloneNode(true));
      nuevos.forEach((tr) => tbl2.appendChild(tr));
      llenarVehiculo(d, nuevos, veh);
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
  pasajeros.forEach((pas, i) => {
    const tr = filaModelo.cloneNode(true);
    tbl5.appendChild(tr);
    d.set(tr, 0, String(i + 1));
    d.set(tr, 1, v(pas.nombre));
    d.set(tr, 2, v(pas.cum));
    d.set(tr, COL_ROSTER[pas.cat] ?? 9, "X");
  });

  zip.file("word/document.xml", d.xml());
  return zip.generateAsync({
    type: "blob",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    compression: "DEFLATE",
  });
}
