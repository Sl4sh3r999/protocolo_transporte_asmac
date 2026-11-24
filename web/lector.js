"use strict";
// Lee la lista de Scouts de un .xlsx (exportado desde CarpetApp: app.scouts.org.mx/scouts) o .csv
// directamente en el navegador. Nada sale del dispositivo.

const NS_SS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

function normalizar(s) {
  return String(s ?? "").normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase().replace(/\s+/g, " ").trim();
}

function colIndex(ref) {
  const letras = /^[A-Z]+/.exec(ref || "")?.[0] || "A";
  let n = 0;
  for (const c of letras) n = n * 26 + (c.charCodeAt(0) - 64);
  return n - 1;
}

function parseXml(txt) {
  const doc = new DOMParser().parseFromString(txt, "application/xml");
  if (doc.getElementsByTagName("parsererror").length) throw new Error("XML inválido dentro del archivo.");
  return doc;
}

async function leerXlsx(buffer) {
  let zip;
  try { zip = await JSZip.loadAsync(buffer); }
  catch { throw new Error("El archivo no parece un Excel .xlsx válido. Si es .xls antiguo, ábrelo y guárdalo como .xlsx."); }

  // Textos compartidos
  const compartidos = [];
  const ssFile = zip.file("xl/sharedStrings.xml");
  if (ssFile) {
    const ss = parseXml(await ssFile.async("string"));
    for (const si of ss.getElementsByTagNameNS(NS_SS, "si")) {
      let t = "";
      for (const n of si.getElementsByTagNameNS(NS_SS, "t")) {
        // Ignora textos fonéticos (rPh)
        if (n.parentNode.localName === "rPh" || n.parentNode.parentNode?.localName === "rPh") continue;
        t += n.textContent;
      }
      compartidos.push(t);
    }
  }

  // Primera hoja según workbook.xml
  let rutaHoja = "xl/worksheets/sheet1.xml";
  const wbFile = zip.file("xl/workbook.xml"), relsFile = zip.file("xl/_rels/workbook.xml.rels");
  if (wbFile && relsFile) {
    const wb = parseXml(await wbFile.async("string"));
    const rels = parseXml(await relsFile.async("string"));
    const hoja = wb.getElementsByTagNameNS(NS_SS, "sheet")[0];
    const rid = hoja?.getAttributeNS(NS_REL, "id");
    for (const r of rels.getElementsByTagName("Relationship")) {
      if (r.getAttribute("Id") === rid) {
        const t = r.getAttribute("Target");
        rutaHoja = t.startsWith("/") ? t.slice(1) : "xl/" + t.replace(/^\.\//, "");
      }
    }
  }
  const hojaFile = zip.file(rutaHoja) || zip.file(/xl\/worksheets\/sheet\d*\.xml$/)[0];
  if (!hojaFile) throw new Error("No se encontró ninguna hoja dentro del Excel.");
  const hoja = parseXml(await hojaFile.async("string"));

  const filas = [];
  for (const row of hoja.getElementsByTagNameNS(NS_SS, "row")) {
    const fila = [];
    for (const c of row.getElementsByTagNameNS(NS_SS, "c")) {
      const tipo = c.getAttribute("t");
      const v = c.getElementsByTagNameNS(NS_SS, "v")[0]?.textContent ?? "";
      let valor;
      if (tipo === "s") valor = compartidos[Number(v)] ?? "";
      else if (tipo === "inlineStr") valor = [...c.getElementsByTagNameNS(NS_SS, "t")].map((n) => n.textContent).join("");
      else valor = v;
      fila[colIndex(c.getAttribute("r"))] = valor;
    }
    const idx = Number(row.getAttribute("r")) - 1;
    filas[Number.isFinite(idx) && idx >= 0 ? idx : filas.length] = fila;
  }
  return Array.from(filas, (f) => Array.from(f || [], (x) => (x ?? "").toString().trim()));
}

function leerCsv(texto) {
  const sep = (texto.split("\n")[0].match(/;/g) || []).length > (texto.split("\n")[0].match(/,/g) || []).length ? ";" : ",";
  const filas = [];
  let fila = [], campo = "", comillas = false;
  for (let i = 0; i < texto.length; i++) {
    const ch = texto[i];
    if (comillas) {
      if (ch === '"' && texto[i + 1] === '"') { campo += '"'; i++; }
      else if (ch === '"') comillas = false;
      else campo += ch;
    } else if (ch === '"') comillas = true;
    else if (ch === sep) { fila.push(campo.trim()); campo = ""; }
    else if (ch === "\n") { fila.push(campo.trim()); filas.push(fila); fila = []; campo = ""; }
    else if (ch !== "\r") campo += ch;
  }
  if (campo || fila.length) { fila.push(campo.trim()); filas.push(fila); }
  return filas;
}

// Columnas reconocidas (encabezados de CarpetApp y variantes comunes)
const COLUMNAS = {
  nombre: [/^nombre( completo)?$/, /^nombre/],
  apellidos: [/^apellidos?$/],
  apPaterno: [/apellido paterno/],
  apMaterno: [/apellido materno/],
  cum: [/^cum$/, /numero scout/, /^num\.? ?scout/, /registro/],
  edad: [/^edad/],
  seccion: [/^seccion/, /^rama/],
  pg: [/pequeno grupo/, /patrulla/, /seisena/, /equipo/],
  cargo: [/^cargo/],
  estatus: [/^estatus/, /^status/, /^estado/, /vigencia/],
};

function mapearEncabezados(fila) {
  const mapa = {};
  fila.forEach((h, i) => {
    const n = normalizar(h);
    if (!n) return;
    for (const [clave, patrones] of Object.entries(COLUMNAS)) {
      if (mapa[clave] !== undefined) continue;
      if (patrones.some((p) => p.test(n))) { mapa[clave] = i; break; }
    }
  });
  return mapa;
}

const CATEGORIA_POR_SECCION = [
  [/manada|lobat/, "ML"],
  [/tropa|scout(s)? ?$/, "TS"],
  [/comunidad|caminante/, "CC"],
  [/clan|rover/, "CR"],
];

function categoriaDeSeccion(seccion) {
  const n = normalizar(seccion);
  for (const [re, cat] of CATEGORIA_POR_SECCION) if (re.test(n)) return cat;
  return "";
}

function esVigente(estatus) {
  const n = normalizar(estatus);
  return /vigente/.test(n) && !/no vigente/.test(n);
}

/** Convierte las filas leídas en una lista de Scouts. */
function interpretarTabla(filas) {
  const inicio = filas.findIndex((f) => f && f.some((c) => /^nombre/.test(normalizar(c))));
  if (inicio < 0) throw new Error('No se encontró la fila de encabezados (debe existir una columna "Nombre").');
  const mapa = mapearEncabezados(filas[inicio]);
  const tieneEstatus = mapa.estatus !== undefined;
  const scouts = [];
  for (let i = inicio + 1; i < filas.length; i++) {
    const f = filas[i];
    if (!f) continue;
    const g = (k) => (mapa[k] !== undefined ? (f[mapa[k]] || "").trim() : "");
    let nombre = g("nombre");
    const apellidos = [g("apellidos"), g("apPaterno"), g("apMaterno")].filter(Boolean).join(" ");
    if (apellidos) nombre = `${nombre} ${apellidos}`.trim();
    if (!nombre) continue;
    const estatus = g("estatus");
    scouts.push({
      id: `x${i}`,
      nombre: nombre.replace(/\s+/g, " "),
      cum: g("cum"),
      edad: g("edad"),
      seccion: g("seccion"),
      pg: g("pg"),
      cargo: g("cargo"),
      estatus,
      vigente: tieneEstatus ? esVigente(estatus) : true,
      cat: categoriaDeSeccion(g("seccion")) || "OTRO",
    });
  }
  return { scouts, tieneEstatus };
}

async function leerArchivoScouts(file) {
  const nombre = file.name.toLowerCase();
  let filas;
  if (nombre.endsWith(".csv") || file.type === "text/csv") filas = leerCsv(await file.text());
  else filas = await leerXlsx(await file.arrayBuffer());
  return interpretarTabla(filas);
}
