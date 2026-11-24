"use strict";

const $ = (id) => document.getElementById(id);
const CLAVE_GUARDADO = "protocolo_transporte_asmac_v1";
const NOMBRE_CAT = { ML: "Manada", TS: "Tropa", CC: "Comunidad", CR: "Clan", S: "Scouter", D: "Dirigente", OTRO: "No scout" };
const ORDEN_CAT = ["ML", "TS", "CC", "CR", "S", "D", "OTRO"];

function estadoVacio() {
  return {
    evento: {}, medios: [], contratado: {}, publico: {}, aereo: { ida: {}, regreso: {} },
    vehiculos: [], scouts: [], tieneEstatus: true, archivo: "", seleccion: [], manuales: [],
  };
}

// Campos que se eligen con calendario y reloj (<input type="datetime-local">).
// Se guardan como "2026-12-12T08:00" y en el documento salen como "12/12/2026 08:00 hrs".
const CAMPOS_FECHA = [
  "evento.salidaFechaHora", "evento.salidaLlegada", "evento.regresoFechaHora", "evento.regresoLlegada",
  "publico.fechaSalida", "publico.fechaRegreso", "aereo.ida.fecha", "aereo.regreso.fecha",
];

let S = cargar();

function cargar() {
  try {
    const x = JSON.parse(localStorage.getItem(CLAVE_GUARDADO) || "null");
    if (x && typeof x === "object") return migrarFechas(Object.assign(estadoVacio(), x));
  } catch { /* almacenamiento no disponible o dañado */ }
  return estadoVacio();
}

/** Convierte fechas escritas a mano en versiones anteriores ("12/12/2026 08:00") al formato del calendario. */
function migrarFechas(estado) {
  for (const ruta of CAMPOS_FECHA) {
    const val = getRuta(estado, ruta);
    if (!val || /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}/.test(val)) continue;
    const m = /^(\d{1,2})\/(\d{1,2})\/(\d{2,4})(?:\D+(\d{1,2}):(\d{2}))?/.exec(String(val).trim());
    if (!m) continue; // se deja como estaba; el documento lo usa tal cual
    const anio = m[3].length === 2 ? "20" + m[3] : m[3];
    const p2 = (n) => String(n || 0).padStart(2, "0");
    setRuta(estado, ruta, `${anio}-${p2(m[2])}-${p2(m[1])}T${p2(m[4])}:${p2(m[5])}`);
  }
  return estado;
}

/** "2026-12-12T08:00" → "12/12/2026 08:00 hrs" (lo que va en el Word). */
function formatearFecha(val) {
  const m = /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})/.exec(val || "");
  return m ? `${m[3]}/${m[2]}/${m[1]} ${m[4]}:${m[5]} hrs` : (val || "");
}

function datosParaDocumento() {
  const copia = JSON.parse(JSON.stringify(S));
  for (const ruta of CAMPOS_FECHA) setRuta(copia, ruta, formatearFecha(getRuta(copia, ruta)));
  return copia;
}

let tGuardar = null;
function guardar() {
  clearTimeout(tGuardar);
  tGuardar = setTimeout(() => {
    try { localStorage.setItem(CLAVE_GUARDADO, JSON.stringify(S)); } catch { /* modo privado */ }
  }, 250);
}

const escapar = (s) => String(s ?? "").replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
function getRuta(obj, ruta) { return ruta.split(".").reduce((o, k) => (o ? o[k] : undefined), obj); }
function setRuta(obj, ruta, valor) {
  const partes = ruta.split(".");
  let o = obj;
  for (const k of partes.slice(0, -1)) o = o[k] ??= {};
  o[partes.at(-1)] = valor;
}

let tAviso = null;
function aviso(el, txt, error = false) {
  el.textContent = txt;
  el.classList.toggle("error", error);
  clearTimeout(tAviso);
  if (txt && !error) tAviso = setTimeout(() => (el.textContent = ""), 5000);
}

// ───────── Documentos "¿Tienes en tu poder…?" (Sí / No + anexo) ─────────
const DOCS_CONTRATADO = [
  ["chkContrato", "Contrato del servicio de transporte"],
  ["chkFeMecanica", "Fe mecánica de la unidad"],
  ["chkSeguroViajero", "Seguro del viajero"],
  ["chkLicencia", "Copia de licencia del conductor"],
  ["chkTarjeta", "Copia de tarjeta de circulación"],
];
const DOCS_VEHICULO = [
  ["chkPoliza", "Póliza del seguro del auto"],
  ["chkLicencia", "Copia de licencia del conductor"],
  ["chkTarjeta", "Copia de tarjeta de circulación"],
];
const ES_TACTIL = matchMedia("(pointer: coarse)").matches;

/**
 * Fila con la pregunta, botones Sí/No y el área para adjuntar la foto o archivo.
 * Adjuntar algo marca "Sí" automáticamente; marcar "Sí" sin adjuntar queda en rojo.
 */
function montarSiNo(div, leer, escribir, claveAdj) {
  const etiqueta = div.textContent.trim();
  div.textContent = "";
  div.classList.add("doc");

  const fila = document.createElement("div");
  fila.className = "doc-fila";
  const sp = document.createElement("span");
  sp.className = "doc-nombre";
  sp.textContent = etiqueta;
  const grupo = document.createElement("span");
  grupo.className = "sino-btns";
  for (const [val, txt] of [["si", "Sí"], ["no", "No"]]) {
    const b = document.createElement("button");
    b.type = "button"; b.textContent = txt; b.dataset.val = val;
    b.addEventListener("click", () => {
      escribir(leer() === val ? "" : val);
      pintar(); actualizarResumen();
    });
    grupo.appendChild(b);
  }
  fila.append(sp, grupo);

  const panel = document.createElement("div");
  panel.className = "adj";
  const msg = document.createElement("p");
  msg.className = "adj-msg";
  const botones = document.createElement("div");
  botones.className = "adj-btns";
  const miniaturas = document.createElement("div");
  miniaturas.className = "adj-mini";

  const crearSelector = (texto, accept, capture) => {
    const lab = document.createElement("label");
    lab.className = "btn sec adj-btn";
    lab.textContent = texto;
    const inp = document.createElement("input");
    inp.type = "file"; inp.accept = accept; inp.hidden = true;
    if (capture) inp.setAttribute("capture", "environment");
    else inp.multiple = true;
    inp.addEventListener("change", async () => {
      const files = [...inp.files];
      inp.value = "";
      await adjuntar(files);
    });
    lab.appendChild(inp);
    botones.appendChild(lab);
  };
  if (ES_TACTIL) crearSelector("📷 Tomar foto", "image/*", true);
  crearSelector(ES_TACTIL ? "🖼️ Galería o archivo" : "📎 Elegir foto o PDF", "image/*,application/pdf,.pdf", false);

  async function adjuntar(files) {
    if (!files.length) return;
    msg.textContent = "Procesando…";
    msg.className = "adj-msg";
    const errores = [];
    for (const f of files) {
      try {
        const { imagenes, omitidas } = await procesarArchivoAnexo(f);
        for (let i = 0; i < imagenes.length; i++) {
          const nombre = imagenes.length > 1 ? `${f.name} (pág. ${i + 1})` : f.name;
          await guardarAdjunto(claveAdj, nombre, imagenes[i]);
        }
        if (omitidas) errores.push(`De "${f.name}" solo se tomaron las primeras ${imagenes.length} páginas.`);
      } catch (err) {
        errores.push(err.message || `No se pudo leer "${f.name}".`);
      }
    }
    if (adjuntosDe(claveAdj).length && leer() !== "si") escribir("si");
    pintar(); actualizarResumen();
    if (errores.length) { msg.textContent = errores.join(" "); msg.className = "adj-msg mal"; }
  }

  function pintarMiniaturas() {
    miniaturas.innerHTML = "";
    for (const a of adjuntosDe(claveAdj)) {
      const fig = document.createElement("figure");
      const img = document.createElement("img");
      img.src = a.url; img.alt = a.nombre;
      const quitar = document.createElement("button");
      quitar.type = "button"; quitar.className = "adj-quitar"; quitar.textContent = "✕";
      quitar.setAttribute("aria-label", `Quitar ${a.nombre}`);
      quitar.addEventListener("click", async () => { await borrarAdjunto(claveAdj, a.id); pintar(); actualizarResumen(); });
      const cap = document.createElement("figcaption");
      cap.textContent = a.nombre;
      fig.append(img, quitar, cap);
      miniaturas.appendChild(fig);
    }
  }

  function pintar() {
    const v = leer(), n = adjuntosDe(claveAdj).length;
    grupo.querySelectorAll("button").forEach((b) => b.classList.toggle("on", b.dataset.val === v));
    panel.hidden = v === "no";
    div.classList.toggle("falta", v === "si" && !n);
    div.classList.toggle("listo", v === "si" && n > 0);
    if (v === "si" && !n) { msg.textContent = "⚠️ Falta adjuntar la foto o el archivo de este documento."; msg.className = "adj-msg mal"; }
    else if (n) { msg.textContent = n === 1 ? "✔ 1 imagen adjunta: se incluirá como anexo en el documento." : `✔ ${n} imágenes adjuntas: se incluirán como anexo en el documento.`; msg.className = "adj-msg ok"; }
    else { msg.textContent = "Si lo tienes, adjunta la foto o el archivo (se marca «Sí» solo)."; msg.className = "adj-msg"; }
    pintarMiniaturas();
  }

  panel.append(msg, botones, miniaturas);
  div.append(fila, panel);
  pintar();
}

/** Lista de documentos que aplican según los medios elegidos, con su respuesta y anexos. */
function documentosRequeridos() {
  const out = [];
  if (S.medios.includes("contratado")) {
    const ctx = ["Transporte contratado", S.contratado.razonSocial].filter(Boolean).join(" – ");
    for (const [k, nombre] of DOCS_CONTRATADO) {
      out.push({ nombre, contexto: ctx, valor: S.contratado[k] || "", adjuntos: adjuntosDe(`contratado.${k}`) });
    }
  }
  if (S.medios.includes("particular")) {
    S.vehiculos.forEach((veh, i) => {
      const desc = [veh.marca, veh.modelo, veh.placas].filter(Boolean).join(" ");
      const ctx = `Vehículo particular ${i + 1}${desc ? ` (${desc})` : ""}`;
      for (const [k, nombre] of DOCS_VEHICULO) {
        out.push({ nombre, contexto: ctx, valor: veh[k] || "", adjuntos: adjuntosDe(`veh:${veh.id}:${k}`) });
      }
    });
  }
  return out;
}

// ───────── Campos simples (data-k) ─────────
function montarCampos() {
  document.querySelectorAll("main [data-k]").forEach((el) => {
    const k = el.dataset.k;
    if (el.classList.contains("sino")) {
      montarSiNo(el, () => getRuta(S, k) || "", (v) => { setRuta(S, k, v); guardar(); }, k);
      return;
    }
    el.value = getRuta(S, k) || "";
    el.addEventListener("input", () => { setRuta(S, k, el.value); guardar(); actualizarResumen(); });
    if (el.type === "datetime-local") {
      // Abre el calendario al tocar cualquier parte del campo, no solo el iconito.
      el.addEventListener("click", () => { try { el.showPicker(); } catch { /* navegador sin showPicker */ } });
    }
  });

  document.querySelectorAll("[data-medio]").forEach((chk) => {
    chk.checked = S.medios.includes(chk.dataset.medio);
    chk.addEventListener("change", () => {
      const m = chk.dataset.medio;
      S.medios = S.medios.filter((x) => x !== m);
      if (chk.checked) S.medios.push(m);
      if (m === "particular" && chk.checked && !S.vehiculos.length) S.vehiculos.push({});
      mostrarMedios(); renderVehiculos(); guardar(); actualizarResumen();
    });
  });
  mostrarMedios();
}

function mostrarMedios() {
  document.querySelectorAll(".sub[data-para]").forEach((s) => (s.hidden = !S.medios.includes(s.dataset.para)));
}

// ───────── Vehículos ─────────
function renderVehiculos() {
  const cont = $("vehiculos");
  cont.innerHTML = "";
  S.vehiculos.forEach((veh, i) => {
    veh.id ??= `${Date.now().toString(36)}${Math.random().toString(36).slice(2, 7)}`;
    const frag = $("tpl-vehiculo").content.cloneNode(true);
    frag.querySelector(".n").textContent = i + 1;
    frag.querySelectorAll("[data-v]").forEach((el) => {
      const k = el.dataset.v;
      if (el.classList.contains("sino")) {
        montarSiNo(el, () => veh[k] || "", (v) => { veh[k] = v; guardar(); }, `veh:${veh.id}:${k}`);
      } else {
        el.value = veh[k] || "";
        el.addEventListener("input", () => { veh[k] = el.value; guardar(); actualizarResumen(); });
      }
    });
    frag.querySelector(".quitar").addEventListener("click", async () => {
      const n = DOCS_VEHICULO.reduce((s, [k]) => s + adjuntosDe(`veh:${veh.id}:${k}`).length, 0);
      if (n && !confirm(`Este vehículo tiene ${n} imagen(es) adjunta(s). ¿Quitarlo junto con sus anexos?`)) return;
      await borrarAdjuntosConPrefijo(`veh:${veh.id}:`);
      S.vehiculos.splice(i, 1); renderVehiculos(); guardar(); actualizarResumen();
    });
    cont.appendChild(frag);
  });
  guardar(); // conserva los id de vehículo (las fotos adjuntas dependen de ellos)
}

// ───────── Excel de Scouts ─────────
async function cargarArchivo(file) {
  const est = $("estado-archivo");
  if (!file) return;
  aviso(est, "Leyendo archivo…");
  try {
    const { scouts, tieneEstatus } = await leerArchivoScouts(file);
    if (!scouts.length) throw new Error("No se encontró ninguna persona en el archivo.");
    S.scouts = scouts; S.tieneEstatus = tieneEstatus; S.archivo = file.name;
    S.seleccion = [];
    const vig = scouts.filter((s) => s.vigente).length;
    aviso(est, `✔ ${file.name}: ${scouts.length} personas leídas, ${vig} con registro vigente` +
      (scouts.length - vig ? ` y ${scouts.length - vig} vencidas (no seleccionables).` : "."));
    if (!tieneEstatus) {
      aviso(est, `✔ ${file.name}: ${scouts.length} personas. ⚠️ El archivo no trae columna "Estatus", así que no se pudo ` +
        "revisar quién tiene el registro vencido: verifícalo tú antes de enviar.", true);
    }
    renderScouts(); guardar(); actualizarResumen();
  } catch (err) {
    aviso(est, "No se pudo leer el archivo: " + err.message, true);
  } finally {
    $("archivo").value = "";
  }
}

function renderScouts() {
  const bloque = $("bloque-scouts");
  bloque.hidden = !S.scouts.length;
  if (!S.scouts.length) return;
  const q = ($("buscar").value || "").normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase();
  const sel = new Set(S.seleccion);
  const vigentes = S.scouts.filter((s) => s.vigente);
  const vencidos = S.scouts.filter((s) => !s.vigente);

  const tbody = $("lista-scouts");
  tbody.innerHTML = "";
  for (const s of vigentes) {
    const texto = `${s.nombre} ${s.cum} ${s.pg} ${s.cargo}`.normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase();
    if (q && !texto.includes(q)) continue;
    const tr = document.createElement("tr");
    tr.className = sel.has(s.id) ? "sel" : "";
    const catSel = ORDEN_CAT.slice(0, 4).concat("OTRO").map((c) =>
      `<option value="${c}"${c === s.cat ? " selected" : ""}>${NOMBRE_CAT[c]}</option>`).join("");
    const seccionTd = s.cat === "OTRO" || !s.seccion
      ? `<select class="cat" aria-label="Sección para el documento">${catSel}</select> <small>${escapar(s.seccion)}</small>`
      : escapar(s.seccion);
    tr.innerHTML = `<td class="c-chk"><input type="checkbox" aria-label="Va al evento"${sel.has(s.id) ? " checked" : ""}></td>
      <td>${escapar(s.nombre)}${s.edad ? ` <small>(${escapar(s.edad)})</small>` : ""}</td><td>${escapar(s.cum)}</td>
      <td>${seccionTd}</td><td>${escapar(s.pg)}</td><td>${escapar(s.cargo)}</td>
      <td><span class="badge ok">${escapar(s.estatus || "—")}</span></td>`;
    const chk = tr.querySelector("input");
    const alternar = (on) => {
      S.seleccion = S.seleccion.filter((x) => x !== s.id);
      if (on) S.seleccion.push(s.id);
      tr.classList.toggle("sel", on); guardar(); actualizarResumen(); contador();
    };
    chk.addEventListener("change", () => alternar(chk.checked));
    tr.addEventListener("click", (e) => {
      if (e.target.closest("input,select")) return;
      chk.checked = !chk.checked; alternar(chk.checked);
    });
    tr.querySelector("select.cat")?.addEventListener("change", (e) => { s.cat = e.target.value; guardar(); actualizarResumen(); });
    tbody.appendChild(tr);
  }
  if (!tbody.children.length) tbody.innerHTML = `<tr><td colspan="7" class="vacio">Nadie coincide con la búsqueda.</td></tr>`;

  $("bloque-vencidos").hidden = !vencidos.length;
  $("lista-vencidos").innerHTML = vencidos.map((s) =>
    `<li><span>${escapar(s.nombre)}</span> <small>${escapar(s.cum)} · ${escapar(s.seccion)}</small> <span class="badge mal">${escapar(s.estatus || "Sin estatus")}</span></li>`).join("");
  contador();
}

function contador() {
  const vig = S.scouts.filter((s) => s.vigente).length;
  $("contador-scouts").textContent = `${S.seleccion.length} de ${vig} vigentes seleccionados para el evento` +
    (S.archivo ? ` · archivo: ${S.archivo}` : "");
}

// ───────── Scouters y externos ─────────
function renderManuales() {
  const ul = $("lista-manual");
  ul.innerHTML = "";
  S.manuales.forEach((p, i) => {
    const li = document.createElement("li");
    li.innerHTML = `<span class="badge cat">${NOMBRE_CAT[p.cat] || p.cat}</span>
      <span class="nom">${escapar(p.nombre)}</span>
      <small>${escapar([p.cum, p.nota].filter(Boolean).join(" · "))}</small>
      <button type="button" class="quitar sec" aria-label="Quitar">✕</button>`;
    li.querySelector("button").addEventListener("click", () => {
      S.manuales.splice(i, 1); renderManuales(); guardar(); actualizarResumen();
    });
    ul.appendChild(li);
  });
}

// ───────── Pasajeros y resumen ─────────
function pasajeros() {
  const sel = new Set(S.seleccion);
  const jovenes = S.scouts.filter((s) => s.vigente && sel.has(s.id));
  const lista = [...jovenes, ...S.manuales].map((p) => ({ nombre: p.nombre, cum: p.cum, cat: p.cat || "OTRO" }));
  return lista.sort((a, b) => ORDEN_CAT.indexOf(a.cat) - ORDEN_CAT.indexOf(b.cat));
}

function avisos(lista) {
  const e = S.evento, out = [];
  if (!(e.nombre || "").trim()) out.push("Falta el nombre del evento.");
  if (!(e.coordNombre || "").trim()) out.push("Falta el coordinador de seguridad (es quien NO asiste al evento).");
  if (!(e.enlaceNombre || "").trim()) out.push("Falta el enlace de seguridad (quien SÍ asiste).");
  const antes = (a, b, msg) => { if (e[a] && e[b] && e[b] < e[a]) out.push(msg); };
  antes("salidaFechaHora", "salidaLlegada", "La llegada estimada al evento es anterior a la hora de salida.");
  antes("salidaFechaHora", "regresoFechaHora", "La fecha de regreso es anterior a la de salida.");
  antes("regresoFechaHora", "regresoLlegada", "La llegada a tu ciudad es anterior a la hora de regreso.");
  if (!S.medios.length) out.push("No has marcado ningún medio de transporte.");
  if (!lista.length) out.push("La lista de pasajeros está vacía.");
  if (lista.length && !lista.some((p) => p.cat === "S" || p.cat === "D")) out.push("No hay ningún Scouter o Dirigente en la lista.");
  if (S.medios.includes("particular")) {
    const cap = S.vehiculos.reduce((s, v) => s + (parseInt(v.capacidad, 10) || 0), 0);
    if (cap && cap < lista.length) out.push(`La capacidad de los vehículos (${cap}) es menor que el número de pasajeros (${lista.length}).`);
  }
  if (S.scouts.length && !S.tieneEstatus) out.push("Tu Excel no tiene columna de estatus: revisa a mano que nadie tenga el registro vencido.");
  return out;
}

function actualizarResumen() {
  const lista = pasajeros();
  const c = conteos(lista);
  const celdas = [["Manada", c.ML], ["Tropa", c.TS], ["Comunidad", c.CC], ["Clan", c.CR],
    ["Scouters", c.S], ["Dirigentes", c.D], ["No scouts", c.OTRO], ["Jóvenes", c.jovenes, "tot"], ["Adultos", c.adultos, "tot"], ["Total", c.total, "tot"]];
  $("resumen").innerHTML = celdas.map(([n, x, cls]) => `<div class="kpi ${cls || ""}"><b>${x}</b><span>${n}</span></div>`).join("");
  $("n-pasajeros").textContent = lista.length;
  $("lista-pasajeros").innerHTML = lista.map((p) =>
    `<li>${escapar(p.nombre)} <small>${escapar(p.cum || "")} · ${p.cat}</small></li>`).join("");
  $("avisos").innerHTML = avisos(lista).map((a) => `<li>${escapar(a)}</li>`).join("");
  pintarEstadoAnexos();
}

function problemasAnexos() {
  const docs = documentosRequeridos();
  return {
    docs,
    sinFoto: docs.filter((d) => d.valor === "si" && !d.adjuntos.length),
    sinResponder: docs.filter((d) => !d.valor),
    noTiene: docs.filter((d) => d.valor === "no"),
  };
}

function pintarEstadoAnexos() {
  const cont = $("anexos-estado");
  const { docs, sinFoto, sinResponder, noTiene } = problemasAnexos();
  if (!docs.length) {
    cont.innerHTML = `<p class="ayuda">No hay documentos que anexar para el medio de transporte elegido.</p>`;
    return;
  }
  const listos = docs.filter((d) => d.valor === "si" && d.adjuntos.length);
  const item = (d, txt) => `<li><b>${escapar(d.nombre)}</b> <small>${escapar(d.contexto)}</small> — ${txt}</li>`;
  let html = `<p class="anexos-cuenta ${sinFoto.length || sinResponder.length ? "mal" : "ok"}">📎 Anexos: <b>${listos.length} de ${docs.length}</b> documentos con foto adjunta.</p>`;
  if (sinFoto.length || sinResponder.length) {
    html += `<ul class="anexos-faltan">${sinFoto.map((d) => item(d, "marcaste «Sí» pero <b>falta la foto o archivo</b>")).join("")}` +
      `${sinResponder.map((d) => item(d, "no has respondido si lo tienes")).join("")}</ul>`;
  }
  if (noTiene.length) {
    html += `<ul class="anexos-no">${noTiene.map((d) => item(d, "respondiste que NO lo tienes")).join("")}</ul>`;
  }
  cont.innerHTML = html;
}

/** Anexos listos para el Word: [{titulo, imagenes:[{bytes, w, h}]}] */
async function anexosParaDocumento() {
  const out = [];
  for (const d of documentosRequeridos()) {
    if (d.valor !== "si" || !d.adjuntos.length) continue;
    const imagenes = [];
    for (const a of d.adjuntos) imagenes.push({ bytes: new Uint8Array(await a.blob.arrayBuffer()), w: a.w, h: a.h });
    out.push({ titulo: `${d.nombre} — ${d.contexto}`, imagenes });
  }
  return out;
}

// ───────── Generar documento ─────────
function plantillaBytes() {
  const bin = atob(PLANTILLA_B64);
  const u8 = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i);
  return u8.buffer;
}

function nombreArchivo() {
  const base = (S.evento.nombre || "evento").normalize("NFD").replace(/[̀-ͯ]/g, "")
    .replace(/[^\w\- ]+/g, "").trim().replace(/\s+/g, "_").slice(0, 60) || "evento";
  return `Protocolo_transporte_${base}.docx`;
}

async function generar() {
  if (!(S.evento.nombre || "").trim()) throw new Error("Escribe al menos el nombre del evento (paso 1).");
  const blob = await generarProtocoloDocx(plantillaBytes(), datosParaDocumento(), pasajeros(), await anexosParaDocumento());
  return new File([blob], nombreArchivo(), { type: blob.type });
}

function confirmarAnexos() {
  const { sinFoto, sinResponder } = problemasAnexos();
  if (!sinFoto.length && !sinResponder.length) return true;
  const lineas = [...sinFoto.map((d) => `• ${d.nombre} (${d.contexto}): falta la foto`),
    ...sinResponder.map((d) => `• ${d.nombre} (${d.contexto}): sin responder`)];
  const ok = confirm(`Faltan anexos del protocolo:\n\n${lineas.join("\n")}\n\n` +
    "CNAR pide anexar estos documentos en imagen. ¿Descargar de todos modos?");
  if (!ok) $("anexos-estado").scrollIntoView({ behavior: "smooth", block: "center" });
  return ok;
}

async function descargar() {
  const est = $("estado");
  if (!confirmarAnexos()) return;
  try {
    aviso(est, "Generando documento…");
    const file = await generar();
    const url = URL.createObjectURL(file);
    const a = Object.assign(document.createElement("a"), { href: url, download: file.name });
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 10000);
    aviso(est, `✔ Listo: ${file.name}`);
  } catch (err) {
    aviso(est, "No se pudo generar: " + err.message, true);
  }
}

async function compartir() {
  const est = $("estado");
  if (!confirmarAnexos()) return;
  try {
    const file = await generar();
    await navigator.share({ files: [file], title: file.name });
  } catch (err) {
    if (err.name !== "AbortError") aviso(est, "No se pudo compartir: " + err.message, true);
  }
}

// ───────── Inicio ─────────
async function iniciar() {
  await cargarAdjuntos();
  montarCampos();
  renderVehiculos();
  renderScouts();
  renderManuales();
  actualizarResumen();

  $("agregar-vehiculo").addEventListener("click", () => { S.vehiculos.push({}); renderVehiculos(); guardar(); });
  $("archivo").addEventListener("change", (e) => cargarArchivo(e.target.files[0]));
  const zona = $("zona-subir");
  zona.addEventListener("dragover", (e) => { e.preventDefault(); zona.classList.add("arrastre"); });
  zona.addEventListener("dragleave", () => zona.classList.remove("arrastre"));
  zona.addEventListener("drop", (e) => { e.preventDefault(); zona.classList.remove("arrastre"); cargarArchivo(e.dataTransfer.files[0]); });
  $("buscar").addEventListener("input", renderScouts);
  $("todos").addEventListener("click", () => { S.seleccion = S.scouts.filter((s) => s.vigente).map((s) => s.id); renderScouts(); guardar(); actualizarResumen(); });
  $("ninguno").addEventListener("click", () => { S.seleccion = []; renderScouts(); guardar(); actualizarResumen(); });

  $("form-manual").addEventListener("submit", (e) => {
    e.preventDefault();
    const f = new FormData(e.target);
    const nombre = (f.get("nombre") || "").toString().trim().replace(/\s+/g, " ");
    if (!nombre) return;
    S.manuales.push({ nombre, cum: (f.get("cum") || "").toString().trim(), cat: f.get("cat"), nota: (f.get("nota") || "").toString().trim() });
    const cat = e.target.elements.cat.value;
    e.target.reset(); e.target.elements.cat.value = cat; e.target.elements.nombre.focus();
    renderManuales(); guardar(); actualizarResumen();
  });

  $("descargar").addEventListener("click", descargar);
  const prueba = new File([""], "x.docx", { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
  if (navigator.canShare && navigator.canShare({ files: [prueba] })) {
    $("compartir").hidden = false;
    $("compartir").addEventListener("click", compartir);
  }

  $("borrar-todo").addEventListener("click", async () => {
    if (!confirm("¿Borrar todos los datos capturados, la lista de Scouts y las fotos adjuntas de este navegador?")) return;
    clearTimeout(tGuardar);
    try { localStorage.removeItem(CLAVE_GUARDADO); } catch { /* nada */ }
    await borrarTodosLosAdjuntos();
    location.reload();
  });
}

iniciar();
