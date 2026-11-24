# Protocolo de transporte ASMAC

Web para llenar el formato oficial **“Información para integrar el protocolo de
seguridad en el traslado”** (Protocolo de transporte eventos 2026 – CNAR) de la
Asociación de Scouts de México, y descargarlo listo en Word (.docx).

**Todo corre en el dispositivo de quien la usa.** No hay servidor: el Excel se
lee y el documento se arma en el navegador. La página declara una política de
seguridad (`connect-src 'none'`) que le impide conectarse a cualquier sitio, así
que aunque quisiera no podría enviar datos. El avance se guarda únicamente en
el navegador (`localStorage`; las fotos adjuntas en `IndexedDB`) y se borra
con el botón **Borrar todo**.

## Cómo se usa

1. **Datos del evento**: nombre, lugar, provincia, grupo, coordinador de
   seguridad (no asiste), enlace de seguridad (sí asiste), salida y regreso.
2. **Medio de transporte**: particular (uno o varios vehículos), contratado,
   público (autobús) y/o aéreo. Solo se piden los datos del medio elegido.
3. **Scouts de tu sección**: desde **CarpetApp**
   (<https://app.scouts.org.mx/scouts>) descarga a todos los
   Scouts de tu sección como **.xlsx** y súbelo tal cual; la lista se llena
   sola. Marca a los que van.
   - Las personas con registro **vencido** (columna *Estatus* distinta de
     “Vigente”) se muestran aparte, **no se pueden seleccionar** y no entran en
     el documento.
   - Columnas que se reconocen: `Nombre`, `CUM`, `Edad`, `Sección`,
     `Pequeño Grupo`, `Cargo`, `Estatus` (también variantes como
     `Apellidos`, `Patrulla`, `Status`…). Se acepta también `.csv`.
   - Hay un [archivo de ejemplo](web/ejemplo/scouts_ejemplo.xlsx) con datos ficticios.
4. **Anexos (“¿Tienes en tu poder…?”)**: para el transporte contratado y cada
   vehículo particular se responde Sí/No y se **adjunta la foto o el PDF** de
   cada documento (contrato, fe mecánica, seguro, licencia, tarjeta de
   circulación, póliza). En computadora se abre el explorador de archivos; en
   celular se puede **tomar la foto con la cámara** o elegirla de la galería.
   Las imágenes (y las páginas de los PDF) se integran como **anexos al final
   del Word**. Si falta alguna foto o una respuesta, la página lo marca en rojo
   y pide confirmación antes de descargar.
5. **Scouters y personas externas**: se agregan a mano (Scouter, Dirigente,
   jóvenes de otro grupo o no scouts). El formato pide anotar a *todos* los
   pasajeros.
6. **Resumen y documento**: conteo por sección (Manada, Tropa, Comunidad, Clan,
   Scouters, Dirigentes, No scouts), avisos de lo que falta y botón
   **Descargar protocolo (Word)**. En celular aparece también **Compartir**.

Para PDF, abre el .docx en Word, Google Docs o WPS y guárdalo como PDF.

## Probar en local

```bash
py -m http.server 8766 --directory web
```

y abrir <http://localhost:8766>. Se publica con GitHub Pages mediante
`.github/workflows/pages.yml` (Settings → Pages → Source: *GitHub Actions*).

## Cómo funciona

| Archivo | Qué hace |
|---|---|
| `web/lector.js` | Lee el .xlsx (o .csv) con JSZip + DOMParser, detecta encabezados, decide vigente/vencido y la sección del roster (ML/TS/CC/CR). |
| `web/protocolo.js` | Llena la plantilla oficial: abre el .docx, edita `word/document.xml` y lo reempaqueta. Es el mismo llenado que hace *scoutdb* con python-docx (celdas por columna de cuadrícula, bloque de vehículo clonado por cada auto, una fila del roster por pasajero), más los apartados de transporte contratado, público y aéreo. |
| `web/adjuntos.js` | Fotos y PDF de los anexos: los convierte a JPEG (PDF con pdf.js), los guarda en IndexedDB. |
| `web/app.js` | Interfaz, guardado local y descarga. |
| `web/plantilla/…docx` | Formato oficial en blanco (sin modificar). |
| `web/plantilla.js` | La misma plantilla en base64, para que funcione sin red. Se regenera con `py tools/embed_plantilla.py` si cambia el .docx. |
| `web/vendor/jszip.min.js` | [JSZip](https://github.com/Stuk/jszip) 3.10.1 (MIT). |
| `web/vendor/pdfjs/` | [pdf.js](https://github.com/mozilla/pdf.js) 4.10.38 (Apache 2.0), se carga solo al adjuntar un PDF. |

Si ASMAC publica una versión nueva del formato, reemplaza el archivo en
`web/plantilla/`, ejecuta `tools/embed_plantilla.py` y revisa los índices de
filas/columnas en `protocolo.js` (el código avisa si la estructura cambió).
