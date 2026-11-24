"""Incrusta la plantilla oficial .docx en web/plantilla.js (base64) para que la
web funcione sin red, incluso abriendo index.html directo desde el disco.
Ejecutar cada vez que se reemplace web/plantilla/*.docx."""
import base64
from pathlib import Path

WEB = Path(__file__).resolve().parent.parent / "web"
docx = next((WEB / "plantilla").glob("*.docx"))
b64 = base64.b64encode(docx.read_bytes()).decode("ascii")
(WEB / "plantilla.js").write_text(
    "// Generado por tools/embed_plantilla.py — no editar a mano.\n"
    f"// Fuente: plantilla/{docx.name}\n"
    f'const PLANTILLA_NOMBRE = "{docx.name}";\n'
    f'const PLANTILLA_B64 = "{b64}";\n',
    encoding="utf-8",
)
print("plantilla.js:", len(b64), "caracteres")
