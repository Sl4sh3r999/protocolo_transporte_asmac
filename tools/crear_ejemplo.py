"""Genera web/ejemplo/scouts_ejemplo.xlsx con datos ficticios y el mismo
formato de columnas que la exportación de Scouts de CarpetApp (app.scouts.org.mx/scouts)."""
from pathlib import Path
import openpyxl

FILAS = [
    ("Cosme Fulanito Martes", "AGS0000000", 11, "Tropa", "TIGRES", "Guía", "Vigente"),
    ("Ana Pérez López", "AGS0000001", 12, "Tropa", "TIGRES", "Subguía", "Vigente"),
    ("Luis Ramírez Soto", "AGS0000002", 13, "Tropa", "TIGRES", "Scout", "Vencido"),
    ("María Hernández Ruiz", "AGS0000003", 11, "Tropa", "HALCONES", "Guía", "Vigente"),
    ("Jorge Díaz Castro", "AGS0000004", 12, "Tropa", "HALCONES", "Scout", "Vigente"),
    ("Sofía Morales Vega", "AGS0000005", 14, "Tropa", "HALCONES", "Scout", "Vencido"),
    ("Diego Torres Luna", "AGS0000006", 13, "Tropa", "LOBOS", "Scout", "Vigente"),
]

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Scouts"
ws.append(["Nombre", "CUM", "Edad", "Sección", "Pequeño Grupo", "Cargo", "Estatus"])
for f in FILAS:
    ws.append(f)
destino = Path(__file__).resolve().parent.parent / "web" / "ejemplo" / "scouts_ejemplo.xlsx"
wb.save(destino)
print("Escrito", destino)
