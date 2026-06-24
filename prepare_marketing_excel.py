"""Prepara marketing: repara tablas y agrega hojas Dynamics."""
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.table import Table, TableStyleInfo

from reparar_marketing_xlsx import SALIDA as REPARADO, ORIGEN, reparar_xlsx

BASE_DIR = Path(__file__).parent
SALIDA = BASE_DIR / "Prueba_OV marketing - Preparado.xlsx"

PANEL_FILAS = [
    ("U2", "Orden de venta (OV)", "V2", ""),
    ("U3", "Cuenta Dynamics", "V3", "C0010"),
    ("U4", "Nombre del cliente (opcional)", "V4", ""),
    ("U5", "Descripcion pedido", "V5", ""),
    ("U6", "Referencia (opcional)", "V6", ""),
    ("U7", "Fecha envio (opcional)", "V7", ""),
    ("U8", "Fecha recepcion (opcional)", "V8", ""),
]

PANEL_CAMPOS = {
    "ov": "V2",
    "cuenta": "V3",
    "nombre": "V4",
    "descripcion": "V5",
    "referencia": "V6",
    "fecha_envio": "V7",
    "fecha_recepcion": "V8",
}

RESULTADO_HEADERS = [
    "Estado",
    "Pedido Dynamics",
    "Fecha de ejecución",
    "Usuario",
    "Error",
]

HISTORIAL_HEADERS = [
    "FechaHora",
    "OrdenVenta",
    "ReferenciaCliente",
    "Cliente",
    "DescripcionPedido",
    "Codigo_Articulo",
    "Cantidad",
    "PrecioUnitario",
    "Porcentaje de descuento",
    "Fecha de envio",
    "Comentario",
]

INSTRUCCIONES = [
    ("A1", "SUBIR ORDENES DE VENTA A DYNAMICS (Marketing)"),
    ("A3", "1. Arranque dynamics-integration (run.ps1) y ngrok http 8080"),
    ("A4", "2. Pegue la URL https de ngrok en Resultado!B1"),
    ("A5", "3. Botones: ProbarConexion, CrearCabeceraMarketing, CrearLineasMarketing"),
    ("A7", "PANEL DYNAMICS (U=etiqueta, V=dato, filas 2-8):"),
    ("A8", "- V3 cuenta | V5 descripcion | V2 OV automatica"),
    ("A9", "- Agregue filas solo dentro de Tabla3 (Costco) o Tabla4 (E-commerce)"),
    ("A12", "FLUJO: Probar conexion -> Crear orden -> Crear lineas"),
]


def quitar_hoja(wb, nombre: str) -> None:
    if nombre in wb.sheetnames:
        del wb[nombre]


def crear_resultado(wb) -> None:
    quitar_hoja(wb, "Resultado")
    ws = wb.create_sheet("Resultado", 0)
    ws["A1"] = "URL API (ngrok)"
    ws["B1"] = ""
    for col, h in enumerate(RESULTADO_HEADERS, 1):
        ws.cell(3, col, h)
    ws.cell(4, 1, "")
    t = Table(displayName="tblResultados", ref="A3:E4")
    t.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2")
    ws.add_table(t)


def crear_historial(wb) -> None:
    quitar_hoja(wb, "Historial")
    ws = wb.create_sheet("Historial")
    for col, h in enumerate(HISTORIAL_HEADERS, 1):
        ws.cell(1, col, h)
        ws.cell(2, col, "")
    lc = get_column_letter(len(HISTORIAL_HEADERS))
    t = Table(displayName="tblHistorial", ref=f"A1:{lc}2")
    t.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2")
    ws.add_table(t)


def crear_instrucciones(wb) -> None:
    quitar_hoja(wb, "Instrucciones Dynamics")
    ws = wb.create_sheet("Instrucciones Dynamics", 0)
    for addr, txt in INSTRUCCIONES:
        ws[addr] = txt
    ws.column_dimensions["A"].width = 95


def celda_abs(hoja: str, celda: str) -> str:
    col = "".join(ch for ch in celda if ch.isalpha())
    row = "".join(ch for ch in celda if ch.isdigit())
    return f"'{hoja}'!${col}${row}"


def panel(ws, prefijo: str) -> None:
    ws["U1"] = "PANEL DYNAMICS"
    ws["U1"].font = Font(bold=True)
    for u, etiq, v, val in PANEL_FILAS:
        ws[u] = etiq
        if val and ws[v].value in (None, ""):
            ws[v] = val
    wb = ws.parent
    for campo, celda in PANEL_CAMPOS.items():
        nom = f"{prefijo}_{campo}"
        if nom in wb.defined_names:
            del wb.defined_names[nom]
        wb.defined_names.add(DefinedName(nom, attr_text=celda_abs(ws.title, celda)))


def main() -> None:
    origen = ORIGEN if ORIGEN.exists() else SALIDA
    destino = SALIDA
    reparar_xlsx(origen, REPARADO)

    wb = load_workbook(REPARADO)
    crear_instrucciones(wb)
    crear_resultado(wb)
    crear_historial(wb)
    if "E-commerce" in wb.sheetnames:
        panel(wb["E-commerce"], "dyn_mkt_ecom")
    if "Costco" in wb.sheetnames:
        panel(wb["Costco"], "dyn_mkt_costco")

    try:
        wb.save(destino)
    except PermissionError:
        destino = BASE_DIR / "Prueba_OV marketing - Preparado-NUEVO.xlsx"
        wb.save(destino)
        print("Cierre Excel y vuelva a ejecutar para guardar en - Preparado.xlsx")

    print("Listo:", destino)
    print("Base reparada:", REPARADO)


if __name__ == "__main__":
    main()
