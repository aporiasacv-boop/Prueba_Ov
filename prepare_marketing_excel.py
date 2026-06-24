"""Prepara marketing: primero repara XML de tablas, luego agrega hojas panel."""
from pathlib import Path
import shutil

from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter, range_boundaries
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.table import Table, TableStyleInfo

from reparar_marketing_xlsx import reparar_xlsx

BASE_DIR = Path(__file__).parent
ORIGEN = BASE_DIR / "Prueba_OV marketing.xlsx"
SALIDA = BASE_DIR / "Prueba_OV marketing - Preparado.xlsx"
TEMP = BASE_DIR / "Prueba_OV marketing - _tmp_reparado.xlsx"

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
    ("A7", "PANEL DYNAMICS (columnas U y V, filas 2-8):"),
    ("A8", "- Etiquetas en U | Datos en V. Agregue filas solo en la TABLA de productos."),
    ("A9", "- V3 = Cuenta (C0010) | V5 = Descripcion | V2 = OV (automatico)"),
    ("A12", "FLUJO: Probar conexion -> Crear orden -> Crear lineas"),
]


def quitar_hoja_si_existe(wb, nombre: str) -> None:
    if nombre in wb.sheetnames:
        del wb[nombre]


def crear_hoja_resultado(wb) -> None:
    quitar_hoja_si_existe(wb, "Resultado")
    ws = wb.create_sheet("Resultado", 0)
    ws["A1"] = "URL API (ngrok)"
    ws["B1"] = ""
    for col, header in enumerate(RESULTADO_HEADERS, start=1):
        ws.cell(3, col, header)
    ws.cell(4, 1, "")
    tabla = Table(displayName="tblResultados", ref="A3:E4")
    tabla.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2")
    ws.add_table(tabla)


def crear_hoja_historial(wb) -> None:
    quitar_hoja_si_existe(wb, "Historial")
    ws = wb.create_sheet("Historial")
    for col, header in enumerate(HISTORIAL_HEADERS, start=1):
        ws.cell(1, col, header)
        ws.cell(2, col, "")
    last_col = get_column_letter(len(HISTORIAL_HEADERS))
    tabla = Table(displayName="tblHistorial", ref=f"A1:{last_col}2")
    tabla.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2")
    ws.add_table(tabla)


def crear_hoja_instrucciones(wb) -> None:
    quitar_hoja_si_existe(wb, "Instrucciones Dynamics")
    ws = wb.create_sheet("Instrucciones Dynamics", 0)
    for addr, texto in INSTRUCCIONES:
        ws[addr] = texto
    ws.column_dimensions["A"].width = 95


def celda_absoluta(hoja: str, celda: str) -> str:
    col = "".join(ch for ch in celda if ch.isalpha())
    row = "".join(ch for ch in celda if ch.isdigit())
    return f"'{hoja}'!${col}${row}"


def definir_rangos_panel(wb, hoja: str, prefijo: str) -> None:
    for campo, celda in PANEL_CAMPOS.items():
        nombre = f"{prefijo}_{campo}"
        if nombre in wb.defined_names:
            del wb.defined_names[nombre]
        wb.defined_names.add(
            DefinedName(nombre, attr_text=celda_absoluta(hoja, celda))
        )


def configurar_panel(ws, prefijo: str) -> None:
    ws["U1"] = "PANEL DYNAMICS"
    ws["U1"].font = Font(bold=True)
    for celda_u, etiqueta, celda_v, valor in PANEL_FILAS:
        ws[celda_u] = etiqueta
        if valor and ws[celda_v].value in (None, ""):
            ws[celda_v] = valor
    definir_rangos_panel(ws.parent, ws.title, prefijo)


def main() -> None:
    origen = ORIGEN if ORIGEN.exists() else SALIDA
    destino = SALIDA

    reparar_xlsx(origen, TEMP)
    shutil.copy2(TEMP, destino)

    wb = load_workbook(destino)
    crear_hoja_instrucciones(wb)
    crear_hoja_resultado(wb)
    crear_hoja_historial(wb)
    if "E-commerce" in wb.sheetnames:
        configurar_panel(wb["E-commerce"], "dyn_mkt_ecom")
    if "Costco" in wb.sheetnames:
        configurar_panel(wb["Costco"], "dyn_mkt_costco")

    try:
        wb.save(destino)
    except PermissionError:
        alt = BASE_DIR / "Prueba_OV marketing - Preparado-NUEVO.xlsx"
        wb.save(alt)
        destino = alt
        print("AVISO: cierre Excel y vuelva a ejecutar para guardar en - Preparado.xlsx")

    if TEMP.exists():
        TEMP.unlink()
    print("Listo:", destino)


if __name__ == "__main__":
    main()
