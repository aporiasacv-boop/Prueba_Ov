"""Prepara el Excel de marketing para Dynamics sin romper tablas existentes."""
from copy import copy
from pathlib import Path
import shutil

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter, range_boundaries
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.table import Table, TableColumn, TableStyleInfo

BASE_DIR = Path(__file__).parent
ORIGEN_LIMPIO = BASE_DIR / "Prueba_OV marketing.xlsx"
ORIGEN_ALTERNATIVO = BASE_DIR / "Prueba_OV marketing - Reparado.xlsx"
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
    ("A7", "PANEL DYNAMICS (columnas U y V, filas 2-8 en Costco / E-commerce):"),
    ("A8", "- Etiquetas en columna U | Sus datos en columna V (al lado)"),
    ("A9", "- V2 = OV (se llena sola) | V3 = Cuenta (C0010) | V5 = Descripcion"),
    ("A10", "- Agregue filas solo dentro de la TABLA de productos (Tabla3 / Tabla4)"),
    ("A12", "FLUJO: 1) Llene V3 y V5  2) Crear orden  3) Crear lineas"),
    ("A13", "Lineas: filas con Dynamics + Cantidad y sin OV en la tabla."),
]


def resolver_origen() -> Path:
    if ORIGEN_ALTERNATIVO.exists():
        return ORIGEN_ALTERNATIVO
    if ORIGEN_LIMPIO.exists():
        return ORIGEN_LIMPIO
    if SALIDA.exists():
        return SALIDA
    raise FileNotFoundError(
        "No se encontro Prueba_OV marketing.xlsx ni - Reparado.xlsx para preparar."
    )


def quitar_hoja_si_existe(wb, nombre: str) -> None:
    if nombre in wb.sheetnames:
        del wb[nombre]


def obtener_tabla(ws, nombre_tabla: str) -> Table | None:
    tabla = ws.tables.get(nombre_tabla)
    if tabla is None or isinstance(tabla, str):
        return None
    return tabla


def sincronizar_autofiltro_tabla(tabla: Table) -> bool:
    if tabla.autoFilter is None:
        return False
    if tabla.autoFilter.ref == tabla.ref:
        return False
    tabla.autoFilter.ref = tabla.ref
    return True


def asegurar_fila_datos_tabla(ws, tabla: Table) -> bool:
    min_col, min_row, max_col, max_row = range_boundaries(tabla.ref)
    if max_row > min_row:
        return False
    max_row = min_row + 1
    col_letra = get_column_letter(max_col)
    tabla.ref = f"{get_column_letter(min_col)}{min_row}:{col_letra}{max_row}"
    for c in range(min_col, max_col + 1):
        if ws.cell(max_row, c).value is None:
            ws.cell(max_row, c, "")
    sincronizar_autofiltro_tabla(tabla)
    return True


def reparar_tablas_workbook(wb) -> list[str]:
    cambios: list[str] = []
    for hoja in wb.sheetnames:
        ws = wb[hoja]
        for nombre in list(ws.tables.keys()):
            tabla = obtener_tabla(ws, nombre)
            if tabla is None:
                continue
            if sincronizar_autofiltro_tabla(tabla):
                cambios.append(f"{hoja}/{nombre}: autoFilter -> {tabla.ref}")
            if asegurar_fila_datos_tabla(ws, tabla):
                cambios.append(f"{hoja}/{nombre}: fila de datos vacia agregada")
    return cambios


def crear_hoja_resultado(wb) -> None:
    quitar_hoja_si_existe(wb, "Resultado")
    ws = wb.create_sheet("Resultado", 0)
    ws["A1"] = "URL API (ngrok)"
    ws["B1"] = ""
    ws["A2"] = "Resultado ultima ejecucion"
    for col, header in enumerate(RESULTADO_HEADERS, start=1):
        ws.cell(3, col, header)
    ws.cell(4, 1, "")

    tabla = Table(displayName="tblResultados", ref="A3:E4")
    tabla.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    ws.add_table(tabla)
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 48
    ws.column_dimensions["E"].width = 50


def crear_hoja_historial(wb) -> None:
    quitar_hoja_si_existe(wb, "Historial")
    ws = wb.create_sheet("Historial")
    for col, header in enumerate(HISTORIAL_HEADERS, start=1):
        ws.cell(1, col, header)
    for col in range(1, len(HISTORIAL_HEADERS) + 1):
        ws.cell(2, col, "")

    last_col = get_column_letter(len(HISTORIAL_HEADERS))
    tabla = Table(displayName="tblHistorial", ref=f"A1:{last_col}2")
    tabla.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
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
        ref = celda_absoluta(hoja, celda)
        if nombre in wb.defined_names:
            del wb.defined_names[nombre]
        wb.defined_names.add(DefinedName(nombre, attr_text=ref))


def configurar_panel_dynamics(ws, prefijo: str) -> None:
    ws["U1"] = "PANEL DYNAMICS"
    ws["U1"].font = Font(bold=True)
    for celda_u, etiqueta, celda_v, valor_defecto in PANEL_FILAS:
        ws[celda_u] = etiqueta
        if valor_defecto and ws[celda_v].value in (None, ""):
            ws[celda_v] = valor_defecto
    definir_rangos_panel(ws.parent, ws.title, prefijo)


def preparar_hojas_captura(wb) -> None:
    if "E-commerce" in wb.sheetnames:
        configurar_panel_dynamics(wb["E-commerce"], "dyn_mkt_ecom")
    if "Costco" in wb.sheetnames:
        configurar_panel_dynamics(wb["Costco"], "dyn_mkt_costco")


def main() -> None:
    origen = resolver_origen()
    destino = SALIDA

    if destino.exists() and origen.resolve() != destino.resolve():
        try:
            shutil.copy2(origen, destino)
        except PermissionError:
            destino = BASE_DIR / "Prueba_OV marketing - Preparado-NUEVO.xlsx"
            shutil.copy2(origen, destino)
    elif not destino.exists():
        try:
            shutil.copy2(origen, destino)
        except PermissionError:
            destino = BASE_DIR / "Prueba_OV marketing - Preparado-NUEVO.xlsx"
            shutil.copy2(origen, destino)

    wb = load_workbook(destino)
    reparaciones = reparar_tablas_workbook(wb)
    crear_hoja_instrucciones(wb)
    crear_hoja_resultado(wb)
    crear_hoja_historial(wb)
    preparar_hojas_captura(wb)
    reparaciones.extend(reparar_tablas_workbook(wb))

    try:
        wb.save(destino)
    except PermissionError:
        destino = BASE_DIR / "Prueba_OV marketing - Preparado-NUEVO.xlsx"
        wb.save(destino)
        print("AVISO: cierre el Excel abierto y vuelva a ejecutar para sobrescribir - Preparado.xlsx")

    print("Listo:", destino)
    print("Origen:", origen)
    if reparaciones:
        print("Reparaciones de tablas:")
        for linea in reparaciones:
            print(" -", linea)
    else:
        print("Tablas OK (sin reparaciones).")


if __name__ == "__main__":
    main()
