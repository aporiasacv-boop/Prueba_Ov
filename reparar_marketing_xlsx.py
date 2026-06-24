"""Repara xlsx marketing corrupto editando solo el XML interno (sin openpyxl)."""
from __future__ import annotations

import re
import shutil
import sys
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

from openpyxl.utils import get_column_letter, range_boundaries

BASE_DIR = Path(__file__).parent
ORIGEN = BASE_DIR / "Prueba_OV marketing.xlsx"
SALIDA = BASE_DIR / "Prueba_OV marketing - REPARADO.xlsx"

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
Q = f"{{{NS}}}"


def qfind(el: ET.Element, tag: str) -> ET.Element | None:
    return el.find(Q + tag)


def reparar_table_xml(xml: str, quitar_nombres: set[str]) -> tuple[str, list[str]]:
    cambios: list[str] = []
    root = ET.fromstring(xml)
    nombre = root.attrib.get("name", "?")
    cols_el = qfind(root, "tableColumns")
    if cols_el is None:
        return xml, cambios

    hijos = list(cols_el)
    kept: list[ET.Element] = []
    for ch in hijos:
        col_name = ch.attrib.get("name", "")
        if col_name in quitar_nombres:
            cols_el.remove(ch)
            cambios.append(f"{nombre}: quitar columna '{col_name}'")
        else:
            kept.append(ch)

    for idx, ch in enumerate(kept, start=1):
        if ch.attrib.get("id") != str(idx):
            ch.set("id", str(idx))

    cols_el.set("count", str(len(kept)))

    ref = root.attrib.get("ref", "")
    if ref and kept:
        min_col, min_row, _, max_row = range_boundaries(ref)
        new_max_col = min_col + len(kept) - 1
        new_ref = (
            f"{get_column_letter(min_col)}{min_row}:"
            f"{get_column_letter(new_max_col)}{max_row}"
        )
        if new_ref != ref:
            root.set("ref", new_ref)
            cambios.append(f"{nombre}: ref {ref} -> {new_ref}")
        ref = new_ref

    af = qfind(root, "autoFilter")
    if af is not None and ref:
        if af.attrib.get("ref") != ref:
            cambios.append(
                f"{nombre}: autoFilter {af.attrib.get('ref')} -> {ref}"
            )
            af.set("ref", ref)

    ids = [ch.attrib.get("id") for ch in kept]
    esperado = [str(i) for i in range(1, len(kept) + 1)]
    if ids != esperado:
        cambios.append(f"{nombre}: reordenar ids de columnas")

    xml_out = ET.tostring(root, encoding="unicode")
    return xml_out, cambios


def reparar_xlsx(origen: Path, destino: Path) -> list[str]:
    shutil.copy2(origen, destino)
    cambios: list[str] = []
    quitar = {"Cuenta Dynamics", "Descripción pedido", "Descripcion pedido"}

    with zipfile.ZipFile(destino, "r") as zin:
        entries = list(zin.infolist())
        contenido = {info.filename: zin.read(info.filename) for info in entries}

    for nombre_archivo in ("xl/tables/table2.xml", "xl/tables/table3.xml"):
        if nombre_archivo not in contenido:
            continue
        xml = contenido[nombre_archivo].decode("utf-8")
        nuevo, delta = reparar_table_xml(xml, quitar)
        if nuevo != xml:
            contenido[nombre_archivo] = nuevo.encode("utf-8")
            cambios.extend(delta)

    with zipfile.ZipFile(destino, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for info in entries:
            data = contenido[info.filename]
            nuevo_info = zipfile.ZipInfo(info.filename)
            nuevo_info.compress_type = zipfile.ZIP_DEFLATED
            zout.writestr(nuevo_info, data)

    return cambios


def main() -> None:
    origen = Path(sys.argv[1]) if len(sys.argv) > 1 else ORIGEN
    destino = Path(sys.argv[2]) if len(sys.argv) > 2 else SALIDA
    if not origen.exists():
        raise FileNotFoundError(origen)

    cambios = reparar_xlsx(origen, destino)
    print("Listo:", destino)
    if cambios:
        print("Cambios:")
        for linea in cambios:
            print(" -", linea)
    else:
        print("Sin cambios en tablas.")


if __name__ == "__main__":
    main()
