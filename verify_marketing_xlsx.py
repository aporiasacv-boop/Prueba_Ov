"""Verifica que las tablas del Excel marketing tengan ref y autoFilter alineados."""
import sys
import zipfile
from pathlib import Path


def verificar(ruta: Path) -> int:
    errores = 0
    with zipfile.ZipFile(ruta) as z:
        for nombre in sorted(z.namelist()):
            if not nombre.startswith("xl/tables/"):
                continue
            xml = z.read(nombre).decode("utf-8", errors="replace")
            if 'ref="' not in xml or "autoFilter ref=" not in xml:
                continue
            ref = xml.split('ref="', 1)[1].split('"', 1)[0]
            af = xml.split('autoFilter ref="', 1)[1].split('"', 1)[0]
            ok = ref == af
            print(f"{nombre}: ref={ref} autoFilter={af} {'OK' if ok else 'MISMATCH'}")
            if not ok:
                errores += 1
    return errores


if __name__ == "__main__":
    path = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("Prueba_OV marketing - Preparado-NUEVO.xlsx")
    sys.exit(verificar(path))
