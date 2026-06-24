"""Analiza XML de tablas en xlsx marketing."""
import re
import sys
import zipfile
from pathlib import Path

from openpyxl.utils import range_boundaries


def analizar(path: Path) -> None:
    print("====", path.name, "====")
    with zipfile.ZipFile(path) as z:
        for t in sorted(n for n in z.namelist() if n.startswith("xl/tables/")):
            xml = z.read(t).decode("utf-8", errors="replace")
            ref_m = re.search(r'\bref="([^"]+)"', xml)
            cnt_m = re.search(r'tableColumns count="(\d+)"', xml)
            names = re.findall(r'tableColumn id="\d+" name="([^"]+)"', xml)
            af_m = re.search(r'autoFilter ref="([^"]+)"', xml)
            ref = ref_m.group(1) if ref_m else "?"
            af = af_m.group(1) if af_m else "none"
            span = 0
            if ref_m:
                min_col, _, max_col, _ = range_boundaries(ref)
                span = max_col - min_col + 1
            print(
                t,
                "ref=",
                ref,
                "span=",
                span,
                "declared=",
                cnt_m.group(1) if cnt_m else "?",
                "names=",
                len(names),
                "autoFilter=",
                af,
                "OK" if ref == af else "AF_MISMATCH",
            )
            if span and len(names) and span != len(names):
                print("  !! column span != name count")
            ids = re.findall(r'tableColumn id="(\d+)"', xml)
            expected = [str(i) for i in range(1, len(ids) + 1)]
            if ids != expected:
                print("  !! bad column ids:", ids)
            if "calculatedColumnFormula" in xml:
                print("  !! has calculated columns")
            for bad in ("#REF!", "null", "NULL"):
                if bad in xml:
                    print("  !! contains", bad)
            print("  columns:", names)


def main() -> None:
    base = Path(__file__).parent
    files = sys.argv[1:] or [
        "Prueba_OV marketing.xlsx",
        "Prueba_OV marketing - Preparado-NUEVO.xlsx",
    ]
    for name in files:
        p = base / name
        if p.exists():
            analizar(p)
        else:
            print("missing", name)


if __name__ == "__main__":
    main()
