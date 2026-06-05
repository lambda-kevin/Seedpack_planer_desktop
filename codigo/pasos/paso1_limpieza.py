import sys
import pandas as pd
import os

_DIR  = os.path.dirname(os.path.abspath(__file__))
BASE  = (os.path.dirname(sys.executable) if getattr(sys, 'frozen', False)
         else os.path.join(os.path.dirname(os.path.dirname(_DIR)), "app")) + "/"
VENTAS = BASE + "Archivos Historicos/ventasXproducto.xlsx"
SALIDA = BASE + "datos_pipeline/paso1_limpieza.xlsx"

os.makedirs(os.path.dirname(SALIDA), exist_ok=True)


def _aplicar_mapeos_usuario(df, clave):
    """Renombra columnas segun el mapeo guardado por el usuario en column_mappings.json."""
    import json as _json
    _path = os.path.normpath(BASE + "column_mappings.json")
    if not os.path.isfile(_path):
        return df
    try:
        with open(_path, encoding="utf-8") as _f:
            _m = _json.load(_f).get(clave, {})
        _rename = {v: k for k, v in _m.items()
                   if v and str(v) in df.columns and str(v) != str(k)}
        return df.rename(columns=_rename) if _rename else df
    except Exception:
        return df

# Algunos exports del ERP guardan números con coma como separador de miles
# (ej: '-2,424.000'). openpyxl no puede parsearlos; este parche lo resuelve.
try:
    import openpyxl.worksheet._reader as _wr
    _orig_cast = _wr._cast_number
    def _safe_cast(value):
        try:
            return _orig_cast(value)
        except (ValueError, TypeError):
            return float(str(value).replace(',', ''))
    _wr._cast_number = _safe_cast
except Exception:
    pass


def run(arch_ventas=None):
    src = arch_ventas if arch_ventas else VENTAS
    df = pd.read_excel(src, sheet_name="Ventas por producto")
    df = _aplicar_mapeos_usuario(df, "arch_ventas")
    df.columns = [str(c).strip() for c in df.columns]

    df = df.rename(columns={
        "Fecha Fra.": "fecha",
        "Factura":    "factura",
        "Producto Terminado": "producto",
        "Producto":   "producto",
        "Cantidad":   "cantidad",
    })

    df["fecha"] = pd.to_datetime(df["fecha"], errors="coerce")
    df = df.dropna(subset=["fecha", "producto"])
    df = df[["producto", "fecha", "cantidad", "factura"]]
    df = df.sort_values("fecha").reset_index(drop=True)

    df.to_excel(SALIDA, index=False)

    print(f"Filas: {len(df)}")
    print(f"Periodo: {df['fecha'].min().date()} - {df['fecha'].max().date()}")
    print(f"Guardado en: {SALIDA}")


if __name__ == "__main__":
    run()
