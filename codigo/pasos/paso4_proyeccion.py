"""
PASO 4 - Proyeccion de ventas por producto
==========================================
Uso:
    from paso4_proyeccion import proyectar

    proyectar(
        FI  = "2026-06-01",   # primer dia del mes actual (o fecha personalizada)
        FF  = "2026-12-31",
        MOD = "Random Forest",       # "Random Forest" | "XGBoost" | "Reg. Regularizada" | "Red Neuronal"
        REF = "NOMBRE PRODUCTO",     # opcional — si se omite proyecta todos los frecuentes
    )

Salida:
    resultados/proyeccion_<FI>_<FF>.xlsx
"""

import os, sys
import pandas as pd
import numpy as np

from paso3_modelos import (
    cargar_datos,
    agregar_mensual,
    limpiar_datos,
    crear_features,
    preparar_train_test,
    modelo_random_forest,
    modelo_xgboost,
    modelo_regresion_regularizada,
    modelo_red_neuronal,
)

_DIR        = os.path.dirname(os.path.abspath(__file__))
_DATA       = (os.path.dirname(sys.executable) if getattr(sys, 'frozen', False)
               else os.path.join(os.path.dirname(os.path.dirname(_DIR)), "app"))
SALIDA_BASE = _DATA + "/datos_pipeline/"

MODELOS_DISPONIBLES = {
    "Random Forest":      modelo_random_forest,
    "XGBoost":            modelo_xgboost,
    "Reg. Regularizada":  modelo_regresion_regularizada,
    "Red Neuronal":       modelo_red_neuronal,
}


def _construir_meses(FI, FF):
    """Genera lista de (año, mes) entre FI y FF inclusive."""
    inicio = pd.Timestamp(FI).replace(day=1)
    fin    = pd.Timestamp(FF).replace(day=1)
    periodos = []
    actual = inicio
    while actual <= fin:
        periodos.append((actual.year, actual.month))
        actual += pd.DateOffset(months=1)
    return periodos


def _predecir(modelo, features, feature_cols):
    """Aplica scaler si el modelo lo tiene y retorna la prediccion."""
    X = np.array([features[c] for c in feature_cols], dtype=float).reshape(1, -1)
    if hasattr(modelo, "_scaler"):
        X = modelo._scaler.transform(X)
        return float(np.maximum(modelo.predict(X), 0)[0])
    return float(np.maximum(modelo.predict(X), 0)[0])


def _rolling_forecast(modelo, df_features, feature_cols, periodos, target="cantidad"):
    """
    Genera predicciones para cada (año, mes) en periodos usando rolling forecast.
    Cada prediccion alimenta los lags del siguiente mes.
    """
    MESES_NOMBRES = {
        1:"Ene", 2:"Feb", 3:"Mar", 4:"Abr", 5:"May", 6:"Jun",
        7:"Jul", 8:"Ago", 9:"Sep", 10:"Oct", 11:"Nov", 12:"Dic"
    }

    df_sorted = df_features.sort_values(["producto", "año", "mes"])
    ultimo    = df_sorted.groupby("producto", as_index=False).last()

    hist_lookup = {
        (r["producto"], int(r["año"]) * 12 + int(r["mes"])): float(r[target])
        for _, r in df_sorted.iterrows()
    }
    historico_6 = {
        prod: list(grp[target].tail(6))
        for prod, grp in df_sorted.groupby("producto")
    }

    resultados = []
    n_productos = len(ultimo)
    print(f"[proyectar] Generando predicciones para {n_productos} productos...")

    for i, (_, row) in enumerate(ultimo.iterrows()):
        if (i + 1) % 25 == 0 or (i + 1) == n_productos:
            print(f"  {i + 1}/{n_productos} productos procesados...")
        producto = row["producto"]
        ventana  = list(historico_6.get(producto, [row[target]]))

        lag1     = float(row.get(f"{target}_lag1",  row[target]) or row[target])
        lag2     = float(row.get(f"{target}_lag2",  row[target]) or row[target])
        lag3     = float(row.get(f"{target}_lag3",  row[target]) or row[target])
        lag12_fb = float(row.get(f"{target}_lag12", row[target]) or row[target])

        for año, mes in periodos:
            periodo_actual = año * 12 + mes
            lag12 = hist_lookup.get((producto, periodo_actual - 12), lag12_fb)

            roll3 = float(np.mean(ventana[-3:])) if ventana else lag1
            roll6 = float(np.mean(ventana[-6:])) if ventana else lag1

            tendencia = (
                float(np.polyfit(np.arange(len(ventana)), ventana, 1)[0])
                if len(ventana) >= 2 else 0.0
            )
            crecimiento = (
                (ventana[-1] - ventana[-2]) / (abs(ventana[-2]) + 1)
                if len(ventana) >= 2 else 0.0
            )

            features = {
                "mes":                   mes,
                "mes_sin":               np.sin(2 * np.pi * mes / 12),
                "mes_cos":               np.cos(2 * np.pi * mes / 12),
                "trimestre":             ((mes - 1) // 3) + 1,
                "año":                   año,
                "producto_enc":          row["producto_enc"],
                f"{target}_lag1":        lag1,
                f"{target}_lag2":        lag2,
                f"{target}_lag3":        lag3,
                f"{target}_lag12":       lag12,
                f"{target}_roll3":       roll3,
                f"{target}_roll6":       roll6,
                f"{target}_crecimiento": crecimiento,
                f"{target}_tendencia":   tendencia,
            }

            pred = _predecir(modelo, features, feature_cols)

            resultados.append({
                "producto":  producto,
                "año":       año,
                "mes":       mes,
                "periodo":   f"{MESES_NOMBRES[mes]}-{año}",
                "cantidad_proyectada": round(pred),
            })

            lag3, lag2, lag1 = lag2, lag1, pred
            hist_lookup[(producto, periodo_actual)] = pred
            ventana.append(pred)
            if len(ventana) > 6:
                ventana.pop(0)

    return pd.DataFrame(resultados)


def proyectar(FI, FF, MOD="Random Forest", REF=None, target="cantidad"):
    """
    Parametros
    ----------
    FI  : str  — fecha inicio  ej. "2026-06-01"
    FF  : str  — fecha fin     ej. "2026-12-31"
    MOD : str  — modelo a usar: "Random Forest" | "XGBoost" |
                                "Reg. Regularizada" | "Red Neuronal"
    REF : str  — producto especifico (opcional). Si se omite proyecta
                 todos los productos frecuentes.

    Retorna
    -------
    DataFrame con columnas: producto, año, mes, periodo, cantidad_proyectada
    Guarda: resultados/proyeccion_<FI>_<FF>.xlsx
    """

    # ── Validar modelo ────────────────────────────────────────────────────────
    if MOD not in MODELOS_DISPONIBLES:
        raise ValueError(
            f"Modelo '{MOD}' no reconocido. "
            f"Opciones: {list(MODELOS_DISPONIBLES.keys())}"
        )

    # ── Preparar datos ────────────────────────────────────────────────────────
    print(f"\n{'='*60}")
    print(f"PROYECCION  |  {FI} → {FF}  |  Modelo: {MOD}")
    print(f"Producto   : {REF if REF else 'todos los frecuentes'}")
    print(f"{'='*60}")

    df_raw      = cargar_datos()
    df_mensual  = agregar_mensual(df_raw)
    df_limpio   = limpiar_datos(df_mensual, target=target)
    df_features = crear_features(df_limpio, target=target)

    # Filtrar por REF si se especifica
    if REF is not None:
        if REF not in df_features["producto"].values:
            raise ValueError(
                f"Producto '{REF}' no encontrado entre los productos frecuentes."
            )
        df_features = df_features[df_features["producto"] == REF].copy()
        print(f"[proyectar] Filtrando por producto: {REF}")

    # ── Entrenar modelo ───────────────────────────────────────────────────────
    X_train, X_test, y_train, y_test, _, feature_cols = preparar_train_test(
        df_features, target=target
    )

    fn_modelo = MODELOS_DISPONIBLES[MOD]
    modelo, _, _ = fn_modelo(X_train, X_test, y_train, y_test)

    # ── Generar proyeccion ────────────────────────────────────────────────────
    periodos = _construir_meses(FI, FF)
    print(f"\n[proyectar] Generando {len(periodos)} meses: "
          f"{periodos[0][1]}/{periodos[0][0]} → {periodos[-1][1]}/{periodos[-1][0]}")

    df_pred = _rolling_forecast(modelo, df_features, feature_cols, periodos, target=target)

    # ── Guardar ───────────────────────────────────────────────────────────────
    nombre   = "paso4_proyeccion.xlsx"
    ruta_out = SALIDA_BASE + nombre

    with pd.ExcelWriter(ruta_out, engine="openpyxl") as writer:
        df_pred.to_excel(writer, sheet_name="Proyeccion", index=False)

        resumen = (
            df_pred.groupby("producto")["cantidad_proyectada"]
                   .sum().reset_index()
                   .rename(columns={"cantidad_proyectada": "total_periodo"})
                   .sort_values("total_periodo", ascending=False)
        )
        resumen.to_excel(writer, sheet_name="Resumen_Producto", index=False)

    print(f"\n[proyectar] {len(df_pred):,} registros | "
          f"{df_pred['producto'].nunique()} productos")
    print(f"[proyectar] Guardado en: {ruta_out}")
    print(f"{'='*60}\n")

    return df_pred


# ══════════════════════════════════════════════════════════════════════════════
# CONFIGURACION — edita aqui los parametros y ejecuta el script
# ══════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":

    # Fecha inicio y fecha fin de la proyeccion (formato "YYYY-MM-DD")
    from datetime import datetime as _dt
    _hoy_p4 = _dt.today()
    FI = _hoy_p4.replace(day=1).strftime("%Y-%m-%d")
    FF = _dt(_hoy_p4.year, 12, 31).strftime("%Y-%m-%d")

    # Modelo a usar — elige UNA de las siguientes opciones:
    #   "Random Forest"
    #   "XGBoost"
    #   "Reg. Regularizada"
    #   "Red Neuronal"
    #   "Croston's Method"
    #   "Facebook Prophet"
    #   "SARIMA"
    MOD = "Random Forest"

    # Producto a proyectar — escribe el nombre exacto del producto
    # Si quieres proyectar TODOS los productos frecuentes pon: None
    REF = None

    # ─────────────────────────────────────────────────────────────────────────
    proyectar(FI=FI, FF=FF, MOD=MOD, REF=REF)