"""
PASO 3 - Entrenamiento Random Forest (modelo seleccionado)
==========================================================
Fuentes:
  - datos_pipeline/paso2_clusterizacion.xlsx  -> clasificacion de productos
  - datos_pipeline/paso1_limpieza.xlsx        -> historial de ventas limpio

Solo productos con cluster = 'frecuente'.
Las predicciones rolling se generan en PASO 4.

Salida:
  - datos_pipeline/paso3_modelos.xlsx  (metricas del modelo)
"""

import warnings
warnings.filterwarnings("ignore")
import os
os.environ["PYTHONWARNINGS"] = "ignore"

import pandas as pd
import numpy as np

from sklearn.ensemble import RandomForestRegressor
from sklearn.linear_model import RidgeCV, LassoCV, ElasticNetCV
from sklearn.neural_network import MLPRegressor
from sklearn.preprocessing import StandardScaler
from sklearn.metrics import mean_absolute_error, mean_squared_error, r2_score
from xgboost import XGBRegressor

# ─── Rutas ────────────────────────────────────────────────────────────────────
import sys
_DIR     = os.path.dirname(os.path.abspath(__file__))
_DATA    = (os.path.dirname(sys.executable) if getattr(sys, 'frozen', False)
            else os.path.join(os.path.dirname(os.path.dirname(_DIR)), "app"))
BASE     = _DATA + "/datos_pipeline/"
CLUSTERS = BASE + "paso2_clusterizacion.xlsx"
VENTAS   = BASE + "paso1_limpieza.xlsx"
SALIDA   = BASE + "paso3_modelos.xlsx"

os.makedirs(BASE, exist_ok=True)


# ══════════════════════════════════════════════════════════════════════════════
# BLOQUE 1 — CARGA Y PREPARACION
# ══════════════════════════════════════════════════════════════════════════════

def cargar_datos():
    clusters   = pd.read_excel(CLUSTERS)
    frecuentes = set(clusters.loc[clusters["cluster"] == "frecuente", "producto"])
    print(f"[cargar_datos] Productos frecuentes: {len(frecuentes)}")

    df = pd.read_excel(VENTAS)
    df["fecha"] = pd.to_datetime(df["fecha"])
    df = df[df["producto"].isin(frecuentes)].copy()
    print(f"[cargar_datos] {len(df):,} registros "
          f"({df['fecha'].min().date()} -> {df['fecha'].max().date()})")
    return df


def agregar_mensual(df):
    df["año"] = df["fecha"].dt.year
    df["mes"]  = df["fecha"].dt.month

    agg = (
        df.groupby(["año", "mes", "producto"], observed=True)
          .agg(cantidad=("cantidad", "sum"))
          .reset_index()
          .sort_values(["producto", "año", "mes"])
          .reset_index(drop=True)
    )
    print(f"[agregar_mensual] {len(agg):,} registros mensuales "
          f"| {agg['producto'].nunique()} productos")
    return agg


def limpiar_datos(df, target="cantidad"):
    n_antes = len(df)
    df = df[df[target] > 0].copy()
    n_ceros = n_antes - len(df)

    grp  = df.groupby("producto")[target]
    Q1   = grp.quantile(0.25).rename("Q1")
    Q3   = grp.quantile(0.75).rename("Q3")
    lims = pd.concat([Q1, Q3], axis=1).reset_index()
    lims["lim_sup"] = lims["Q3"] + 3 * (lims["Q3"] - lims["Q1"])

    df = (
        df.merge(lims[["producto", "lim_sup"]], on="producto", how="left")
          .loc[lambda d: d[target] <= d["lim_sup"]]
          .drop(columns="lim_sup")
          .reset_index(drop=True)
    )
    n_outliers = (n_antes - n_ceros) - len(df)
    print(f"[limpiar_datos] Eliminados: {n_ceros} ceros | "
          f"{n_outliers} outliers | Restantes: {len(df):,}")
    return df


def crear_features(df, target="cantidad"):
    df = df.copy()
    grp = df.groupby("producto")[target]

    df[f"{target}_lag1"]  = grp.shift(1)
    df[f"{target}_lag2"]  = grp.shift(2)
    df[f"{target}_lag3"]  = grp.shift(3)
    df[f"{target}_lag12"] = grp.shift(12)

    df[f"{target}_roll3"] = grp.transform(
        lambda x: x.shift(1).rolling(3, min_periods=1).mean()
    )
    df[f"{target}_roll6"] = grp.transform(
        lambda x: x.shift(1).rolling(6, min_periods=1).mean()
    )

    df[f"{target}_crecimiento"] = (
        (df[f"{target}_lag1"] - df[f"{target}_lag2"])
        / (df[f"{target}_lag2"].abs() + 1)
    )
    df[f"{target}_tendencia"] = grp.transform(
        lambda x: x.shift(1).rolling(6, min_periods=2).apply(
            lambda v: np.polyfit(np.arange(len(v)), v, 1)[0], raw=True
        )
    )

    df["mes_sin"]      = np.sin(2 * np.pi * df["mes"] / 12)
    df["mes_cos"]      = np.cos(2 * np.pi * df["mes"] / 12)
    df["trimestre"]    = ((df["mes"] - 1) // 3) + 1
    df["producto_enc"] = df["producto"].astype("category").cat.codes

    print(f"[crear_features] {df.shape[1]} columnas | target='{target}'")
    return df


def preparar_train_test(df, año_corte=None, target="cantidad"):
    if año_corte is None:
        año_corte = int(df["año"].max())
    FEATURE_COLS = [
        "mes", "mes_sin", "mes_cos", "trimestre", "año", "producto_enc",
        f"{target}_lag1",  f"{target}_lag2",  f"{target}_lag3",
        f"{target}_lag12", f"{target}_roll3",  f"{target}_roll6",
        f"{target}_crecimiento", f"{target}_tendencia",
    ]

    df_clean = df.dropna(subset=[target]).copy()
    antes = len(df_clean)
    df_clean = df_clean.dropna(subset=[
        f"{target}_lag1", f"{target}_lag2", f"{target}_lag3"
    ])
    print(f"[preparar_train_test] Eliminadas {antes - len(df_clean):,} filas sin historial")

    mask_lag12 = df_clean[f"{target}_lag12"].isna()
    df_clean.loc[mask_lag12, f"{target}_lag12"] = df_clean.loc[mask_lag12, f"{target}_roll6"]
    df_clean[f"{target}_crecimiento"] = df_clean[f"{target}_crecimiento"].fillna(0)
    df_clean[f"{target}_tendencia"]   = df_clean[f"{target}_tendencia"].fillna(0)

    train = df_clean[df_clean["año"] <  año_corte]
    test  = df_clean[df_clean["año"] == año_corte]

    X_train = train[FEATURE_COLS].astype(float)
    y_train = train[target].astype(float)
    X_test  = test[FEATURE_COLS].astype(float)
    y_test  = test[target].astype(float)

    print(f"[preparar_train_test] Train={len(X_train):,} | Test={len(X_test):,}")
    return X_train, X_test, y_train, y_test, test, FEATURE_COLS


# ══════════════════════════════════════════════════════════════════════════════
# BLOQUE 2 — METRICAS
# ══════════════════════════════════════════════════════════════════════════════

def calcular_metricas(y_true, y_pred, nombre_modelo, tolerancia=0.20):
    y_true = np.array(y_true, dtype=float)
    y_pred = np.maximum(np.array(y_pred, dtype=float), 0)

    mae   = mean_absolute_error(y_true, y_pred)
    rmse  = np.sqrt(mean_squared_error(y_true, y_pred))
    r2    = r2_score(y_true, y_pred)
    smape = np.mean(
        200 * np.abs(y_true - y_pred) / (np.abs(y_true) + np.abs(y_pred) + 1e-8)
    )
    accuracy = np.mean(
        np.abs(y_true - y_pred) / (np.abs(y_true) + 1) <= tolerancia
    ) * 100

    return {
        "modelo":    nombre_modelo,
        "MAE":       round(mae,      2),
        "RMSE":      round(rmse,     2),
        "R2":        round(r2,       4),
        "SMAPE%":    round(smape,    2),
        "Accuracy%": round(accuracy, 2),
    }


# ══════════════════════════════════════════════════════════════════════════════
# BLOQUE 3 — MODELOS  (los 4 estan disponibles para paso4_proyeccion)
# ══════════════════════════════════════════════════════════════════════════════

def modelo_random_forest(X_train, X_test, y_train, y_test):
    print("\n" + "-"*50)
    print("[Random Forest] Entrenando...")
    rf = RandomForestRegressor(
        n_estimators=300, max_depth=10, min_samples_leaf=3,
        random_state=42, n_jobs=1
    )
    rf.fit(X_train, y_train)
    y_pred = np.maximum(rf.predict(X_test), 0)
    m = calcular_metricas(y_test, y_pred, "Random Forest")
    print(f"  MAE={m['MAE']:>10,.1f}  RMSE={m['RMSE']:>10,.1f}  "
          f"R²={m['R2']:.4f}  SMAPE={m['SMAPE%']:.1f}%  Accuracy={m['Accuracy%']:.1f}%")
    return rf, y_pred, m


def modelo_xgboost(X_train, X_test, y_train, y_test):
    print("\n" + "-"*50)
    print("[XGBoost] Entrenando...")
    xgb = XGBRegressor(
        n_estimators=400, learning_rate=0.05, max_depth=5,
        subsample=0.8, colsample_bytree=0.8,
        reg_alpha=0.1, reg_lambda=1.0,
        random_state=42, n_jobs=1, verbosity=0, eval_metric="rmse"
    )
    xgb.fit(X_train, y_train, eval_set=[(X_test, y_test)], verbose=False)
    y_pred = np.maximum(xgb.predict(X_test), 0)
    m = calcular_metricas(y_test, y_pred, "XGBoost")
    print(f"  MAE={m['MAE']:>10,.1f}  RMSE={m['RMSE']:>10,.1f}  "
          f"R²={m['R2']:.4f}  SMAPE={m['SMAPE%']:.1f}%  Accuracy={m['Accuracy%']:.1f}%")
    return xgb, y_pred, m


def modelo_regresion_regularizada(X_train, X_test, y_train, y_test):
    print("\n" + "-"*50)
    print("[Reg. Regularizada] Entrenando (Ridge / Lasso / ElasticNet)...")
    scaler     = StandardScaler()
    X_train_sc = scaler.fit_transform(X_train)
    X_test_sc  = scaler.transform(X_test)

    candidatos = {
        "Ridge":      RidgeCV(alphas=[0.01, 0.1, 1, 10, 100, 500]),
        "Lasso":      LassoCV(cv=5, max_iter=10_000, random_state=42),
        "ElasticNet": ElasticNetCV(cv=5, max_iter=10_000, random_state=42),
    }

    mejor_r2, mejor_nombre, mejor_modelo, mejor_pred = -np.inf, None, None, None
    for nombre, modelo in candidatos.items():
        modelo.fit(X_train_sc, y_train)
        pred = np.maximum(modelo.predict(X_test_sc), 0)
        r2   = r2_score(y_test, pred)
        print(f"  {nombre:<12}: R²={r2:.4f}")
        if r2 > mejor_r2:
            mejor_r2, mejor_nombre, mejor_modelo, mejor_pred = r2, nombre, modelo, pred

    print(f"  >> Ganador: {mejor_nombre}")
    mejor_modelo._scaler = scaler
    m = calcular_metricas(y_test, mejor_pred, f"Reg. Regularizada ({mejor_nombre})")
    print(f"  MAE={m['MAE']:>10,.1f}  RMSE={m['RMSE']:>10,.1f}  "
          f"R²={m['R2']:.4f}  SMAPE={m['SMAPE%']:.1f}%  Accuracy={m['Accuracy%']:.1f}%")
    return mejor_modelo, mejor_pred, m


def modelo_red_neuronal(X_train, X_test, y_train, y_test):
    print("\n" + "-"*50)
    print("[Red Neuronal MLP] Entrenando...")
    scaler     = StandardScaler()
    X_train_sc = scaler.fit_transform(X_train)
    X_test_sc  = scaler.transform(X_test)

    mlp = MLPRegressor(
        hidden_layer_sizes=(128, 64, 32), activation="relu",
        solver="adam", alpha=0.001, learning_rate="adaptive",
        max_iter=500, random_state=42,
        early_stopping=True, validation_fraction=0.1, n_iter_no_change=20
    )
    mlp.fit(X_train_sc, y_train)
    y_pred = np.maximum(mlp.predict(X_test_sc), 0)
    mlp._scaler = scaler
    print(f"  Iteraciones: {mlp.n_iter_}")
    m = calcular_metricas(y_test, y_pred, "Red Neuronal (MLP)")
    print(f"  MAE={m['MAE']:>10,.1f}  RMSE={m['RMSE']:>10,.1f}  "
          f"R²={m['R2']:.4f}  SMAPE={m['SMAPE%']:.1f}%  Accuracy={m['Accuracy%']:.1f}%")
    return mlp, y_pred, m


# ══════════════════════════════════════════════════════════════════════════════
# BLOQUE 4 — MAIN  (solo entrena RF y guarda metricas; predicciones en paso 4)
# ══════════════════════════════════════════════════════════════════════════════

def main(target="cantidad"):
    print("=" * 65)
    print("PASO 3 — Entrenamiento Random Forest")
    print(f"         Target: {target}")
    print("=" * 65)

    df_raw      = cargar_datos()
    df_mensual  = agregar_mensual(df_raw)
    df_limpio   = limpiar_datos(df_mensual, target=target)
    df_features = crear_features(df_limpio, target=target)

    X_train, X_test, y_train, y_test, _, feature_cols = preparar_train_test(
        df_features, target=target
    )

    rf_model, rf_pred, rf_m = modelo_random_forest(X_train, X_test, y_train, y_test)

    print("\n" + "="*65)
    print(f"  MAE={rf_m['MAE']:>10,.1f}  RMSE={rf_m['RMSE']:>10,.1f}  "
          f"R²={rf_m['R2']:.4f}  SMAPE={rf_m['SMAPE%']:.1f}%  Accuracy={rf_m['Accuracy%']:.1f}%")
    print("="*65)

    print(f"\n[paso3] Guardando metricas en {SALIDA} ...")
    pd.DataFrame([rf_m]).to_excel(SALIDA, index=False, engine="openpyxl")

    print("\n" + "=" * 65)
    print(f"PASO 3 COMPLETADO — Random Forest  |  Accuracy={rf_m['Accuracy%']:.1f}%")
    print("  Las predicciones rolling se generan en PASO 4.")
    print(f"  Reporte en: {SALIDA}")
    print("=" * 65)


if __name__ == "__main__":
    main()
