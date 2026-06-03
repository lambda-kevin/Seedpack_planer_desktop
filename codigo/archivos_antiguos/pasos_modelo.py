"""
PASO 2 - Modelos ML de prediccion de ventas SeedPack
=====================================================
4 modelos comparados (cada uno como funcion independiente):
  1. Random Forest
  2. XGBoost
  3. Regresion Regularizada (Ridge / Lasso / ElasticNet - seleccion automatica)
  4. Red Neuronal (MLP - Multi Layer Perceptron)

Fuente de datos:
  ventas_consolidado.xlsx (generado por el script de consolidacion)
  Hoja: Transacciones

Predicciones a dos niveles:
  - Por producto individual
  - Total empresa (agregado mensual + por canal)

Uso standalone:
    python pasos_modelo.py

Uso modular (testing):
    from pasos_modelo import cargar_datos, modelo_random_forest, ...
"""

import warnings
warnings.filterwarnings("ignore")
import os
os.environ["PYTHONWARNINGS"] = "ignore"

import pandas as pd
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker

from sklearn.ensemble import RandomForestRegressor
from sklearn.linear_model import RidgeCV, LassoCV, ElasticNetCV
from sklearn.neural_network import MLPRegressor
from sklearn.preprocessing import StandardScaler
from sklearn.metrics import mean_absolute_error, mean_squared_error, r2_score
from xgboost import XGBRegressor


warnings.filterwarnings(
    "ignore",
    message=".*sklearn.utils.parallel.delayed.*",
    category=UserWarning
)

# ─── Rutas ────────────────────────────────────────────────────────────────────
BASE     = "C:/Users/EQUIPO/Desktop/SEEDPACK/"
ENTRADA  = BASE + "ventas_consolidado.xlsx"
GRAFICAS = BASE + "graficas/"
SALIDA   = BASE + "paso3_predicciones.xlsx"

MESES_NOMBRES = {
    1:"Ene", 2:"Feb", 3:"Mar", 4:"Abr", 5:"May", 6:"Jun",
    7:"Jul", 8:"Ago", 9:"Sep", 10:"Oct", 11:"Nov", 12:"Dic"
}

COLORES = ["#2E86AB", "#F18F01", "#C73E1D", "#44BBA4", "#A23B72"]

os.makedirs(GRAFICAS, exist_ok=True)


# ══════════════════════════════════════════════════════════════════════════════
# BLOQUE 1 — CARGA Y PREPARACION DE DATOS
# ══════════════════════════════════════════════════════════════════════════════

def cargar_datos(ruta=ENTRADA):
    """
    Carga la hoja Transacciones de ventas_consolidado.xlsx.
    Retorna: DataFrame con columnas fecha, año, mes, canal, linea,
             codigo, producto, cantidad, valor_venta.
    """
    df = pd.read_excel(ruta, sheet_name="Transacciones")
    df["fecha"] = pd.to_datetime(df["fecha"])
    print(f"[cargar_datos] {len(df):,} transacciones cargadas "
          f"({df['fecha'].min().date()} -> {df['fecha'].max().date()})")
    return df


def agregar_mensual(df):
    """
    Colapsa transacciones individuales a nivel mensual por producto+canal.
    Retorna: DataFrame con una fila por (año, mes, canal, linea, codigo, producto).
    """
    agg = (
        df.groupby(["año", "mes", "canal", "linea", "codigo", "producto"],
                   observed=True)
          .agg(
              cantidad    = ("cantidad",    "sum"),
              valor_venta = ("valor_venta", "sum"),
              num_facturas= ("valor_venta", "count"),
          )
          .reset_index()
          .sort_values(["codigo", "año", "mes"])
          .reset_index(drop=True)
    )
    print(f"[agregar_mensual] {len(agg):,} registros mensuales "
          f"| {agg['codigo'].nunique()} productos")
    return agg


def preparar_para_modelo(df_mensual):
    """
    Prepara df_mensual para el pipeline ML:

    1. Excluye canal='Plan': son proyecciones propias, no demanda real.
       Si quedan en el dataset, el test-set de 2026 mezcla datos reales
       con datos planeados y las métricas MAE/R² pierden significado.

    2. Desduplicar (codigo, año, mes): un mismo pedido puede generar
       una OP en INEDITTO (Produccion) Y una factura en ventasXproducto
       (Ventas) para el mismo mes. Si ambas entran al modelo, la cantidad
       se cuenta el doble y los lags calculados por codigo quedan cruzados.
       Regla: si existe fila de Ventas para ese mes → conservar Ventas.
              si solo existe Produccion → conservar Produccion.
              (Ventas = demanda ya facturada = señal más confiable)
    """
    n_orig = len(df_mensual)

    # 1. Excluir Plan
    df = df_mensual[df_mensual["canal"] != "Plan"].copy()
    n_plan = n_orig - len(df)

    # 2. Desduplicar: ordenar por prioridad y quedarse con la primera fila
    prioridad = {"Ventas": 0, "Produccion": 1}
    df["_prio"] = df["canal"].map(prioridad).fillna(99)
    df = (
        df.sort_values(["codigo", "año", "mes", "_prio"])
          .drop_duplicates(subset=["codigo", "año", "mes"], keep="first")
          .drop(columns="_prio")
          .reset_index(drop=True)
    )
    n_dup = n_orig - n_plan - len(df)

    print(f"[preparar_para_modelo] Excluidos: {n_plan} filas Plan  |  "
          f"{n_dup} duplicados canal  |  Restantes para ML: {len(df):,}")
    return df


def limpiar_datos(df_mensual, target="cantidad"):
    """
    1. Quita registros con valor cero o negativo en el target
       (ceros inconsistentes: meses sin venta real que distorsionan los lags).
    2. Quita outliers por producto usando IQR con factor 3x
       (conservador para no eliminar picos legitimos de temporada).

    Nota: no usa groupby().apply() para evitar comportamientos distintos
    entre versiones de pandas (2.2+ excluye la columna de agrupacion).

    Retorna: DataFrame limpio.
    """
    n_antes = len(df_mensual)

    # — Paso 1: eliminar ceros / negativos ——————————————————————————————————
    df = df_mensual[df_mensual[target] > 0].copy()
    n_ceros = n_antes - len(df)

    # — Paso 2: calcular limite superior por producto (IQR x3) ———————————————
    # Se hace con un merge para no perder ninguna columna del DataFrame.
    grp = df.groupby("codigo")[target]
    Q1 = grp.quantile(0.25).rename("Q1")
    Q3 = grp.quantile(0.75).rename("Q3")
    limites = pd.concat([Q1, Q3], axis=1).reset_index()
    limites["lim_sup"] = limites["Q3"] + 3 * (limites["Q3"] - limites["Q1"])

    df = (
        df.merge(limites[["codigo", "lim_sup"]], on="codigo", how="left")
          .loc[lambda d: d[target] <= d["lim_sup"]]
          .drop(columns="lim_sup")
          .reset_index(drop=True)
    )
    n_outliers = (n_antes - n_ceros) - len(df)

    print(f"[limpiar_datos] Eliminados: {n_ceros} ceros  |  "
          f"{n_outliers} outliers  |  Restantes: {len(df):,}")
    return df


def segmentar_productos(df_mensual, umbral_frecuencia=0.60):
    """
    Clasifica cada producto en dos grupos segun su frecuencia de aparicion:

      Recurrentes  (frecuencia >= umbral):
        Tienen comportamiento predecible → van al pipeline ML.

      Esporadicos  (frecuencia <  umbral):
        Pedidos irregulares → manejo determinístico bajo pedido.

    Frecuencia = meses con venta del producto / total meses en el dataset.
    Se usa el rango global del dataset como denominador para no penalizar
    productos recientes que venden consistentemente.

    Retorna:
      df_recurrentes   : subset de df_mensual con los productos recurrentes
      df_esporadicos   : subset de df_mensual con los productos esporadicos
      df_clasificacion : DataFrame resumen con la clasificacion de cada producto
    """
    # Total de periodos (año-mes) únicos en el dataset completo
    total_periodos = df_mensual[["año", "mes"]].drop_duplicates().shape[0]

    # Periodos con venta por producto:
    # primero eliminar duplicados (codigo, año, mes), luego contar con size().
    # No usa apply() para evitar problemas con pandas 2.2+.
    meses_por_producto = (
        df_mensual[["codigo", "año", "mes"]]
        .drop_duplicates()
        .groupby("codigo")
        .size()
        .reset_index(name="meses_con_venta")
    )
    meses_por_producto["frecuencia"]  = (
        meses_por_producto["meses_con_venta"] / total_periodos
    )
    meses_por_producto["grupo"] = np.where(
        meses_por_producto["frecuencia"] >= umbral_frecuencia,
        "PLANEADO", "BAJO_PEDIDO"
    )

    # Enriquecer con nombre de producto y linea
    meta = (df_mensual.groupby("codigo")[["producto", "linea", "canal"]]
            .first().reset_index())
    df_clasificacion = meses_por_producto.merge(meta, on="codigo", how="left")
    df_clasificacion["umbral_usado"] = umbral_frecuencia
    df_clasificacion = df_clasificacion[[
        "codigo", "producto", "linea", "canal",
        "meses_con_venta", "frecuencia", "grupo", "umbral_usado"
    ]].sort_values("frecuencia", ascending=False).reset_index(drop=True)

    codigos_rec = set(
        df_clasificacion.loc[df_clasificacion["grupo"] == "PLANEADO", "codigo"]
    )
    codigos_esp = set(
        df_clasificacion.loc[df_clasificacion["grupo"] == "BAJO_PEDIDO", "codigo"]
    )

    df_recurrentes  = df_mensual[df_mensual["codigo"].isin(codigos_rec)].copy()
    df_esporadicos  = df_mensual[df_mensual["codigo"].isin(codigos_esp)].copy()

    n_rec = len(codigos_rec)
    n_esp = len(codigos_esp)
    print(f"\n[segmentar_productos] Umbral frecuencia: {umbral_frecuencia:.0%} "
          f"| Periodos totales: {total_periodos}")
    print(f"  PLANEADOS   (van al modelo ML) : {n_rec:>4} productos")
    print(f"  BAJO PEDIDO (deterministico)   : {n_esp:>4} productos")
    print(f"  TOTAL                          : {n_rec + n_esp:>4} productos")

    return df_recurrentes, df_esporadicos, df_clasificacion


def perfil_bajo_pedido(df_esporadicos, target="cantidad"):
    """
    Calcula el perfil de demanda historica para productos esporadicos
    y genera una recomendacion de stock de seguridad.

    Stock sugerido:
      Si el producto tiene >= 2 registros:
        stock = promedio + 1.28 * desviacion_std
        (cubre el 90% de los pedidos historicos, nivel de servicio Z=1.28)
      Si tiene solo 1 registro:
        stock = cantidad_unica * 1.25  (buffer fijo 25%)

    Columnas del resultado:
      codigo, producto, linea, canal,
      n_pedidos, qty_promedio, qty_std, qty_mediana, qty_max,
      frecuencia_anual, ultimo_año, ultimo_mes,
      meses_entre_pedidos_prom, stock_sugerido, programa
    """
    if df_esporadicos.empty:
        print("[perfil_bajo_pedido] Sin productos esporadicos.")
        return pd.DataFrame()

    filas = []
    for codigo, grp in df_esporadicos.groupby("codigo"):
        grp = grp.sort_values(["año", "mes"])
        qty       = grp[target].values
        n         = len(qty)
        promedio  = float(np.mean(qty))
        std       = float(np.std(qty, ddof=1)) if n >= 2 else 0.0
        mediana   = float(np.median(qty))
        maximo    = float(np.max(qty))

        # Años cubiertos (para frecuencia anual)
        años_cubiertos = grp["año"].nunique()
        frec_anual     = round(n / max(años_cubiertos, 1), 2)

        # Meses promedio entre pedidos
        if n >= 2:
            # Crear un indice ordinal de periodo (año*12 + mes)
            periodos = grp["año"].values * 12 + grp["mes"].values
            diffs    = np.diff(periodos)
            meses_entre = round(float(np.mean(diffs)), 1)
        else:
            meses_entre = None

        # Stock sugerido
        if n >= 2:
            stock = max(round(promedio + 1.28 * std), 1)
        else:
            stock = max(round(promedio * 1.25), 1)

        ultimo = grp.iloc[-1]

        filas.append({
            "codigo":                    codigo,
            "producto":                  ultimo["producto"],
            "linea":                     ultimo.get("linea", ""),
            "canal":                     ultimo.get("canal", ""),
            "n_pedidos":                 n,
            "qty_promedio":              round(promedio, 1),
            "qty_std":                   round(std, 1),
            "qty_mediana":               round(mediana, 1),
            "qty_max":                   round(maximo, 1),
            "frecuencia_anual":          frec_anual,
            "ultimo_año":                int(ultimo["año"]),
            "ultimo_mes":                MESES_NOMBRES[int(ultimo["mes"])],
            "meses_entre_pedidos_prom":  meses_entre,
            "stock_sugerido":            stock,
            "programa":                  "BAJO PEDIDO",
        })

    df_perfil = pd.DataFrame(filas).sort_values("n_pedidos", ascending=False)
    print(f"[perfil_bajo_pedido] {len(df_perfil)} productos esporadicos procesados")
    print(f"  Stock sugerido promedio : {df_perfil['stock_sugerido'].mean():,.0f} unidades")
    print(f"  Pedidos/año promedio    : {df_perfil['frecuencia_anual'].mean():.1f}")
    return df_perfil.reset_index(drop=True)


def crear_features(df_mensual, target="cantidad"):
    """
    Agrega features de series de tiempo al DataFrame mensual:
      - Lags: t-1, t-2, t-3, t-12 (mismo mes año anterior)
      - Rolling: media movil 3m y 6m (sin data leak)
      - Crecimiento: tasa de cambio mes a mes  (lag1 - lag2) / (lag2 + 1)
      - Tendencia: pendiente lineal de los ultimos 6 meses por producto
      - Temporales: mes ciclico (sin/cos), trimestre
      - Categoricos: canal y linea codificados numericamente

    Parametro target: 'cantidad' o 'valor_venta'
    Retorna: DataFrame enriquecido con features.
    """
    df = df_mensual.copy()

    # — Lags por producto ————————————————————————————————————————————————————
    grp = df.groupby("codigo")[target]
    df[f"{target}_lag1"]  = grp.shift(1)
    df[f"{target}_lag2"]  = grp.shift(2)
    df[f"{target}_lag3"]  = grp.shift(3)
    df[f"{target}_lag12"] = grp.shift(12)

    # — Rolling mean (shift 1 para no usar el mes actual) ————————————————————
    df[f"{target}_roll3"] = grp.transform(
        lambda x: x.shift(1).rolling(3, min_periods=1).mean()
    )
    df[f"{target}_roll6"] = grp.transform(
        lambda x: x.shift(1).rolling(6, min_periods=1).mean()
    )

    # — Crecimiento: tasa de cambio mensual ——————————————————————————————————
    # (lag1 - lag2) / (|lag2| + 1)  →  positivo = crece, negativo = cae
    df[f"{target}_crecimiento"] = (
        (df[f"{target}_lag1"] - df[f"{target}_lag2"])
        / (df[f"{target}_lag2"].abs() + 1)
    )

    # — Tendencia: pendiente lineal de los ultimos 6 meses ———————————————————
    # Usa polyfit de grado 1 sobre la ventana de 6 meses previos.
    # min_periods=2 para no perder demasiados registros al inicio.
    df[f"{target}_tendencia"] = grp.transform(
        lambda x: x.shift(1).rolling(6, min_periods=2).apply(
            lambda v: np.polyfit(np.arange(len(v)), v, 1)[0],
            raw=True
        )
    )

    # — Encoding ciclico del mes (captura la estacionalidad) —————————————————
    df["mes_sin"]   = np.sin(2 * np.pi * df["mes"] / 12)
    df["mes_cos"]   = np.cos(2 * np.pi * df["mes"] / 12)
    df["trimestre"] = ((df["mes"] - 1) // 3) + 1

    # — Encoding categorico ——————————————————————————————————————————————————
    df["canal_enc"] = df["canal"].astype("category").cat.codes
    df["linea_enc"] = df["linea"].astype("category").cat.codes

    print(f"[crear_features] Features creadas para target='{target}' "
          f"| {df.shape[1]} columnas total")
    return df


def preparar_train_test(df_features, año_corte=2026, target="cantidad"):
    """
    Divide en train (años < año_corte) y test (año == año_corte).

    Politica de NaN:
      - lag1 / lag2 / lag3 NaN  →  se ELIMINAN (sin historial = sin prediccion fiable)
      - lag12 NaN               →  se rellena con roll6 del producto (mejor que 0)
      - crecimiento / tendencia →  se rellenan con 0 si aun son NaN tras los drops
        (solo ocurre en los primeros 2 meses de un producto)

    Retorna: X_train, X_test, y_train, y_test, df_test_completo, lista_features
    """
    FEATURE_COLS = [
        "mes", "mes_sin", "mes_cos", "trimestre", "año",
        "canal_enc", "linea_enc",
        f"{target}_lag1",  f"{target}_lag2",  f"{target}_lag3",
        f"{target}_lag12", f"{target}_roll3",  f"{target}_roll6",
        f"{target}_crecimiento", f"{target}_tendencia",
    ]

    df_clean = df_features.dropna(subset=[target]).copy()

    # Eliminar filas sin historial suficiente (lag1/2/3 NaN)
    antes = len(df_clean)
    df_clean = df_clean.dropna(subset=[
        f"{target}_lag1", f"{target}_lag2", f"{target}_lag3"
    ])
    print(f"[preparar_train_test] Eliminadas {antes - len(df_clean):,} filas "
          f"sin historial suficiente (lag1/2/3 NaN)")

    # lag12: rellenar con roll6 (mismo producto, mismo periodo aprox.)
    mask_lag12 = df_clean[f"{target}_lag12"].isna()
    df_clean.loc[mask_lag12, f"{target}_lag12"] = (
        df_clean.loc[mask_lag12, f"{target}_roll6"]
    )

    # crecimiento / tendencia: rellenar con 0 los pocos NaN restantes
    df_clean[f"{target}_crecimiento"] = df_clean[f"{target}_crecimiento"].fillna(0)
    df_clean[f"{target}_tendencia"]   = df_clean[f"{target}_tendencia"].fillna(0)

    train = df_clean[df_clean["año"] <  año_corte]
    test  = df_clean[df_clean["año"] == año_corte]

    X_train = train[FEATURE_COLS].astype(float)
    y_train = train[target].astype(float)
    X_test  = test[FEATURE_COLS].astype(float)
    y_test  = test[target].astype(float)

    print(f"[preparar_train_test] Train={len(X_train):,} obs "
          f"| Test={len(X_test):,} obs | target='{target}'")
    return X_train, X_test, y_train, y_test, test, FEATURE_COLS


def calcular_metricas(y_true, y_pred, nombre_modelo, tolerancia=0.20):
    """
    Calcula MAE, RMSE, R², SMAPE y Accuracy para un conjunto de predicciones.

    SMAPE (Symmetric Mean Absolute Percentage Error):
      Mas estable que MAPE porque penaliza igual sobre- y sub-estimacion
      y no explota cuando el valor real es cercano a cero.
      Formula: mean( 200 * |y - ŷ| / (|y| + |ŷ| + ε) )
      Rango: [0, 200] — menor es mejor.

    Accuracy: % de predicciones dentro del margen de tolerancia del valor real.
      tolerancia=0.20 => aceptada si |pred - real| <= 20% del real.

    Retorna diccionario con las metricas.
    """
    y_true = np.array(y_true, dtype=float)
    y_pred = np.maximum(np.array(y_pred, dtype=float), 0)

    mae   = mean_absolute_error(y_true, y_pred)
    rmse  = np.sqrt(mean_squared_error(y_true, y_pred))
    r2    = r2_score(y_true, y_pred)
    smape = np.mean(
        200 * np.abs(y_true - y_pred) / (np.abs(y_true) + np.abs(y_pred) + 1e-8)
    )
    error_rel = np.abs(y_true - y_pred) / (np.abs(y_true) + 1)
    accuracy  = np.mean(error_rel <= tolerancia) * 100

    return {
        "modelo":    nombre_modelo,
        "MAE":       round(mae,      2),
        "RMSE":      round(rmse,     2),
        "R2":        round(r2,       4),
        "SMAPE%":    round(smape,    2),
        "Accuracy%": round(accuracy, 2),
    }


# ══════════════════════════════════════════════════════════════════════════════
# BLOQUE 2 — MODELOS
# ══════════════════════════════════════════════════════════════════════════════

def modelo_random_forest(X_train, X_test, y_train, y_test,
                         n_estimators=300, max_depth=10, random_state=42):
    """
    Entrena un Random Forest Regressor.

    Parametros configurables para testing:
      n_estimators : numero de arboles (default 300)
      max_depth    : profundidad maxima (default 10)
      random_state : semilla de aleatoriedad

    Retorna: (modelo_entrenado, predicciones_test, dict_metricas)
    """
    print("\n" + "-"*50)
    print("[1/4] Random Forest — entrenando...")

    rf = RandomForestRegressor(
        n_estimators     = n_estimators,
        max_depth        = max_depth,
        min_samples_leaf = 3,
        random_state     = random_state,
        n_jobs           = 1,
    )
    rf.fit(X_train, y_train)
    y_pred = np.maximum(rf.predict(X_test), 0)

    m = calcular_metricas(y_test, y_pred, "Random Forest")
    print(f"  MAE={m['MAE']:>10,.1f}  RMSE={m['RMSE']:>10,.1f}  "
          f"R²={m['R2']:.4f}  SMAPE={m['SMAPE%']:.1f}%  Accuracy={m['Accuracy%']:.1f}%")

    importancia = pd.Series(rf.feature_importances_, index=X_train.columns)
    top3 = importancia.nlargest(3)
    print(f"  Top features: {', '.join(f'{f}({v:.3f})' for f, v in top3.items())}")

    return rf, y_pred, m


def modelo_xgboost(X_train, X_test, y_train, y_test,
                   n_estimators=400, learning_rate=0.05,
                   max_depth=5, random_state=42):
    """
    Entrena un XGBoost Regressor.

    Parametros configurables para testing:
      n_estimators  : numero de rondas de boosting (default 400)
      learning_rate : tasa de aprendizaje eta (default 0.05)
      max_depth     : profundidad de cada arbol (default 5)

    Retorna: (modelo_entrenado, predicciones_test, dict_metricas)
    """
    print("\n" + "-"*50)
    print("[2/4] XGBoost — entrenando...")

    xgb = XGBRegressor(
        n_estimators     = n_estimators,
        learning_rate    = learning_rate,
        max_depth        = max_depth,
        subsample        = 0.8,
        colsample_bytree = 0.8,
        reg_alpha        = 0.1,
        reg_lambda       = 1.0,
        random_state     = random_state,
        n_jobs           = 1,
        verbosity        = 0,
        eval_metric      = "rmse",
    )
    xgb.fit(
        X_train, y_train,
        eval_set=[(X_test, y_test)],
        verbose=False,
    )
    y_pred = np.maximum(xgb.predict(X_test), 0)

    m = calcular_metricas(y_test, y_pred, "XGBoost")
    print(f"  MAE={m['MAE']:>10,.1f}  RMSE={m['RMSE']:>10,.1f}  "
          f"R²={m['R2']:.4f}  SMAPE={m['SMAPE%']:.1f}%  Accuracy={m['Accuracy%']:.1f}%")

    importancia = pd.Series(xgb.feature_importances_, index=X_train.columns)
    top3 = importancia.nlargest(3)
    print(f"  Top features: {', '.join(f'{f}({v:.3f})' for f, v in top3.items())}")

    return xgb, y_pred, m


def modelo_regresion_regularizada(X_train, X_test, y_train, y_test):
    """
    Evalua Ridge, Lasso y ElasticNet con validacion cruzada automatica.
    Selecciona el de mejor R² en test y lo devuelve.
    El scaler queda guardado en modelo._scaler para predicciones futuras.

    Retorna: (mejor_modelo, predicciones_test, dict_metricas)
    """
    print("\n" + "-"*50)
    print("[3/4] Regresion Regularizada (Ridge / Lasso / ElasticNet)...")

    scaler = StandardScaler()
    X_train_sc = scaler.fit_transform(X_train)
    X_test_sc  = scaler.transform(X_test)

    candidatos = {
        "Ridge":      RidgeCV(alphas=[0.01, 0.1, 1, 10, 100, 500]),
        "Lasso":      LassoCV(cv=5, max_iter=10_000, random_state=42),
        "ElasticNet": ElasticNetCV(cv=5, max_iter=10_000, random_state=42),
    }

    mejor_r2     = -np.inf
    mejor_nombre = None
    mejor_modelo = None
    mejor_pred   = None

    for nombre, modelo in candidatos.items():
        modelo.fit(X_train_sc, y_train)
        pred = np.maximum(modelo.predict(X_test_sc), 0)
        r2   = r2_score(y_test, pred)
        print(f"  {nombre:<12}: R²={r2:.4f}")
        if r2 > mejor_r2:
            mejor_r2     = r2
            mejor_nombre = nombre
            mejor_modelo = modelo
            mejor_pred   = pred

    print(f"  >> Ganador: {mejor_nombre}")
    mejor_modelo._scaler = scaler

    m = calcular_metricas(y_test, mejor_pred,
                          f"Reg. Regularizada ({mejor_nombre})")
    print(f"  MAE={m['MAE']:>10,.1f}  RMSE={m['RMSE']:>10,.1f}  "
          f"R²={m['R2']:.4f}  SMAPE={m['SMAPE%']:.1f}%  Accuracy={m['Accuracy%']:.1f}%")

    return mejor_modelo, mejor_pred, m


def modelo_red_neuronal(X_train, X_test, y_train, y_test,
                        capas=(128, 64, 32), max_iter=500, random_state=42):
    """
    Entrena una Red Neuronal MLP (Multi Layer Perceptron) para regresion.

    Arquitectura: 3 capas ocultas (128 -> 64 -> 32 neuronas), activacion relu.
    Escala los datos con StandardScaler (obligatorio para redes neuronales).
    El scaler queda guardado en modelo._scaler para predicciones futuras.

    Parametros configurables:
      capas        : tupla con neuronas por capa oculta (default (128, 64, 32))
      max_iter     : iteraciones maximas de entrenamiento (default 500)
      random_state : semilla de aleatoriedad

    Retorna: (modelo_entrenado, predicciones_test, dict_metricas)
    """
    print("\n" + "-"*50)
    print("[4/4] Red Neuronal (MLP) — entrenando...")

    scaler = StandardScaler()
    X_train_sc = scaler.fit_transform(X_train)
    X_test_sc  = scaler.transform(X_test)

    mlp = MLPRegressor(
        hidden_layer_sizes  = capas,
        activation          = "relu",
        solver              = "adam",
        alpha               = 0.001,
        batch_size          = "auto",
        learning_rate       = "adaptive",
        max_iter            = max_iter,
        random_state        = random_state,
        early_stopping      = True,
        validation_fraction = 0.1,
        n_iter_no_change    = 20,
        tol                 = 1e-4,
    )
    mlp.fit(X_train_sc, y_train)
    y_pred = np.maximum(mlp.predict(X_test_sc), 0)

    mlp._scaler = scaler

    print(f"  Iteraciones realizadas: {mlp.n_iter_}")

    m = calcular_metricas(y_test, y_pred, "Red Neuronal (MLP)")
    print(f"  MAE={m['MAE']:>10,.1f}  RMSE={m['RMSE']:>10,.1f}  "
          f"R²={m['R2']:.4f}  SMAPE={m['SMAPE%']:.1f}%  Accuracy={m['Accuracy%']:.1f}%")

    return mlp, y_pred, m


def comparar_modelos(lista_metricas):
    """
    Recibe lista de dicts de metricas y muestra tabla comparativa.
    Retorna DataFrame con la comparacion y nombre del mejor modelo.
    """
    df_m = pd.DataFrame(lista_metricas)
    print("\n" + "="*80)
    print("COMPARACION DE MODELOS")
    print("="*80)
    print(df_m.to_string(index=False))

    idx_mejor_r2  = df_m["R2"].idxmax()
    idx_mejor_acc = df_m["Accuracy%"].idxmax()
    mejor         = df_m.loc[idx_mejor_r2, "modelo"]
    mejor_acc     = df_m.loc[idx_mejor_acc, "modelo"]
    print(f"\n>> Mejor modelo por R²:        {mejor}")
    print(f">> Mejor modelo por Accuracy%: {mejor_acc}")
    print("="*80)
    return df_m, mejor


# ══════════════════════════════════════════════════════════════════════════════
# BLOQUE 3 — PREDICCIONES FUTURAS
# ══════════════════════════════════════════════════════════════════════════════

def _predecir_con_modelo(modelo, X, feature_cols):
    """Wrapper interno: aplica scaler si el modelo lo tiene adjunto."""
    X_df = pd.DataFrame([X])[feature_cols].astype(float).fillna(0)
    if hasattr(modelo, "_scaler"):
        X_tr = modelo._scaler.transform(X_df)
        return float(np.maximum(modelo.predict(X_tr), 0)[0])
    return float(np.maximum(modelo.predict(X_df), 0)[0])


def predecir_por_producto(modelo, df_features, feature_cols,
                           meses_futuros, año_futuro=2026,
                           target="cantidad"):
    """
    Genera predicciones mensuales por producto usando rolling forecast.
    Cada mes predicho alimenta los lags y la ventana del siguiente mes.

    Rolling window real:
      Se mantiene una ventana de los ultimos 6 valores reales + predichos
      por producto. roll3 y roll6 se calculan desde esa ventana.
      tendencia se calcula con polyfit sobre la ventana completa.
      crecimiento se calcula como (pred_actual - pred_anterior) / (pred_anterior + 1).

    Retorna: DataFrame con columnas
      codigo, producto, canal, linea, año, mes, pred_{target}
    """
    # Ultimo registro historico por producto (punto de arranque)
    df_sorted = df_features.sort_values(["codigo", "año", "mes"])
    ultimo = df_sorted.groupby("codigo", as_index=False).last()

    # Lookup historico: (codigo, periodo_ordinal) → valor real
    # periodo_ordinal = año * 12 + mes  (entero comparable entre años)
    # Se usa para resolver lag12 correctamente mes a mes.
    # Las predicciones futuras se van agregando aqui para soportar
    # horizontes de mas de 12 meses si algun dia se extiende el forecast.
    hist_lookup = {
        (r["codigo"], int(r["año"]) * 12 + int(r["mes"])): float(r[target])
        for _, r in df_sorted.iterrows()
    }

    # Ultimos 6 valores historicos por producto (ventana inicial).
    # Usa for-loop en lugar de apply() para evitar problemas con pandas 2.2+.
    historico_6 = {}
    for codigo_h, grp_h in df_sorted.groupby("codigo"):
        historico_6[codigo_h] = list(grp_h[target].tail(6))

    resultados = []

    for _, row in ultimo.iterrows():
        codigo = row["codigo"]

        # Ventana real: lista de los ultimos 6 valores (historicos)
        ventana = list(historico_6.get(codigo, [row[target]]))

        # Estado inicial con los lags del ultimo registro historico
        lag1 = float(row.get(f"{target}_lag1", row[target]) or row[target])
        lag2 = float(row.get(f"{target}_lag2", row[target]) or row[target])
        lag3 = float(row.get(f"{target}_lag3", row[target]) or row[target])
        # lag12_fb: fallback si no hay dato historico 12 meses atras
        lag12_fb = float(row.get(f"{target}_lag12", row[target]) or row[target])

        for mes in meses_futuros:
            # — lag12: valor del mismo mes del año anterior ———————————————————
            # Se resuelve desde hist_lookup (datos reales o predichos).
            # Para Apr-Dic 2026 esto apunta a Apr-Dic 2025 (en historial).
            # Si el dataset no cubre ese mes, cae al fallback lag12_fb.
            periodo_actual = año_futuro * 12 + mes
            periodo_lag12  = periodo_actual - 12
            lag12 = hist_lookup.get((codigo, periodo_lag12), lag12_fb)

            # — roll3 y roll6 desde la ventana real ——————————————————————————
            roll3 = float(np.mean(ventana[-3:])) if len(ventana) >= 1 else lag1
            roll6 = float(np.mean(ventana[-6:])) if len(ventana) >= 1 else lag1

            # — Tendencia: pendiente lineal de la ventana ————————————————————
            if len(ventana) >= 2:
                tendencia = float(np.polyfit(np.arange(len(ventana)), ventana, 1)[0])
            else:
                tendencia = 0.0

            # — Crecimiento: tasa de cambio entre los dos ultimos ————————————
            if len(ventana) >= 2:
                crecimiento = (ventana[-1] - ventana[-2]) / (abs(ventana[-2]) + 1)
            else:
                crecimiento = 0.0

            features = {
                "mes":                      mes,
                "mes_sin":                  np.sin(2 * np.pi * mes / 12),
                "mes_cos":                  np.cos(2 * np.pi * mes / 12),
                "trimestre":                ((mes - 1) // 3) + 1,
                "año":                      año_futuro,
                "canal_enc":                row["canal_enc"],
                "linea_enc":                row["linea_enc"],
                f"{target}_lag1":           lag1,
                f"{target}_lag2":           lag2,
                f"{target}_lag3":           lag3,
                f"{target}_lag12":          lag12,
                f"{target}_roll3":          roll3,
                f"{target}_roll6":          roll6,
                f"{target}_crecimiento":    crecimiento,
                f"{target}_tendencia":      tendencia,
            }

            pred = _predecir_con_modelo(modelo, features, feature_cols)

            resultados.append({
                "codigo":         codigo,
                "producto":       row["producto"],
                "canal":          row["canal"],
                "linea":          row.get("linea", ""),
                "año":            año_futuro,
                "mes":            mes,
                f"pred_{target}": round(pred),
            })

            # — Actualizar lags para el siguiente mes ————————————————————————
            lag3 = lag2
            lag2 = lag1
            lag1 = pred

            # — Guardar prediccion en lookup (por si se extiende > 12 meses) —
            hist_lookup[(codigo, periodo_actual)] = pred

            # — Actualizar ventana real ——————————————————————————————————————
            ventana.append(pred)
            if len(ventana) > 6:
                ventana.pop(0)

    df_pred = pd.DataFrame(resultados)
    print(f"[predecir_por_producto] {len(df_pred):,} predicciones "
          f"| {df_pred['codigo'].nunique()} productos "
          f"| meses {meses_futuros[0]}-{meses_futuros[-1]} {año_futuro}")
    return df_pred


def predecir_total_empresa(df_pred_producto, target="cantidad"):
    """
    Agrega predicciones de producto a nivel empresa total por mes.

    Retorna dos DataFrames:
      total_mes  : (año, mes, pred_total, pred_por_canal...)
      por_linea  : (año, mes, linea, pred_{target})
    """
    col = f"pred_{target}"

    total_mes = (
        df_pred_producto
        .groupby(["año", "mes"])[col]
        .sum()
        .reset_index()
        .rename(columns={col: f"total_{target}"})
    )

    pivot_canal = (
        df_pred_producto
        .groupby(["año", "mes", "canal"])[col]
        .sum()
        .unstack("canal")
        .reset_index()
    )
    pivot_canal.columns.name = None
    total_mes = total_mes.merge(pivot_canal, on=["año", "mes"], how="left")

    por_linea = (
        df_pred_producto
        .groupby(["año", "mes", "linea"])[col]
        .sum()
        .reset_index()
    )

    print(f"[predecir_total_empresa] Prediccion total Abr-Dic {df_pred_producto['año'].iloc[0]}:")
    for _, r in total_mes.iterrows():
        print(f"  {MESES_NOMBRES[int(r['mes'])]}: {r[f'total_{target}']:>10,.0f}")
    print(f"  TOTAL: {total_mes[f'total_{target}'].sum():>10,.0f}")

    return total_mes, por_linea


# ══════════════════════════════════════════════════════════════════════════════
# BLOQUE 4 — GRAFICAS
# ══════════════════════════════════════════════════════════════════════════════

def graficar_comparacion_modelos(df_metricas, ruta_graficas=GRAFICAS):
    """Barras comparativas de MAE, RMSE, R² y SMAPE para los 4 modelos."""
    fig, axes = plt.subplots(1, 4, figsize=(22, 5))
    fig.suptitle("Comparacion de Modelos ML — SeedPack",
                 fontsize=14, fontweight="bold")

    for ax, col, titulo in zip(
        axes,
        ["MAE", "RMSE", "R2", "SMAPE%"],
        ["MAE (menor es mejor)", "RMSE (menor es mejor)",
         "R² (mayor es mejor)", "SMAPE% (menor es mejor)"]
    ):
        bars = ax.bar(df_metricas["modelo"], df_metricas[col],
                      color=COLORES[:len(df_metricas)], alpha=0.85, edgecolor="white")
        ax.set_title(titulo, fontsize=11)
        ax.tick_params(axis="x", rotation=20)
        for bar, val in zip(bars, df_metricas[col]):
            ax.text(bar.get_x() + bar.get_width() / 2,
                    bar.get_height() * 1.01,
                    f"{val:.3f}", ha="center", va="bottom", fontsize=9)

    plt.tight_layout()
    ruta = ruta_graficas + "7_comparacion_modelos.png"
    plt.savefig(ruta, dpi=150, bbox_inches="tight")
    plt.close()
    print("  Guardada: 7_comparacion_modelos.png")


def graficar_prediccion_total(total_mes_por_modelo, target="cantidad",
                               ruta_graficas=GRAFICAS):
    """Lineas de prediccion mensual total empresa para los 4 modelos."""
    fig, ax = plt.subplots(figsize=(14, 6))
    col = f"total_{target}"

    for (nombre, df_total), color in zip(total_mes_por_modelo.items(), COLORES):
        x = [MESES_NOMBRES[int(m)] for m in df_total["mes"]]
        ax.plot(x, df_total[col] / 1000, "o-", color=color,
                linewidth=2.2, markersize=7, label=nombre)

    ax.set_ylabel(f"{target.replace('_', ' ').title()} (miles)")
    ax.set_title(f"Prediccion Total Empresa — Abr-Dic 2026\nComparacion de los 4 modelos",
                 fontsize=13, fontweight="bold")
    ax.yaxis.set_major_formatter(
        mticker.FuncFormatter(lambda v, _: f"{v:.0f}K"))
    ax.legend()
    ax.grid(True, alpha=0.35)
    plt.tight_layout()

    ruta = ruta_graficas + "8_prediccion_total_empresa.png"
    plt.savefig(ruta, dpi=150, bbox_inches="tight")
    plt.close()
    print("  Guardada: 8_prediccion_total_empresa.png")


def graficar_top_productos(df_pred_producto, nombre_modelo,
                            target="cantidad", top_n=15,
                            ruta_graficas=GRAFICAS):
    """Top N productos por prediccion acumulada Abr-Dic."""
    col = f"pred_{target}"
    top = (df_pred_producto
           .groupby(["codigo", "producto"])[col]
           .sum()
           .sort_values(ascending=False)
           .head(top_n)
           .reset_index())

    _, ax = plt.subplots(figsize=(14, 7))
    bars = ax.barh(top["producto"], top[col] / 1000,
                   color=COLORES[0], alpha=0.85)
    ax.set_xlabel(f"{target.replace('_', ' ').title()} (miles)")
    ax.set_title(f"Top {top_n} Productos — Prediccion Abr-Dic 2026\n({nombre_modelo})",
                 fontsize=13, fontweight="bold")
    ax.invert_yaxis()
    ax.xaxis.set_major_formatter(
        mticker.FuncFormatter(lambda v, _: f"{v:.0f}K"))
    for bar, val in zip(bars, top[col]):
        ax.text(bar.get_width() + 0.3,
                bar.get_y() + bar.get_height() / 2,
                f"{val/1000:.1f}K", va="center", fontsize=8)
    plt.tight_layout()

    ruta = ruta_graficas + "9_top_productos_prediccion.png"
    plt.savefig(ruta, dpi=150, bbox_inches="tight")
    plt.close()
    print("  Guardada: 9_top_productos_prediccion.png")


def graficar_real_vs_pred_test(y_test_values, preds_por_modelo,
                                ruta_graficas=GRAFICAS):
    """Scatter real vs predicho en test para cada modelo (subplots)."""
    n = len(preds_por_modelo)
    fig, axes = plt.subplots(1, n, figsize=(6 * n, 5))
    if n == 1:
        axes = [axes]
    fig.suptitle("Real vs Predicho — Conjunto de Test",
                 fontsize=13, fontweight="bold")

    for ax, (nombre, y_pred), color in zip(axes, preds_por_modelo.items(), COLORES):
        lim = max(y_test_values.max(), np.max(y_pred)) * 1.05
        ax.scatter(y_test_values, y_pred, alpha=0.3, s=15, color=color)
        ax.plot([0, lim], [0, lim], "k--", linewidth=1, label="Perfecto")
        ax.set_xlim(0, lim)
        ax.set_ylim(0, lim)
        ax.set_xlabel("Real")
        ax.set_ylabel("Predicho")
        ax.set_title(nombre, fontsize=11)
        ax.legend(fontsize=8)

    plt.tight_layout()
    ruta = ruta_graficas + "10_real_vs_pred_test.png"
    plt.savefig(ruta, dpi=150, bbox_inches="tight")
    plt.close()
    print("  Guardada: 10_real_vs_pred_test.png")


# ══════════════════════════════════════════════════════════════════════════════
# BLOQUE 5 — REPORTE EXCEL
# ══════════════════════════════════════════════════════════════════════════════

def guardar_reporte(df_metricas, preds_producto_por_modelo,
                    totales_por_modelo, df_clasificacion,
                    df_bajo_pedido, ruta=SALIDA):
    """
    Exporta a Excel:
      - Hoja 'Clasificacion'          : todos los productos con su grupo
      - Hoja 'Metricas_Modelos'       : comparacion MAE/RMSE/R2/SMAPE (solo recurrentes)
      - Hoja 'Pred_Prod_{modelo}'     : prediccion por producto (1 hoja x modelo)
      - Hoja 'Total_{modelo}'         : prediccion total empresa (1 hoja x modelo)
      - Hoja 'Linea_{modelo}'         : prediccion por linea (1 hoja x modelo)
      - Hoja 'Resumen_Comparativo'    : los 4 modelos lado a lado por mes
      - Hoja 'Bajo_Pedido'            : perfil y stock sugerido de esporadicos
    """
    print(f"\n[guardar_reporte] Exportando a {ruta} ...")

    with pd.ExcelWriter(ruta, engine="openpyxl") as writer:

        # — Clasificacion general de todos los productos ——————————————————————
        df_clasificacion.to_excel(writer, sheet_name="Clasificacion", index=False)

        # — Metricas de los 4 modelos (solo productos planeados) —————————————
        df_metricas.to_excel(writer, sheet_name="Metricas_Modelos", index=False)

        # — Predicciones por producto (una hoja por modelo) ——————————————————
        for nombre, df_prod in preds_producto_por_modelo.items():
            hoja = f"Pred_Prod_{nombre[:10].replace(' ','_')}"
            df_prod.to_excel(writer, sheet_name=hoja, index=False)

        # — Totales por modelo ————————————————————————————————————————————————
        for nombre, (df_total, df_linea) in totales_por_modelo.items():
            hoja_t = f"Total_{nombre[:10].replace(' ','_')}"
            df_total.to_excel(writer, sheet_name=hoja_t, index=False)

            hoja_l = f"Linea_{nombre[:9].replace(' ','_')}"
            df_linea.to_excel(writer, sheet_name=hoja_l, index=False)

        # — Resumen comparativo mensual (todos los modelos juntos) ————————————
        filas = []
        for nombre, (df_total, _) in totales_por_modelo.items():
            col_total = [c for c in df_total.columns if c.startswith("total_")][0]
            tgt = col_total.replace("total_", "")
            for _, r in df_total.iterrows():
                filas.append({
                    "modelo":       nombre,
                    "año":          int(r["año"]),
                    "mes":          int(r["mes"]),
                    "mes_nombre":   MESES_NOMBRES[int(r["mes"])],
                    f"pred_{tgt}":  r[col_total],
                })
        pd.DataFrame(filas).to_excel(
            writer, sheet_name="Resumen_Comparativo", index=False)

        # — Perfil bajo pedido con stock sugerido ————————————————————————————
        if not df_bajo_pedido.empty:
            df_bajo_pedido.to_excel(writer, sheet_name="Bajo_Pedido", index=False)

    print("  Archivo guardado correctamente.")
    print("  Hojas:")
    print("    Clasificacion         — todos los productos + grupo asignado")
    print("    Metricas_Modelos      — comparacion de los 4 modelos (planeados)")
    for nombre in preds_producto_por_modelo:
        print(f"    Pred_Prod_{nombre[:10]:<12} — prediccion por producto")
        print(f"    Total_{nombre[:10]:<14} — prediccion total empresa")
        print(f"    Linea_{nombre[:9]:<14} — prediccion por linea")
    print("    Resumen_Comparativo   — los 4 modelos juntos por mes")
    if not df_bajo_pedido.empty:
        print("    Bajo_Pedido           — perfil historico + stock sugerido")


# ══════════════════════════════════════════════════════════════════════════════
# BLOQUE 6 — MAIN
# ══════════════════════════════════════════════════════════════════════════════

def main(target="cantidad", umbral_frecuencia=0.60):
    """
    Ejecuta el pipeline completo:
      1. Carga y prepara datos desde ventas_consolidado.xlsx
      2. Limpia ceros inconsistentes y outliers
      3. Segmenta productos: PLANEADOS (>=umbral) vs BAJO PEDIDO (<umbral)
      4. PLANEADOS → features + 4 modelos ML + predicciones Abr-Dic 2026
      5. BAJO PEDIDO → perfil de demanda historica + stock sugerido
      6. Guarda graficas y reporte Excel (hojas separadas por grupo)

    Parametros:
      target            : 'cantidad' o 'valor_venta'
      umbral_frecuencia : fraccion minima de meses con venta para ser planeado
                          (default 0.60 = 60%)
    """
    print("=" * 65)
    print("PASO 2 — Modelos ML de prediccion de ventas SeedPack")
    print(f"         Target: {target}  |  Umbral frecuencia: {umbral_frecuencia:.0%}")
    print("=" * 65)

    # ── 1. Carga y limpieza ───────────────────────────────────────────────────
    df_raw     = cargar_datos()
    df_mensual = agregar_mensual(df_raw)
    df_modelo  = preparar_para_modelo(df_mensual)
    df_limpio  = limpiar_datos(df_modelo, target=target)

    # ── 2. Segmentacion ───────────────────────────────────────────────────────
    df_recurrentes, df_esporadicos, df_clasificacion = segmentar_productos(
        df_limpio, umbral_frecuencia=umbral_frecuencia
    )

    # ── 3. Modulo BAJO PEDIDO (deterministico) ────────────────────────────────
    print("\n" + "─" * 65)
    print("MODULO A — Programacion Bajo Pedido (esporadicos)")
    print("─" * 65)
    df_bajo_pedido = perfil_bajo_pedido(df_esporadicos, target=target)

    # ── 4. Modulo PLANEADO — features ─────────────────────────────────────────
    print("\n" + "─" * 65)
    print("MODULO B — Programacion Planeada (recurrentes → ML)")
    print("─" * 65)
    df_features = crear_features(df_recurrentes, target=target)

    X_train, X_test, y_train, y_test, _, feature_cols = (
        preparar_train_test(df_features, target=target)
    )

    # ── 2. Entrenar los 4 modelos ─────────────────────────────────────────────
    rf_model,  rf_pred,  rf_m  = modelo_random_forest(
        X_train, X_test, y_train, y_test)

    xgb_model, xgb_pred, xgb_m = modelo_xgboost(
        X_train, X_test, y_train, y_test)

    reg_model, reg_pred, reg_m = modelo_regresion_regularizada(
        X_train, X_test, y_train, y_test)

    mlp_model, mlp_pred, mlp_m = modelo_red_neuronal(
        X_train, X_test, y_train, y_test)

    # ── 3. Comparar ───────────────────────────────────────────────────────────
    df_metricas, mejor_nombre = comparar_modelos([rf_m, xgb_m, reg_m, mlp_m])

    # ── 4. Predicciones futuras Abr-Dic 2026 (los 4 modelos) ─────────────────
    print("\n[Generando predicciones Abr-Dic 2026...]")
    meses_futuros = list(range(4, 13))

    modelos = {
        "Random Forest":      rf_model,
        "XGBoost":            xgb_model,
        reg_m["modelo"]:      reg_model,
        "Red Neuronal (MLP)": mlp_model,
    }

    preds_producto = {}
    totales        = {}

    for nombre, modelo in modelos.items():
        print(f"\n  >> {nombre}")
        df_prod = predecir_por_producto(
            modelo, df_features, feature_cols,
            meses_futuros, target=target
        )
        df_total, df_linea = predecir_total_empresa(df_prod, target=target)
        preds_producto[nombre] = df_prod
        totales[nombre]        = (df_total, df_linea)

    # ── 5. Graficas ───────────────────────────────────────────────────────────
    print("\n[Generando graficas...]")
    graficar_comparacion_modelos(df_metricas)

    graficar_prediccion_total(
        {n: t for n, (t, _) in totales.items()},
        target=target,
    )

    graficar_top_productos(
        preds_producto[mejor_nombre],
        nombre_modelo=mejor_nombre,
        target=target,
    )

    graficar_real_vs_pred_test(
        y_test.values,
        {
            "Random Forest":      rf_pred,
            "XGBoost":            xgb_pred,
            reg_m["modelo"]:      reg_pred,
            "Red Neuronal (MLP)": mlp_pred,
        },
    )

    # ── 6. Reporte Excel ──────────────────────────────────────────────────────
    guardar_reporte(
        df_metricas, preds_producto, totales,
        df_clasificacion, df_bajo_pedido,
    )

    print("\n" + "=" * 65)
    print(f"PASO 2 COMPLETADO — Mejor modelo: {mejor_nombre}")
    print(f"Graficas en : {GRAFICAS}")
    print(f"Reporte en  : {SALIDA}")
    print("=" * 65)

    return {
        "metricas":        df_metricas,
        "preds_producto":  preds_producto,
        "totales":         totales,
        "mejor_modelo":    mejor_nombre,
        "clasificacion":   df_clasificacion,
        "bajo_pedido":     df_bajo_pedido,
    }


# ─── Entry point ──────────────────────────────────────────────────────────────
if __name__ == "__main__":
    main(target="cantidad")
