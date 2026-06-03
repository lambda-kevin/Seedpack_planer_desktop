import os, sys
import pandas as pd
from sklearn.cluster import KMeans

_DIR  = os.path.dirname(os.path.abspath(__file__))
BASE  = (os.path.dirname(sys.executable) if getattr(sys, 'frozen', False)
         else os.path.join(os.path.dirname(os.path.dirname(_DIR)), "app")) + "/"
ENTRADA = BASE + "datos_pipeline/paso1_limpieza.xlsx"
SALIDA  = BASE + "datos_pipeline/paso2_clusterizacion.xlsx"


def run():
    df = pd.read_excel(ENTRADA)
    df["fecha"] = pd.to_datetime(df["fecha"])
    df["año"]   = df["fecha"].dt.year

    freq = (
        df.groupby(["producto", "año"])
          .size()
          .reset_index(name="pedidos_ese_año")
    )

    resumen = (
        freq.groupby("producto")["pedidos_ese_año"]
            .mean()
            .reset_index(name="frecuencia_anual_promedio")
    )

    X = resumen[["frecuencia_anual_promedio"]]
    kmeans = KMeans(n_clusters=2, random_state=42, n_init=10)
    resumen["cluster_id"] = kmeans.fit_predict(X)

    centroides = resumen.groupby("cluster_id")["frecuencia_anual_promedio"].mean()
    cluster_frecuente  = centroides.idxmax()   # noqa: F841
    cluster_esporadico = centroides.idxmin()   # noqa: F841

    resumen["cluster"] = resumen["frecuencia_anual_promedio"].apply(
        lambda x: "esporadico" if x <= 6 else "frecuente"
    )

    resumen = resumen.drop(columns="cluster_id").sort_values("frecuencia_anual_promedio")

    resumen.to_excel(SALIDA, index=False)

    print(f"Productos clasificados: {len(resumen)}")
    print(resumen["cluster"].value_counts().to_string())
    print(f"Guardado en: {SALIDA}")


if __name__ == "__main__":
    run()
