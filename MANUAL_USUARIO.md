# SEEDPACK PLANNER — Manual de Usuario

**Versión:** 1.0  
**Fecha:** Mayo 2026  
**Sistema:** Planificación de Producción Basada en Machine Learning

---

## Tabla de Contenidos

1. [¿Qué es SEEDPACK Planner?](#1-qué-es-seedpack-planner)
2. [Requisitos del sistema](#2-requisitos-del-sistema)
3. [Estructura de carpetas](#3-estructura-de-carpetas)
4. [Archivos de entrada](#4-archivos-de-entrada)
5. [Cómo usar la aplicación](#5-cómo-usar-la-aplicación)
   - [Pestaña 1 — Ordenes de Producción](#pestaña-1--ordenes-de-producción)
   - [Pestaña 2 — Dashboard Ejecutivo](#pestaña-2--dashboard-ejecutivo)
   - [Pestaña 3 — Plan Diario](#pestaña-3--plan-diario)
   - [Pestaña 4 — Tablas del Excel](#pestaña-4--tablas-del-excel)
   - [Pestaña 5 — Gestión de Pedidos](#pestaña-5--gestión-de-pedidos)
   - [Pestaña 6 — Saldo de OPs](#pestaña-6--saldo-de-ops)
6. [El pipeline paso a paso](#6-el-pipeline-paso-a-paso)
7. [Archivos de salida](#7-archivos-de-salida)
8. [Exportar resultados](#8-exportar-resultados)
9. [Resolución de problemas](#9-resolución-de-problemas)
10. [Glosario](#10-glosario)

---

## 1. ¿Qué es SEEDPACK Planner?

SEEDPACK Planner es una aplicación de escritorio que genera automáticamente un **plan de producción diario** para el período que usted defina. Combina:

- El **histórico de ventas** de los últimos años para entrenar modelos de Machine Learning.
- El **inventario actual** (stock en bodega).
- Las **Órdenes de Producción en proceso** (lo que ya está fabricando).
- Los **pedidos pendientes** de clientes.
- Los **lotes mínimos** de producción por referencia.

Con toda esa información, el sistema calcula cuánto producir de cada referencia, cada día hábil, dentro del período elegido.

### ¿Qué obtiene usted al final?

Un archivo Excel (`orden_produccion_final.xlsx`) con 10+ hojas que incluyen:

| Hoja | ¿Qué contiene? |
|------|----------------|
| **Plan_Diario** | Qué producir, de qué referencia, cada día |
| **Plan_Mensual** | Resumen por mes y producto |
| **Calendario_Semanal** | Vista de semana × producto |
| **Inventario_Actual** | Stock consolidado con costo total actualizado |
| **Inventario_Detalle** | Desglose por bodega y lote |
| **OPs_En_Proceso** | Órdenes activas cargadas |
| **Pedidos_Ocasionales** | Análisis de pedidos de referencias especiales |
| **Saldo_OPs** | Cruce de OPs con entradas reales de inventario |
| **Gestion_Pedidos** | Análisis de déficit por referencia vs pedidos |
| **Comparativa_Comercial** | Plan ML vs plan comercial (si aplica) |
| **Dashboard** | KPIs ejecutivos resumidos |

---

## 2. Requisitos del sistema

- **Sistema Operativo:** Windows 10 / Windows 11
- **Versión .exe:** No requiere instalación adicional.
- **Versión script Python:** Python 3.10+ con el entorno virtual `venv/` del proyecto.
- **Microsoft Excel:** Recomendado para visualizar los resultados (no obligatorio para ejecutar).
- **RAM mínima:** 4 GB (8 GB recomendado para catálogos grandes).
- **Espacio en disco:** 500 MB libres en la carpeta del proyecto.

---

## 3. Estructura de carpetas

```
SEEDPACK\
│
├── app\                              ← DATOS DEL PROYECTO
│   ├── Archivos Historicos\          ← Archivos de entrada
│   │   ├── ventasXproducto.xlsx           (pre-cargado)
│   │   ├── INEDITTO_OP.xlsx               (pre-cargado)
│   │   ├── Codigos lote minimo.xlsx       (pre-cargado)
│   │   ├── plan_comercial.xlsx            (pre-cargado, opcional)
│   │   ├── INEDITTO_Listado_Pedidos.xlsx  (cargar manualmente)
│   │   ├── ordenes de produccion en proceso.xlsx  (cargar manualmente)
│   │   └── Bodega.xlsx                    (cargar manualmente)
│   │
│   ├── datos_pipeline\               ← Archivos intermedios (generados automáticamente)
│   │   ├── paso1_limpieza.xlsx
│   │   ├── paso2_clusterizacion.xlsx
│   │   ├── paso3_modelos.xlsx
│   │   └── paso4_proyeccion.xlsx
│   │
│   ├── Resultado Final\              ← SALIDA PRINCIPAL
│   │   └── orden_produccion_final.xlsx
│   │
│   └── graficas\                     ← Gráficas del Dashboard
│
├── codigo\                           ← Código fuente (no modificar)
│
├── dist\                             ← Ejecutable compilado
│   └── SeedPack Planner.exe
│
└── crear_ejecutable.bat              ← Recompila el .exe
```

> **Importante:** Los archivos en `datos_pipeline\` y `Resultado Final\` se sobreescriben en cada ejecución. No guarde información importante en esas carpetas.

---

## 4. Archivos de entrada

La aplicación necesita **6 archivos obligatorios** y acepta **2 opcionales**.

### 4.1 Archivos pre-cargados (ya disponibles en Archivos Historicos)

Estos archivos se actualizan con menor frecuencia y la app los carga automáticamente al iniciar.

---

#### 01 — Historial de ventas
**Nombre de archivo:** `ventasXproducto.xlsx`  
**Cuándo actualizar:** Mensualmente, con el histórico acumulado.

| Columna requerida | Descripción |
|-------------------|-------------|
| `Fecha Fra.` | Fecha de la factura (formato fecha de Excel) |
| `Producto` | Nombre completo del producto |
| `Cantidad` | Unidades vendidas (número entero positivo) |
| `Factura` | Número de factura o remisión |

> **Nota:** El sistema limpia automáticamente filas con fechas inválidas, productos vacíos o cantidades nulas. Cuanto más historial contenga este archivo (mínimo 2 años recomendado), mejores serán las predicciones.

---

#### 02 — Histórico de Órdenes de Producción
**Nombre de archivo:** `INEDITTO_OP.xlsx`  
**Cuándo actualizar:** Cuando cambien tamaños de lote habituales de producción.

| Columna requerida | Descripción |
|-------------------|-------------|
| `Cód. Producto` | Código del producto terminado (PT) |
| `Referencia` | Nombre/descripción del producto |
| `Cant. Aprobada` | Cantidad aprobada en la OP |

> **Uso interno:** El sistema calcula la **mediana de lotes históricos** por referencia. Estos tamaños se usan para distribuir la producción en días hábiles del mes.

---

#### 06 — Lotes mínimos de producción
**Nombre de archivo:** `Codigos lote minimo.xlsx`  
**Cuándo actualizar:** Cuando cambien los lotes mínimos o se agreguen/eliminen referencias.

**Estructura** (el sistema lee desde la fila 3, sin encabezado):

| Columna | Contenido |
|---------|-----------|
| A | Código PT |
| B | Referencia / Nombre |
| C | Línea de producción |
| D | Número de cavidades |
| E | **Lote mínimo** (unidades mínimas por OP) |

> **Columna Línea:** Las referencias con `Línea = "Genérico"` son productos que siempre deben mantener stock (se planifican con proyección ML aunque no tengan pedidos activos). Las demás referencias solo se producen si hay pedido.

---

#### 08 — Plan Comercial *(opcional)*
**Nombre de archivo:** `plan_comercial.xlsx`  
**Cuándo actualizar:** Al inicio de cada período comercial.

El sistema detecta automáticamente las columnas de meses. Genera la hoja `Comparativa_Comercial` en el resultado para contrastar la proyección del equipo comercial vs la del modelo ML.

---

### 4.2 Archivos que se cargan manualmente (datos del período actual)

Estos archivos cambian con frecuencia y deben cargarse cada vez que se ejecuta el plan.

---

#### 03 — Listado de pedidos pendientes
**Nombre de archivo:** Exportar desde INEDDITO como `INEDITTO_Listado_Pedidos.xlsx`  
**Cuándo obtenerlo:** Antes de cada ejecución del plan (datos del día).

**Hoja requerida:** `Maestro`  
**El sistema lee desde la fila 4 (encabezado en fila 4).**

| Columna requerida | Descripción |
|-------------------|-------------|
| `Pedido` | Número de pedido |
| `Código` o `Código PT` | Código del producto |
| `Descripción` o `Descripcion` | Nombre del producto |
| `Saldo Pendiente` | Unidades aún sin despachar |
| `Fecha de entrega` | Fecha compromiso con el cliente |

> **Filtro automático:** Solo se toman pedidos con `Saldo Pendiente > 0`. Los pedidos ya despachados (saldo = 0) se ignoran.

---

#### 04 — OPs en proceso
**Nombre de archivo:** Exportar desde INEDDITO como `ordenes de produccion en proceso.xlsx`  
**Cuándo obtenerlo:** Antes de cada ejecución del plan (datos del día).

| Columna requerida | Descripción |
|-------------------|-------------|
| `Cód. Producto` | Código del producto |
| `Cant. Aprobada` | Unidades en fabricación |
| `Fecha Programada` | Fecha estimada de entrega al almacén |
| `Compromiso Cliente` | Fecha compromiso con el cliente |
| `OP` | Número de orden de producción |
| `Tipo de Trabajo` | Tipo de OP |
| `Referencia` | Nombre del producto |
| `Cliente` | Cliente asociado a la OP |

> **Uso en MRP:** Las unidades de cada OP se descuentan de la demanda en el mes de `Fecha Programada`, evitando producir en exceso lo que ya está en fabricación.

---

#### 05 — Existencias en bodega
**Nombre de archivo:** Exportar desde INEDDITO (módulo ExistBodega) como `Bodega.xlsx`  
**Cuándo obtenerlo:** Antes de cada ejecución del plan (datos del día).

> **Importante:** Use el reporte **ExistBodega** (saldo real actual), **NO** el reporte de entradasInventario. La diferencia es que ExistBodega muestra el saldo vigente, mientras que entradas muestra el acumulado histórico.

| Columna requerida | Descripción |
|-------------------|-------------|
| `Cód. PT` | Código del producto terminado |
| `Bodega` | Nombre de la bodega |
| `Saldo` | Unidades disponibles en ese lote |
| `Costo` | Costo unitario del lote |
| `Cant. Ingresada` | Cantidad original del lote (referencia) |
| `OP/OC` | Número de OP u OC del lote |

> **Bodegas excluidas:** La bodega `Obsoletos` se excluye automáticamente del cálculo de stock disponible.

---

#### 07 — Entradas de inventario *(opcional)*
**Nombre de archivo:** Exportar desde INEDDITO.  
**Para qué sirve:** Permite calcular el saldo real pendiente de cada OP en proceso, cruzando lo que la OP aprobó vs lo que ya entró efectivamente a bodega. Genera la hoja `Saldo_OPs`.

---

## 5. Cómo usar la aplicación

Al abrir `SeedPack Planner.exe` (o ejecutar `app.py`), verá una ventana con **6 pestañas** en la parte superior.

---

### Pestaña 1 — Ordenes de Producción

Esta es la pestaña principal. Desde aquí se configura y ejecuta el plan.

#### Panel izquierdo (Parámetros)

| Campo | Descripción | Ejemplo |
|-------|-------------|---------|
| **Fecha inicial** | Primer día del período a planificar | `2026-04-01` |
| **Fecha final** | Último día del período a planificar | `2026-12-31` |

> **Formato obligatorio:** `AAAA-MM-DD` (año de 4 dígitos, mes y día con dos dígitos).

**Botones disponibles:**

- **Generar Plan de Producción:** Inicia el pipeline completo. El proceso puede tardar entre 2 y 10 minutos según el tamaño del catálogo.
- **Exportar Excel:** Activo solo después de generar el plan. Guarda una copia del resultado donde usted indique.
- **Exportar PDF:** Genera un informe PDF con KPIs y gráficas.
- **Cancelar:** Cierra la aplicación.

#### Panel derecho (Archivos de entrada)

Muestra 8 tarjetas con el estado de cada archivo:

| Indicador | Significado |
|-----------|-------------|
| 🔵 **Pre-cargado** | El archivo se encontró automáticamente en `Archivos Historicos\`. |
| ✅ **Cargado** | Archivo cargado manualmente por el usuario. |
| 🔴 **Requerido** | Archivo obligatorio, aún no cargado. El plan no puede ejecutarse sin él. |
| ⚪ **Opcional** | No es obligatorio; activa funciones adicionales si se carga. |

**Para cargar un archivo manualmente:**
1. Haga clic en el botón **"Examinar"** de la tarjeta correspondiente.
2. Navegue hasta el archivo exportado desde INEDDITO.
3. La tarjeta cambiará a color verde ✅.

**Para actualizar un archivo pre-cargado:**
- Haga clic en **"Actualizar"** para seleccionar una versión más reciente.

#### Ventana de progreso

Al hacer clic en "Generar Plan", aparece una ventana con:

- **Barra de progreso** (0% → 100%).
- **Lista de 6 pasos** con estado (activo ⏳ / completado ✓).
- **Tiempo transcurrido** y **tiempo estimado** restante.
- **Botón "Cancelar proceso"**: Detiene la ejecución (el proceso en curso termina antes de cerrarse).

Los 6 pasos del pipeline son:

| Paso | Nombre | Descripción |
|------|--------|-------------|
| 1 | Limpieza de datos | Normaliza y filtra el histórico de ventas |
| 2 | Clusterización | Clasifica productos en frecuentes vs esporádicos |
| 3 | Modelos ML de predicción | Entrena los modelos de Machine Learning |
| 4 | Proyección de ventas | Genera demanda proyectada mes a mes |
| 5 | Ordenes de producción | Crea el plan diario con MRP y lotes |
| 6 | Gestión de pedidos | Cruza pedidos pendientes con stock y proyección |

---

### Pestaña 2 — Dashboard Ejecutivo

Se carga automáticamente después de generar el plan (o al abrir la app si ya existe un resultado).

#### KPIs del dashboard (10 indicadores en 2 filas)

| KPI | Descripción |
|-----|-------------|
| Días planificados | Días hábiles totales en el período |
| Referencias | Cantidad de productos en el plan |
| Unidades a producir | Total de unidades en el plan |
| Refs. proyección | Productos con proyección ML activa |
| Refs. ocasionales | Productos solo con pedidos activos |
| Uds. ocasionales | Unidades de pedidos ocasionales |
| Cubiertos por stock | Referencias cubiertas sin producir |
| OPs en proceso | Número de OPs activas cargadas |
| Inventario (uds) | Total de unidades en bodega |
| Costo inventario | Valor total del inventario (en millones $) |

#### Gráficas

El dashboard genera hasta **8 gráficas** automáticamente:

1. Top 10 referencias por unidades a producir
2. Producción semanal total (unidades por semana)
3. Distribución por tipo (Proyección vs Ocasional) — gráfico donut
4. Top 15 inventario por unidades en bodega
5. OPs en proceso por tipo de trabajo
6. Top 10 pedidos ocasionales por unidades
7. Producción mensual por tipo (barras apiladas)
8. Top 10 inventario por costo total

**Botones:**
- **Actualizar:** Recarga el dashboard desde el archivo de resultados.
- **PDF:** Exporta el dashboard como informe PDF.

---

### Pestaña 3 — Plan Diario

Muestra la tabla del plan de producción con herramientas de filtrado.

#### Filtros disponibles

| Filtro | Opciones |
|--------|----------|
| **Búsqueda** | Texto libre — filtra por cualquier columna |
| **Tipo** | Todos / Proyección / Ocasional |

#### Columnas de la tabla

| Columna | Descripción |
|---------|-------------|
| Fecha | Día hábil de producción |
| Codigo PT | Código del producto terminado |
| Referencia | Nombre del producto |
| Cantidad a Producir | Unidades a fabricar ese día |
| Tipo | `Proyeccion` (ML) o `Ocasional` (pedido específico) |
| Fecha Entrega | Para ocasionales: fecha compromiso con cliente |

**Ordenamiento:** Haga clic en cualquier encabezado de columna para ordenar de forma ascendente o descendente.

**Colores:**
- Fondo azul claro → Proyección ML
- Fondo amarillo → Ocasional (pedido específico)

---

### Pestaña 4 — Tablas del Excel

Permite explorar todas las hojas del archivo de resultados dentro de la aplicación, sin necesidad de abrir Excel.

Las hojas disponibles se muestran como sub-pestañas:

- Plan_Diario
- Plan_Mensual
- Calendario_Semanal
- OPs_En_Proceso
- Pedidos_Ocasionales
- Inventario_Actual
- Inventario_Detalle
- Gestion_Pedidos
- Saldo_OPs
- Comparativa_Comercial *(si se cargó plan comercial)*
- Dashboard KPIs

**Funciones disponibles en cada hoja:**
- Campo de búsqueda por texto.
- Scroll horizontal y vertical.
- Botón **"Exportar Excel"**: guarda la hoja actual.
- Botón **"Exportar PDF"**: exporta la vista actual como PDF.

---

### Pestaña 5 — Gestión de Pedidos

Análisis cruzado de los pedidos pendientes vs recursos disponibles (stock + OPs + proyección ML).

#### KPIs superiores

| KPI | Color | Descripción |
|-----|-------|-------------|
| Pedidos activos | Azul | Total de pedidos con saldo pendiente |
| Saldo pendiente total | Violeta | Suma de unidades pendientes |
| Refs. con OP requerida | Rojo | Referencias que necesitan nueva OP |
| Total a producir | Naranja | Unidades totales a producir por déficit |

#### Tabla de análisis

| Columna | Descripción |
|---------|-------------|
| Codigo | Código PT de la referencia |
| Descripcion | Nombre del producto |
| Pedidos | Número de pedidos activos |
| Saldo Pendiente | Unidades sin despachar |
| Stock Bodega | Unidades en inventario físico |
| OPs Proceso | Unidades en fabricación activa |
| Proyeccion ML | Unidades que el modelo proyecta producir |
| Disponible | Stock + OPs en proceso |
| Demanda Total | Saldo pendiente + Proyección ML |
| Deficit | Máximo(0, Demanda - Disponible) |
| Lote Minimo | Tamaño mínimo de lote de producción |
| A Producir | Déficit ajustado al múltiplo del lote mínimo |
| **Estado** | **OK** (cubierto) o **Solicitar OP** (requiere acción) |
| Justificacion | Resumen textual del análisis para esa referencia |

**Cómo interpretar el Estado:**

- ✅ **OK (fondo verde):** El stock disponible + OPs cubre la demanda total. No se requiere nueva orden de producción.
- 🟡 **Solicitar OP (fondo amarillo):** Hay déficit. La columna "A Producir" indica cuántas unidades ordenar, ya ajustadas al lote mínimo. Divida por el lote mínimo para saber el número de lotes a solicitar.

---

### Pestaña 6 — Saldo de OPs

Muestra el avance real de cada OP en proceso, cruzando la cantidad aprobada vs las entradas reales que ya llegaron a bodega.

> Esta pestaña solo muestra datos completos si se cargó el archivo opcional **07 — Entradas de inventario (OC)**.

#### KPIs

| KPI | Descripción |
|-----|-------------|
| OPs en proceso | Número total de OPs activas |
| Total aprobado (uds) | Suma de cantidades aprobadas en todas las OPs |
| Ya producido (uds) | Suma de entradas confirmadas en inventario |
| Pendiente por producir | Aprobado − Producido |

#### Tabla

| Columna | Descripción |
|---------|-------------|
| OP | Número de orden de producción |
| Tipo | Tipo de trabajo |
| Cód. Producto | Código PT |
| Referencia | Nombre del producto |
| Cant. Aprobada | Unidades de la OP |
| Cant. Producida | Entradas reales confirmadas en bodega |
| Cant. Pendiente | Aprobada − Producida |
| **Estado** | Completada / En Proceso / Sin Iniciar |
| Fecha Prog. | Fecha programada de entrega |
| Compromiso | Fecha compromiso con el cliente |

**Estados de la OP:**
- 🟢 **Completada:** Cant. Producida ≥ Cant. Aprobada.
- 🟡 **En Proceso:** Cant. Producida > 0 pero < Aprobada.
- 🔴 **Sin Iniciar:** Cant. Producida = 0.

---

## 6. El pipeline paso a paso

Internamente, al presionar "Generar Plan", el sistema ejecuta estos 6 pasos en secuencia:

### Paso 1 — Limpieza de datos

**Qué hace:**
- Lee `ventasXproducto.xlsx` (hoja "Ventas por producto").
- Convierte la columna de fechas a formato estándar.
- Descarta filas con fecha inválida, producto vacío o cantidad nula.
- Ordena cronológicamente.
- Guarda en `datos_pipeline/paso1_limpieza.xlsx`.

**Duración típica:** 5–15 segundos.

---

### Paso 2 — Clusterización de productos

**Qué hace:**
- Calcula con qué frecuencia se vendió cada producto (promedio de pedidos por año).
- Clasifica en dos grupos usando KMeans:
  - **Frecuente:** Más de 6 pedidos/año en promedio → sujeto a proyección ML.
  - **Esporádico:** 6 o menos pedidos/año → solo se cubre si hay pedido activo.

**Por qué importa:** Un producto que solo se vende 2 veces al año no tiene suficiente patrón histórico para que el modelo ML proyecte bien. Separar estos productos evita sobreproducción.

**Duración típica:** 3–10 segundos.

---

### Paso 3 — Entrenamiento de modelos ML

**Qué hace:**
- Prepara los datos históricos de ventas mensuales por producto frecuente.
- Crea 14 variables predictoras (mes, estacionalidad, lags, medias móviles, tendencia).
- Entrena 4 modelos alternativos:

| Modelo | Descripción | Mejor para |
|--------|-------------|------------|
| **Random Forest** *(predeterminado)* | Ensamble de 300 árboles de decisión | Patrones no lineales, robustez |
| **XGBoost** | Gradient boosting con regularización | Alta precisión en datos estructurados |
| **Reg. Regularizada** | Ridge / Lasso / ElasticNet | Datos con pocas variables |
| **Red Neuronal** | MLP con capas 128-64-32 neuronas | Relaciones complejas, más lento |

- Guarda métricas de evaluación en `datos_pipeline/paso3_modelos.xlsx`.

**Duración típica:** 1–5 minutos (depende del tamaño del catálogo y la máquina).

---

### Paso 4 — Proyección de ventas

**Qué hace:**
- Usa el modelo entrenado para proyectar la demanda mensual de cada producto **frecuente** desde la Fecha Inicial hasta la Fecha Final.
- Utiliza **rolling forecast**: cada predicción alimenta los valores de períodos anteriores (lags) para la siguiente predicción.
- Guarda resultados en `datos_pipeline/paso4_proyeccion.xlsx`.

**Ejemplo de resultado:**

| Producto | Año | Mes | Proyectado |
|----------|-----|-----|-----------|
| Bandeja plana #7 | 2026 | 4 | 73.600 |
| Bandeja plana #7 | 2026 | 5 | 71.200 |
| ... | ... | ... | ... |

**Duración típica:** 30–120 segundos.

---

### Paso 5 — Ordenes de producción (MRP)

**Qué hace:**
Para cada producto en cada mes del período:

```
Demanda del mes     = Proyección ML (frecuentes) o pedido (esporádicos)
Stock disponible    = Inventario actual + OPs programadas para ese mes
Neto a producir     = máx(0, Demanda − Stock disponible)
Neto ajustado       = Redondear al múltiplo del lote mínimo (hacia arriba)
```

Luego distribuye el neto ajustado de cada mes en los días hábiles (lunes a sábado, sin festivos colombianos), agrupando en lotes cuando aplica.

Adicional:
- Integra **pedidos ocasionales**: referencias sin proyección ML que tienen pedidos activos. Se programan 3 días hábiles antes de la fecha de entrega comprometida.
- Genera todas las hojas de análisis en `Resultado Final/orden_produccion_final.xlsx`.

**Duración típica:** 30–90 segundos.

---

### Paso 6 — Gestión de pedidos

**Qué hace:**
- Cruza cada referencia del listado de pedidos pendientes con:
  - Stock en bodega
  - OPs en proceso
  - Proyección ML del período
- Calcula déficit y cuántas unidades solicitar (ajustado a lote mínimo).
- Escribe la hoja `Gestion_Pedidos` en el Excel de resultados.

---

## 7. Archivos de salida

### orden_produccion_final.xlsx

Ubicación: `app\Resultado Final\orden_produccion_final.xlsx`

#### Hoja Plan_Diario

La hoja más importante. Una fila por cada día/producto a producir.

| Columna | Tipo | Descripción |
|---------|------|-------------|
| Fecha | Fecha | Día hábil de producción |
| Codigo PT | Texto | Código del producto |
| Referencia | Texto | Nombre del producto |
| Cantidad a Producir | Número | Unidades ese día |
| Tipo | Texto | `Proyeccion` o `Ocasional` |
| Fecha Entrega | Fecha | Solo para ocasionales |

#### Hoja Inventario_Actual

Consolidado de stock por producto (excluye Obsoletos).

| Columna | Descripción |
|---------|-------------|
| Cod. PT | Código del producto |
| Referencia | Nombre |
| # Bodegas | En cuántas bodegas hay stock |
| Saldo | Total de unidades disponibles |
| Costo Unit. | Costo unitario ponderado (promedio por saldo) |
| Costo Total | **Saldo × Costo Unitario** (valor real del inventario) |

> **Nota sobre Costo Total:** El cálculo es `Saldo × Costo` por lote, sumado por producto. Esto refleja el valor actual del inventario, no el valor histórico de entrada.

#### Hoja Gestion_Pedidos

Análisis de pedidos pendientes. Ver [Pestaña 5](#pestaña-5--gestión-de-pedidos) para descripción de columnas.

#### Hoja Comparativa_Comercial *(solo si se cargó plan comercial)*

Tabla cruzada: cada fila es un producto, cada columna es un mes. Muestra:
- Proyección ML
- Plan comercial
- Diferencia % (semáforo: verde ≤10%, amarillo ≤30%, rojo >30%)

---

## 8. Exportar resultados

### Exportar a Excel

1. Haga clic en **"Exportar Excel"** (disponible después de generar el plan).
2. Seleccione la carpeta y nombre de destino.
3. El archivo se copia y se abre automáticamente en Excel si está disponible.

> Si Excel ya tiene abierto el archivo `orden_produccion_final.xlsx`, el sistema genera automáticamente una copia con la fecha y hora en el nombre (ej: `orden_produccion_final_20260526_1435.xlsx`).

### Exportar a PDF

1. Haga clic en **"Exportar PDF"** para el informe ejecutivo del Dashboard.
2. El PDF incluye:
   - Portada con KPIs principales
   - Gráficas del Dashboard
   - Tabla resumen del Plan Diario (primeras filas)
3. También puede exportar en PDF desde la pestaña **"Tablas del Excel"** para cualquier hoja individual.

---

## 9. Resolución de problemas

### El plan no se genera / Error al iniciar

| Mensaje de error | Causa | Solución |
|------------------|-------|----------|
| "Archivo no encontrado" | La ruta del archivo cambió o fue eliminado | Volver a seleccionar el archivo con "Examinar" |
| "Fecha inválida" | Formato de fecha incorrecto | Usar formato `AAAA-MM-DD`, ej: `2026-04-01` |
| "Fecha inicial debe ser menor que fecha final" | Las fechas están invertidas | Corregir el orden |
| "Falta el archivo [nombre]" | Un archivo obligatorio no fue cargado | Cargar el archivo faltante en la tarjeta correspondiente |
| Error al cargar hoja "Ventas por producto" | El Excel de ventas no tiene esa hoja | Verificar que el archivo es el correcto de INEDDITO |
| Error al cargar hoja "Maestro" | El Excel de pedidos no tiene esa hoja | Verificar que el archivo es el correcto de INEDDITO |

---

### El archivo de resultados no se puede guardar

**Síntoma:** Aparece un archivo con nombre `orden_produccion_final_YYYYMMDD_HHMM.xlsx`.

**Causa:** El archivo `orden_produccion_final.xlsx` está abierto en Microsoft Excel.

**Solución:** Cierre el archivo en Excel y vuelva a ejecutar el plan, o use la copia con timestamp que se generó.

---

### El Dashboard no muestra gráficas

**Causa posible 1:** El plan generó muy pocos datos (período muy corto o sin productos frecuentes).  
**Solución:** Ampliar el período de proyección o revisar que `ventasXproducto.xlsx` tenga suficiente historial.

**Causa posible 2:** El archivo de resultados está corrupto o incompleto.  
**Solución:** Regenerar el plan.

---

### Los valores del inventario parecen incorrectos

**Verifique que está usando ExistBodega**, no entradasInventario. La diferencia:

| Reporte INEDDITO | Columna Saldo | ¿Usar para SEEDPACK? |
|-----------------|---------------|----------------------|
| **ExistBodega** | Saldo actual real | ✅ Sí |
| entradasInventario | Cantidad acumulada histórica | ❌ No |

---

### Las predicciones parecen muy altas o muy bajas

El modelo ML se entrena únicamente con el historial disponible en `ventasXproducto.xlsx`. Si las predicciones son inexactas:

1. **Verificar que el histórico sea suficiente:** Mínimo 2 años de datos, con ventas regulares.
2. **Revisar outliers:** Meses con ventas anómalas (picos extraordinarios por evento puntual) pueden distorsionar el modelo. Considere limpiarlos manualmente del archivo de ventas.
3. **Cambiar el modelo:** El modelo predeterminado es **Random Forest**. Si los datos tienen una tendencia muy lineal, pruebe **Reg. Regularizada**. (Requiere modificar el parámetro `MOD` en `main.py`.)

---

### El proceso es muy lento

El paso más lento es el **Paso 3 (Entrenamiento ML)** y el **Paso 4 (Proyección)**. El tiempo depende de:

- Número de productos frecuentes (a mayor catálogo, más tiempo).
- Longitud del período de proyección (más meses = más tiempo).
- Recursos de la máquina (RAM y CPU).

Tiempos típicos de referencia:

| Catálogo | Período | Tiempo estimado |
|----------|---------|-----------------|
| < 50 productos | 6 meses | 1–3 minutos |
| 50–200 productos | 9 meses | 3–7 minutos |
| > 200 productos | 12 meses | 7–15 minutos |

---

### Necesito regenerar el ejecutable (.exe)

Cuando se actualice el código, ejecute `crear_ejecutable.bat` (doble clic). Este script:

1. Usa el entorno virtual `venv\` del proyecto (asegura las dependencias correctas).
2. Compila todo el código en un único archivo `.exe` en la carpeta `dist\`.
3. Muestra un mensaje de éxito o error al finalizar.

> **Requisito:** El script debe ejecutarse desde la carpeta raíz de SEEDPACK (donde está el archivo `crear_ejecutable.bat`).

---

## 10. Glosario

| Término | Definición |
|---------|-----------|
| **Cluster "Frecuente"** | Producto con promedio ≥ 6 pedidos por año. Recibe proyección ML completa. |
| **Cluster "Esporádico"** | Producto con < 6 pedidos/año. Solo se produce cuando hay un pedido activo. |
| **MRP** (Material Requirements Planning) | Cálculo de cantidad a producir = Demanda − Stock − OPs programadas. |
| **Lote mínimo** | Cantidad mínima de unidades que puede producir una línea de fabricación en una OP. La producción siempre se redondea al múltiplo hacia arriba. |
| **Rolling Forecast** | Método de proyección donde cada mes predicho alimenta los valores históricos del siguiente. Evita acumular errores de un único pronóstico a largo plazo. |
| **Lag** | Valor de un período anterior. `lag1` = mes anterior, `lag12` = mismo mes del año anterior. |
| **Ocasional** | Referencia sin proyección ML regular que aparece en el plan únicamente por un pedido activo del cliente. |
| **Lead Time** | Tiempo de anticipación. El sistema usa 3 días hábiles antes de la fecha de entrega para programar la producción de ocasionales. |
| **ExistBodega** | Reporte de INEDDITO que muestra el saldo actual real de cada lote en bodega. |
| **Pre-cargado** | Archivo que el sistema encuentra automáticamente en `Archivos Historicos\` al iniciar. |
| **KPI** (Key Performance Indicator) | Indicador clave de desempeño. Métricas resumidas que aparecen en el Dashboard. |
| **Costo Total (inventario)** | Valor calculado como Saldo × Costo Unitario, reflejando el valor actual del inventario disponible. |
| **Saldo_OPs** | Hoja que muestra cuánto de cada OP ya ingresó a bodega vs cuánto está pendiente (requiere archivo de Entradas de inventario). |
| **Días hábiles** | Lunes a sábado, excluyendo festivos oficiales colombianos. |

---

*Manual generado para SeedPack Planner v1.0 — Mayo 2026*
