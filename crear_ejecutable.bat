@echo off
echo ============================================================
echo  SeedPack Planner - Generador de ejecutable
echo ============================================================
echo.

REM ── 1. Verificar que Python este instalado en el sistema ──────────────────
python --version >nul 2>&1
if errorlevel 1 (
    echo  Python no encontrado. Intentando instalar automaticamente con winget...
    echo.
    winget install --id Python.Python.3.12 --source winget --silent --accept-package-agreements --accept-source-agreements
    if errorlevel 1 (
        echo.
        echo  No se pudo instalar Python automaticamente.
        echo  Por favor instala Python manualmente desde:
        echo    https://www.python.org/downloads/
        echo  Asegurate de marcar "Add Python to PATH" durante la instalacion.
        pause
        exit /b 1
    )
    echo.
    echo  Python instalado. Reinicia este .bat para continuar.
    pause
    exit /b 0
)

REM ── 2. Crear el entorno virtual si no existe ──────────────────────────────
if not exist "venv\Scripts\python.exe" (
    echo  Creando entorno virtual...
    python -m venv venv
    if errorlevel 1 (
        echo  ERROR: No se pudo crear el entorno virtual.
        pause
        exit /b 1
    )
    echo  Entorno virtual creado.
    echo.
)

set PYTHON=venv\Scripts\python.exe

REM ── 3. Instalar dependencias desde requirements.txt ───────────────────────
echo  Instalando dependencias (esto puede tardar varios minutos)...
%PYTHON% -m pip install --upgrade pip --quiet
%PYTHON% -m pip install -r requirements.txt --quiet
if errorlevel 1 (
    echo  ERROR: Fallo la instalacion de dependencias.
    pause
    exit /b 1
)
echo  Dependencias instaladas.
echo.

REM ── 4. Instalar PyInstaller si no esta ───────────────────────────────────
%PYTHON% -m PyInstaller --version >nul 2>&1
if errorlevel 1 (
    echo  Instalando PyInstaller...
    %PYTHON% -m pip install pyinstaller --quiet
    echo.
)

REM ── 5. Compilar el ejecutable ─────────────────────────────────────────────
echo  Compilando SeedPack Planner.exe ...
echo.

%PYTHON% -m PyInstaller ^
  --onefile ^
  --windowed ^
  --name "SeedPack Planner" ^
  --paths "codigo" ^
  --paths "codigo\pasos" ^
  --hidden-import=main ^
  --hidden-import=paso1_limpieza ^
  --hidden-import=paso2_clusterizacion ^
  --hidden-import=paso3_modelos ^
  --hidden-import=paso4_proyeccion ^
  --hidden-import=paso5_orden_produccion ^
  --hidden-import=sklearn ^
  --hidden-import=sklearn.ensemble ^
  --hidden-import=sklearn.linear_model ^
  --hidden-import=sklearn.neural_network ^
  --hidden-import=sklearn.preprocessing ^
  --hidden-import=sklearn.metrics ^
  --hidden-import=xgboost ^
  --hidden-import=openpyxl ^
  --hidden-import=pandas ^
  --hidden-import=numpy ^
  --hidden-import=matplotlib ^
  --hidden-import=matplotlib.backends.backend_pdf ^
  --hidden-import=matplotlib.backends.backend_tkagg ^
  --hidden-import=holidays ^
  --collect-submodules=sklearn ^
  --collect-all=xgboost ^
  --collect-all=holidays ^
  codigo\app.py

echo.
if exist "dist\SeedPack Planner.exe" (
    echo  ============================================================
    echo   Ejecutable generado exitosamente.
    echo   Copiando a app\ ...
    echo  ============================================================
    echo.
    copy /Y "dist\SeedPack Planner.exe" "app\SeedPack Planner.exe" >nul
    if errorlevel 1 (
        echo  ADVERTENCIA: No se pudo copiar a app\. Copia manual desde dist\.
    ) else (
        echo  Listo: app\SeedPack Planner.exe actualizado.
    )
    echo.
) else (
    echo  ERROR: No se genero el ejecutable. Revisa los mensajes arriba.
)

pause
