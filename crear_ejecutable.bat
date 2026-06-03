@echo off
echo ============================================================
echo  SeedPack Planner - Generador de ejecutable
echo ============================================================
echo.

REM Usar el Python del entorno virtual del proyecto
set PYTHON=venv\Scripts\python.exe

if not exist "%PYTHON%" (
    echo  ERROR: No se encontro el entorno virtual en venv\
    echo  Ejecuta este .bat desde la carpeta raiz de SEEDPACK
    pause
    exit /b 1
)

REM Verificar que pyinstaller este instalado en el venv
%PYTHON% -m PyInstaller --version >nul 2>&1
if errorlevel 1 (
    echo  PyInstaller no encontrado en el venv. Instalando...
    %PYTHON% -m pip install pyinstaller
    echo.
)

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
    echo   Ejecutable generado exitosamente:
    echo   dist\SeedPack Planner.exe
    echo  ============================================================
    echo.
    echo  ENTREGA AL CLIENTE: Copia el .exe a la carpeta app\
    echo  junto con las subcarpetas Archivos Historicos\,
    echo  Resultado Final\, datos_pipeline\ y graficas\.
    echo.
) else (
    echo  ERROR: No se genero el ejecutable. Revisa los mensajes arriba.
)

pause
