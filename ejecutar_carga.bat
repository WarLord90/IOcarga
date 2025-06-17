@echo off
setlocal

REM Guardamos la ruta actual
set DIR=%~dp0

REM Cargamos la imagen desde el .tar si no está cargada
docker image inspect cargaio-py312 >nul 2>&1
if errorlevel 1 (
    echo Cargando imagen Docker...
    docker load -i "%DIR%cargaio.tar"
)

REM Ejecutamos el contenedor (sin mostrar nueva consola)
start "" /b docker run --rm ^
  -v "%DIR%.env:/app/.env" ^
  -v "%DIR%carga.xlsx:/app/carga.xlsx" ^
  -v "%DIR%output:/app/output" ^
  cargaio-py312

exit
