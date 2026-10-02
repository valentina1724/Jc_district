@echo off
setlocal
cd /d "%~dp0"

echo ============================================================
echo    JC DISTRICT - Actualizar catalogo
echo ============================================================
echo.

where python >nul 2>nul
if errorlevel 1 (
  echo [ERROR] No se encontro Python en este computador.
  echo Instala Python desde python.org y vuelve a intentar.
  echo.
  pause
  exit /b 1
)

python generar_catalogo.py
if errorlevel 1 goto error

echo.
echo --------------------------------------------------------
echo  Publicando cambios en la web...
echo --------------------------------------------------------
echo.

git add productos.xlsx generar_catalogo.py datos.js index.html actualizar.bat
git commit -m "Actualiza catalogo desde productos.xlsx"
if errorlevel 1 echo (No habia cambios nuevos que publicar.)

git push
if errorlevel 1 (
  echo.
  echo [AVISO] No se pudo subir a GitHub.
  echo La pagina sigue igual. Revisa tu internet o publica con GitHub Desktop.
) else (
  echo.
  echo [LISTO] Catalogo actualizado. La web cambia en 1 o 2 minutos.
)

echo.
pause
exit /b 0

:error
echo.
echo --------------------------------------------------------
echo  [ERROR] El Excel tiene problemas. La pagina NO se modifico.
echo --------------------------------------------------------
echo.
echo  Corrige lo que indico arriba en productos.xlsx, guardalo
echo  con Ctrl+S y vuelve a hacer doble clic en este archivo.
echo.
pause
exit /b 1
