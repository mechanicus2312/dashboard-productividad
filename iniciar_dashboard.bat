@echo off
title Dashboard de Productividad
echo =============================================
echo    Dashboard de Productividad - MSN
echo =============================================
echo.

:: Verificar Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python no esta instalado.
    echo Descargalo en https://www.python.org/downloads/
    pause
    exit /b
)

:: Instalar dependencias si faltan
echo Verificando dependencias...
pip install -r "%~dp0requirements.txt" --quiet

:: Obtener IP local
for /f "tokens=2 delims=:" %%a in ('ipconfig ^| findstr /i "IPv4" ^| findstr /v "169.254"') do (
    set IP=%%a
    goto :found
)
:found
set IP=%IP: =%

echo.
echo Dashboard listo. Comparte este link con tus companeros:
echo.
echo     http://%IP%:8501
echo.
echo Deja esta ventana abierta mientras lo usen.
echo Para cerrar el dashboard presiona Ctrl+C
echo.

:: Iniciar Streamlit
streamlit run "%~dp0dashboard_productividad.py" --server.address 0.0.0.0 --server.port 8501 --server.headless true

pause
