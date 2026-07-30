@echo off
cd /d "%~dp0"
REM ============================================================
REM  Gera o AtualizarDashboard.exe (Windows) - rodar UMA vez,
REM  numa maquina com Python 3.9+ (marque "Add Python to PATH").
REM  Usa MODO SCRIPT: o PyInstaller acha os modulos locais sozinho.
REM ============================================================

echo Conferindo os arquivos necessarios nesta pasta...
set FALTA=0
for %%F in (atualizar.py forecast_accuracy.py dashboard_build.py dashboard_template.html) do (
    if not exist "%%~F" ( echo   *** FALTANDO: %%~F & set FALTA=1 ) else ( echo   ok: %%~F )
)
if "%FALTA%"=="1" (
    echo.
    echo ============================================================
    echo  ERRO: os arquivos marcados FALTANDO nao estao nesta pasta.
    echo  Coloque TODOS na MESMA pasta deste build_exe.bat e rode de novo.
    echo  Pasta atual: %CD%
    echo ============================================================
    pause
    exit /b 1
)
echo Todos presentes.
echo.

echo Instalando dependencias de build...
python -m pip install --upgrade pip
python -m pip install pyinstaller pandas numpy openpyxl pillow
if errorlevel 1 goto erro

echo.
echo Gerando icone a partir do logo (se existir data\logo.png)...
set ICONE=
python -c "from PIL import Image; import os; p='data/logo.png'; (Image.open(p).convert('RGBA').save('data/logo.ico', sizes=[(256,256),(64,64),(32,32),(16,16)])) if os.path.exists(p) else None" 2>NUL
if exist "data\logo.ico" set ICONE=--icon "data\logo.ico"

echo.
echo Empacotando (modo script - encontra forecast_accuracy e dashboard_build automaticamente)...
python -m PyInstaller --clean --noconfirm --onefile --console --name AtualizarDashboard ^
  --add-data "dashboard_template.html;." ^
  --hidden-import openpyxl.cell._writer ^
  --exclude-module matplotlib --exclude-module tkinter --exclude-module scipy ^
  %ICONE% atualizar.py
if errorlevel 1 goto erro

echo.
echo ============================================================
echo  Pronto! O executavel esta em:  dist\AtualizarDashboard.exe
echo  Use SEMPRE esse arquivo recem-gerado (nao uma copia antiga).
echo.
echo  Para usar: coloque o .exe numa pasta com uma subpasta 'data'
echo  contendo os Order NN 2026.xlsx, o Received_itens_2026.xlsx e
echo  (opcional) o logo.png. Rode o .exe -> resultados em 'saida'.
echo ============================================================
pause
exit /b 0

:erro
echo.
echo ERRO durante o build. Verifique se o Python esta instalado e no PATH.
pause
exit /b 1
