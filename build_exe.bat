@echo off
cd /d "%~dp0"
REM ============================================================
REM  Gera o AtualizarDashboard.exe (Windows) - rodar UMA vez,
REM  numa maquina que tenha Python 3.9+ instalado.
REM  Depois de gerado, o .exe roda sem Python.
REM ============================================================

echo Conferindo os arquivos necessarios nesta pasta...
set FALTA=0
for %%F in (atualizar.py forecast_accuracy.py dashboard_build.py dashboard_template.html atualizar.spec) do (
    if not exist "%%~F" (
        echo   *** FALTANDO: %%~F
        set FALTA=1
    ) else (
        echo   ok: %%~F
    )
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
echo Todos presentes. Continuando...
echo.

echo Instalando dependencias de build...
python -m pip install --upgrade pip
python -m pip install pyinstaller pandas numpy openpyxl pillow
if errorlevel 1 goto erro

echo.
echo Gerando icone a partir do logo (se existir data\logo.png)...
python -c "from PIL import Image; import os; p='data/logo.png'; (Image.open(p).convert('RGBA').save('data/logo.ico', sizes=[(256,256),(64,64),(32,32),(16,16)])) if os.path.exists(p) else None"

echo.
echo Empacotando com PyInstaller...
python -m PyInstaller --clean --noconfirm atualizar.spec
if errorlevel 1 goto erro

echo.
echo ============================================================
echo  Pronto!  O executavel esta em:  dist\AtualizarDashboard.exe
echo.
echo  Para usar: coloque o .exe numa pasta com uma subpasta 'data'
echo  contendo os Order_NN_2026.xlsx, o Received_itens_2026.xlsx
echo  e (opcional) o logo.png. Rode o .exe e pegue os resultados
echo  na subpasta 'saida'.
echo ============================================================
pause
exit /b 0

:erro
echo.
echo ERRO durante o build. Verifique se o Python esta instalado e no PATH.
pause
exit /b 1
