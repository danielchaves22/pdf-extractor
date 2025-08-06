@echo off
REM ================================================================
REM Build Manual - PDF para Excel Updater v3.2
REM ================================================================
REM Use este arquivo quando pyinstaller nao estiver no PATH
REM Comando: build_manual.bat
REM ================================================================

echo.
echo ==========================================
echo  PDF para Excel Updater - Build Manual
echo ==========================================
echo.

REM Verifica se dependencias estao instaladas
echo Verificando PyInstaller...
python -c "import PyInstaller; print('PyInstaller:', PyInstaller.__version__)" 2>nul
if %errorlevel% neq 0 (
    echo ERRO: PyInstaller nao instalado
    echo Execute: pip install -r requirements_desktop.txt
    pause
    exit /b 1
)
echo OK: PyInstaller encontrado

REM Verifica e remove pacote pathlib problematico
echo.
echo Verificando conflitos com pathlib...
python -c "import pathlib" 2>nul
if %errorlevel% equ 0 (
    echo AVISO: Pacote pathlib obsoleto detectado
    echo Removendo pathlib incompativel com PyInstaller...
    python -m pip uninstall pathlib -y >nul 2>&1
    if %errorlevel% equ 0 (
        echo OK: pathlib removido com sucesso
    ) else (
        echo AVISO: Nao foi possivel remover pathlib automaticamente
        echo Execute manualmente: python -m pip uninstall pathlib
    )
) else (
    echo OK: Sem conflitos de pathlib
)

REM Verifica arquivos necessarios
echo.
echo Verificando arquivos...
if not exist "desktop_app.py" (
    echo ERRO: desktop_app.py nao encontrado
    pause
    exit /b 1
)
if not exist "pdf_processor_core.py" (
    echo ERRO: pdf_processor_core.py nao encontrado
    pause
    exit /b 1
)
echo OK: Arquivos encontrados

REM Limpa builds anteriores
echo.
echo Limpando builds anteriores...
if exist "build" rmdir /s /q "build" >nul 2>&1
if exist "dist" rmdir /s /q "dist" >nul 2>&1
if exist "__pycache__" rmdir /s /q "__pycache__" >nul 2>&1
if exist "*.spec" del /q "*.spec" >nul 2>&1
echo OK: Limpeza concluida

REM Executa PyInstaller como modulo
echo.
echo Construindo executavel...
echo Isso pode demorar alguns minutos...
echo.

python -m PyInstaller --onefile --windowed --name=PDFExcelUpdater --clean --hidden-import=customtkinter --hidden-import=PIL --hidden-import=PIL._tkinter_finder --hidden-import=pandas --hidden-import=openpyxl --hidden-import=pdfplumber --hidden-import=dotenv --hidden-import=tkinter --hidden-import=tkinter.filedialog --hidden-import=tkinter.messagebox --hidden-import=tkinterdnd2 desktop_app.py

if %errorlevel% neq 0 (
    echo.
    echo ERRO: Falha no build
    echo Verifique se todas as dependencias estao instaladas:
    echo pip install -r requirements_desktop.txt
    pause
    exit /b 1
)

REM Verifica resultado
echo.
echo ==========================================
echo  Verificando resultado...
echo ==========================================
echo.

if exist "dist\PDFExcelUpdater.exe" (
    echo OK: BUILD CONCLUIDO COM SUCESSO!
    echo Executavel: dist\PDFExcelUpdater.exe
    echo.
    echo Para testar: dist\PDFExcelUpdater.exe
    echo.
    
) else (
    echo ERRO: Executavel nao foi criado
    echo.
    echo Verifique se:
    echo 1. Todas as dependencias estao instaladas
    echo 2. PyInstaller funciona: python -m PyInstaller --version
    echo 3. Nao ha antivirus bloqueando
    echo.
    pause
    exit /b 1
)

echo Pressione qualquer tecla para sair...
pause >nul