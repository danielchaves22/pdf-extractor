@echo off
:: ================================================================
:: PDF para Excel Updater - Setup Automatizado v2.2
:: ================================================================
:: Instala todas as dependencias necessarias para um novo ambiente
:: ================================================================

echo.
echo ==========================================
echo  PDF para Excel Updater - Setup v2.2
echo ==========================================
echo.

:: Verifica se Python esta instalado
echo [1/6] Verificando instalacao do Python...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERRO: Python nao esta instalado ou nao esta no PATH
    echo.
    echo Baixe e instale Python em: https://www.python.org/downloads/
    echo Certifique-se de marcar "Add Python to PATH" durante a instalacao
    pause
    exit /b 1
)

for /f "tokens=2" %%i in ('python --version 2^>^&1') do set PYTHON_VERSION=%%i
echo OK: Python %PYTHON_VERSION% encontrado

:: Verifica se pip esta funcionando
echo.
echo [2/6] Verificando pip...
pip --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERRO: pip nao esta disponivel
    echo Reinstale Python com pip incluido
    pause
    exit /b 1
)

for /f "tokens=2" %%i in ('pip --version 2^>^&1') do set PIP_VERSION=%%i
echo OK: pip %PIP_VERSION% funcionando

:: Atualiza pip para versao mais recente
echo.
echo [3/6] Atualizando pip...
python -m pip install --upgrade pip --quiet
if %errorlevel% neq 0 (
    echo AVISO: Nao foi possivel atualizar pip, continuando...
) else (
    echo OK: pip atualizado
)

:: Instala dependencias principais
echo.
echo [4/6] Instalando dependencias principais...

echo   - Instalando pandas...
pip install pandas>=1.5.0 --quiet
if %errorlevel% neq 0 (
    echo ERRO: Falha ao instalar pandas
    goto :error
)

echo   - Instalando openpyxl...
pip install openpyxl>=3.0.0 --quiet
if %errorlevel% neq 0 (
    echo ERRO: Falha ao instalar openpyxl
    goto :error
)

echo   - Instalando pdfplumber...
pip install pdfplumber>=0.7.0 --quiet
if %errorlevel% neq 0 (
    echo ERRO: Falha ao instalar pdfplumber
    goto :error
)

echo   - Instalando python-dotenv...
pip install python-dotenv>=1.0.0 --quiet
if %errorlevel% neq 0 (
    echo ERRO: Falha ao instalar python-dotenv
    goto :error
)

echo   - Instalando PyInstaller...
pip install pyinstaller>=5.0.0 --quiet
if %errorlevel% neq 0 (
    echo ERRO: Falha ao instalar PyInstaller
    goto :error
)

echo OK: Todas as dependencias instaladas com sucesso

:: Cria arquivo .env exemplo se nao existir
echo.
echo [5/6] Configurando arquivo .env...
if not exist ".env" (
    echo # Configuracao do PDF para Excel Updater v3.0 > .env
    echo # ===================================================== >> .env
    echo # DIRETORIO DE TRABALHO (obrigatorio) >> .env
    echo # >> .env
    echo # O MODELO_DIR agora e o diretorio de trabalho onde devem estar: >> .env
    echo # - MODELO.xlsm (planilha modelo) >> .env
    echo # - arquivo.pdf (PDF a processar) >> .env
    echo # - DADOS/ (pasta criada automaticamente para resultados) >> .env
    echo. >> .env
    echo # Use barras normais / ou escape as barras invertidas \\ >> .env
    echo MODELO_DIR=C:/trabalho/folhas_pagamento >> .env
    echo # ou >> .env
    echo # MODELO_DIR=C:\\trabalho\\folhas_pagamento >> .env
    echo. >> .env
    echo # Exemplos de configuracao: >> .env
    echo # MODELO_DIR=D:/empresa/processamento_folhas >> .env
    echo # MODELO_DIR=./diretorio_trabalho >> .env
    echo # MODELO_DIR=../folhas_compartilhadas >> .env
    
    echo OK: Arquivo .env criado com configuracao exemplo
    echo     Edite o arquivo .env para configurar o diretorio de trabalho
) else (
    echo OK: Arquivo .env ja existe
)

:: Verifica se script principal existe
echo.
echo [6/6] Verificando arquivos...
if not exist "pdf_to_excel_updater.py" (
    echo AVISO: pdf_to_excel_updater.py nao encontrado na pasta atual
    echo        Certifique-se de que este arquivo esteja na mesma pasta do setup.bat
) else (
    echo OK: pdf_to_excel_updater.py encontrado
)

:: Testa instalacao
echo.
echo ==========================================
echo  Teste de Instalacao
echo ==========================================
echo.

echo Testando imports das bibliotecas...
python -c "import pandas; print('  pandas:', pandas.__version__)" 2>nul
if %errorlevel% neq 0 (
    echo ERRO: pandas nao pode ser importado
    goto :error
)

python -c "import openpyxl; print('  openpyxl:', openpyxl.__version__)" 2>nul
if %errorlevel% neq 0 (
    echo ERRO: openpyxl nao pode ser importado
    goto :error
)

python -c "import pdfplumber; print('  pdfplumber:', pdfplumber.__version__)" 2>nul
if %errorlevel% neq 0 (
    echo ERRO: pdfplumber nao pode ser importado
    goto :error
)

python -c "import dotenv; print('  python-dotenv: OK')" 2>nul
if %errorlevel% neq 0 (
    echo ERRO: python-dotenv nao pode ser importado
    goto :error
)

echo.
echo ==========================================
echo  INSTALACAO CONCLUIDA COM SUCESSO!
echo ==========================================
echo.
echo Como usar (v3.0 - Diretorio de Trabalho):
echo.
echo  1. Configure o diretorio de trabalho no .env:
echo     MODELO_DIR=C:/seu/diretorio/de/trabalho
echo.
echo  2. Coloque no diretorio de trabalho:
echo     - MODELO.xlsm (planilha modelo)
echo     - arquivo.pdf (PDF a processar)
echo.
echo  3. Execute de qualquer local:
echo     python pdf_to_excel_updater.py arquivo.pdf
echo.
echo  4. O resultado aparecera em:
echo     DIRETORIO_TRABALHO/DADOS/arquivo.xlsm
echo.
echo Configuracao obrigatoria:
echo  - Edite o arquivo .env para configurar MODELO_DIR
echo  - Coloque MODELO.xlsm no diretorio de trabalho
echo.
echo Para mais informacoes use: python pdf_to_excel_updater.py --help
echo.
goto :end

:error
echo.
echo ==========================================
echo  ERRO NA INSTALACAO
echo ==========================================
echo.
echo Algumas dependencias nao puderam ser instaladas.
echo.
echo Solucoes:
echo  1. Execute como Administrador
echo  2. Verifique conexao com internet
echo  3. Tente instalar manualmente:
echo     pip install pandas openpyxl pdfplumber python-dotenv
echo.
pause
exit /b 1

:end
echo Pressione qualquer tecla para continuar...
pause >nul