@echo off
:: ================================================================
:: PDF para Excel Updater - Setup Automatizado v3.2
:: ================================================================
:: Instala todas as dependências para interface gráfica e CLI
:: ================================================================

echo.
echo ==========================================
echo  PDF para Excel Updater - Setup v3.2
echo ==========================================
echo.

:: Verifica se Python esta instalado
echo [1/7] Verificando instalacao do Python...
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
echo [2/7] Verificando pip...
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
echo [3/7] Atualizando pip...
python -m pip install --upgrade pip --quiet
if %errorlevel% neq 0 (
    echo AVISO: Nao foi possivel atualizar pip, continuando...
) else (
    echo OK: pip atualizado
)

:: Instala dependências CORE (obrigatórias)
echo.
echo [4/7] Instalando dependencias CORE...

echo   - Instalando pandas...
pip install "pandas>=1.5.0" --quiet
if %errorlevel% neq 0 (
    echo ERRO: Falha ao instalar pandas
    goto :error
)

echo   - Instalando openpyxl...
pip install "openpyxl>=3.0.0" --quiet
if %errorlevel% neq 0 (
    echo ERRO: Falha ao instalar openpyxl
    goto :error
)

echo   - Instalando pdfplumber...
pip install "pdfplumber>=0.7.0" --quiet
if %errorlevel% neq 0 (
    echo ERRO: Falha ao instalar pdfplumber
    goto :error
)

echo   - Instalando python-dotenv...
pip install "python-dotenv>=1.0.0" --quiet
if %errorlevel% neq 0 (
    echo ERRO: Falha ao instalar python-dotenv
    goto :error
)

echo OK: Dependencias CORE instaladas

:: Instala dependências GUI (obrigatórias para interface gráfica)
echo.
echo [5/7] Instalando dependencias GUI...

echo   - Instalando customtkinter...
pip install "customtkinter>=5.0.0" --quiet
if %errorlevel% neq 0 (
    echo ERRO: Falha ao instalar customtkinter
    goto :error
)

echo   - Instalando pillow...
pip install "pillow>=9.0.0" --quiet
if %errorlevel% neq 0 (
    echo ERRO: Falha ao instalar pillow
    goto :error
)

echo OK: Dependencias GUI instaladas

:: Instala dependências opcionais (não críticas)
echo.
echo [6/7] Instalando dependencias opcionais...

echo   - Instalando tkinterdnd2 (drag ^& drop)...
pip install "tkinterdnd2>=0.3.0" --quiet
if %errorlevel% neq 0 (
    echo AVISO: tkinterdnd2 nao instalado (drag ^& drop desabilitado)
    echo         Funcionalidade principal nao afetada
) else (
    echo OK: tkinterdnd2 instalado (drag ^& drop habilitado)
)

echo   - Instalando PyInstaller (build executavel)...
pip install "pyinstaller>=5.0.0" --quiet
if %errorlevel% neq 0 (
    echo AVISO: PyInstaller nao instalado (build de executavel desabilitado)
) else (
    echo OK: PyInstaller instalado
)

echo OK: Dependencias opcionais processadas

:: Cria arquivo .env exemplo se nao existir
echo.
echo [7/7] Configurando arquivo .env...
if not exist ".env" (
    echo # Configuracao do PDF para Excel Updater v3.2 > .env
    echo # ===================================================== >> .env
    echo # DIRETORIO DE TRABALHO (obrigatorio) >> .env
    echo # >> .env
    echo # O MODELO_DIR e o diretorio de trabalho onde devem estar: >> .env
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

:: Verifica se scripts principais existem
if not exist "desktop_app.py" (
    echo AVISO: desktop_app.py nao encontrado (interface grafica indisponivel)
)
if not exist "pdf_to_excel_updater.py" (
    echo AVISO: pdf_to_excel_updater.py nao encontrado (linha de comando indisponivel)
)
if not exist "pdf_processor_core.py" (
    echo AVISO: pdf_processor_core.py nao encontrado (modulo core ausente)
)

:: Testa instalacao completa
echo.
echo ==========================================
echo  Teste Completo de Instalacao
echo ==========================================
echo.

echo Testando dependencias CORE...
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
echo Testando dependencias GUI...
python -c "import customtkinter; print('  customtkinter:', customtkinter.__version__)" 2>nul
if %errorlevel% neq 0 (
    echo ERRO: customtkinter nao pode ser importado
    goto :error
)

python -c "import PIL; print('  pillow: OK')" 2>nul
if %errorlevel% neq 0 (
    echo ERRO: pillow nao pode ser importado
    goto :error
)

echo.
echo Testando dependencias opcionais...
python -c "import tkinterdnd2; print('  tkinterdnd2: OK (drag ^& drop disponivel)')" 2>nul
if %errorlevel% neq 0 (
    echo AVISO: tkinterdnd2 nao disponivel (drag ^& drop desabilitado)
)

python -c "import PyInstaller; print('  pyinstaller: OK')" 2>nul
if %errorlevel% neq 0 (
    echo AVISO: PyInstaller nao disponivel (build de executavel desabilitado)
)

echo.
echo ==========================================
echo  INSTALACAO CONCLUIDA COM SUCESSO!
echo ==========================================
echo.
echo PDF para Excel Updater v3.2 - Como usar:
echo.
echo  === INTERFACE GRAFICA (Recomendado) ===
echo  python desktop_app.py
echo.
echo  Funcionalidades:
echo  - Interface moderna com abas organizadas
echo  - Drag ^& Drop de arquivos PDF
echo  - Historico persistido de processamentos
echo  - Configuracoes automaticas
echo  - Popup de progresso em tempo real
echo.
echo  === LINHA DE COMANDO (Alternativa) ===
echo  python pdf_to_excel_updater.py               # Seletor de arquivo
echo  python pdf_to_excel_updater.py arquivo.pdf   # Arquivo especifico
echo.
echo Configuracao obrigatoria:
echo  1. Edite o arquivo .env para configurar MODELO_DIR
echo  2. Coloque MODELO.xlsm no diretorio de trabalho
echo  3. Coloque arquivos PDF no mesmo diretorio
echo.
echo Estrutura esperada:
echo  DIRETORIO_TRABALHO/
echo  ├── MODELO.xlsm          ← Planilha modelo
echo  ├── arquivo.pdf          ← PDF a processar
echo  └── DADOS/               ← Resultados (criado automaticamente)
echo      └── NOME_PESSOA.xlsm ← Arquivo final
echo.
echo Para mais informacoes: README.md
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
echo     pip install -r requirements.txt
echo  4. Para apenas linha de comando (sem GUI):
echo     pip install pandas openpyxl pdfplumber python-dotenv
echo.
echo Se o problema persistir:
echo  - Verifique se Python e pip estao atualizados
echo  - Tente criar um ambiente virtual (venv)
echo  - Consulte README.md para solucao de problemas
echo.
pause
exit /b 1

:end
echo Pressione qualquer tecla para continuar...
pause >nul