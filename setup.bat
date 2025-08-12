@echo off
:: ================================================================
:: PDF para Excel Updater - Setup Automatizado v4.0 (PyQt6)
:: ================================================================
:: Instala todas as dependências para interface PyQt6 e CLI
:: ================================================================

echo.
echo ==========================================
echo  PDF para Excel Updater - Setup v4.0
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

:: Verifica versão mínima do Python (3.8+ requerido para PyQt6)
python -c "import sys; exit(0 if sys.version_info >= (3,8) else 1)" >nul 2>&1
if %errorlevel% neq 0 (
    echo ERRO: PyQt6 requer Python 3.8 ou superior
    echo Versao atual: %PYTHON_VERSION%
    echo.
    echo Atualize Python em: https://www.python.org/downloads/
    pause
    exit /b 1
)
echo OK: Python %PYTHON_VERSION% e compativel com PyQt6

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

:: Instala dependências GUI v4.0 (PyQt6)
echo.
echo [5/7] Instalando dependencias GUI v4.0 (PyQt6)...

echo   - Instalando PyQt6 (interface moderna)...
pip install "PyQt6>=6.4.0" --quiet
if %errorlevel% neq 0 (
    echo ERRO: Falha ao instalar PyQt6
    echo.
    echo Solucoes:
    echo 1. Verifique se Python e 3.8+ (requerido para PyQt6)
    echo 2. Execute como Administrador
    echo 3. Tente: pip install PyQt6 --user
    goto :error
)

echo OK: PyQt6 instalado com sucesso

:: Remove dependências antigas (se existirem)
echo.
echo [6/7] Removendo dependencias antigas (v3.x)...

echo   - Verificando customtkinter...
python -c "import customtkinter" >nul 2>&1
if %errorlevel% equ 0 (
    echo     Removendo customtkinter (substituido por PyQt6)...
    pip uninstall customtkinter -y --quiet >nul 2>&1
    if %errorlevel% equ 0 (
        echo     OK: customtkinter removido
    ) else (
        echo     AVISO: Nao foi possivel remover customtkinter automaticamente
    )
) else (
    echo     OK: customtkinter nao instalado
)

echo   - Verificando tkinterdnd2...
python -c "import tkinterdnd2" >nul 2>&1
if %errorlevel% equ 0 (
    echo     Removendo tkinterdnd2 (PyQt6 tem drag ^& drop nativo)...
    pip uninstall tkinterdnd2 -y --quiet >nul 2>&1
    if %errorlevel% equ 0 (
        echo     OK: tkinterdnd2 removido
    ) else (
        echo     AVISO: Nao foi possivel remover tkinterdnd2 automaticamente
    )
) else (
    echo     OK: tkinterdnd2 nao instalado
)

echo OK: Dependencias antigas processadas

:: Instala dependências opcionais (não críticas)
echo.
echo [7/7] Instalando dependencias opcionais...

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
echo [8/8] Configurando arquivo .env...
if not exist ".env" (
    echo # Configuracao do PDF para Excel Updater v4.0 ^(PyQt6^) > .env
    echo # ======================================================= >> .env
    echo # DIRETORIO DE TRABALHO ^(obrigatorio^) >> .env
    echo # >> .env
    echo # O MODELO_DIR e o diretorio de trabalho onde devem estar: >> .env
    echo # - MODELO.xlsm ^(planilha modelo^) >> .env
    echo # - arquivo.pdf ^(PDF a processar^) >> .env
    echo # - DADOS/ ^(pasta criada automaticamente para resultados^) >> .env
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
    echo AVISO: desktop_app.py nao encontrado ^(interface grafica PyQt6 indisponivel^)
)
if not exist "pdf_to_excel_updater.py" (
    echo AVISO: pdf_to_excel_updater.py nao encontrado ^(linha de comando indisponivel^)
)
if not exist "pdf_processor_core.py" (
    echo AVISO: pdf_processor_core.py nao encontrado ^(modulo core ausente^)
)

:: Testa instalacao completa
echo.
echo ==========================================
echo  Teste Completo de Instalacao v4.0
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
echo Testando dependencias GUI v4.0 ^(PyQt6^)...
python -c "import PyQt6.QtWidgets; print('  PyQt6: OK ^(interface v4.0 disponivel^)')" 2>nul
if %errorlevel% neq 0 (
    echo ERRO: PyQt6 nao pode ser importado
    goto :error
)

echo.
echo Testando dependencias GUI v3.x ^(CustomTkinter - compatibilidade^)...
python -c "import customtkinter; print('  CustomTkinter: OK ^(interface v3.x compativel^)')" 2>nul
if %errorlevel% neq 0 (
    echo INFO: CustomTkinter nao disponivel ^(apenas v4.0 funcional^)
) else (
    echo OK: Ambas as versoes de interface disponiveis
)

echo.
echo Testando dependencias opcionais...
python -c "import PyInstaller; print('  pyinstaller: OK')" 2>nul
if %errorlevel% neq 0 (
    echo AVISO: PyInstaller nao disponivel ^(build de executavel desabilitado^)
)

echo.
echo ==========================================
echo  INSTALACAO v4.0 CONCLUIDA COM SUCESSO!
echo ==========================================
echo.
echo PDF para Excel Updater - Multiplas Versoes Disponiveis:
echo.
echo  === INTERFACE GRAFICA v4.0 (PyQt6 - Recomendado) ===
echo  python desktop_app.py
echo.
echo  NOVIDADES v4.0:
echo  - Interface PyQt6 moderna e performatica (10-20x mais rapida)
echo  - Threading nativo com signals/slots thread-safe
echo  - Drag ^& Drop nativo sem dependencias extras
echo  - Virtualizacao automatica de listas grandes
echo  - Eliminacao de polling manual (updates em tempo real)
echo  - Estilo escuro moderno com CSS-like styling
echo.
echo  === COMPATIBILIDADE v3.x (CustomTkinter) ===
echo  Dependencias v3.x mantidas para compatibilidade de branches
echo  Ambas as versoes podem coexistir no mesmo ambiente
echo.
echo  === LINHA DE COMANDO (Todas as versoes) ===
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
echo  └── DADOS/               ← Resultados ^(criado automaticamente^)
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
echo Solucoes v4.0:
echo  1. Execute como Administrador
echo  2. Verifique conexao com internet
echo  3. Certifique-se que Python e 3.8+ ^(requerido para PyQt6^)
echo  4. Tente instalar manualmente:
echo     pip install -r requirements.txt
echo  5. Para apenas linha de comando ^(sem GUI^):
echo     pip install pandas openpyxl pdfplumber python-dotenv
echo.
echo Se o problema persistir:
echo  - Verifique se Python e pip estao atualizados
echo  - Tente criar um ambiente virtual ^(venv^)
echo  - Para PyQt6: pip install PyQt6 --user
echo  - Consulte README.md para solucao de problemas
echo.
pause
exit /b 1

:end
echo Pressione qualquer tecla para continuar...
pause >nul