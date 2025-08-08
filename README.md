# 🔧 PDF para Excel Updater v3.2

Aplicação Python para extrair dados de PDFs de folha de pagamento e preencher planilhas Excel com **interface gráfica moderna** e **sistema de histórico persistido**, preservando formatação, fórmulas e macros VBA.

## ✨ Funcionalidades Principais v3.2

- ✅ **Interface gráfica moderna** com CustomTkinter e abas organizadas
- ✅ **Sistema de histórico persistido** entre sessões
- ✅ **Processamento de FOLHA NORMAL e 13º SALÁRIO** com regras específicas
- ✅ **Diretório de trabalho** configurado via .env (MODELO_DIR)
- ✅ **Drag & Drop de arquivos PDF** (opcional com tkinterdnd2)
- ✅ **Detecta nome da pessoa** no PDF para nomear arquivo automaticamente
- ✅ **Preserva macros VBA** (.xlsm) e formatação completa
- ✅ **Processamento offline** (sem internet)
- ✅ **Logs detalhados** com popup de progresso em tempo real
- ✅ **Fallback inteligente** para códigos de produção

## 📊 Mapeamento de Dados Completo

### 🔵 FOLHA NORMAL (Linhas 1-65)

| PDF Código | Descrição | Excel Coluna | Fonte | Regras Especiais |
|------------|-----------|--------------|-------|------------------|
| `09090301` | SALARIO CONTRIB INSS | **B** (REMUNERAÇÃO RECEBIDA) | Último número | - |
| `01003601` | PREMIO PROD. MENSAL | **X** (PRODUÇÃO) | Penúltimo número | ⚡ **Fallback**: Se vazio, usa último número |
| `01007301` | HORAS EXT.100%-180 | **Y** (INDICE HE 100%) | Penúltimo número | 🕐 Suporta formato horas |
| `01003501` | HORAS EXT.75%-180 | **AA** (INDICE HE 75%) | Penúltimo número | 🕐 Suporta formato horas |
| `02007501` | DIFER.PROV. HORAS EXTRAS 75% | **AA** (INDICE HE 75%) | Penúltimo número | 🕐 Código alternativo |
| `01009001` | ADIC.NOT.25%-180 | **AC** (INDICE ADC. NOT.) | Penúltimo número | 🕐 Suporta formato horas |

### 🔴 13º SALÁRIO (Linhas 67+)

| PDF Código | Descrição | Excel Coluna | Fonte | Regras Especiais |
|------------|-----------|--------------|-------|------------------|
| `09090301` | SALARIO CONTRIB INSS | **B** (REMUNERAÇÃO RECEBIDA) | Último número | **Prioridade 1** |
| `09090101` | REMUNERACAO BRUTA | **B** (REMUNERAÇÃO RECEBIDA) | Último número | **Fallback** se 09090301 não encontrado |

### 🕐 Formato de Horas
- **Detecção automática**: `06:34` → `06,34`
- **Aplicável a**: HE 100%, HE 75%, ADIC. NOT.
- **Conversão**: Substitui `:` por `,` automaticamente

## 🚀 Instalação e Configuração

### Opção 1: Automática (Windows) - Recomendada
```batch
# Execute o setup que instala tudo automaticamente
setup.bat
```

### Opção 2: Manual
```bash
# 1. Verifique Python 3.7+
python --version

# 2. Instale todas as dependências
pip install -r requirements.txt
```

### Pré-requisitos
- **Python 3.7+**
- **pip** (incluído no Python)

## ⚙️ Configuração Obrigatória (.env)

Crie arquivo `.env` na pasta do script:

```bash
# Diretório de trabalho (obrigatório)
MODELO_DIR=C:/trabalho/folhas_pagamento
```

### Estrutura do Diretório de Trabalho:
```
C:/trabalho/folhas_pagamento/     ← MODELO_DIR
├── MODELO.xlsm                   ← Planilha modelo (obrigatório)
├── relatorio.pdf                 ← PDF a processar
└── DADOS/                        ← Criado automaticamente
    └── NOME_DA_PESSOA.xlsm       ← Resultado processado
```

## 💻 Como Usar

### 🎯 Interface Gráfica (Recomendado) - v3.2
```bash
# Abre aplicação gráfica moderna com abas
python desktop_app.py
```

**Funcionalidades da Interface:**
- **📄 Aba Processamento**: Configuração, seleção de PDF e processamento
- **📊 Aba Histórico**: Histórico persistido de todos os processamentos
- **⚙️ Aba Configurações**: Planilha personalizada e logs detalhados
- **🎯 Drag & Drop**: Arraste PDFs diretamente na interface
- **📝 Logs em tempo real**: Popup com progresso e logs detalhados
- **💾 Persistência**: Configurações e histórico salvos entre sessões

### 🎯 Linha de Comando (Alternativa)
```bash
# Abre seletor de arquivo
python pdf_to_excel_updater.py

# Processa arquivo específico
python pdf_to_excel_updater.py "relatorio.pdf"

# Com planilha específica
python pdf_to_excel_updater.py "relatorio.pdf" -s "DADOS"

# Modo verboso (diagnóstico)
python pdf_to_excel_updater.py "relatorio.pdf" -v
```

### 📋 Exemplo Completo de Uso

```bash
# 1. Configure .env
echo "MODELO_DIR=C:/trabalho/folhas" > .env

# 2. Estrutura no diretório:
# C:/trabalho/folhas/
# ├── MODELO.xlsm
# └── EdsonGoulart-Jan2025.pdf

# 3. Execute interface gráfica
python desktop_app.py

# 4. Na interface:
# - Configure diretório: C:/trabalho/folhas
# - Arraste EdsonGoulart-Jan2025.pdf
# - Clique "Processar PDF"

# 5. Resultado automático:
# C:/trabalho/folhas/DADOS/EDSON GOULART.xlsm ✓
```

## 📱 Interface Gráfica v3.2 - Recursos

### 🎨 Design Moderno
- **Tema escuro** com CustomTkinter
- **Layout responsivo** com abas organizadas
- **Ícones intuitivos** e cores de status
- **Animações suaves** e feedback visual

### 📊 Sistema de Histórico Persistido
- **Histórico automático** de todos os processamentos
- **Dados persistidos** entre sessões da aplicação
- **Detalhes completos**: logs, resultados, timestamps
- **Abertura direta** dos arquivos processados
- **Limpeza de histórico** com confirmação

### 🎯 Processamento Inteligente
- **Detecção automática** do nome da pessoa no PDF
- **Validação em tempo real** do diretório de trabalho
- **Lista automática** de PDFs disponíveis
- **Popup de progresso** com logs em tempo real

### ⚙️ Configurações Avançadas
- **Planilha personalizada** (padrão: "LEVANTAMENTO DADOS")
- **Modo verboso** para diagnóstico detalhado
- **Reset de configurações** para valores padrão
- **Persistência automática** de todas as configurações

## 📈 Resultado Esperado

### ✅ Interface Gráfica - Sucesso:
- Popup de progresso em tempo real
- Mensagem de sucesso com estatísticas completas
- Entrada automática no histórico
- Opção de abrir arquivo criado

### ⚠️ Interface Gráfica - Falha:
- Popup com logs detalhados do erro
- Navegação automática para aba apropriada
- Entrada no histórico com detalhes da falha
- Sugestões de correção contextuais

### 🖥️ Linha de Comando:
```
✅ Processamento concluído: 54 períodos processados
   📄 FOLHA NORMAL: 45 períodos
   💰 13 SALÁRIO: 9 períodos

👤 Nome detectado: EDSON GOULART

💾 Arquivo criado: DADOS/EDSON GOULART.xlsm
```

## 🔧 Funcionalidades Avançadas v3.2

### 🎯 Processamento Dual (FOLHA NORMAL + 13º SALÁRIO)
- **FOLHA NORMAL**: Linhas 1-65 com códigos específicos
- **13º SALÁRIO**: Linhas 67+ com fallback inteligente entre códigos
- **Filtro automático** por "Tipo da folha"
- **Categorização inteligente** de páginas PDF

### 📊 Sistema de Fallback Robusto
- **Produção (01003601)**: ÍNDICE → VALOR se vazio
- **13º Salário**: 09090301 → 09090101 se primeiro não encontrado
- **Formato horas**: Detecção automática e conversão

### 💾 Persistência e Histórico
- **config.json**: Configurações da aplicação
- **history.json**: Histórico completo de processamentos
- **Sessões múltiplas**: Mantém histórico entre reinicializações
- **Limpeza automática**: Limita histórico às últimas 10 sessões

## 🏗️ Build de Executável (Opcional)

### Windows - PyInstaller:
```batch
# Build automático
build.bat

# Resultado: dist/PDFExcelUpdater.exe
```

### Manual:
```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name=PDFExcelUpdater desktop_app.py
```

## 🐛 Solução de Problemas v3.2

### Interface Gráfica não abre:
```bash
# Verifique dependências GUI
pip install customtkinter pillow

# Opcional para drag & drop
pip install tkinterdnd2
```

### Erro: "pdf_processor_core.py não encontrado":
```bash
# Certifique-se de que ambos arquivos estão na mesma pasta:
# - desktop_app.py
# - pdf_processor_core.py
```

### Configuração não persiste:
```bash
# Verifique permissões de escrita na pasta do script
# Os arquivos config.json e history.json devem ser criáveis
```

### Drag & Drop não funciona:
```bash
# Instale dependência opcional
pip install tkinterdnd2

# Ou use botão "Selecionar" como alternativa
```

## 📁 Arquivos do Projeto v3.2

```
pdf-updater/
├── desktop_app.py              # ← Interface gráfica moderna v3.2
├── pdf_processor_core.py       # ← Lógica central de processamento
├── pdf_to_excel_updater.py     # ← Interface linha de comando
├── requirements.txt            # ← Todas as dependências consolidadas
├── setup.bat                   # ← Instalação automática
├── build.bat                   # ← Build do executável
├── .env                        # ← Configuração (MODELO_DIR)
├── config.json                 # ← Configurações persistidas (criado automaticamente)
├── history.json                # ← Histórico persistido (criado automaticamente)
└── README.md                   # ← Esta documentação
```

## 🔒 Preservação e Segurança

### ✅ O que é Preservado
- **Macros VBA** (.xlsm) - Preservação completa
- **Fórmulas existentes** - Mantidas intactas
- **Formatação** - Cores, bordas, fontes preservadas
- **Estrutura** - Layout da planilha mantido
- **Modelo original** - Nunca é alterado

### ✅ O que é Preenchido
- **Apenas** colunas B, X, Y, AA, AC
- **Apenas** células vazias (não sobrescreve)
- **Apenas** dados extraídos com sucesso do PDF

## 🎯 Casos de Uso v3.2

### **Uso Individual com Interface Gráfica:**
```bash
# 1. Configure uma vez via interface
python desktop_app.py

# 2. Use drag & drop para processar PDFs
# 3. Consulte histórico de processamentos
# 4. Todas as configurações são salvas automaticamente
```

### **Uso Corporativo:**
```bash
# .env corporativo em rede
MODELO_DIR=//servidor/rh/processamento_folhas

# Interface gráfica funcionará para qualquer usuário
python desktop_app.py
```

### **Automação por Linha de Comando:**
```bash
# Para scripts automatizados
python pdf_to_excel_updater.py "arquivo.pdf" -v
```

## 📞 Comandos de Diagnóstico

### Teste Completo de Dependências:
```bash
# Testa todas as dependências
python -c "
import pandas, openpyxl, pdfplumber, dotenv, customtkinter
print('✅ Todas as dependências principais OK')
try:
    import tkinterdnd2
    print('✅ Drag & Drop disponível')
except:
    print('⚠️ tkinterdnd2 não instalado (opcional)')
"
```

### Teste de Configuração:
```bash
# Interface gráfica com validação automática
python desktop_app.py

# Ou linha de comando
python pdf_to_excel_updater.py --help
```

## 📝 Changelog v3.2

### v3.2 (Atual) - Interface Gráfica + Histórico Persistido
- ✅ **Interface gráfica moderna** com CustomTkinter
- ✅ **Sistema de abas** organizadas (Processamento/Histórico/Configurações)
- ✅ **Histórico persistido** entre sessões
- ✅ **Drag & Drop** de arquivos PDF
- ✅ **Popup de progresso** com logs em tempo real
- ✅ **Processamento dual** (FOLHA NORMAL + 13º SALÁRIO)
- ✅ **Fallback inteligente** para códigos de 13º salário
- ✅ **Persistência de configurações** automática
- ✅ **Validação em tempo real** de configurações

### v3.1 - Interface Gráfica + Diretório de Trabalho
- ✅ **Interface gráfica** básica para seleção de PDF
- ✅ **Diretório de trabalho** obrigatório via .env
- ✅ **Execução de qualquer local**
- ✅ **Organização padronizada** (DADOS/)

### v3.0 - Diretório de Trabalho
- ✅ **Sistema de diretório de trabalho**
- ✅ **Configuração obrigatória** via .env
- ✅ **Modo único** (removidos modos alternativos)

---

## 🎯 Guia Rápido v3.2

```bash
# 1. Instalação automática
setup.bat

# 2. Configuração inicial (.env)
echo "MODELO_DIR=C:/trabalho/folhas" > .env

# 3. Estrutura mínima:
# C:/trabalho/folhas/MODELO.xlsm ← Obrigatório

# 4. Execução (Interface Gráfica)
python desktop_app.py
# - Configure diretório na aba Processamento
# - Arraste PDF ou use botão Selecionar
# - Clique "Processar PDF"
# - Consulte histórico na aba Histórico

# 5. Alternativa (Linha de Comando)
python pdf_to_excel_updater.py
```

**💡 Novidade v3.2:** A aplicação agora funciona como uma **suite completa** com interface gráfica moderna, sistema de histórico persistido e configurações automáticas - ideal tanto para uso individual quanto corporativo!