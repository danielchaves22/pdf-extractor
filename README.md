# 🔧 PDF para Excel Updater v3.1

Aplicação Python para extrair dados de PDFs de folha de pagamento e preencher planilhas Excel usando **diretório de trabalho configurado**, preservando formatação, fórmulas e macros VBA.

## ✨ Funcionalidades Principais

- ✅ **Diretório de trabalho** configurado via .env (MODELO_DIR)
- ✅ **Interface gráfica** para seleção de PDF (opcional)
- ✅ **Preserva macros VBA** (.xlsm) e formatação completa
- ✅ **Copia modelo automaticamente** para pasta DADOS/
- ✅ **Filtro inteligente** por "Tipo da folha: FOLHA NORMAL"
- ✅ **Processamento offline** (sem internet)
- ✅ **Execução de qualquer local** (não precisa estar na pasta do PDF)

## 📊 Mapeamento de Dados Completo

| PDF Código | Descrição | Excel Coluna | Fonte | Regras Especiais |
|------------|-----------|--------------|-------|------------------|
| `09090301` | SALARIO CONTRIB INSS | **B** (REMUNERAÇÃO RECEBIDA) | Último número | - |
| `01003601` | PREMIO PROD. MENSAL | **X** (PRODUÇÃO) | Penúltimo número | ⚡ **Fallback**: Se vazio, usa último número |
| `01007301` | HORAS EXT.100%-180 | **Y** (INDICE HE 100%) | Penúltimo número | 🕐 Suporta formato horas |
| `01003501` | HORAS EXT.75%-180 | **AA** (INDICE HE 75%) | Penúltimo número | 🕐 Suporta formato horas |
| `02007501` | DIFER.PROV. HORAS EXTRAS 75% | **AA** (INDICE HE 75%) | Penúltimo número | 🕐 Código alternativo |
| `01009001` | ADIC.NOT.25%-180 | **AC** (INDICE ADC. NOT.) | Penúltimo número | 🕐 Suporta formato horas |

### 🕐 Formato de Horas
- **Detecção automática**: `06:34` → `06,34`
- **Aplicável a**: HE 100%, HE 75%, ADIC. NOT.
- **Conversão**: Substitui `:` por `,` automaticamente

## 🚀 Instalação Rápida

### Opção 1: Automática (Windows)
```batch
# Execute o setup automático
setup.bat
```

### Opção 2: Manual
```bash
# 1. Verifique Python 3.7+
python --version

# 2. Instale dependências
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
    └── relatorio.xlsm            ← Resultado processado
```

## 💻 Como Usar

### 🎯 Uso com Interface Gráfica (Recomendado)
```bash
# Abre seletor de arquivo no diretório de trabalho
python pdf_to_excel_updater.py
```

### 🎯 Uso por Linha de Comando
```bash
# Processa arquivo específico
python pdf_to_excel_updater.py "relatorio.pdf"

# Com planilha específica
python pdf_to_excel_updater.py "relatorio.pdf" -s "DADOS"

# Modo verboso (diagnóstico)
python pdf_to_excel_updater.py "relatorio.pdf" -v
```

### 📋 Exemplo Real
```bash
# 1. Configure .env
MODELO_DIR=C:/trabalho/folhas

# 2. Estrutura no diretório:
# C:/trabalho/folhas/
# ├── MODELO.xlsm
# └── EdsonGoulart-Jan2025.pdf

# 3. Execute (interface gráfica)
python pdf_to_excel_updater.py

# 4. Resultado automático:
# C:/trabalho/folhas/DADOS/EdsonGoulart-Jan2025.xlsm
```

## 📈 Resultado Esperado

### ✅ Sucesso Total:
```
Processando: EdsonGoulart-Jan2025.pdf
[OK] Processamento concluído: 54 períodos atualizados
OK: Concluído: 54 períodos processados
Arquivo criado: DADOS/EdsonGoulart-Jan2025.xlsm
```

### ⚠️ Sucesso Parcial:
```
Processando: relatorio.pdf
[AVISO] Processamento concluído: 45/54 períodos atualizados
[ERRO] Falhas em 9 períodos:
   out/12 (linha não encontrada)
   nov/14 (células já preenchidas)
   dez/15 (linha não encontrada)
OK: Concluído: 45 períodos processados
Arquivo criado: DADOS/relatorio.xlsm
```

### ❌ Erro de Configuração:
```
ERRO: Arquivo .env não encontrado. Configure MODELO_DIR no arquivo .env
```

## 🔧 Funcionalidades Avançadas

### 📊 Filtro Inteligente de Folhas
- ✅ **Procura especificamente** por "Tipo da folha: FOLHA NORMAL"
- ✅ **Ignora automaticamente** 13º salário, férias, rescisão
- ✅ **Fallback inteligente** para PDFs sem linha "Tipo da folha"

### 🎯 Detecção de Planilha
1. **Padrão**: "LEVANTAMENTO DADOS" (obrigatório se não especificado)
2. **Manual**: Use `-s "Nome_Da_Planilha"`

### 📅 Mapeamento de Períodos Flexível
- **Texto**: `nov/12`, `dez/12`, `jan/13`
- **DateTime**: `2012-11-10 00:00:00`
- **Serial Date**: Números do Excel

## 🐛 Solução de Problemas

### Erro: "Arquivo .env não encontrado"
```bash
# Crie arquivo .env
echo "MODELO_DIR=C:/trabalho/folhas" > .env
```

### Erro: "MODELO.xlsm não encontrado"
```bash
# Coloque MODELO.xlsm no diretório de trabalho
# Verifique se o caminho em MODELO_DIR está correto
```

### Erro: "Arquivo PDF não encontrado"
```bash
# Use aspas para nomes com espaços
python pdf_to_excel_updater.py "Relatório Janeiro 2025.pdf"
```

### Nenhum dado extraído
```bash
# Use modo verboso para diagnóstico
python pdf_to_excel_updater.py arquivo.pdf -v
```

## 📁 Arquivos do Projeto

```
pdf-updater/
├── pdf_to_excel_updater.py     # ← Aplicação principal v3.1
├── requirements.txt            # ← Dependências
├── setup.bat                   # ← Instalação automática
├── .env                        # ← Configuração (MODELO_DIR)
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

## 🎯 Casos de Uso

### **Uso Corporativo:**
```bash
# .env corporativo
MODELO_DIR=//servidor/rh/processamento_folhas

# Uso por qualquer usuário na rede
python pdf_to_excel_updater.py
```

### **Uso Individual:**
```bash
# .env local
MODELO_DIR=D:/meus_documentos/folhas

# Processamento local
python pdf_to_excel_updater.py "folha_janeiro.pdf"
```

## 📞 Comandos de Diagnóstico

### Teste de Configuração:
```bash
# Verifica se .env está correto
python -c "from dotenv import load_dotenv; import os; load_dotenv(); print('MODELO_DIR:', os.getenv('MODELO_DIR'))"
```

### Modo Verbose:
```bash
# Diagnóstico completo
python pdf_to_excel_updater.py arquivo.pdf -v
```

### Ajuda:
```bash
# Menu de ajuda
python pdf_to_excel_updater.py --help
```

## 📝 Changelog

### v3.1 (Atual) - Interface Gráfica + Diretório de Trabalho
- ✅ **Interface gráfica** para seleção de PDF
- ✅ **Diretório de trabalho** obrigatório via .env
- ✅ **Execução de qualquer local**
- ✅ **Organização padronizada** (DADOS/)
- ✅ **Modo único** simplificado (sempre usa modelo)

### v3.0 - Diretório de Trabalho
- ✅ **Sistema de diretório de trabalho**
- ✅ **Configuração obrigatória** via .env
- ✅ **Modo único** (removidos modos alternativos)

### v2.x - Funcionalidades Completas
- ✅ **Todos os códigos** implementados (incluindo 02007501)
- ✅ **Fallback inteligente** para PRODUÇÃO
- ✅ **Formato de horas** automático
- ✅ **Filtro aprimorado** de folhas

---

## 🎯 Exemplo Completo de Uso

```bash
# 1. Configuração inicial
echo "MODELO_DIR=C:/trabalho/folhas" > .env

# 2. Estrutura necessária:
# C:/trabalho/folhas/
# ├── MODELO.xlsm              ← Seu template
# └── funcionario-jan2025.pdf  ← PDF a processar

# 3. Execução (escolha uma):
python pdf_to_excel_updater.py                           # Interface gráfica
python pdf_to_excel_updater.py "funcionario-jan2025.pdf" # Linha de comando

# 4. Resultado:
# C:/trabalho/folhas/DADOS/funcionario-jan2025.xlsm ✓ Criado e preenchido
```

**💡 Dica:** A aplicação funciona como ferramenta corporativa - configure uma vez e use em qualquer projeto!