# 🔧 PDF para Excel Updater

Aplicação Python para extrair dados de PDFs de folha de pagamento e preencher **diretamente** em planilhas Excel existentes (.xlsx/.xlsm), preservando formatação, fórmulas e macros VBA.

## ✨ Funcionalidades Principais

- ✅ **Preenche Excel existente** - não cria arquivo novo
- ✅ **Preserva macros VBA** (.xlsm) e formatação
- ✅ **Ignora campos fantasma** do PDF automaticamente
- ✅ **Detecta arquivo Excel** com mesmo nome do PDF
- ✅ **Mapeia datas** automaticamente (formato texto ou datetime)
- ✅ **Processamento offline** (sem internet)
- ✅ **Filtro inteligente** (ignora 13º salário, férias)

## 📊 Mapeamento de Dados (v2.1)

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

# 3. Teste a instalação
python test_instalacao.py
```

### Pré-requisitos
- **Python 3.7+** 
- **pip** (incluído no Python)

## 💻 Como Usar

### 📁 Estrutura de Arquivos
```
pasta_projeto/
├── relatorio.pdf          # Seu PDF
├── relatorio.xlsm          # Excel com MESMO NOME
└── pdf_to_excel_updater.py # Aplicação
```

### 🎯 Uso Básico
```bash
# Comando mais simples (usa "LEVANTAMENTO DADOS" automaticamente)
python pdf_to_excel_updater.py EdsonGoulartLeonardo2.pdf

# ⚠️ IMPORTANTE: Se "LEVANTAMENTO DADOS" não existir, especifique outra:
python pdf_to_excel_updater.py arquivo.pdf -s "Nome_Da_Planilha_Correta"

# Modo verboso (mostra detalhes de debug)
python pdf_to_excel_updater.py arquivo.pdf -v

# Especifica Excel manualmente
python pdf_to_excel_updater.py arquivo.pdf -e planilha.xlsm
```

### 📋 Exemplo Real
```bash
# Seus arquivos:
# - EdsonGoulartLeonardo2.pdf
# - EdsonGoulartLeonardo2.xlsm (com aba "LEVANTAMENTO DADOS")

python pdf_to_excel_updater.py EdsonGoulartLeonardo2.pdf

# Output esperado:
# 🔄 Processando: EdsonGoulartLeonardo2.pdf
# ✅ Concluído: 54 períodos processados  
# 📁 Arquivo: EdsonGoulartLeonardo2.xlsm (macros preservados)
```

## 📈 Resultado Esperado (v2.1)

### ✅ Sucesso Total:
```
🔄 Processando: EdsonGoulartLeonardo2.pdf
✅ Concluído: 54 períodos processados
📁 Arquivo: EdsonGoulartLeonardo2.xlsm (macros preservados)
```

### ⚠️ Sucesso Parcial:
```
🔄 Processando: EdsonGoulartLeonardo2.pdf  
✅ Processamento concluído: 45/54 períodos atualizados
❌ Falhas em 9 períodos:
   out/12 (linha não encontrada)
   nov/14 (células já preenchidas)
   dez/15 (linha não encontrada)
   ...
📁 Arquivo: EdsonGoulartLeonardo2.xlsm (macros preservados)
```

### ❌ Erro Crítico:
```
🔄 Processando: arquivo.pdf
❌ Erro: Planilha 'LEVANTAMENTO DADOS' não encontrada. Use -s para especificar outra planilha.
```

## 🔧 Funcionalidades Avançadas

### 📊 Suporte a Múltiplos Formatos
- **.xlsm** - Excel com macros (preserva VBA)
- **.xlsx** - Excel padrão
- **.xls** - Excel legado

### 🎯 Detecção Inteligente de Planilhas
1. Procura por nome: "LEVANTAMENTO DADOS"
2. Procura por palavra-chave: "LEVANTAMENTO" ou "DADOS"
3. Usa segunda aba como fallback
4. Permite especificação manual com `-s`

### 📅 Mapeamento de Datas Flexível
- **Texto**: `nov/12`, `dez/12`, `jan/13`
- **DateTime**: `2012-11-10 00:00:00`
- **Serial Date**: Números do Excel (41224, 41254, etc.)

## 🐛 Solução de Problemas

### Erro: "Períodos não encontrados"
```bash
# Diagnostique a planilha
python diagnose_excel.py arquivo.xlsm

# Especifique a planilha correta
python pdf_to_excel_updater.py arquivo.pdf -s "Nome_Aba_Correta"
```

### Erro: "ModuleNotFoundError"
```bash
# Reinstale dependências
pip install --upgrade -r requirements.txt

# Ou instale individualmente
pip install pandas openpyxl pdfplumber
```

### Erro: "Excel não encontrado"
- ✅ Verifique se PDF e Excel têm **mesmo nome**
- ✅ Confirme extensão: `.xlsm`, `.xlsx` ou `.xls`
- ✅ Use `-e` para especificar caminho manualmente

### Dados não são extraídos
- ✅ Use `-v` para modo verboso
- ✅ Verifique se PDF contém os códigos corretos
- ✅ Confirme que não é folha de férias/13º salário

## 📁 Arquivos do Projeto

```
pdf-extractor/
├── pdf_to_excel_updater.py  # ← Aplicação principal
├── requirements.txt         # ← Dependências
├── setup.bat               # ← Instalação automática (Windows)
├── test_instalacao.py      # ← Teste de instalação
├── README.md               # ← Esta documentação
└── diagnose_excel.py       # ← Ferramenta de diagnóstico
```

## 🔒 Segurança e Preservação

### ✅ O que é Preservado
- **Macros VBA** (.xlsm)
- **Fórmulas existentes**
- **Formatação** (cores, bordas, fontes)
- **Estrutura** da planilha
- **Dados existentes** (não sobrescreve)

### ✅ O que é Preenchido
- **Apenas** colunas B, X, Y, AA, AC
- **Apenas** se célula estiver vazia
- **Apenas** dados extraídos do PDF

## 📞 Suporte

### Comandos de Diagnóstico
```bash
# Testa instalação
python test_instalacao.py

# Analisa estrutura do Excel
python diagnose_excel.py arquivo.xlsm

# Testa com modo verboso
python pdf_to_excel_updater.py arquivo.pdf -v
```

### Logs Importantes
- `✓ Código encontrado:` - PDF sendo processado corretamente
- `Período X/Y encontrado na linha Z` - Mapeamento funcionando
- `Atualizado B5: None → 1234.56` - Dados sendo inseridos
- `Excel .xlsm atualizado com N alterações` - Sucesso

## 📝 Versões

### v2.1 (Atual) - Conformidade Total com Regras de Negócio
- ✅ **Código 02007501** adicionado (DIFER.PROV. HORAS EXTRAS 75%)
- ✅ **Fallback inteligente** para PRODUÇÃO (índice → valor)
- ✅ **Formato de horas** automático (`06:34` → `06,34`)
- ✅ **Planilha padrão obrigatória** ("LEVANTAMENTO DADOS")
- ✅ **Output simplificado** (resumo conciso, erros detalhados)
- ✅ **Preservação total** de macros VBA e formatação

### v2.0 (Anterior)
- ✅ Preenche Excel existente (.xlsm/.xlsx)
- ✅ Preserva macros VBA completamente
- ✅ Mapeamento de datas automático
- ✅ Detecção inteligente de planilhas
- ✅ Compatibilidade com Windows

### v1.0 (Legado)
- ✅ Gerava novo Excel/CSV
- ❌ Não preservava formatação

---

## 🎯 Exemplo Completo

```bash
# 1. Prepare os arquivos
EdsonGoulartLeonardo2.pdf       # ← Seu PDF
EdsonGoulartLeonardo2.xlsm      # ← Excel existente

# 2. Execute
python pdf_to_excel_updater.py EdsonGoulartLeonardo2.pdf -s "LEVANTAMENTO DADOS" -v

# 3. Resultado
# ✓ 54 períodos processados
# ✓ Colunas B, X, Y, AA, AC preenchidas
# ✓ Macros VBA preservados
# ✓ Formatação mantida
```

**💡 Dica:** Sempre mantenha backup do Excel original antes de processar!