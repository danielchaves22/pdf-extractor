# 📄 Sistema de Planilha Modelo - v2.2

## 🆕 Nova Funcionalidade

Agora o **PDF para Excel Updater** suporta **dois modos de operação**:

1. **Modo Tradicional**: Busca arquivo Excel existente com mesmo nome do PDF
2. **Modo Modelo**: Copia planilha modelo padronizada e renomeia conforme PDF

---

## 🚀 Como Configurar

### 1. **Criar arquivo .env**

Crie um arquivo chamado `.env` na mesma pasta do script:

```bash
# .env
MODELO_DIR=C:/caminho/para/pasta/do/modelo
```

**Exemplos válidos:**
```bash
# Windows
MODELO_DIR=C:/modelos/planilhas
MODELO_DIR=C:\\modelos\\planilhas
MODELO_DIR=D:\\empresa\\templates

# Relativo
MODELO_DIR=./modelos
MODELO_DIR=../modelos_compartilhados
```

### 2. **Criar arquivo MODELO.xlsm**

Na pasta configurada em `MODELO_DIR`, coloque seu arquivo modelo:

```
C:/modelos/planilhas/
└── MODELO.xlsm  ← Nome obrigatório!
```

**⚠️ IMPORTANTE**: O arquivo deve se chamar exatamente `MODELO.xlsm`

### 3. **Instalar nova dependência**

```bash
pip install python-dotenv
```

Ou use o requirements.txt atualizado:
```bash
pip install -r requirements.txt
```

---

## 💻 Como Usar

### **Modo Tradicional** (padrão):
```bash
# Busca arquivo Excel existente
python pdf_to_excel_updater.py relatorio.pdf
```

### **Modo Modelo** (novo):
```bash
# Usa planilha modelo
python pdf_to_excel_updater.py relatorio.pdf --modelo
python pdf_to_excel_updater.py relatorio.pdf -m
```

---

## 📁 Estrutura de Arquivos

### **Antes do processamento:**
```
projeto/
├── relatorio.pdf
├── pdf_to_excel_updater.py
├── .env
└── (sem pasta DADOS)
```

### **Depois do processamento (modo modelo):**
```
projeto/
├── relatorio.pdf
├── pdf_to_excel_updater.py  
├── .env
└── DADOS/
    └── relatorio.xlsm  ← Cópia do MODELO.xlsm
```

---

## 🎯 Exemplo Completo

### **1. Configuração inicial:**
```bash
# Crie .env
echo "MODELO_DIR=C:/modelos" > .env

# Estrutura necessária:
# C:/modelos/MODELO.xlsm ✓ (existe)
# projeto/relatorio.pdf ✓ (existe)
```

### **2. Execução:**
```bash
python pdf_to_excel_updater.py relatorio.pdf --modelo
```

### **3. Output esperado:**
```
Processando: relatorio.pdf (usando modelo)
[OK] Processamento concluído: 54 períodos atualizados
Arquivo: DADOS/relatorio.xlsm (criado a partir do modelo)
```

### **4. Resultado:**
- ✅ Pasta `DADOS/` criada automaticamente
- ✅ Arquivo `DADOS/relatorio.xlsm` criado (cópia do modelo)
- ✅ Dados do PDF inseridos na planilha
- ✅ Formatação, fórmulas e macros preservados

---

## ⚠️ Tratamento de Erros

### **Arquivo .env não encontrado:**
```
ERRO: MODELO_DIR não está definido no arquivo .env
```
**Solução:** Crie arquivo `.env` com `MODELO_DIR=caminho`

### **Diretório não encontrado:**
```
ERRO: Diretório do modelo não encontrado: C:/modelos
```
**Solução:** Verifique se o caminho em `MODELO_DIR` existe

### **MODELO.xlsm não encontrado:**
```
ERRO: Arquivo MODELO.xlsm não encontrado em: C:/modelos
```
**Solução:** Coloque arquivo `MODELO.xlsm` na pasta configurada

### **Conflito de argumentos:**
```
ERRO: Não é possível usar --modelo (-m) e --excel (-e) simultaneamente
```
**Solução:** Use apenas um dos dois modos

---

## 🔧 Argumentos de Linha de Comando

| Argumento | Descrição | Exemplo |
|-----------|-----------|---------|
| `pdf_path` | Arquivo PDF (obrigatório) | `relatorio.pdf` |
| `-m, --modelo` | Usar planilha modelo | `--modelo` |
| `-e, --excel` | Arquivo Excel específico | `-e planilha.xlsm` |
| `-s, --sheet` | Nome da planilha | `-s "DADOS"` |
| `-v, --verbose` | Modo verboso | `--verbose` |

**⚠️ Conflitos:** `--modelo` e `--excel` não podem ser usados juntos

---

## ✅ Vantagens do Modo Modelo

1. **📄 Padronização**: Todos os arquivos seguem o mesmo template
2. **🔒 Preservação**: Modelo original nunca é alterado
3. **📁 Organização**: Arquivos processados ficam na pasta DADOS
4. **🔄 Reprodutibilidade**: Processo consistente e repetível
5. **⚙️ Flexibilidade**: Fácil trocar de modelo via .env

---

## 🚨 Checklist de Validação

Antes de usar o modo modelo, verifique:

- [ ] Arquivo `.env` existe na pasta do script
- [ ] `MODELO_DIR` está definido no `.env`
- [ ] Pasta do modelo existe no sistema
- [ ] Arquivo `MODELO.xlsm` existe na pasta do modelo
- [ ] Planilha "LEVANTAMENTO DADOS" existe no modelo (ou use `-s`)
- [ ] `python-dotenv` está instalado

---

## 📞 Solução de Problemas

### **Teste de configuração:**
```bash
# Teste se .env está correto
python -c "from dotenv import load_dotenv; import os; load_dotenv(); print('MODELO_DIR:', os.getenv('MODELO_DIR'))"
```

### **Verificação de arquivos:**
```bash
# Lista conteúdo da pasta do modelo
ls "C:/modelos/"  # Deve mostrar MODELO.xlsm
```

### **Modo verbose para diagnóstico:**
```bash
python pdf_to_excel_updater.py relatorio.pdf --modelo --verbose
```

A nova funcionalidade torna o sistema muito mais flexível e adequado para ambientes empresariais que precisam de padronização! 🎯