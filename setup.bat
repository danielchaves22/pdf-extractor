# 🔧 Extrator de Dados de Folha de Pagamento

Aplicação Python para extrair dados de PDFs de folha de pagamento e gerar planilhas estruturadas automaticamente.

## 📋 Funcionalidades

- ✅ Extração automática de dados de PDFs
- ✅ Processamento offline (sem internet)
- ✅ Geração de planilhas Excel (.xlsx) ou CSV
- ✅ Mapeamento inteligente de códigos específicos
- ✅ Filtro automático de folhas especiais (13º salário, férias)
- ✅ Preenchimento automático de fórmulas
- ✅ Período completo (Nov/12 a Nov/17) com linhas em branco para meses sem dados

## 🚀 Instalação

### Pré-requisitos
- Python 3.7 ou superior
- pip (gerenciador de pacotes Python)

### Passos de Instalação

1. **Clone ou baixe os arquivos:**
   ```bash
   # Baixe os seguintes arquivos:
   # - pdf_extractor.py
   # - requirements.txt
   # - exemplo_uso.py
   ```

2. **Instale as dependências:**
   ```bash
   pip install -r requirements.txt
   ```

3. **Verifique a instalação:**
   ```bash
   python exemplo_uso.py
   ```

## 💻 Como Usar

### Uso Básico (Linha de Comando)

```bash
# Comando simples
python pdf_extractor.py seu_arquivo.pdf

# Especificando arquivo de saída
python pdf_extractor.py folha_pagamento.pdf -o resultado.xlsx

# Modo verboso (mostra detalhes do processamento)
python pdf_extractor.py folha_pagamento.pdf -v
```

### Uso Programático

```python
from pdf_extractor import PDFDataExtractor

# Cria extrator
extractor = PDFDataExtractor()

# Processa PDF e gera planilha
df = extractor.process_pdf("meu_arquivo.pdf", "resultado.xlsx")

# Trabalha com os dados
print(f"Extraídos {len(df)} períodos")
```

## 📊 Estrutura da Saída

A aplicação gera uma planilha com as seguintes colunas:

| Coluna | Descrição | Fonte |
|--------|-----------|-------|
| PERÍODO | Mês/Ano (formato: nov/12) | Calculado |
| REMUNERAÇÃO RECEBIDA | Código 09090301 | Coluna Valor |
| PRODUÇÃO | Código 01003601 | Coluna Índice |
| INDICE HE 100% | Código 01007301 | Coluna Índice |
| FORMULA_1 | =Y[linha]/100 | Calculado |
| INDICE HE 75% | Código 01003501 | Coluna Índice |
| FORMULA_2 | =AC[linha]/10000 | Calculado |
| INDICE ADC. NOT. | Código 01009001 | Coluna Índice |
| FORMULA_3 | =AC[linha]/10000 | Calculado |

## ⚙️ Regras de Mapeamento

### Códigos Processados

- **01003601** (PREMIO PROD. MENSAL) → Coluna PRODUÇÃO
- **01007301** (HORAS EXT.100%-180) → Coluna INDICE HE 100%
- **01009001** (ADIC.NOT.25%-180) → Coluna INDICE ADC. NOT.
- **01003501** (HORAS EXT.75%-180) → Coluna INDICE HE 75%
- **09090301** (SALARIO CONTRIB INSS) → Coluna REMUNERAÇÃO RECEBIDA

### Filtros Aplicados

- ❌ Folhas de 13º salário
- ❌ Folhas de férias
- ✅ Apenas folhas normais de pagamento

## 📁 Estrutura de Arquivos

```
projeto/
├── pdf_extractor.py      # Aplicação principal
├── requirements.txt      # Dependências
├── exemplo_uso.py       # Exemplos de uso
├── README.md            # Esta documentação
└── seus_pdfs/           # Seus arquivos PDF
```

## 🐛 Solução de Problemas

### Erro: "Arquivo não encontrado"
```bash
# Verifique se o caminho está correto
ls seu_arquivo.pdf

# Use caminho absoluto se necessário
python pdf_extractor.py /caminho/completo/para/arquivo.pdf
```

### Erro: "ModuleNotFoundError"
```bash
# Instale as dependências
pip install -r requirements.txt

# Se ainda der erro, tente:
pip install pandas pdfplumber openpyxl
```

### PDF não está sendo processado corretamente
- ✅ Verifique se o PDF não está protegido por senha
- ✅ Confirme se o PDF contém texto (não é só imagem)
- ✅ Use o modo verboso (`-v`) para ver detalhes

### Dados não aparecem na planilha
- ✅ Verifique se os códigos estão presentes no PDF
- ✅ Confirme se o formato do PDF está correto
- ✅ Use modo verboso para ver quais dados foram encontrados

## 📈 Exemplo de Resultado

A aplicação gera uma planilha similar a esta:

```
PERÍODO    | REMUNERAÇÃO | PRODUÇÃO | INDICE HE 100% | FORMULA_1  | ...
nov/12     | 6176,41     | 1203,30  | 4224,00        | =Y5/100    | ...
dez/12     | 5918,34     | 745,79   | 8058,00        | =Y6/100    | ...
jan/13     | 4895,82     | 362,35   | 6405,00        | =Y7/100    | ...
...        | ...         | ...      | ...            | ...        | ...
```

## 🔧 Personalização

### Modificar Período de Extração

No arquivo `pdf_extractor.py`, altere os parâmetros:

```python
# Altere estas linhas na função generate_complete_table
start_date: Tuple[int, int] = (11, 2012),  # (mês, ano) inicial
end_date: Tuple[int, int] = (11, 2017)     # (mês, ano) final
```

### Adicionar Novos Códigos

No arquivo `pdf_extractor.py`, modifique o dicionário `mapping_rules`:

```python
self.mapping_rules = {
    # Códigos existentes...
    'NOVO_CODIGO': {
        'code': 'DESCRIÇÃO', 
        'target': 'NOME_COLUNA', 
        'source': 'indice' # ou 'valor'
    }
}
```

## 📞 Suporte

Se encontrar problemas:

1. **Verifique os logs:** Use `-v` para modo verboso
2. **Teste com exemplo:** Execute `python exemplo_uso.py`
3. **Valide o PDF:** Certifique-se que contém os códigos esperados
4. **Verifique dependências:** Execute `pip list` para ver pacotes instalados

## 📝 Changelog

### Versão 1.0
- ✅ Extração básica de PDFs
- ✅ Mapeamento de códigos específicos
- ✅ Geração de planilhas Excel/CSV
- ✅ Filtro de folhas especiais
- ✅ Período completo Nov/12 a Nov/17

---

**💡 Dica:** Mantenha seus PDFs organizados em uma pasta específica para facilitar o processamento em lote!
