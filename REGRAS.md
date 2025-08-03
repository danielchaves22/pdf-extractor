# 📋 Regras de Negócio - PDF para Excel Updater

## 1. 🗂️ Filtro de Folhas

### ✅ Processar Apenas:
- **FOLHA NORMAL** - Folhas de pagamento regulares

### ❌ Ignorar Completamente:
- **13º SALÁRIO** - Folhas de décimo terceiro
- **FÉRIAS** - Folhas de férias
- **RESCISÃO** - Folhas de rescisão
- **ADIANTAMENTO** - Adiantamentos salariais

---

## 2. 📊 Mapeamento de Códigos PDF → Excel

### 🔵 Obter da Coluna ÍNDICE (Penúltimo número da linha)

| Código PDF | Descrição | Excel Coluna | Observações |
|------------|-----------|--------------|-------------|
| `01003601` | PREMIO PROD. MENSAL | **X** (PRODUÇÃO) | ⚠️ **FALLBACK**: Se ÍNDICE vazio, usar VALOR |
| `01007301` | HORAS EXT.100%-180 | **Y** (INDICE HE 100%) | - |
| `01009001` | ADIC.NOT.25%-180 | **AC** (INDICE ADC. NOT.) | 🕐 Pode virar formato horas |
| `01003501` | HORAS EXT.75%-180 | **AA** (INDICE HE 75%) | 🕐 Pode virar formato horas |
| `02007501` | DIFER.PROV. HORAS EXTRAS 75% | **AA** (INDICE HE 75%) | 🕐 Código alternativo para HE 75% |

### 🔴 Obter da Coluna VALOR (Último número da linha)

| Código PDF | Descrição | Excel Coluna | Observações |
|------------|-----------|--------------|-------------|
| `09090301` | SALARIO CONTRIB INSS | **B** (REMUNERAÇÃO RECEBIDA) | - |

---

## 3. 🕐 Tratamento de Formato de Horas

### Situação:
- **Período inicial**: Valores numéricos normais (ex: `123.45`)
- **Período posterior**: Formato de horas (ex: `06:34`)

### Regra de Conversão:
```
Entrada: '06:34'
Saída:   '06,34'
```

**Implementação**: Substituir `:` por `,` quando detectado formato de horas.

**Colunas Afetadas**:
- **Y** (INDICE HE 100%)
- **AA** (INDICE HE 75%) 
- **AC** (INDICE ADC. NOT.)

---

## 4. 📂 Regras de Planilha Excel

### Planilha Padrão Obrigatória:
- **Nome**: `"LEVANTAMENTO DADOS"`
- **Comportamento**: Se não especificada no comando, usa esta automaticamente
- **Erro**: Se não existir e não for especificada outra, **PARAR EXECUÇÃO**

### Estrutura Requerida:
- **Coluna A**: PERÍODO (datas - não modificar)
- **Coluna B**: REMUNERAÇÃO RECEBIDA
- **Coluna X**: PRODUÇÃO  
- **Coluna Y**: INDICE HE 100%
- **Coluna AA**: INDICE HE 75%
- **Coluna AC**: INDICE ADC. NOT.

---

## 5. 🔄 Lógica de Fallback

### Para Código 01003601 (PRODUÇÃO):
1. **Prioridade 1**: Tentar coluna ÍNDICE (penúltimo número)
2. **Prioridade 2**: Se ÍNDICE vazio/zero, usar coluna VALOR (último número)
3. **Resultado**: Sempre tentar preencher a coluna X (PRODUÇÃO)

### Exemplo:
```
Linha PDF: P 01003601 PREMIO PROD. MENSAL 1.2 00,30,030 0,00 1.203,30
                                                        ^^^^  ^^^^^^^^
                                                      ÍNDICE   VALOR
                                                     (vazio)  (usar este)
```

---

## 6. 🎯 Códigos Duplicados

### INDICE HE 75% (Coluna AA):
Pode ser preenchido por **dois códigos diferentes**:
- `01003501` - HORAS EXT.75%-180
- `02007501` - DIFER.PROV. HORAS EXTRAS 75%

**Comportamento**: Aceitar ambos, última ocorrência sobrescreve.

---

## 7. 📤 Regras de Output

### Resumo Conciso:
```
✅ Processamento concluído: X/Y períodos atualizados
❌ Falhas: Z períodos não encontrados
```

### Detalhes Apenas para Erros:
- Mostrar **apenas** períodos que falharam
- Não mostrar sucessos individuais
- Foco em problemas que precisam atenção

---

## 8. 🔒 Preservação de Dados

### ✅ Sempre Preservar:
- **Macros VBA** (.xlsm)
- **Fórmulas existentes**
- **Formatação** (cores, bordas, fontes)
- **Dados existentes** (não sobrescrever se já preenchido)

### ✅ Apenas Preencher:
- **Células vazias** (`None`, `''`, `0`)
- **Colunas específicas**: B, X, Y, AA, AC
- **Dados extraídos** do PDF com sucesso

---

## 9. 🚨 Tratamento de Erros

### Erros Críticos (Parar Execução):
- PDF não encontrado
- Excel não encontrado
- Planilha "LEVANTAMENTO DADOS" não existe (quando não especificada)
- Dependências não instaladas

### Avisos (Continuar Execução):
- Período não encontrado na planilha
- Código não encontrado no PDF
- Formato de número inválido

---

## 10. 📋 Resumo de Implementação

### ✅ Já Implementado:
- Filtro de folhas especiais
- Mapeamento básico de códigos
- Preservação de macros (.xlsm)
- Detecção automática de Excel
- Mapeamento de datas flexível

### ❌ Pendente de Implementação:
- Código `02007501` (DIFER.PROV. HORAS EXTRAS 75%)
- Fallback para `01003601` (ÍNDICE → VALOR)
- Conversão formato horas (`06:34` → `06,34`)
- Planilha padrão obrigatória
- Output simplificado