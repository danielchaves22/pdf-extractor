# 📋 Regras de Negócio - PDF para Excel Updater v3.1

## 1. 🗂️ Filtro de Folhas ✅ IMPLEMENTADO

### ✅ Processar Apenas:
- **"Tipo da folha: FOLHA NORMAL"** - Folhas de pagamento regulares

### ❌ Ignorar Completamente:
- **"Tipo da folha: 13º SALÁRIO"** - Folhas de décimo terceiro
- **"Tipo da folha: FÉRIAS"** - Folhas de férias
- **"Tipo da folha: RESCISÃO"** - Folhas de rescisão
- **"Tipo da folha: ADIANTAMENTO"** - Adiantamentos salariais

### 🔍 Lógica de Filtro:
1. **Prioridade 1**: Procura linha específica "Tipo da folha:"
2. **Prioridade 2**: Se não encontrar, aplica filtro no cabeçalho (primeiras 10 linhas)
3. **Comportamento**: Termos como "FÉRIAS" ou "13º SALÁRIO" em valores são ignorados

---

## 2. 📊 Mapeamento de Códigos PDF → Excel ✅ IMPLEMENTADO

### 🔵 Obter da Coluna ÍNDICE (Penúltimo número da linha)

| Código PDF | Descrição | Excel Coluna | Status | Observações |
|------------|-----------|--------------|--------|-------------|
| `01003601` | PREMIO PROD. MENSAL | **X** (PRODUÇÃO) | ✅ | ⚡ **FALLBACK**: Se ÍNDICE vazio, usar VALOR |
| `01007301` | HORAS EXT.100%-180 | **Y** (INDICE HE 100%) | ✅ | 🕐 Suporta formato horas |
| `01009001` | ADIC.NOT.25%-180 | **AC** (INDICE ADC. NOT.) | ✅ | 🕐 Suporta formato horas |
| `01003501` | HORAS EXT.75%-180 | **AA** (INDICE HE 75%) | ✅ | 🕐 Suporta formato horas |
| `02007501` | DIFER.PROV. HORAS EXTRAS 75% | **AA** (INDICE HE 75%) | ✅ | 🕐 Código alternativo para HE 75% |

### 🔴 Obter da Coluna VALOR (Último número da linha)

| Código PDF | Descrição | Excel Coluna | Status | Observações |
|------------|-----------|--------------|--------|-------------|
| `09090301` | SALARIO CONTRIB INSS | **B** (REMUNERAÇÃO RECEBIDA) | ✅ | - |

---

## 3. 🕐 Tratamento de Formato de Horas ✅ IMPLEMENTADO

### Situação:
- **Período inicial**: Valores numéricos normais (ex: `123.45`)
- **Período posterior**: Formato de horas (ex: `06:34`)

### Regra de Conversão:
```
Entrada: '06:34'
Saída:   '06,34'
```

**✅ Implementação**: Detecta automaticamente formato `\d{1,2}:\d{2}` e substitui `:` por `,`

**Colunas Afetadas**:
- **Y** (INDICE HE 100%) ✅
- **AA** (INDICE HE 75%) ✅
- **AC** (INDICE ADC. NOT.) ✅

---

## 4. 📂 Regras de Planilha Excel ✅ IMPLEMENTADO

### Planilha Padrão Obrigatória:
- **Nome**: `"LEVANTAMENTO DADOS"` ✅
- **Comportamento**: Se não especificada no comando, usa esta automaticamente ✅
- **Erro**: Se não existir e não for especificada outra, **PARAR EXECUÇÃO** ✅

### Estrutura Requerida:
- **Coluna A**: PERÍODO (datas - não modificar) ✅
- **Coluna B**: REMUNERAÇÃO RECEBIDA ✅
- **Coluna X**: PRODUÇÃO ✅
- **Coluna Y**: INDICE HE 100% ✅
- **Coluna AA**: INDICE HE 75% ✅
- **Coluna AC**: INDICE ADC. NOT. ✅

---

## 5. 🔄 Lógica de Fallback ✅ IMPLEMENTADO

### Para Código 01003601 (PRODUÇÃO):
1. **Prioridade 1**: Tentar coluna ÍNDICE (penúltimo número) ✅
2. **Prioridade 2**: Se ÍNDICE vazio/zero, usar coluna VALOR (último número) ✅
3. **Resultado**: Sempre tentar preencher a coluna X (PRODUÇÃO) ✅

### Exemplo:
```
Linha PDF: P 01003601 PREMIO PROD. MENSAL 1.2 00,30,030 0,00 1.203,30
                                                        ^^^^  ^^^^^^^^
                                                      ÍNDICE   VALOR
                                                     (vazio)  (usar este) ✅
```

---

## 6. 🎯 Códigos Duplicados ✅ IMPLEMENTADO

### INDICE HE 75% (Coluna AA):
Pode ser preenchido por **dois códigos diferentes**:
- `01003501` - HORAS EXT.75%-180 ✅
- `02007501` - DIFER.PROV. HORAS EXTRAS 75% ✅

**Comportamento**: Aceitar ambos, última ocorrência sobrescreve ✅

---

## 7. 📤 Regras de Output ✅ IMPLEMENTADO

### Resumo Conciso:
```
[OK] Processamento concluído: X períodos atualizados          ← Sucesso total
[AVISO] Processamento concluído: X/Y períodos atualizados     ← Sucesso parcial
OK: Concluído: X períodos processados                         ← Resumo final
```

### Detalhes Apenas para Erros:
- Mostrar **apenas** períodos que falharam ✅
- Não mostrar sucessos individuais ✅
- Foco em problemas que precisam atenção ✅
- Máximo 3 erros mostrados (+ contador do restante) ✅

---

## 8. 🔒 Preservação de Dados ✅ IMPLEMENTADO

### ✅ Sempre Preservar:
- **Macros VBA** (.xlsm) - `keep_vba=True` ✅
- **Fórmulas existentes** - Não sobrescrever ✅
- **Formatação** (cores, bordas, fontes) - Preservação completa ✅
- **Dados existentes** - Não sobrescrever se já preenchido ✅

### ✅ Apenas Preencher:
- **Células vazias** (`None`, `''`, `0`) ✅
- **Colunas específicas**: B, X, Y, AA, AC ✅
- **Dados extraídos** do PDF com sucesso ✅

---

## 9. 🚨 Tratamento de Erros ✅ IMPLEMENTADO

### Erros Críticos (Parar Execução):
- Arquivo .env não encontrado ✅
- MODELO_DIR não definido no .env ✅
- Diretório de trabalho não encontrado ✅
- MODELO.xlsm não encontrado ✅
- PDF não encontrado no diretório de trabalho ✅
- Planilha "LEVANTAMENTO DADOS" não existe (quando não especificada) ✅
- Dependências não instaladas ✅

### Avisos (Continuar Execução):
- Período não encontrado na planilha ✅
- Código não encontrado no PDF ✅
- Células já preenchidas ✅
- Páginas sem período identificado ✅
- Páginas sem códigos mapeados ✅

---

## 10. 🏗️ Arquitetura v3.1 ✅ IMPLEMENTADO

### ✅ Modo Único:
- **Sempre usa modelo** MODELO.xlsm ✅
- **Sempre cria arquivo** em DADOS/ ✅
- **Sem modos alternativos** (simplificado) ✅

### ✅ Diretório de Trabalho:
- **Configuração obrigatória** via MODELO_DIR no .env ✅
- **PDF e MODELO.xlsm** no mesmo diretório ✅
- **Pasta DADOS/** criada automaticamente ✅
- **Execução de qualquer local** ✅

### ✅ Interface Gráfica:
- **Seleção visual de PDF** (tkinter) ✅
- **Filtra apenas PDFs** do diretório de trabalho ✅
- **Fallback para linha de comando** ✅
- **Compatibilidade total** (opcional) ✅

---

## 11. 📋 Resumo de Implementação v3.1

### ✅ TOTALMENTE IMPLEMENTADO:
- ✅ Filtro de folhas por "Tipo da folha: FOLHA NORMAL"
- ✅ Mapeamento completo de todos os códigos (incluindo 02007501)
- ✅ Fallback ÍNDICE → VALOR para 01003601 (PRODUÇÃO)
- ✅ Conversão automática formato horas (`06:34` → `06,34`)
- ✅ Planilha padrão obrigatória "LEVANTAMENTO DADOS"
- ✅ Preservação total de macros VBA (.xlsm)
- ✅ Output simplificado e conciso
- ✅ Tratamento completo de erros
- ✅ Diretório de trabalho obrigatório
- ✅ Interface gráfica para seleção
- ✅ Modo único simplificado

### 🎯 REGRAS DE NEGÓCIO: 100% CONFORMES

Todas as regras especificadas foram implementadas e testadas na versão 3.1 do PDF para Excel Updater.