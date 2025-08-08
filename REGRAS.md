# 📋 Regras de Negócio - PDF para Excel Updater v3.2

## 1. 🗂️ Filtro de Folhas ✅ IMPLEMENTADO

### ✅ Processar por Tipo:
- **"Tipo da folha: FOLHA NORMAL"** - Folhas de pagamento regulares (Linhas 1-65)
- **"Tipo da folha: 13º SALÁRIO"** - Décimo terceiro salário (Linhas 67+)

### ❌ Ignorar Completamente:
- **"Tipo da folha: FÉRIAS"** - Folhas de férias
- **"Tipo da folha: RESCISÃO"** - Folhas de rescisão
- **"Tipo da folha: ADIANTAMENTO"** - Adiantamentos salariais

### 🔍 Lógica de Filtro:
1. **Prioridade 1**: Procura linha específica "Tipo da folha:"
2. **Prioridade 2**: Se não encontrar, aplica filtro no cabeçalho (primeiras 10 linhas)
3. **Comportamento**: Termos como "FÉRIAS" ou "13º SALÁRIO" em valores são ignorados

---

## 2. 📊 Mapeamento de Códigos PDF → Excel ✅ IMPLEMENTADO v3.2

### 🔵 FOLHA NORMAL - Obter da Coluna ÍNDICE (Penúltimo número da linha)

| Código PDF | Descrição | Excel Coluna | Status | Observações |
|------------|-----------|--------------|--------|-------------|
| `01003601` | PREMIO PROD. MENSAL | **X** (PRODUÇÃO) | ✅ | ⚡ **FALLBACK**: Se ÍNDICE vazio, usar VALOR |
| `01007301` | HORAS EXT.100%-180 | **Y** (INDICE HE 100%) | ✅ | 🕐 Suporta formato horas |
| `01009001` | ADIC.NOT.25%-180 | **AC** (INDICE ADC. NOT.) | ✅ | 🕐 Suporta formato horas |
| `01003501` | HORAS EXT.75%-180 | **AA** (INDICE HE 75%) | ✅ | 🕐 Suporta formato horas |
| `02007501` | DIFER.PROV. HORAS EXTRAS 75% | **AA** (INDICE HE 75%) | ✅ | 🕐 Código alternativo para HE 75% |

### 🔴 FOLHA NORMAL - Obter da Coluna VALOR (Último número da linha)

| Código PDF | Descrição | Excel Coluna | Status | Observações |
|------------|-----------|--------------|--------|-------------|
| `09090301` | SALARIO CONTRIB INSS | **B** (REMUNERAÇÃO RECEBIDA) | ✅ | - |

### 🟡 13º SALÁRIO - Obter da Coluna VALOR com Fallback ✅ NOVO v3.2

| Código PDF | Descrição | Excel Coluna | Status | Observações |
|------------|-----------|--------------|--------|-------------|
| `09090301` | SALARIO CONTRIB INSS | **B** (REMUNERAÇÃO RECEBIDA) | ✅ | **PRIORIDADE 1** |
| `09090101` | REMUNERACAO BRUTA | **B** (REMUNERAÇÃO RECEBIDA) | ✅ | **FALLBACK** se 09090301 não encontrado |

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

### 🆕 Mapeamento de Linhas v3.2:
- **Linhas 1-65**: FOLHA NORMAL ✅
- **Linhas 67+**: 13º SALÁRIO ✅

---

## 5. 🔄 Lógica de Fallback ✅ IMPLEMENTADO v3.2

### Para Código 01003601 (PRODUÇÃO) - FOLHA NORMAL:
1. **Prioridade 1**: Tentar coluna ÍNDICE (penúltimo número) ✅
2. **Prioridade 2**: Se ÍNDICE vazio/zero, usar coluna VALOR (último número) ✅
3. **Resultado**: Sempre tentar preencher a coluna X (PRODUÇÃO) ✅

### 🆕 Para 13º SALÁRIO - Códigos 09090301/09090101:
1. **Prioridade 1**: Procurar 09090301 (SALARIO CONTRIB INSS) ✅
2. **Prioridade 2**: Se não encontrado ou valor zero, usar 09090101 (REMUNERACAO BRUTA) ✅
3. **Resultado**: Sempre tentar preencher a coluna B (REMUNERAÇÃO RECEBIDA) ✅

### Exemplo FOLHA NORMAL:
```
Linha PDF: P 01003601 PREMIO PROD. MENSAL 1.2 00,30,030 0,00 1.203,30
                                                        ^^^^  ^^^^^^^^
                                                      ÍNDICE   VALOR
                                                     (vazio)  (usar este) ✅
```

### 🆕 Exemplo 13º SALÁRIO:
```
Páginas 13º SALÁRIO: Procurar 09090301 primeiro
Se não encontrar ou for zero: Usar 09090101 como fallback ✅
```

---

## 6. 🎯 Códigos Duplicados ✅ IMPLEMENTADO

### INDICE HE 75% (Coluna AA):
Pode ser preenchido por **dois códigos diferentes**:
- `01003501` - HORAS EXT.75%-180 ✅
- `02007501` - DIFER.PROV. HORAS EXTRAS 75% ✅

**Comportamento**: Aceitar ambos, última ocorrência sobrescreve ✅

---

## 7. 📤 Regras de Output ✅ IMPLEMENTADO v3.2

### 🖥️ Interface Gráfica (desktop_app.py):
- **Popup de progresso** em tempo real com logs detalhados ✅
- **Mensagem de sucesso** com estatísticas completas ✅
- **Histórico persistido** de todos os processamentos ✅
- **Navegação automática** para aba apropriada em caso de erro ✅

### 🖥️ Linha de Comando (pdf_to_excel_updater.py):
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

## 10. 🏗️ Arquitetura v3.2 ✅ IMPLEMENTADO

### ✅ Modo Duplo v3.2:
- **Interface Gráfica**: `desktop_app.py` - Interface moderna com abas ✅
- **Linha de Comando**: `pdf_to_excel_updater.py` - CLI tradicional ✅
- **Core Processamento**: `pdf_processor_core.py` - Lógica compartilhada ✅

### ✅ Interface Gráfica v3.2:
- **CustomTkinter** com tema escuro moderno ✅
- **Sistema de abas**: Processamento/Histórico/Configurações ✅
- **Drag & Drop** de arquivos PDF (tkinterdnd2) ✅
- **Popup de progresso** com logs em tempo real ✅
- **Histórico persistido** entre sessões ✅
- **Validação automática** de configurações ✅

### ✅ Diretório de Trabalho:
- **Configuração obrigatória** via MODELO_DIR no .env ✅
- **PDF e MODELO.xlsm** no mesmo diretório ✅
- **Pasta DADOS/** criada automaticamente ✅
- **Execução de qualquer local** ✅

### ✅ Persistência v3.2:
- **config.json**: Configurações da aplicação ✅
- **history.json**: Histórico completo de processamentos ✅
- **Sessões múltiplas**: Mantém dados entre reinicializações ✅

---

## 11. 🆕 Funcionalidades v3.2 ✅ IMPLEMENTADO

### ✅ Sistema de Histórico Persistido:
- **Histórico automático** de todos os processamentos ✅
- **Persistência** em arquivo JSON entre sessões ✅
- **Detalhes completos**: logs, resultados, timestamps ✅
- **Abertura direta** dos arquivos processados ✅
- **Limpeza de histórico** com confirmação ✅

### ✅ Processamento Dual:
- **FOLHA NORMAL**: Linhas 1-65 com mapeamento completo ✅
- **13º SALÁRIO**: Linhas 67+ com fallback inteligente ✅
- **Categorização automática** de páginas por tipo ✅
- **Regras específicas** para cada tipo de folha ✅

### ✅ Interface Moderna:
- **Design responsivo** com CustomTkinter ✅
- **Abas organizadas** para diferentes funcões ✅
- **Validação em tempo real** de configurações ✅
- **Feedback visual** com ícones e cores de status ✅

### ✅ Detecção Inteligente:
- **Nome da pessoa** extraído automaticamente do PDF ✅
- **Nomeação automática** do arquivo final ✅
- **Validação de diretório** em tempo real ✅
- **Lista automática** de PDFs disponíveis ✅

---

## 12. 📋 Resumo de Implementação v3.2

### ✅ TOTALMENTE IMPLEMENTADO:
- ✅ Filtro de folhas por "Tipo da folha" (FOLHA NORMAL + 13º SALÁRIO)
- ✅ Mapeamento completo de todos os códigos para FOLHA NORMAL
- ✅ **NOVO**: Mapeamento específico para 13º SALÁRIO com fallback
- ✅ Fallback ÍNDICE → VALOR para 01003601 (PRODUÇÃO)
- ✅ **NOVO**: Fallback 09090301 → 09090101 para 13º SALÁRIO
- ✅ Conversão automática formato horas (`06:34` → `06,34`)
- ✅ Planilha padrão obrigatória "LEVANTAMENTO DADOS"
- ✅ **NOVO**: Mapeamento de linhas (1-65: FOLHA NORMAL, 67+: 13º SALÁRIO)
- ✅ Preservação total de macros VBA (.xlsm)
- ✅ **NOVO**: Interface gráfica moderna com CustomTkinter
- ✅ **NOVO**: Sistema de histórico persistido
- ✅ **NOVO**: Drag & Drop de arquivos PDF
- ✅ **NOVO**: Popup de progresso com logs em tempo real
- ✅ **NOVO**: Persistência de configurações
- ✅ Output simplificado e conciso
- ✅ Tratamento completo de erros
- ✅ Diretório de trabalho obrigatório
- ✅ Detecção automática do nome da pessoa

### 🎯 REGRAS DE NEGÓCIO: 100% CONFORMES v3.2

Todas as regras especificadas foram implementadas e testadas na versão 3.2 do PDF para Excel Updater, incluindo:

- **Processamento dual** (FOLHA NORMAL + 13º SALÁRIO)
- **Interface gráfica completa** com histórico persistido
- **Fallbacks inteligentes** para todos os códigos críticos
- **Arquitetura modular** com core de processamento compartilhado

---

## 🚀 Novidades da Versão 3.2

1. **🖥️ Interface Gráfica Moderna**: CustomTkinter com tema escuro
2. **📊 Histórico Persistido**: Todas as operações são salvas automaticamente
3. **🎯 Drag & Drop**: Arraste PDFs diretamente na interface
4. **⚙️ Configurações Automáticas**: Persistência de todas as configurações
5. **📱 Sistema de Abas**: Organização intuitiva das funcionalidades
6. **🔄 Processamento Dual**: FOLHA NORMAL + 13º SALÁRIO simultaneamente
7. **🎨 Validação em Tempo Real**: Feedback imediato das configurações
8. **📝 Logs Detalhados**: Popup com progresso e logs em tempo real

**A versão 3.2 representa a evolução completa do sistema para uma aplicação moderna e robusta, mantendo 100% de compatibilidade com as regras de negócio estabelecidas.**