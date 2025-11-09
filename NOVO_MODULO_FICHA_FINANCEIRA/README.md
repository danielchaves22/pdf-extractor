# 📄 Módulo "Ficha Financeira"

Este módulo processa as fichas financeiras quadrimestrais e gera arquivos CSV padronizados. A primeira entrega contempla o `PROVENTOS.csv`, respeitando o mesmo layout disponibilizado na pasta `NOVO_MODULO_FICHA_FINANCEIRA`.

## 🧭 Fluxo de uso na interface desktop

1. Abra a aplicação normalmente. Após o splash screen, selecione **Ficha Financeira**.
2. Informe o período desejado (mês/ano inicial e final).
3. Adicione um ou mais PDFs da ficha financeira. Todos os valores encontrados são consolidados, mês a mês.
4. Clique em **Gerar PROVENTOS.csv** para criar o arquivo no diretório do primeiro PDF.

A janela mantém um painel de logs para acompanhar o andamento da extração.

## 🧮 Regras de extração

- Cada página é dividida em blocos de quatro meses. O último bloco do ano possui uma coluna adicional de totais que é ignorada.
- Os valores são extraídos pela posição das colunas `Comp.` e `Valor` identificadas no PDF. A rotina reconhece automaticamente as colunas, mesmo quando determinados campos estão vazios.
- Para o `PROVENTOS.csv`, são coletados os valores da verba **`3123 - Base INSS (Folha)`** na coluna `Valor`. Se o mês estiver ausente no PDF, o resultado é preenchido com `0`.
- Meses fora do intervalo solicitado são descartados, mesmo que existam no documento.
- As colunas `FGTS`, `FGTS_REC.`, `CONTRIBUICAO_SOCIAL` e `CONTRIBUICAO_SOCIAL_REC.` são preenchidas com `N`, conforme especificação.

## 🗂️ Estrutura do arquivo gerado

O arquivo segue o padrão de separador `;` e contém dez colunas:

```
MES_ANO;VALOR;FGTS;FGTS_REC.;CONTRIBUICAO_SOCIAL;CONTRIBUICAO_SOCIAL_REC.;;;;
```

Cada linha representa um mês no formato `MM/AAAA`, com os valores convertidos para vírgula como separador decimal.

## 🔁 Reutilização futura

O processador `FichaFinanceiraProcessor` centraliza a lógica de leitura das fichas e facilita a inclusão de novas verbas ou CSVs. Os métodos responsáveis por mapear colunas e agrupar meses foram pensados para servir outros arquivos de saída que venham a ser necessários.

