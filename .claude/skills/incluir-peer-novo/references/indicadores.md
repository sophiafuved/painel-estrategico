# Definições dos indicadores (como a planilha Análise de Competidores calcula)

As abas de cada empresa (ex.: `Icasa`, `Deca - Louças`) têm duas partes:

- **Bloco de entrada (a partir da linha ~124), na moeda original e em milhares**: DRE, Balanço
  (BP) e Conversão de Caixa. É aqui que se digitam os números da demonstração.
- **Bloco de cima (linhas 1–110), em R$ mil**: convertido automaticamente pela aba `Câmbio`
  (coluna "Média do ano" do par Ano&Moeda, ex.: `2025USD`) quando a moeda da empresa não é BRL.
  Os indicadores são calculados a partir daí.

As abas "Gráficos" puxam esses indicadores em **R$ milhões** (valores) ou em % / x / dias.

## DRE (entrada)

| Linha | Conteúdo | Sinal |
|---|---|---|
| Receita Líquida | receita líquida consolidada do segmento/empresa | + |
| Custo de Venda | CPV / custo dos produtos vendidos | − |
| Lucro Bruto | Receita − CPV | |
| G&A | despesas gerais e administrativas | − |
| Despesa de Vendas | despesas com vendas | − |
| Outros Resultados Operacionais | outras receitas/despesas operacionais | ± |
| EBIT | resultado operacional antes do financeiro | |
| Outros Resultados Não-operacionais, Receitas Financeiras, Despesas Financeiras | | ± |
| EBT, IR/CSLL, Lucro Líquido | | |
| D&A | depreciação e amortização (nota explicativa ou DFC) | + |
| EBITDA | EBIT + D&A | |
| Capex (Imobilizado), Capex (Intangível) | aquisições de imobilizado e intangível (DFC) | + (valor absoluto) |

## Balanço (entrada)

Ativo: Caixa, Clientes, Estoque, Outros, Imobilizado, Intangível, Investimentos…
Passivo: Fornecedores, Empréstimos CP, Empréstimos LP, Outros; Patrimônio Líquido.
A linha **Check** (Ativo − Passivo − PL) precisa dar ~0.

## Conversão de caixa (entrada)

EBITDA; (+/−) Variação WK; (−) Capex; (−) Imposto; (+/−) Fluxo Financeiro; (+/−) Outros;
**FCL = soma das linhas acima**. Consumo de caixa é negativo.

## Indicadores calculados (não digitar; conferir se fazem sentido)

| Indicador | Fórmula |
|---|---|
| Margem bruta / EBITDA / líquida | Lucro bruto, EBITDA, Lucro líquido ÷ Receita |
| Capital investido | PL + Empréstimos CP + Empréstimos LP − Caixa |
| NOPAT | EBIT × (1 − alíquota do país, aba `Global Tax Rate`) |
| ROIC (teórico) | NOPAT ÷ Capital investido |
| Giro de ativos | Receita ÷ Capital investido |
| Alavancagem (DL/EBITDA) | (Empréstimos CP + LP − Caixa) ÷ EBITDA |
| Dias de recebíveis | Clientes × 365 ÷ Receita |
| Dias de estoque | Estoque × 365 ÷ CPV |
| Dias de fornecedores | Fornecedores × 365 ÷ CPV |
| Capital de giro / ROL | (Clientes + Estoque − Fornecedores) ÷ Receita |
| CPV caixa / ROL | −(CPV + D&A) ÷ Receita (CPV sem depreciação) |
| Desp. vendas + outras / ROL | −(Desp. vendas + Outros operacionais) ÷ Receita |
| G&A / ROL, D&A / ROL, Capex / ROL | ÷ Receita |
| FCL / EBITDA, Var. WK / EBITDA, Capex / EBITDA, Outras / EBITDA | ÷ EBITDA; "NA" quando EBITDA < 0 |
| FCL / ROL, Var. WK / ROL, Outras Desp. / ROL | ÷ Receita (blocos opcionais nas abas "Gráficos") |

## Formato dos blocos nas abas "Gráficos"

```
      B (país)   C (título/empresa)     D      E      F   ...
r     -          ROIC                   2019   2020   2021 ...
r+1   BR         Deca - Louças          0,05   0,03   ...
r+2   BR         Icasa                  ...
r+3   JP         Toto                   ...
```

- Título na coluna C e anos nas colunas ao lado, na mesma linha; uma linha por empresa logo
  abaixo, sem linha em branco no meio. País na coluna B (BR, JP, ES, US, IT, SUI, AL, MX).
- Valores em R$ milhões (blocos de valor) ou em fração (0,05 = 5%). Célula vazia = sem dado; o
  painel também lê zero como sem dado.
- Para incluir uma empresa num bloco existente, acrescente a linha **no fim do bloco**; para um
  ano novo, acrescente a coluna **à direita**. O painel lê os dois automaticamente.
