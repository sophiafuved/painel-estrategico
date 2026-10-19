---
name: incluir-peer-novo
description: Extrai os números de um concorrente (peer) da Dexco a partir de demonstrações financeiras, releases de resultados, ITR/DFP, relatórios anuais ou PDFs/planilhas de RI e os prepara no formato da planilha Análise de Competidores (aba da empresa + blocos das abas "Gráficos"), com fonte e página de cada número. Use sempre que pedirem para incluir um peer novo, atualizar um ano de um concorrente, lançar a demonstração de 2025 de uma empresa, preencher dados faltantes de peers (como Docol, Fabrimar, Gessi, Berneck, Porcelanosa) ou montar o benchmarking de uma empresa nova de louças, metais, madeira/painéis ou revestimentos.
---

# Incluir ou atualizar um peer

A planilha Análise de Competidores é a fonte única do benchmarking. O trabalho aqui é **levar os
números da fonte para o formato da planilha, com rastreabilidade**, para que alguém do time cole
os valores e qualquer pessoa consiga conferir de onde veio cada número depois.

Converse em português. Leia `references/indicadores.md` antes de começar: ele tem as linhas de
entrada da aba da empresa, as fórmulas dos indicadores e o formato dos blocos das abas "Gráficos".

## Passo a passo

1. **Entenda o pedido**: empresa, negócio (Louças, Metais, Madeira, Revestimento), anos e se é
   peer novo (precisa de aba nova) ou atualização de ano (aba já existe).
2. **Leia as fontes** que o usuário enviou (use as skills de PDF/planilha quando for o caso). Prefira,
   nesta ordem: DFP/ITR ou demonstrações auditadas → release de resultados → apresentação de RI.
   Se o segmento relevante for só parte da empresa (ex.: divisão de louças de um grupo), use a
   abertura por segmento e diga isso.
3. **Monte o bloco de entrada** (moeda original, em milhares, na ordem das linhas da aba da
   empresa): DRE, D&A, Capex, Balanço e Conversão de caixa. Cada número com **fonte, página e linha
   da demonstração**. Respeite os sinais (custos e consumo de caixa negativos).
4. **Confira antes de entregar**:
   - Balanço fecha (Ativo = Passivo + PL) e EBITDA = EBIT + D&A.
   - FCL = EBITDA + Var. WK − Capex − Imposto ± Fluxo financeiro ± Outros.
   - Calcule os indicadores pelas fórmulas da referência e compare com o ano anterior da própria
     empresa e com os pares; saltos grandes (margem ± 15 p.p., prazos ± 60 dias, receita ± 40%)
     precisam de explicação (aquisição, mudança contábil, câmbio, evento não recorrente).
   - Moeda: anote a moeda original; a conversão para R$ é feita pela planilha (aba `Câmbio`, média
     do ano). Se o par Ano&Moeda não existir na aba `Câmbio`, avise que precisa ser incluído.
5. **Entregue** (veja o formato abaixo) e diga o que faltou.

## Formato da entrega

Uma planilha Excel (ou tabela no chat, se for pouca coisa) com:

- **Aba "Entrada"**: linhas na ordem da aba da empresa, uma coluna por ano, valores na moeda
  original em milhares, e as colunas `Fonte`, `Página` e `Observação`.
- **Aba "Blocos Gráficos"**: os valores já no formato das abas "Gráficos" (R$ milhões ou fração),
  uma linha por indicador e ano, indicando em qual bloco entram (ex.: "Gráficos - Metais → ROIC").
  Use isso para conferência; na planilha, esses blocos se preenchem pelas fórmulas da aba da empresa.
- **Aba "Checagens"**: balanço fecha, EBITDA, FCL, indicadores vs ano anterior, e alertas.

No chat, resuma: empresa, anos, fontes usadas, principais números (receita, EBITDA, margem, ROIC) e
os pontos que o time precisa decidir.

## Cuidados

- Nunca estime nem "complete" um número que não está na fonte. Deixe vazio e liste em "o que
  faltou"; o painel trata vazio como sem dado. Números redondos (ex.: R$ 1.100 mi) chamam atenção
  na conferência justamente porque costumam ser estimativas.
- Não edite a planilha do time diretamente; entregue os valores para alguém colar, a não ser que
  o usuário peça explicitamente e haja uma cópia de trabalho.
- Se a empresa reporta EBITDA ajustado e reportado, use o mesmo critério dos demais peers do bloco
  (pergunte se não souber) e anote qual foi usado.
- Depois que o time colar os números, sugira rodar a skill `conferir-peers` com a versão anterior
  da planilha para ver exatamente o que mudou.
