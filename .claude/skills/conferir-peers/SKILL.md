---
name: conferir-peers
description: Confere os números da Análise de Peers da Dexco (planilha Análise de Competidores) — confere cada valor do painel com a célula de origem, aponta saltos, inconsistências, números redondos, dados faltando e o que mudou desde a versão anterior — e entrega uma planilha de conferência para o time preencher. Use sempre que pedirem para conferir, revisar, checar, validar ou fazer double check dos números de peers/concorrentes, quando sair uma demonstração financeira nova de um peer, quando a planilha de competidores for atualizada, ou antes de publicar o painel ou apresentar o benchmarking.
---

# Conferir os números de peers

O objetivo é dar ao time uma lista curta e priorizada do que precisa ser checado na fonte, em vez
de reler milhares de células. O painel só mostra o que está na planilha; então a conferência é
sobre **a planilha estar certa em relação às demonstrações financeiras**.

Converse em português.

## Como rodar

1. Localize a planilha Análise de Competidores pelo conteúdo (abas `Gráficos - Louças`,
   `Gráficos - Metais`, `Gráficos Madeira 1`, `Gráficos - Revestimento`). Nunca a edite.
2. Precisa da pasta do projeto `painel-estrategico` (o script usa os mesmos leitores do painel,
   em `pipeline/peers`, para que a conferência veja exatamente o que o painel mostra). Se não
   estiver disponível, peça ao usuário para conectar a pasta ou enviar o ZIP do projeto.
3. Rode o script da skill:

```bash
python <pasta desta skill>/scripts/conferir_peers.py --projeto <pasta do painel-estrategico> \
  --planilha "<Análise de Competidores>.xlsx" \
  [--anterior "<versão anterior>.xlsx"] [--em-aberto em_aberto.json] --saida Conferencia_Peers.xlsx
```

- `--anterior`: use quando o usuário tiver a versão antiga (ex.: a do mês passado). Gera a aba
  "Mudanças" com valores novos, alterados e que saíram do painel; é o mais útil quando entra a
  demonstração de um peer.
- `--em-aberto`: JSON com os pontos que dependem de decisão do time
  (`[{"tema": "...", "visto": "...", "decisao": "..."}]`). Se o projeto tiver uma lista de decisões
  em aberto (no `CLAUDE.md`/`README.md` ou `PENDENCIAS.md`), transforme-a nesse arquivo.

O script imprime um resumo em JSON e grava a planilha. Entregue o arquivo ao usuário.

## O que a planilha traz

- **Leia-me**: como usar, critérios dos alertas e um quadro de andamento que se atualiza conforme
  o time marca o status.
- **Pontos de atenção**: um alerta por linha, com prioridade (Alta = Dexco), célula de origem,
  e as colunas amarelas Status/Comentário para o time preencher.
- **Mudanças** (com `--anterior`).
- **Base completa**: todos os valores do painel com célula de origem e tratamento (ok, sem dado,
  n.s. por erro/outlier, n.s. por EBITDA negativo, calculado, consolidada).
- **Em aberto** (com `--em-aberto`).

## Como apresentar o resultado

Comece pelo que dá confiança e depois vá ao que precisa de ação:

1. **Leitura**: quantos valores foram conferidos com a célula e quantas divergências houve
   (`divergencias_de_leitura`; o esperado é 0 — se não for, é bug de leitura do painel: investigue
   antes de qualquer outra coisa).
2. **Prioridade alta** (todos da Dexco): liste os principais em linguagem de negócio (ex.: "Deca
   Metais 2025: capital de giro/ROL sobe de 32% para 49% e FCL/EBITDA vai a −98%"), agrupando por
   negócio. Não liste os 200 alertas no chat; eles estão na planilha.
3. **Lacunas**: empresas sem dado no último ano — diga quais e lembre que, se a demonstração já
   saiu, basta preencher.
4. **Mudanças** (se houver versão anterior): quantos valores mudaram e os mais relevantes.
5. **Próximo passo**: quando o time corrigir a planilha, gerar de novo e republicar (skill
   `atualizar-painel-estrategico`).

## Cuidados de interpretação

- Um salto não é erro por si só: 2021 foi um ano de boom em Madeira; aquisições mudam escala; o
  câmbio mexe nos números das estrangeiras. Aponte, não conclua.
- Diferença de 3–7% entre o giro da planilha e receita/capital investido em estrangeiras vem do
  câmbio usado em cada bloco — não é erro (o script só alerta acima de 10%).
- Índices sobre EBITDA com EBITDA negativo já aparecem como n.s. no painel; não são lacunas.
- Não "corrija" números por conta própria nem invente fonte. Se o usuário pedir para checar um
  valor contra um release/PDF, leia o documento e cite página e linha.
