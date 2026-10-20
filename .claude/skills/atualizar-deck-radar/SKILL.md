---
name: atualizar-deck-radar
description: Atualiza o PowerPoint do Radar Estratégia da Dexco (Modelo 2) com os dados atuais da Grande Base — os mesmos do painel — e com a leitura do time (títulos, textos de apoio e caixas), mantendo o vínculo dos gráficos com a Grande Base no SharePoint, a formatação dos slides e os gráficos combinados de dois eixos, e entrega uma lista de conferência. Use sempre que pedirem para atualizar, montar ou gerar a apresentação/deck/slides do Radar, levar o painel e os comentários para o PowerPoint, preparar o Radar do mês, ou "passar o dashboard para a apresentação".
---

# Atualizar o deck do Radar (Modelo 2)

O deck do Radar já é o formato que a liderança conhece; por isso esta skill **atualiza o deck
existente** em vez de criar outro. Os 20 gráficos do Modelo 2 estão vinculados à Grande Base no
SharePoint (atualização manual de links); o script grava os dados novos nos gráficos **sem
quebrar esse vínculo** e ainda estende as faixas vinculadas quando o painel ganhou meses novos.

Converse em português.

## O que você precisa

1. **Deck base**: a edição mais recente do "Radar Estratégia – Modelo 2" (.pptx). Use sempre a
   última edição como base, para herdar os ajustes manuais do time.
2. **Grande Base** (.xlsx) atualizada.
3. **Pasta do projeto `painel-estrategico`**: o script usa as faixas do painel
   (`pipeline/radar/faixas.py`), garantindo que deck e painel mostrem os mesmos dados.
4. **Leitura do time** (`pipeline/radar/leitura_time.json`), escrita com a skill
   `comentarios-radar`. Se ainda não existir para esta edição, ofereça escrevê-la antes; sem ela,
   rode com `--sem-textos` (só dados).

## Como rodar

```bash
python <pasta desta skill>/scripts/atualizar_deck.py --projeto <painel-estrategico> \
  --grande-base "<Grande Base>.xlsx" --deck "<Radar ... Modelo 2>.pptx" \
  --saida "Radar_Estrategia_<mes>_Modelo2.pptx" [--capa "novembro, 2026"] [--sem-textos]
```

- Gera o .pptx e uma conferência em Markdown (`<saida>_conferencia.md`) e imprime um resumo.
- `--listar` mostra os ids das caixas de texto e dos gráficos de um deck (use se o modelo mudar e
  for preciso atualizar `scripts/deck_map_modelo2.json`).

## O que o script faz

| Item | Como |
|---|---|
| Dados dos gráficos | Relê cada série na Grande Base; se a faixa do painel cresceu, estende a faixa vinculada do gráfico (o "Atualizar links" do PowerPoint também passa a pegar o mês novo) |
| Gráficos combinados (colunas + linha, dois eixos) | Mantém tipo, cores e eixos; só troca os valores |
| Rótulo do último ponto | Move para o último dado novo |
| "Dados atualizados até mmm/aa" | Troca pela data do último dado observado (sem projeções); com duas séries, associa cada data à série certa |
| Títulos, textos e caixas | Da leitura do time, com a formatação do slide (caixa alta, destaques em negrito/laranja). Texto em lista vira um bullet por item, copiando marcador, recuo e alinhamento dos bullets do modelo; caixas de leitura ficam centralizadas como no original |
| Ajustes de layout | Título que ocupa 2 linhas: alinha à esquerda, termina antes do marcador verde e desce os bullets o necessário; a fonte só diminui se o texto não couber (e a conferência avisa) |
| Leitura em rascunho | Põe o carimbo "RASCUNHO · texto aguardando revisão do time" nos slides |

O que **não** faz e fica na conferência para revisar à mão: Brent (aba "Daily", fora da Grande
Base), impacto fiscal (dados fixos), rótulos digitados à mão sobre os gráficos (ex.: "5,9%"),
slide "Principais mensagens" e capa (use `--capa`).

## Como entregar

1. Entregue o .pptx e a conferência.
2. No chat, resuma: séries atualizadas e estendidas, dados revisados na fonte, datas trocadas,
   textos trocados (e se estão em rascunho) e a lista "Revisar à mão" — principalmente os rótulos
   digitados à mão, que ficam desatualizados quando os dados mudam.
3. Lembre que, ao abrir no PowerPoint, aparece o aviso de vínculos: pode **atualizar** (puxa da
   Grande Base no SharePoint, mesmo resultado) ou **cancelar** (mantém o que o script gravou).
4. Se a leitura estiver em rascunho, diga que o deck não deve circular antes da revisão do time;
   depois da aprovação, gere de novo (o carimbo some).

## Cuidados

- Nunca invente dado para preencher um gráfico; se uma aba ou faixa não existir, o script mantém
  o que estava e avisa.
- Títulos do Modelo 2 cabem em até 2 linhas a 20 pt (cerca de 12–18 palavras). Se a conferência
  avisar fonte reduzida, peça para encurtar o título na leitura em vez de aceitar letra pequena.
- Textos de apoio do modelo são 2–3 bullets; caixas de leitura têm ~15–20 palavras. Siga esse
  tamanho na leitura (skill `comentarios-radar`).
- O LibreOffice (usado para gerar imagens de conferência) mostra alguns rótulos e eixos de forma
  diferente do PowerPoint (ex.: "0.6174", eixo "1220%"); compare com o original renderizado do
  mesmo jeito antes de concluir que algo quebrou.
