---
name: atualizar-deck-peers
description: Atualiza os PowerPoints de Análise de Peers da Dexco (Louças, Metais, Madeira, Revestimento; decks do Offsite/benchmarking) com os números atuais da planilha Análise de Competidores — os mesmos do painel — reconhecendo cada série pelo nome da empresa e pelos valores (a planilha muda de lugar com o tempo), mantendo o vínculo e a formatação dos gráficos, e entrega uma conferência com valores revisados, séries não reconhecidas, alertas de outlier/n.s. e textos provisórios ou copiados de outro deck. Use sempre que pedirem para atualizar, refazer ou conferir o deck/apresentação/slides de peers, de competidores ou de benchmarking de um negócio, ou levar os números novos da planilha de competidores para o PowerPoint.
---

# Atualizar os decks de Análise de Peers

Os decks de peers (um por negócio) têm dezenas de gráficos vinculados à planilha de competidores
("Book Competidores.xlsx" / "Análise de Competidores"). Com o tempo a planilha ganha linhas e blocos,
e os endereços antigos passam a apontar para **outro dado**: clicar em "Atualizar links" pode trocar
a margem de uma empresa pela de outra sem ninguém perceber. Esta skill resolve isso: reconhece cada
série pelo conteúdo e regrava valores e endereços corretos.

Converse em português.

## O que você precisa

1. **Deck** do negócio (.pptx), a edição mais recente.
2. **Planilha Análise de Competidores** (.xlsx) atualizada. Nunca edite a planilha.
3. **Pasta do projeto `painel-estrategico`** (opcional, recomendada): com ela, a conferência avisa
   quando o deck mostra um valor que o painel trata como erro, outlier ou n.s.

## Como rodar

```bash
python <pasta desta skill>/scripts/atualizar_deck_peers.py --projeto <painel-estrategico> \
  --planilha "<Análise de Competidores>.xlsx" --deck "<Análise de Peers - Louças>.pptx" \
  --saida "Analise_de_Peers_Loucas_<mes>.pptx"
```

Gera o .pptx e a conferência (`<saida>_conferencia.md`) e imprime um resumo. O negócio é detectado
pelas abas que os gráficos usam. Rode uma vez por deck.

## Como cada série é reconhecida (em ordem)

| Regra | Quando |
|---|---|
| Nome e valores | Mesma empresa e ≥ 60% dos valores iguais numa linha de um bloco das abas "Gráficos" |
| Valores (nome diferente) | Todos os valores batem com uma única linha (ex.: "Deca" × "Deca Louças") |
| Mesmo endereço e nome | A linha antiga ainda tem o nome da empresa; valores revisados |
| Nome a poucas linhas | Nenhuma série do gráfico bateu, mas todas acham o próprio nome deslocadas do mesmo número de linhas |
| Deslocamento do gráfico | As demais séries do gráfico seguem o deslocamento das já reconhecidas, se o nome bater |
| Não reconhecida | Valores antigos ficam **fixos** no gráfico, sem vínculo (para o "Atualizar links" não pôr dado de outra empresa); vai para a conferência |

Também: períodos que a planilha não tem mais (ex.: coluna 1T26) saem do gráfico; anos novos com dado
(ex.: 2026) entram automaticamente; o rótulo do último ponto acompanha o dado novo.

## Como entregar

1. Entregue o .pptx e a conferência.
2. No chat, resuma por deck: séries atualizadas, com valores revisados (cite as principais), não
   reconhecidas, alertas de outlier/n.s. e a lista de textos para revisar.
3. Avise que os **textos** (títulos, caixas de comentário, tabela-resumo do slide 2) não são
   atualizados: a conferência aponta textos provisórios ("[WIP]", "Lorem ipsum", "x") e textos que citam
   empresas de outro negócio (sinal de cópia de outro deck). Ofereça escrever a leitura nova.
4. Vínculo: o deck continua apontando para o arquivo original (a conferência mostra o nome). Se a
   planilha do time mudou de nome ou de pasta, no PowerPoint use **Arquivo → Informações → Editar
   links para arquivos → Alterar fonte** antes de atualizar links.

## Cuidados

- Valores revisados em massa (ex.: Madeira) costumam ser decisão de base, não erro do script: aponte
  e pergunte se o time quer manter a versão da planilha (veja "Decisões em aberto" no CLAUDE.md).
- Gráficos com dados colados ou de outros arquivos (econométrico, INCAF, desempenho de mercado) não
  são atualizados; a conferência lista os slides.
- O LibreOffice (prévias) mostra alguns rótulos diferente do PowerPoint; confira no PowerPoint antes
  de concluir que algo quebrou.
