---
name: padrao-apresentacao-dexco
description: Padrão visual e regras de gráfico das apresentações do time de estratégia da Dexco — o deck do Radar Estratégia (cenário macro), o Template PPT 2026 da Dexco (padrão atual dos decks de peers/benchmarking) e o template antigo do Offsite — com templates PowerPoint, script que leva um deck existente para o Template 2026, cores, fontes, layouts e convenções (Dexco em preto, dois eixos, escalas, notas e fontes no rodapé). Use sempre que alguém do time pedir slides, deck, apresentação, gráfico para PowerPoint, "no padrão do Radar", "no template do Offsite", um slide de benchmarking de peers, transformar o painel/dashboard em apresentação, ou pedir para "passar para o template novo/2026", mesmo que não mencione o template.
---

# Padrão de apresentação Dexco (time de estratégia)

Há dois padrões, e escolher o certo é o primeiro passo:

| Pedido | Padrão | Referência |
|---|---|---|
| Cenário macro, juros, renda, crédito, INCC, imobiliário, varejo; "Radar" | **Radar Estratégia** (fundo escuro, Trebuchet, laranja) | `references/radar.md` |
| Concorrentes, benchmarking, peers, negócios (Louças, Metais, Madeira, Revestimento) | **Template PPT 2026** (padrão atual, out/2026): barra preta estreita com logo 75 anos, etiqueta cinza no topo, título 20 pt | `assets/Template_PPT_2026.potx` + seção abaixo; regras de gráfico em `references/offsite_peers.md` |
| Edições antigas do Offsite 2026 | Offsite / Peers antigo (barra "OFFSITE 2026", títulos com realce amarelo) | `references/offsite_peers.md` + `assets/template_offsite_peers.pptx` |

Se não estiver claro, pergunte. Leia a referência do padrão escolhido antes de montar qualquer
slide. Para gerar o arquivo .pptx, use também a skill de PowerPoint disponível (pptx).

Converse em português.

## Template PPT 2026 (padrão atual dos decks de peers)

Slides-modelo do arquivo: capa (TÍTULO / Sub-título / | mês/ano |), "Principais mensagens", agenda
de seções (números em caixas cinzas, seção atual em branco e as demais em cinza), slide de conteúdo
(etiqueta cinza no topo à esquerda, ex.: "1) Louças: rentabilidade", e título em caixa alta 20 pt a
partir de x ≈ 1,13 pol.) e encerramento com as marcas. Conteúdo começa em y ≈ 1,2 pol., à direita da
barra preta (0,86 pol.).

**Levar um deck existente para o Template 2026** (ex.: os decks de peers):

```bash
python <pasta desta skill>/scripts/aplicar_template.py --deck "<deck>.pptx" --saida "<deck>_2026.pptx" \
  [--etiqueta "Análise de Peers: Louças"] [--data "Offsite 2026"]
```

- Capa → capa do template (título, negócio e data tirados da capa antiga); slide vazio → encerramento;
  divisor de seção → agenda; demais → slide de conteúdo, com o título antigo (negrito preservado,
  realce amarelo retirado) e o restante copiado na mesma posição.
- Fontes como no template: texto em **Roboto Light** e negritos em **Roboto Black** (sem negrito
  sintético por cima), inclusive nos gráficos; títulos e etiquetas usam as fontes do tema
  (+mn-lt = Roboto Light, +mj-lt = Roboto Black). Fórmulas e marcadores não mudam. `--manter-fontes`
  desliga a troca.
- Gráficos continuam vinculados à planilha (rode antes a skill `atualizar-deck-peers` para os
  números); imagens, tabelas, notas e links entre slides vêm junto.
- A conferência (`<saida>_template.md`) lista o que foi para onde e os slides com formas na faixa do
  título (lembretes de revisão, caixas altas): ajuste à mão.
- O conteúdo interno dos slides (caixas de comentário com realce amarelo, cores dos gráficos) é
  mantido como estava; se o time quiser também mudar esse estilo, combine antes.

## Princípios comuns aos padrões

- **Uma mensagem por slide**: o título diz o "so what"; o gráfico prova. Para títulos do Radar,
  use a skill `comentarios-radar`.
- **Números rastreáveis**: todo gráfico tem unidade, "Dados atualizados até mmm/aa" quando for
  série temporal, e fonte no rodapé. Notas numeradas (¹ ² ³) explicam definições.
- **Dados da fonte oficial do time**: Grande Base (Radar) e Análise de Competidores (Peers). Não
  redigite números de memória nem de versões antigas de slides.
- **Dexco sempre identificável**: no Peers, preta e mais grossa; no Radar, a série principal em
  laranja.
- **Mantenha os gráficos que a liderança já conhece**: quando a apresentação original usa dois
  eixos ou combinação colunas + linha, preserve essa visão (decisão do gestor, out/2026), em vez de
  separar em dois gráficos.
- **Escalas honestas**: colunas começam no zero; linhas usam do menor ao maior valor; nunca corte
  um eixo de colunas.

## Fluxo sugerido

1. Confirme o padrão, o público (Offsite, comitê, time) e a lista de slides.
2. Puxe os dados da fonte (ou do painel publicado) e anote a data dos dados.
3. Monte os slides no template/edição base, um por mensagem.
4. Confira: títulos completos (sem texto escondido pela barra lateral), legendas legíveis,
   fonte/nota em todos os gráficos, Dexco destacada, nenhum dado de exemplo ou placeholder.
5. Entregue o .pptx e liste no chat o que foi usado como fonte e o que ficou pendente.
