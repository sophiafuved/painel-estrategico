---
name: padrao-apresentacao-dexco
description: Padrão visual e regras de gráfico das apresentações do time de estratégia da Dexco — o deck do Radar Estratégia (cenário macro) e o template do Offsite/Análise de Peers — com template PowerPoint, cores, fontes, layouts e convenções (Dexco em preto, dois eixos, escalas, notas e fontes no rodapé). Use sempre que alguém do time pedir slides, deck, apresentação, gráfico para PowerPoint, "no padrão do Radar", "no template do Offsite", um slide de benchmarking de peers, ou transformar o painel/dashboard em apresentação, mesmo que não mencione o template.
---

# Padrão de apresentação Dexco (time de estratégia)

Há dois padrões, e escolher o certo é o primeiro passo:

| Pedido | Padrão | Referência |
|---|---|---|
| Cenário macro, juros, renda, crédito, INCC, imobiliário, varejo; "Radar" | **Radar Estratégia** (fundo escuro, Trebuchet, laranja) | `references/radar.md` |
| Concorrentes, benchmarking, peers, Offsite, negócios (Louças, Metais, Madeira, Revestimento) | **Offsite / Peers** (template com barra lateral, Roboto, Dexco em preto) | `references/offsite_peers.md` + `assets/template_offsite_peers.pptx` |

Se não estiver claro, pergunte. Leia a referência do padrão escolhido antes de montar qualquer
slide. Para gerar o arquivo .pptx, use também a skill de PowerPoint disponível (pptx).

Converse em português.

## Princípios comuns aos dois padrões

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
