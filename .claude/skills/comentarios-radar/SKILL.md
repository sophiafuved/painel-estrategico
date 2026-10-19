---
name: comentarios-radar
description: Redige os títulos de leitura, textos de apoio e caixas de comentário do Radar Estratégia da Dexco (cenário macro, juros, renda, crédito, INCC, imobiliário, varejo de material de construção) no estilo do time de estratégia, a partir dos dados do painel/Grande Base e das projeções da consultoria. Use sempre que pedirem comentários, títulos, mensagens-chave, "so what", leitura dos gráficos ou o texto dos slides do Radar, quando o painel for virar apresentação, ou quando pedirem para atualizar os comentários de uma edição anterior do Radar.
---

# Comentários do Radar Estratégia

O painel não tem comentários de propósito: a leitura é do time e entra só quando o painel vira
apresentação. Esta skill escreve essa leitura no tom das edições anteriores, apoiada nos números.

Converse em português. Leia `references/exemplos.md` para calibrar tom e tamanho.

## Antes de escrever

1. Reúna os números de cada tópico: último dado observado e variação (o resumo de cada tópico do
   painel já traz isso), projeções (Focus, Tendências, Itaú) e o que mudou desde a edição anterior.
   Se o usuário não enviar dados, pergunte; não escreva leitura sem número que a sustente.
2. Se houver edição anterior dos comentários, compare: o que se confirmou, o que mudou de direção.
   Leitores do Radar valorizam ver a mudança de tese explicitada ("Porém…", "A conta deve chegar
   em 2027").

## Estrutura de cada tópico

- **Título (mensagem-chave)**: uma frase, até ~30 palavras, que diz o "so what" para a Dexco/setor
  de construção, não só o fato. Tom de cenário: "deve", "segue", "permanece". Marque 1 a 3
  expressões-chave com `**` (no slide, viram destaque em laranja).
- **Texto de apoio**: 1 a 2 frases que explicam os vetores e citam números (ex.: "A massa de renda
  deve expandir 3,8% em 2026…"). Fonte das projeções quando citar uma (Focus, Tendências, Itaú).
- **Caixas de leitura** (opcional, uma por gráfico): uma frase curta sobre o que aquele gráfico
  mostra, também com destaques.

Os tópicos seguem a ordem do Radar e formam uma narrativa: Global → Brasil → Incerteza → Renda →
Crédito → Custos → Imobiliário → Varejo. Use conectores entre tópicos quando a tese virar ("Porém,
grande parte da renda está comprometida…").

## Formato da entrega

Para cada tópico:

```
## <Tópico>
Título: <frase com **destaques**>
Texto: <1–2 frases com números e fonte>
Caixas: (opcional)
- <gráfico>: <frase>
Números usados: <indicador = valor (mês/ano, fonte)> ...
```

A linha "Números usados" permite que o time confira cada afirmação antes de ir para o slide.

## Cuidados

- Todo número citado precisa estar nos dados fornecidos; se faltar, escreva sem o número ou
  pergunte. Não "arredonde para cima" uma tendência que os dados não mostram.
- Diferencie dado observado de projeção ("deve", "projeção Tendências de 1,2% em 2027").
- Evite adjetivos vazios ("forte", "importante") sem número; prefira o vetor e a consequência.
- Se o usuário quiser os comentários direto nos slides, use a skill `padrao-apresentacao-dexco`
  para aplicar o visual do Radar.
