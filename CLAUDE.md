# Instruções para o Claude: Painel Estratégico Dexco

Este projeto gera o **Painel Estratégico** do time de estratégia: um site com senha que junta o
**Radar macro** (planilha *Grande Base*) e a **Análise de Peers** (planilha *Análise de
Competidores*), e opcionalmente o PowerPoint de peers no template do Offsite. Converse com o time em
**português**.

## Onde estão as planilhas

As duas planilhas ficam nesta mesma pasta (normalmente sincronizada com o SharePoint). Os nomes
podem variar; identifique cada uma pelo conteúdo:

- **Grande Base**: tem as abas `Dados_Tendências`, `Dados_BACEN`, `Apoio_Radar`.
- **Análise de Competidores**: tem as abas `Gráficos - Louças`, `Gráficos - Metais`,
  `Gráficos Madeira 1`, `Gráficos - Revestimento`.

- **Projeções da consultoria (opcional, só Radar)**: planilha da Tendências (ex.: `102026_Dexco.xlsx`,
  edição MMAAAA no nome), com a primeira linha "Índices" + anos e um indicador por linha. Se houver
  mais de uma edição, use a mais recente (pergunte em caso de dúvida).

Se houver mais de uma versão de alguma planilha, pergunte qual usar. Nunca edite as planilhas.

## Como gerar o painel

Instale as dependências na primeira vez: `python -m pip install -r requirements.txt`.

```bash
python pipeline/build_site.py --radar "<Grande Base>.xlsx" --peers "<Análise de Competidores>.xlsx" --out site --senha "<senha>" [--deck Analise_de_Peers_Dashboard.pptx]
```

- O comando imprime um resumo em JSON (`radar_ate`, `peers_ate`, `ajustes_peers`). Mostre ao
  usuário até quando vão os dados e quantos valores foram retirados como erro/outlier.
- Se falhar com `Falha em ...`, a estrutura de alguma planilha provavelmente mudou (aba renomeada,
  bloco movido, faixa de células diferente). Explique ao usuário o que mudou e proponha o ajuste
  no script; não invente dados nem pule a etapa.
- `--deck` só quando pedirem o PowerPoint.
- Projeções: acrescente `--proj "<arquivo da Tendências>.xlsx"` quando o arquivo estiver na pasta.
  A projeção começa no ano da edição (nome do arquivo); para mudar, use `--proj-desde AAAA`. O resumo
  em JSON traz `projecoes.sem_topico`: indicadores novos que não estão em
  `pipeline/radar/projecoes_map.json`. Pergunte ao usuário em qual tópico do Radar entram e
  acrescente a regra no mapa.

## Senha

- **Peça a senha ao usuário a cada geração.** Nunca a escreva em arquivos, commits, nomes de
  arquivo, mensagens de log ou no CLAUDE.md, e não a repita no resumo final.
- Sem `--senha`, o site sai **sem proteção**: use apenas para conferência local, nunca publique
  essa versão.

## Como publicar

- Publique `site/index.html` como página de um artifact, com `site/radar.html` e `site/peers.html`
  como arquivos de apoio (`files`), com esses mesmos nomes.
- Para atualizar, republique **no mesmo link**. O link oficial do time fica anotado na linha
  "Site publicado" do `README.md`; na primeira publicação, anote o link novo ali.
- Para o time abrir sem login, o dono do link precisa deixar o acesso como "Qualquer pessoa com o
  link" no menu Share (só a pessoa consegue fazer isso; a senha protege o conteúdo).
- Antes de publicar, confira que os arquivos gerados com senha não contêm dados legíveis
  (ex.: procure "SELIC" e "Deca" nos três HTML; devem aparecer 0 vezes).

## Regras de tratamento dos dados (já implementadas; não mude sem o usuário pedir)

- Peers: dados das abas "Gráficos" de cada negócio; células zeradas = sem dado.
- Erros de origem retirados: lucro igual à receita (e margens de 100% decorrentes) e prazos
  médios negativos.
- Outliers retirados: |DL/EBITDA| > 10x; índices sobre EBITDA > 500% em módulo; prazos > 365 dias;
  giro > 6x; margens e ROIC < −50%; capex > 50% da receita. Na tabela-resumo aparecem como **n.s.**
- Tabelas-resumo: uma no início de cada tópico (não há resumo geral no topo); cada coluna usa o
  ano da receita mais recente da empresa; indicadores de outros anos aparecem como "–".
- **Alavancagem da Dexco é sempre a consolidada**, em todos os negócios.
- Madeira: EBITDA = receita × margem EBITDA (a aba não traz o valor absoluto).
- Metais: nos blocos "Mkt Competitivo/Médio/Luxo", participação = empresa ÷ total do segmento (as
  primeiras linhas desses blocos estão em R$ milhões).
- Escalas dos gráficos vão do menor ao maior valor; colunas começam no zero; botão "Foco Dexco"
  ajusta o eixo à Dexco.
- Sem textos de comentário/análise nos painéis: os títulos dos tópicos são neutros. Comentários
  só entram quando o painel vira apresentação.
- Radar: cada tópico abre com um resumo dos seus indicadores (último dado observado, sem
  projeções, e variação em 12 meses / 4 trimestres / 4 semanas no Focus). Os gráficos que na
  apresentação têm dois eixos mantêm os dois eixos (a legenda marca "esq."/"dir.").
- Peers: faixa dos peers (1º–3º quartil e mediana dos peers do filtro, sem a Dexco, com 3+ valores
  no ano) nos gráficos de índices; não aparece nos de valores absolutos (R$). Matriz de
  posicionamento no início (eixos escolhidos pelo usuário, tamanho da bolha escolhido pelo usuário, receita por padrão; linhas na mediana dos
  peers; "Crescimento da receita" é calculado a partir da receita).
- Radar: projeções da consultoria aparecem em um cartão por tópico (realizado em azul, projeção em
  laranja; colunas para variação % a/a, linha tracejada para taxas em nível), com botão
  Mostrar/Ocultar no topo; também entram no "Monte seu gráfico". Indicadores com "%" no nome são
  taxas em nível; os demais, variação % a/a. Células sem formato de % estão em % (5,7 = 5,7%).
- Radar: seção "Monte seu gráfico" (até 3 séries, em nível, variação em 12 meses ou base 100).
- Legendas clicáveis: clicar esconde/mostra a série.
- Radar: Brent e impacto fiscal vêm de `pipeline/radar/static_apresentacao.json` (dados da
  apresentação, não da Grande Base).

## Decisões em aberto

- Madeira usa os números da planilha atual, que divergem do deck de junho/2026 (ex.: capital
  investido da Berneck, EBITDA da Floraplac). Próxima fase: conferência dos valores.
- Docol 2022–23 (receita de R$ 1.100 e R$ 1.200 milhões, números redondos) mantida por enquanto.

## Ao fazer alterações

- Mudanças de visual, textos ou gráficos vão nos modelos (`pipeline/radar/template.html`,
  `pipeline/peers/template.html`) ou no `build_site.py`, nunca editando à mão os arquivos em
  `site/`, que são sobrescritos a cada geração.
- Mudanças de regra de dados vão em `pipeline/peers/build.py` / `clean.py` (Peers) ou
  `pipeline/radar/extract.py` (faixas de células do Radar). Atualize esta lista de regras e o
  `README.md` quando mudar alguma.
- Depois de mudar, gere de novo e confira no navegador antes de publicar.
