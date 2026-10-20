---
name: atualizar-painel-estrategico
description: Gera, confere e publica o Painel Estratégico da Dexco (Radar macro + Análise de Peers, site com senha) a partir das planilhas Grande Base, Análise de Competidores e, opcionalmente, das projeções da Tendências. Use sempre que alguém do time de estratégia pedir para atualizar, gerar, republicar ou "colocar no ar" o painel, o radar, o dashboard de peers ou o site do time, quando uma dessas planilhas for atualizada, ou quando pedirem o PowerPoint de peers no template do Offsite.
---

# Atualizar o Painel Estratégico

O painel é gerado por scripts do projeto `painel-estrategico` (pasta `pipeline/`). Esta skill
descreve o ritual de atualização; as regras de tratamento dos dados ficam no `CLAUDE.md` e no
`README.md` do projeto, que devem ser lidos antes de mexer em qualquer regra.

Converse em português.

## 1. Localizar os arquivos

Na pasta do projeto (normalmente sincronizada com o SharePoint), identifique cada planilha pelo
conteúdo, não pelo nome:

| Planilha | Como reconhecer | Parâmetro |
|---|---|---|
| Grande Base | abas `Dados_Tendências`, `Dados_BACEN`, `Apoio_Radar` | `--radar` |
| Análise de Competidores | abas `Gráficos - Louças`, `Gráficos - Metais`, `Gráficos Madeira 1`, `Gráficos - Revestimento` | `--peers` |
| Projeções Tendências (opcional) | primeira linha "Índices" + anos; nome com a edição MMAAAA (ex.: `102026_Dexco.xlsx`) | `--proj` |

Se houver mais de uma versão de alguma planilha, pergunte qual usar. Gerar o painel não altera as
planilhas: o painel só mostra o que está nelas, e é isso que dá confiança ao time. (Lançar dados
novos numa planilha é outra tarefa, feita só a pedido e com as regras de edição do `CLAUDE.md`.)

Se a pasta do projeto não estiver conectada (não existe `pipeline/build_site.py`), peça ao usuário
para conectar a pasta ou enviar o ZIP do projeto; sem os scripts esta skill não funciona.

## 2. Gerar

Na primeira vez: `python -m pip install -r requirements.txt`.

**Peça a senha do painel ao usuário a cada geração.** Não a escreva em arquivos, commits, nomes de
arquivo nem no resumo final: quem tem o arquivo publicado e a senha vê todos os dados.

```bash
python pipeline/build_site.py --radar "<Grande Base>.xlsx" --peers "<Análise de Competidores>.xlsx" \
  [--proj "<Tendências>.xlsx"] --out site --senha "<senha>" [--deck Analise_de_Peers_Dashboard.pptx]
```

- `--deck` só quando pedirem o PowerPoint.
- Sem `--senha` o site sai aberto: serve só para conferência local, nunca para publicar.

## 3. Conferir antes de publicar

1. Leia o resumo em JSON que o comando imprime e conte ao usuário:
   - até quando vão os dados (`radar_ate`, `peers_ate`);
   - quantos valores foram retirados como erro/outlier/n.s. (`ajustes_peers`), comparando com a
     geração anterior se souber o número;
   - se houver `projecoes`: edição, ano de início da projeção e `sem_topico` (indicadores novos sem
     tópico no Radar; pergunte onde entram e acrescente a regra em `pipeline/radar/projecoes_map.json`).
2. Confirme que os HTML com senha não têm dados legíveis: procure "SELIC" e "Deca" em
   `site/index.html`, `site/radar.html` e `site/peers.html`; o resultado deve ser 0.
3. Se o comando falhar com `Falha em ...`, a estrutura de alguma planilha mudou (aba renomeada,
   bloco movido, faixa de células diferente). Explique o que mudou e proponha o ajuste no script
   (`pipeline/radar/extract.py` para faixas do Radar; `pipeline/peers/build.py` para títulos de
   blocos de Peers). Não pule a etapa nem invente números.
4. Se a planilha de peers mudou, ofereça rodar a skill `conferir-peers` para listar o que mudou e
   os pontos de atenção.

## 4. Publicar

- Publique `site/index.html` como página de um artifact e `site/radar.html` e `site/peers.html`
  como arquivos de apoio, com esses mesmos nomes.
- Republique sempre **no mesmo link**: o link oficial está na linha "Site publicado" do
  `README.md`. Na primeira publicação de uma conta nova, anote o link novo ali.
- Para o time abrir sem login, o dono do link deixa o acesso como "Qualquer pessoa com o link" no
  menu Share. Só o dono consegue mudar isso; a senha é o que protege o conteúdo.

## 5. Mudanças pedidas pelo time

- Visual, textos e gráficos: nos modelos `pipeline/radar/template.html` e
  `pipeline/peers/template.html` ou em `pipeline/build_site.py`. Nunca edite à mão os arquivos em
  `site/`, que são sobrescritos a cada geração.
- Regras de dados: `pipeline/peers/build.py` e `clean.py`, `pipeline/radar/extract.py`. Ao mudar
  uma regra, atualize a lista de regras do `CLAUDE.md` e do `README.md`.
- Depois de qualquer mudança: gere de novo, abra no navegador e confira antes de publicar.
- Os painéis não têm textos de comentário/análise. Se pedirem comentários, use a skill
  `comentarios-radar` e coloque-os na apresentação, não no painel.
