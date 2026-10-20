---
title: "Manual do Claude no time de Planejamento Estratégico"
subtitle: "Painel Estratégico, Radar, Análise de Peers e skills"
date: "Versão de outubro/2026"
---

# Sobre este manual

Este manual explica como o time de Planejamento Estratégico usa o Claude para atualizar o Painel
Estratégico, o Radar Estratégia e a Análise de Peers: onde estão os arquivos, quem faz o quê, quais
pedidos fazer ao Claude e o que conferir antes de entregar.

- **Público:** qualquer pessoa do time, inclusive quem está chegando.
- **Onde fica:** a versão em Word fica na pasta do time no SharePoint (`Planejamento Estratégico >
  Claude`). A fonte do texto é o arquivo `MANUAL.md` no projeto `painel-estrategico`.
- **Como atualizar:** edite o `MANUAL.md` (ou peça ao Claude: *"atualize o manual com …"*) e gere o
  Word de novo com `python docs/gerar_manual.py`. Não edite só o Word, senão as duas versões se
  separam.
- Itens marcados com **[a confirmar]** ainda dependem de uma decisão ou informação do time.

# 1. Visão geral

O sistema tem três entregas, todas alimentadas pelas planilhas do time:

| Entrega | O que é | Fonte dos dados |
|---|---|---|
| **Painel Estratégico** | Site com senha com o Radar macro e a Análise de Peers, gráficos interativos e projeções da Tendências | Grande Base, Análise de Competidores, projeções da Tendências |
| **Radar Estratégia** (PowerPoint) | Deck de cenário macro, Modelo 1 ou Modelo 2, com gráficos vinculados à Grande Base e a leitura do time | Grande Base + leitura do time |
| **Análise de Peers** (PowerPoint) | Um deck por negócio (Louças, Metais, Madeira, Revestimento) no Template PPT 2026 | Análise de Competidores |

O Claude faz o trabalho repetitivo: lê as planilhas, atualiza gráficos e datas, escreve rascunhos de
texto e confere números. **As pessoas do time decidem e aprovam:** escolhem o modelo, revisam a
leitura e conferem os pontos que o Claude aponta.

**Fluxo de uma edição do Radar**

1. Atualizar a Grande Base (e baixar o Excel de projeções mais recente da Tendências).
2. Pedir ao Claude a leitura do Radar (rascunho).
3. Henrique Fantin revisa e aprova a leitura.
4. Gerar o painel com senha e republicar no mesmo link.
5. Gerar o deck do Radar (Modelo 1 ou 2, conforme o pedido), conferir e enviar.

**Fluxo de uma atualização de peers**

1. Atualizar a Análise de Competidores (demonstrações novas dos peers).
2. Pedir a conferência dos números e resolver os pontos de atenção.
3. Gerar o painel e republicar.
4. Atualizar os decks de peers e aplicar o Template PPT 2026.

# 2. Quem faz o quê

| Papel | Pessoa |
|---|---|
| Responsável pelo painel (gera, publica, guarda o link) | Nivaldo |
| Aprova a leitura do Radar antes de ela aparecer como "revisada" | Henrique Fantin (gerente) |
| Atualiza a Grande Base e a Análise de Competidores | Time de Planejamento Estratégico |
| Acesso ao Claude (plano Team, Claude Code e Claude Desktop) | Nivaldo e Ramon |
| Administrador do Claude da Dexco (publica skills para a organização) | **[a confirmar]** |
| Criou o sistema; disponível para dúvidas até dez/2026 | Sophia |

O computador de referência é o do Nivaldo: é nele que o projeto está instalado e que as entregas
são geradas.

# 3. Onde estão as coisas

**SharePoint do time**

```
Duratex SA
└── Controladoria Dexco - General
    └── Planejamento Estratégico
        ├── _2026, Análises, Artigos, FP&A, Fornecedores, Mercado,
        │   OffSite, Projetos, Varejo, zEquipe, zHistórico
        └── Claude                      ← pasta do time para o Claude
            ├── 1. Semana em foco
            ├── 2. Radar Estratégia
            ├── 3. Análise de Peers
            ├── 4. Divulgação Eucatex
            ├── 5. Divulgação Portobello
            └── noticias_mundo
```

A pasta `Claude` está em organização (o time começou a usar o Claude da Dexco em out/2026); se a
estrutura mudar, atualize esta seção.

| O quê | Onde |
|---|---|
| Grande Base | **[a confirmar]** (sugestão: `Claude > 2. Radar Estratégia`) |
| Projeções da Tendências (Excel novo toda quinta-feira) | **[a confirmar]** (sugestão: `Claude > 2. Radar Estratégia > Projeções`) |
| Decks do Radar (Modelo 1 e Modelo 2) | **[a confirmar]** (sugestão: `Claude > 2. Radar Estratégia`) |
| Análise de Competidores | **[a confirmar]** (sugestão: `Claude > 3. Análise de Peers`) |
| Decks de peers por negócio | **[a confirmar]** (sugestão: `Claude > 3. Análise de Peers`) |
| Projeto `painel-estrategico` (scripts, skills, manual) | GitHub do time de PE **[a criar]**; cópia de trabalho no computador do Nivaldo |
| Link do painel | Na linha "Site publicado" do `README.md` do projeto |
| Senha do painel | Planilha **"Acessos"** do time. Nunca escreva a senha em e-mails, arquivos do projeto ou no chat do Claude fora do momento de gerar o painel |
| Pendências e próximos passos | `PENDENCIAS.md` no projeto |
| Regras de dados e instruções para o Claude | `CLAUDE.md` e `README.md` no projeto |
| Pacotes das skills para o Claude da empresa | Pasta `skills_para_empresa/` no projeto |

# 4. Rotina

| Quando | O quê | Skill |
|---|---|---|
| Toda quinta-feira | Chega o Excel de projeções da Tendências: guardar na pasta e usar a versão mais recente na próxima geração do painel | `atualizar-painel-estrategico` |
| A cada dois meses ou sob demanda | Edição do Radar: leitura, aprovação, painel e deck | `comentarios-radar`, `atualizar-painel-estrategico`, `atualizar-deck-radar` |
| Sob demanda | Escolha entre Radar Modelo 1 e Modelo 2, conforme quem pediu | `atualizar-deck-radar` |
| Quando sai demonstração nova de um peer **[a confirmar a frequência]** | Lançar os números, conferir, atualizar painel e decks | `incluir-peer-novo`, `conferir-peers`, `atualizar-deck-peers` |
| Antes de qualquer apresentação de peers | Conferência dos números | `conferir-peers` |

# 5. Como pedir ao Claude: tarefas e prompts prontos

Abra o Claude **na pasta do projeto `painel-estrategico`** (Claude Code ou Claude Desktop). As skills
são carregadas automaticamente e o Claude lê o `CLAUDE.md` com as regras. Coloque as planilhas
necessárias na pasta do projeto (ou indique o caminho) antes de pedir.

Os prompts abaixo podem ser copiados e colados. Ajuste o que estiver entre `< >`.

## 5.1 Atualizar o painel

> Atualize o painel estratégico com a Grande Base e a Análise de Competidores que estão na pasta,
> incluindo as projeções mais recentes da Tendências. Me mostre até quando vão os dados e quantos
> valores foram retirados como erro ou outlier.

- O Claude vai **pedir a senha** (está na planilha "Acessos"). Ele não grava a senha em lugar nenhum.
- Antes de publicar, ele confere que o site gerado não tem dados legíveis sem a senha.
- Republicação: sempre **no mesmo link** (o do `README.md`), para o time não precisar de link novo.
- Se aparecer "projeções sem tópico" (indicador novo da Tendências), o Claude pergunta em qual tópico
  do Radar ele entra.

## 5.2 Escrever a leitura do Radar

> Escreva a leitura do Radar desta edição (títulos, textos e caixas) com os dados atuais do painel e
> as projeções da Tendências, no estilo do time. Deixe como rascunho para o Henrique revisar.

- O texto fica em `pipeline/radar/leitura_time.json`, com status **rascunho**. No painel, ele aparece
  no seletor "Leitura: Do time" com o aviso de rascunho; no deck, os slides levam um carimbo
  "RASCUNHO".
- Depois da revisão do Henrique Fantin:

> O Henrique aprovou a leitura do Radar com estes ajustes: `<ajustes>`. Aplique os ajustes, marque
> como revisada por Henrique Fantin em `<data>` e gere o painel de novo.

## 5.3 Gerar o deck do Radar

> Atualize o deck do Radar Estratégia `<Modelo 1 | Modelo 2>` a partir da última edição que está na
> pasta, com a Grande Base atual e a leitura do time. Na capa, coloque `<mês, ano>`. Me entregue o
> PowerPoint e a lista do que preciso revisar à mão.

- Os gráficos continuam vinculados à Grande Base no SharePoint. Ao abrir, o PowerPoint pergunta sobre
  os vínculos: **Atualizar** puxa da Grande Base (mesmo resultado) e **Cancelar** mantém o que o
  Claude gravou.
- Revise à mão o que a conferência apontar: rótulos digitados sobre os gráficos (ex.: "5,9%"), Brent
  e impacto fiscal (não estão na Grande Base) e o slide "Principais mensagens".
- Com a leitura em rascunho, o deck **não deve circular**: aprove a leitura e gere de novo.

## 5.4 Conferir os números de peers

> Confira os números da Análise de Competidores que está na pasta e compare com a versão anterior
> `<nome do arquivo anterior>`. Me entregue a planilha de conferência com os pontos de atenção.

- A planilha de conferência lista o que mudou, saltos, números redondos e dados faltando, com
  prioridade. Preencha as colunas de decisão e devolva ao Claude se quiser que ele acompanhe.

## 5.5 Incluir um peer novo ou um ano novo

> Inclua os números de 2025 da `<empresa>` a partir deste release/demonstração `<arquivo>`, no formato
> da Análise de Competidores, com fonte e página de cada número.

- O Claude prepara os números com fonte e página. Se você pedir, ele **lança direto na Análise de
  Competidores** da pasta `Claude`, guardando antes uma cópia da versão anterior e listando cada
  célula alterada para conferência (veja a regra na seção 7).

## 5.6 Atualizar os decks de peers

> Atualize o deck de Análise de Peers de `<Louças | Metais | Madeira | Revestimento>` com a Análise
> de Competidores atual e depois aplique o Template PPT 2026. Me entregue o PowerPoint e as
> conferências.

- O Claude reconhece cada série pelo nome da empresa e pelos valores (a planilha muda de lugar com o
  tempo) e aponta séries não reconhecidas, valores revisados, textos provisórios ("[WIP]", "x") e
  textos que citam empresas de outro negócio.
- Os textos dos slides e a tabela-resumo **não** são atualizados automaticamente: revise-os.

## 5.7 Slide avulso no padrão Dexco

> Monte um slide no Template PPT 2026 com `<conteúdo>`, usando os dados de `<fonte>`.

## 5.8 Lançar dados novos numa planilha

> Lance os dados do arquivo `<Excel da Tendências / release>` que está na pasta `Claude` na
> `<Grande Base | Análise de Competidores>`, na aba `<aba>`. Antes, guarde uma cópia da versão
> anterior e, no fim, me liste cada célula alterada (antes → depois).

- Confira a lista de mudanças antes de gerar o painel ou os decks.

# 6. As skills

Skills são "manuais" que o Claude segue para cada tarefa. Ficam em `.claude/skills/` no projeto e são
carregadas sozinhas quando o Claude abre a pasta.

| Skill | Use para | Entrega |
|---|---|---|
| `atualizar-painel-estrategico` | Gerar, conferir e publicar o painel com senha | Site atualizado no mesmo link e resumo das datas dos dados |
| `comentarios-radar` | Escrever títulos, textos e caixas do Radar no estilo do time | `leitura_time.json` em rascunho |
| `atualizar-deck-radar` | Atualizar o Radar Modelo 1 ou 2 com dados e leitura, sem quebrar o vínculo com a Grande Base | PowerPoint + conferência |
| `conferir-peers` | Conferir os números da Análise de Competidores | Planilha de conferência |
| `incluir-peer-novo` | Levar números de releases/DFPs para o formato da planilha | Tabela com números, fontes e páginas |
| `atualizar-deck-peers` | Atualizar os decks de peers com a planilha atual | PowerPoint + conferência |
| `padrao-apresentacao-dexco` | Padrão visual (Radar e Template PPT 2026) e levar decks para o template novo | PowerPoint no padrão |

# 7. Regras do time

- **Dados só da fonte oficial:** Grande Base e Análise de Competidores. O Claude não inventa número;
  quando falta dado, ele avisa.
- **Edição de planilhas com controle:** o Claude pode editar as planilhas da pasta `Claude` quando
  alguém pede e com dados que o time disponibilizou (ex.: Excel da Tendências, demonstração de um peer).
  Ele guarda antes uma cópia da versão anterior com a data, mexe só no que foi pedido (sem mudar abas,
  blocos ou fórmulas sem combinar) e lista o que mudou para o time conferir. Gerar painel, decks e
  conferências só lê as planilhas.
- **Senha:** só na planilha "Acessos". Informe ao Claude apenas na hora de gerar o painel.
- **Painel limpo por padrão:** comentários ficam ocultos; a leitura aparece só no seletor e só como
  "revisada" depois da aprovação do Henrique Fantin.
- **Leitura em rascunho não circula.**
- **Regras de tratamento dos peers** (outliers, n.s. com EBITDA negativo, alavancagem consolidada da
  Dexco, etc.) estão no `CLAUDE.md`. Não mude sem decisão do time.
- **Gráficos de dois eixos do Radar são mantidos** (pedido do gestor).
- **Template:** decks de peers no Template PPT 2026, texto em Roboto Light e negritos em Roboto Black.

# 8. Quando algo dá errado

| Situação | O que fazer |
|---|---|
| O Claude diz "Falha em …" ao gerar o painel | Alguma planilha mudou de estrutura (aba renomeada, bloco movido). Peça: *"explique o que mudou e proponha o ajuste no script"*. Não force a geração |
| Indicador novo da Tendências "sem tópico" | Diga em qual tópico do Radar ele entra; o Claude registra a regra |
| Deck de peers com séries "não reconhecidas" | A linha pode ter sido apagada ou renomeada na planilha. Confira na conferência e apague a série no deck se não fizer mais sentido |
| PowerPoint pede para atualizar vínculos | Pode atualizar (puxa do SharePoint) ou cancelar (mantém o que o Claude gravou) |
| Fontes diferentes ao abrir o deck | Instale Roboto Light e Roboto Black no computador |
| O painel não abre para o time | O dono do link precisa deixar o acesso como "Qualquer pessoa com o link" (a senha protege o conteúdo) |
| O Claude não "acha" as skills | Confira se o Claude foi aberto na pasta do projeto |

# 9. Instalação (uma vez por computador)

1. **Claude:** Claude Code e/ou Claude Desktop com a conta do plano Team da Dexco.
2. **Python 3** instalado.
3. **Projeto:** copiar o projeto do GitHub do time de PE **[a criar]** para o computador (ou pedir ao
   Claude: *"clone o projeto painel-estrategico do GitHub do time"*).
4. **Pacotes:** na pasta do projeto, rodar `python -m pip install -r requirements.txt` (ou pedir ao
   Claude para instalar).
5. **Fontes:** Roboto Light e Roboto Black.
6. **Teste:** abrir o Claude na pasta do projeto e pedir *"gere o painel sem senha só para
   conferência local"* (essa versão nunca é publicada).

**Skills no Claude da empresa (fora do projeto):** quem usa o Claude no navegador ou no app sem a pasta
do projeto pode instalar as skills que não dependem dos scripts (`comentarios-radar`,
`incluir-peer-novo`, `padrao-apresentacao-dexco`, `atualizar-deck-peers`) a partir dos arquivos `.skill`
da pasta `skills_para_empresa/`. Para o time todo, o administrador do Claude da Dexco
**[a confirmar]** pode publicá-las para a organização. O passo a passo está no `README.md` dessa pasta.

# 10. Como pedir melhorias

1. Anote a ideia ou o problema em `PENDENCIAS.md` (ou peça: *"registre nas pendências: …"*).
2. Peça a mudança ao Claude, que altera os modelos, scripts ou skills (nunca os arquivos gerados).
3. Gere sem senha e confira no navegador ou no PowerPoint antes de publicar.
4. Se mudou uma skill, peça *"empacote a skill `<nome>` de novo"* para atualizar `skills_para_empresa/`.
5. Peça para salvar a mudança no GitHub do time, com uma descrição curta.
6. Se a mudança altera o jeito de trabalhar, atualize este manual.

# 11. Transição (até dez/2026)

- [ ] Criar o GitHub do time de PE e mover o projeto para lá.
- [ ] Publicar o painel pela conta do Nivaldo e anotar o link novo no `README.md`.
- [ ] Definir as pastas definitivas das planilhas e dos decks (seção 3).
- [ ] Descobrir o administrador do Claude da Dexco e publicar as skills para o time.
- [ ] Confirmar a frequência de atualização dos peers.
- [ ] Sessão de passagem com o próximo estagiário usando este manual e uma edição real do Radar.
