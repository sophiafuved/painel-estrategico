# Pendências do Painel Estratégico (atualizado em 20/10/2026)

Lista para retomar o trabalho. **Prioridade desta semana: o manual (seção 6).** A planilha de conferência (`Conferencia_Peers.xlsx`) foi enviada no chat em 19/10; ela não fica no repositório porque o projeto não guarda planilhas. Se precisar, peça para gerar de novo.
Site publicado: https://claude.ai/artifact/N8CHHxXKwp8fjJsHYg15bb

## 1. Planilha Análise de Competidores

- [ ] **Criar os blocos sobre a receita** nas abas "Gráficos" de cada negócio, no mesmo formato dos
      blocos sobre EBITDA (título na coluna C com os anos ao lado; país na coluna B e empresa na C
      nas linhas abaixo; mesma convenção de sinal, consumo de caixa negativo):
      `Var. WK/ROL`, `Outras Desp./ROL`, `FCL/ROL`. Ex.: Deca Louças 2023, FCL/ROL = −406 ÷ 633 = −64%.
- [ ] **Conferir os 27 pontos de prioridade alta** (aba "Pontos de atenção", todos da Dexco):
  - Deca Louças 2023: EBITDA −R$ 96 mi; capex/ROL 30%.
  - Deca Metais 2025: capital de giro/ROL 49%; FCL/EBITDA −98%.
  - Dexco Madeira 2025: capital de giro/ROL −0,9%; dias de fornecedores 139.
  - DX Revestimentos: EBITDA 2023 de R$ 8 mi; salto de 2020 (receita +85%).
  - (Fluxo de caixa vazio da Deca Louças 2025 e da DX Revestimentos 2025: explicado pelo EBITDA
    negativo; já aparece como n.s.)
- [ ] **Preencher 2025** (se já publicado): Roca, Docol (também 2024), Fabrimar, Gessi, Berneck,
      Repinho, Sudati, Biancogres, Porcelanosa, Villagres.
- [ ] **Corrigir a linha da Roca (Espanha)**: lucro líquido 2023–24 puxando a receita.
- [ ] Opcional, ponte do caixa completa:
  - Revestimento: criar `Var. WK/EBITDA`, `Capex/EBITDA`, `Outras Desp./EBITDA`.
  - Madeira: o bloco de FCL só tem a Arauco BR; preencher as demais empresas e criar a abertura.

## 2. Decisões do time

- [ ] Pontos da aba "Em aberto": Madeira × deck de junho, Docol 2022–23 (receitas redondas),
      Deca Metais 2025, EBITDA 2025 da Lamosa (manter, corrigir ou tirar).
- [ ] Projeções Tendências (Radar): confirmar que a projeção começa em 2026 (ano da edição) e que o
      atraso PF de 5,73 na planilha é 5,73%.

## 3. Para colocar no ar

- [ ] Mandar a planilha atualizada e a senha → gerar de novo, conferir, republicar no mesmo link e
      refazer a planilha de conferência.

## 4. Com o time

- [ ] Reenviar o ZIP atualizado ao colega (o anterior é antigo).
- [ ] Definir o link oficial do time (conta de trabalho), anotar no README e deixar o acesso como
      "Qualquer pessoa com o link".

## 5. Skill atualizar-deck-radar: ajustes estéticos (FEITO em 20/10)

Retorno da Sophia sobre o deck de teste (Radar Modelo 2, 20/10): "quase perfeito"; ajustes estéticos nos textos:
- [x] **Bullets**: alguns slides do modelo têm vários marcadores (ex.: Renda, slide 5, com 2 bullets). Hoje o
      script junta tudo em um parágrafo só. Ajuste: aceitar `texto` como lista na leitura (um item por bullet) e
      recriar um parágrafo por item copiando o formato (`pPr`, marcador, recuo) de cada parágrafo original.
- [x] **Alinhamento/centralização**: preservar o alinhamento e a ancoragem de cada parágrafo original (hoje só o
      primeiro parágrafo é mantido), e conferir as caixas cinzas de leitura (centralizadas no modelo).
- [x] Títulos longos: encurtar na leitura (15–20 palavras) em vez de reduzir a fonte.
- [x] Depois do ajuste: gerar de novo, comparar slide a slide com o original e reempacotar a skill.

## 6. A FAZER ESTA SEMANA: manual de uso do Claude no time

Objetivo: qualquer pessoa do time (e o próximo estagiário) consegue operar o sistema sem depender da Sophia.
Formato sugerido: um documento curto e prático (`MANUAL.md` no projeto e uma versão para compartilhar com o time).

Rascunho do conteúdo:
- [ ] **Visão geral**: o que é o sistema (painel com senha, deck do Radar, conferência de peers, skills) e o
      fluxo de cada mês/edição, em uma página.
- [ ] **Onde estão as coisas**: planilhas no SharePoint (Grande Base, Análise de Competidores, Tendências),
      projeto no GitHub/pasta do time, link do painel, quem é dono do link, onde fica a senha (fora de arquivos).
- [ ] **Skills**: para que serve cada uma das 7, quando usar e o que entregam.
- [ ] **Prompts prontos** (copiar e colar), por tarefa: atualizar o painel; conferir peers (com versão
      anterior); incluir peer novo; escrever comentários do Radar; revisar e aprovar a leitura; gerar o deck do
      Radar; montar slide no padrão Offsite.
- [ ] **Calendário/rotina**: quando cada planilha atualiza (indicadores ao sair dado novo; peers na
      demonstração anual; Tendências a cada edição) e o que fazer em cada caso.
- [ ] **Regras e decisões do time**: comentários só com revisão; painel limpo por padrão; dados sem invenção;
      o que fazer quando o script falha ("Falha em ...").
- [ ] **Instalação**: Claude da empresa, pasta do projeto, Python, pacotes `.skill`.
- [ ] **Como pedir melhorias**: onde anotar (PENDENCIAS.md), como testar uma mudança antes de publicar.

## 7. Aperfeiçoar as skills (contínuo)

- [x] `atualizar-deck-radar`: ajustes estéticos da seção 5 (20/10).
- [ ] Republicar o painel com a leitura em bullets (precisa da senha).
- [x] `atualizar-deck-radar` também para o **Modelo 1** (20/10): mapa próprio, textos em `deck_modelo1`,
      datas associadas pelo nome entre parênteses.
- [x] Nova skill `atualizar-deck-peers` (20/10): testada nos 4 decks do estágio (850 de 898 séries
      vinculadas atualizadas). Ver no chat/conferências os pontos abaixo.
- [ ] Decks de peers: revisar as conferências — textos copiados de outro negócio no deck de Louças
      (Arauco/Eucatex, Placas do Brasil), "[WIP]"/"x"/"Lorem ipsum"; séries não reconhecidas (Louças slide 16
      participação por segmento; Madeira slide 5 bolhas; Revestimento "Portobello - Ajustado", que não existe
      mais na planilha); Madeira com 57 séries revisadas (mesma decisão "Madeira × deck de junho").
- [x] Decks de peers no **Template PPT 2026** (20/10): script `aplicar_template.py` na skill
      `padrao-apresentacao-dexco` (template salvo em `assets/`). Revisar à mão os slides apontados na
      conferência de cada deck (formas na faixa do título).
- [ ] Decks de peers: tabela-resumo do slide 2 e textos não são atualizados; avaliar uma leitura de peers
      (como a `comentarios-radar`) e a atualização da tabela.
- [ ] `comentarios-radar`: títulos mais curtos para o deck; coletar o retorno do gestor sobre o tom.
- [ ] `incluir-peer-novo` e `padrao-apresentacao-dexco`: testar num caso real (ex.: release 2025 de um peer
      faltante) e ajustar.
- [ ] Testes com e sem skill para mostrar o ganho ao time (opcional).
- [ ] Reempacotar as skills alteradas e manter `skills_para_empresa/` sempre atualizado.

## 8. Transição (a Sophia sai do time em ~2 meses, por volta de dez/2026)

Princípio: deixar o sistema adequado, documentado e atualizado para o próximo estagiário e para quem fica.
- [ ] Mover o projeto, o link do painel e as skills para contas/pastas do time (não pessoais).
- [ ] Definir quem fica responsável pelo painel, pela senha e pela revisão dos comentários.
- [ ] Sessão de passagem com o novo estagiário usando o manual (seção 6) e um ciclo real de atualização.
- [ ] Revisar `CLAUDE.md`, `README.md` e skills para que não dependam de conhecimento que só está na cabeça da Sophia.
