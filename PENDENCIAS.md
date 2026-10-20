# Pendências do Painel Estratégico (atualizado em 20/10/2026)

Lista para retomar o trabalho. **Prioridade desta semana: o manual (seção 6).** A planilha de conferência (`Conferencia_Peers.xlsx`) foi enviada no chat em 19/10; ela não fica no repositório porque o projeto não guarda planilhas. Se precisar, peça para gerar de novo.
Site publicado: https://claude.ai/artifact/N8CHHxXKwp8fjJsHYg15bb

## 1. Planilha Análise de Competidores

> Combinado em 20/10: estes itens serão feitos **direto no Claude do time** (pasta `Claude` no SharePoint),
> seguindo a regra de edição do `CLAUDE.md` (cópia da versão anterior + lista de células alteradas).
> Prompt pronto para os blocos sobre a receita:
> *"Na Análise de Competidores da pasta Claude, crie nas abas Gráficos de cada negócio os blocos
> Var. WK/ROL, Outras Desp./ROL e FCL/ROL, no mesmo formato dos blocos sobre EBITDA e com a mesma convenção
> de sinal, calculando índice sobre EBITDA × EBITDA ÷ receita. Guarde antes uma cópia da versão anterior
> e me liste as células criadas. Depois gere o painel sem senha para eu conferir o botão Sobre receita."*

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

- [ ] (A cada planilha nova) Mandar a planilha atualizada e a senha → gerar de novo, conferir, republicar no mesmo link e
      refazer a planilha de conferência.

## 4. Com o time

- [x] ZIP atualizado enviado à Sophia para o teste com o time (20/10), junto com o manual.
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

## 6. Manual de uso do Claude no time (FEITO em 20/10, versão 1)

- [x] `MANUAL.md` no projeto (fonte única) e Word gerado com `python docs/gerar_manual.py`
      (`Manual_Claude_PE.docx`, para a pasta `Planejamento Estratégico > Claude` no SharePoint).
- [ ] Preencher os **[a confirmar]** do manual: pastas definitivas das planilhas e decks, administrador do
      Claude da Dexco, frequência de atualização dos peers, GitHub do time de PE.
- [ ] Revisão do manual pelo Nivaldo (é quem vai operar) e subir o Word no SharePoint.

## 7. Aperfeiçoar as skills (contínuo)

- [x] `atualizar-deck-radar`: ajustes estéticos da seção 5 (20/10).
- [x] Painel republicado no mesmo link com a leitura em bullets e os títulos novos (20/10; leitura ainda em rascunho).
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
- [ ] **Voltar depois**: melhorias nos decks de peers no Template 2026 (a Sophia ainda vai pensar no que mudar;
      fontes Roboto Light/Black já corrigidas em 20/10).
- [ ] Decks de peers: tabela-resumo do slide 2 e textos não são atualizados; avaliar uma leitura de peers
      (como a `comentarios-radar`) e a atualização da tabela.
- [x] `comentarios-radar`: títulos mais curtos para o deck (20/10).
- [ ] `comentarios-radar`: coletar o retorno do Henrique Fantin sobre o tom da leitura.
- [ ] `incluir-peer-novo` e `padrao-apresentacao-dexco`: testar num caso real (ex.: release 2025 de um peer
      faltante) e ajustar.
- [ ] Testes com e sem skill para mostrar o ganho ao time (opcional).
- [ ] Reempacotar as skills alteradas e manter `skills_para_empresa/` sempre atualizado.

## 8. Transição (a Sophia sai do time em ~2 meses, por volta de dez/2026)

Princípio: deixar o sistema adequado, documentado e atualizado para o próximo estagiário e para quem fica.
- [ ] Mover o projeto, o link do painel e as skills para contas/pastas do time (não pessoais).
- [x] Responsáveis definidos (20/10): painel e senha com o Nivaldo; leitura aprovada pelo Henrique Fantin;
      Nivaldo e Ramon com Claude (plano Team). Senha guardada na planilha "Acessos".
- [ ] Sessão de passagem com o novo estagiário usando o manual (seção 6) e um ciclo real de atualização.
- [ ] Revisar `CLAUDE.md`, `README.md` e skills para que não dependam de conhecimento que só está na cabeça da Sophia.

## 9. Próximas frentes de IA (proposta de 20/10, para discutir com o Henrique)

Premissas: sem assinatura de Capital IQ e Valor Pro (só fontes públicas); atualização automática da
Grande Base já feita no Claude do time; o time já tem um modelo de demanda por regressão múltipla.

- [ ] **A. Inteligência competitiva com fontes públicas**: DFP/ITR da CVM (Portobello, Eucatex, Dexco),
      SEC EDGAR (Masco), relatórios anuais em PDF (Roca, Geberit, Villeroy & Boch, Lixil), releases e
      teleconferências nos sites de RI. Extrair indicadores para a planilha (skill `incluir-peer-novo`),
      comparar o tom e o guidance trimestre a trimestre e alertar sobre capacidade, preço e capex.
      Candidato a piloto.
- [ ] **B. Evoluir o modelo de demanda existente** (não refazer): teste fora da amostra, monitoramento do
      erro, modelo desafiante para comparar ou combinar, e rodada automática com a Tendências da semana
      → projeção própria de demanda por negócio.
- [ ] **C. Cenários e estresse** a partir do modelo de demanda (depende de B): SELIC, MCMV, renda → receita.
- [ ] **D. Triagem de alvos de M&A** com fontes públicas (Receita/CNPJ, notícias, sites, associações).
      Teste feito em 20/10 para Madeira (planilha + 2 slides no Template 2026, entregues no chat e fora do
      repositório por serem confidenciais), para a Sophia sugerir na reunião do time. Próximo: validar a
      régua com o Henrique e rodar no Claude do time com os dados de CNPJ (sucessão).
- [ ] **E. Simulação de reação dos concorrentes** para o Offsite (provocação, não previsão).
- Fora por enquanto: avaliador cético de apresentações e perguntas à base/memória do time.

