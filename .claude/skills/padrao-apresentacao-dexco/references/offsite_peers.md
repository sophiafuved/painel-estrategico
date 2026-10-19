# Padrão visual do template do Offsite / Análise de Peers

Arquivo: `assets/template_offsite_peers.pptx` (16:9, 13,33 × 7,5 pol.). Use sempre os layouts do
próprio template; não crie slides em branco.

## Layouts do template

| Layout | Uso |
|---|---|
| `Capa de início` / `2.Início_2` | capa e abertura de seção |
| `1_Slide de Título` | menu/sumário (botões por negócio) |
| `4. Slide Análise_base` | slide de análise com barra lateral |
| `Corpo_para gráficos` | slides de gráficos (o mais usado) |
| `1_Conteúdo com Legenda` | texto + imagem |

Cuidado técnico (python-pptx): crie slides a partir do layout de um slide existente do template
(`slide.slide_layout`), não de `prs.slide_layouts[i]` por posição; o template tem mais de um
master e a posição muda, o que já trouxe o fundo preto da capa para slides de gráfico.

## Identidade

- **Fonte**: Roboto (títulos e corpo). Texto de eixo 8 pt `#595959`; legenda 7,5 pt.
- **Barra lateral preta** à esquerda: o conteúdo começa em x ≈ 1,05 pol. para não ficar escondido.
- **Navegação** (quando o deck é interativo): linha de botões no topo (y ≈ 0,64 pol.), botão ativo
  preto com texto branco, demais cinza claro com texto `#404040`; hyperlinks entre slides.
- **Cores**:
  - Dexco sempre em **preto** `#000000`, linha mais grossa (2,75 pt) e com marcadores;
  - nos slides, cada peer tem **cor fixa** (a mesma dos decks originais de cada negócio), para o
    leitor reconhecer a empresa de um slide para outro:
    - Louças: Icasa `#70AD47`, Roca do Brasil / Grupo Roca `#0070C0`, Astra `#FF0000`,
      Toto `#FFC000`, Lixil `#ED7D31`, Roca `#A5A5A5`;
    - Metais: Docol `#4BACC6`, Fabrimar `#FF3300`, Masco `#EA9618`, Gessi `#FFC000`,
      Geberit `#4472C4`, Villeroy & Boch `#70AD47`;
    - Madeira: Berneck `#4BACC6`, Eucatex `#FF3300`, Arauco BR `#9BBB59`, Guararapes `#8064A2`,
      Placas do Brasil `#D7428B`, Repinho `#747474`, Sudati `#FFC000`, Floraplac `#764B0B`,
      Fibraplac `#AF2468`; Dexco Painéis Verticalizados `#404040`; MDF `#2F5597` e MDP `#ED7D31`
      na taxa de ocupação;
    - Revestimento: Portobello `#0070C0`, Villagres `#70AD47`, Biancogres `#4BACC6`,
      Embramaco `#ED7D31`, Mohawk `#FF0000`, Porcelanosa `#8064A2`, Grupo Lamosa `#FFC000`;
    - empresa nova: escolha uma cor que ainda não esteja em uso no negócio e registre aqui;
  - cor de cada negócio (capas e menus): Louças `#203864`, Metais `#2F5597`, Madeira `#548235`,
    Revestimento `#ED7D31`;
  - mercado, média e WACC tracejados em cinza `#7F7F7F`;
  - (o painel web usa outra paleta, escolhida na hora pelo usuário; nos slides vale a cor fixa);
  - grade `#E7E6E6` (0,5 pt), eixo `#BFBFBF`.

## Regras de gráfico do time

- Valores absolutos (R$ milhões) em **colunas** começando no zero; índices (%, x, dias) em **linhas**.
- Escala do menor ao maior valor plotado (arredondada), para não achatar as diferenças; quando
  pedirem "Foco Dexco", ajuste o eixo à Dexco e liste quem ficou fora da escala.
- Valores retirados (erro/outlier) não entram no gráfico; vão numa nota "Excluídos do gráfico:
  empresa, ano (motivo)". Índice sobre EBITDA com EBITDA negativo = n.s.
- Unidade abaixo do título do gráfico ("R$ milhões", "%", "vezes (x)", "dias").
- Fonte no rodapé: "Fonte: demonstrações financeiras das empresas; Análise de Competidores".
- Tabelas-resumo: empresas nas colunas (Dexco destacada com fundo `#ECEBE7`), indicadores nas
  linhas, ano de referência indicado; "–" para sem dado e "n.s." para retirado.
