# Painel Estratégico Dexco

Gera o site do time de estratégia, que junta o **Radar macro** (planilha *Grande Base*) e a
**Análise de Peers** (planilha *Análise de Competidores*), e o PowerPoint de peers no template
do Offsite.

- Site publicado: https://claude.ai/artifact/N8CHHxXKwp8fjJsHYg15bb (versão da conta pessoal da Sophia; quando o time publicar pela conta de trabalho, substitua este link pelo novo)
- Usando com o Claude: conecte esta pasta ao Claude e peça, por exemplo, "gere e publique o painel". As instruções que o Claude segue estão no `CLAUDE.md`.

## Como atualizar

Instale as dependências uma vez:

```bash
pip install -r requirements.txt
```

Depois, sempre que uma das planilhas mudar:

```bash
python pipeline/build_site.py \
  --radar "Grande Base.xlsx" \
  --peers "Analise de Competidores.xlsx" \
  --out site \
  --proj "102026_Dexco.xlsx" \             # opcional: projeções da Tendências no Radar
  --deck Analise_de_Peers_Dashboard.pptx   # opcional: também gera o PowerPoint
```

Para proteger o site com senha (necessário se o link for aberto a "qualquer pessoa com o link"),
passe a senha só na hora de gerar, pela variável `PAINEL_SENHA` ou por `--senha`. **Nunca salve a
senha no repositório.** O conteúdo de cada página é criptografado (AES-256); sem a senha, o arquivo
publicado mostra apenas a tela de acesso.

```bash
PAINEL_SENHA='a senha do time' python pipeline/build_site.py --radar ... --peers ... --out site
```

Isso recria `site/index.html` (página inicial), `site/radar.html` e `site/peers.html`. Para colocar
no ar, publique `site/index.html` como página do artifact acima, com `radar.html` e `peers.html`
como arquivos de apoio (é o que o Claude faz quando você pede para republicar).

## Primeira vez rodando (passo a passo no Windows)

1. **Instale o Python** (versão 3.10 ou mais nova) em https://www.python.org/downloads/. Na
   instalação, marque **"Add python.exe to PATH"**.
2. **Baixe o projeto**: na página do repositório no GitHub, clique em **Code → Download ZIP** e
   descompacte numa pasta (ex.: `Documentos\painel-estrategico`). Quem tiver o Git instalado pode
   usar `git clone` no lugar.
3. **Copie as duas planilhas** do SharePoint para essa mesma pasta, com os nomes
   `Grande Base.xlsx` e `Analise de Competidores.xlsx`.
4. **Abra o terminal na pasta**: no Explorador de Arquivos, clique na barra de endereço, digite
   `powershell` e tecle Enter.
5. **Instale as dependências** (só na primeira vez):

   ```powershell
   python -m pip install -r requirements.txt
   ```

6. **Gere o site e o PowerPoint** (troque a senha pela senha do time):

   ```powershell
   python pipeline/build_site.py --radar "Grande Base.xlsx" --peers "Analise de Competidores.xlsx" --out site --deck Analise_de_Peers_Dashboard.pptx --senha "a senha do time"
   ```

   Sem `--senha`, o site sai sem proteção (use só para conferir no próprio computador).
7. **Confira**: abra `site\index.html` no navegador (Chrome ou Edge). Precisa de internet para
   carregar a biblioteca dos gráficos.

Se aparecer `Falha em ...` com uma mensagem de erro, normalmente é porque a estrutura de alguma
planilha mudou (aba renomeada, bloco movido de lugar). O gerador lê blocos em posições e nomes
específicos; veja "Onde o gerador procura os dados" abaixo.

### Onde o gerador procura os dados

- **Grande Base**: as faixas de células usadas por cada gráfico estão em `pipeline/radar/extract.py`
  (ex.: `Dados_Tendências!N155:N328`). Se uma série crescer para novas linhas, ajuste a faixa ali.
- **Análise de Competidores**: os blocos das abas `Gráficos - Louças`, `Gráficos - Metais`,
  `Gráficos Madeira 1` e `Gráficos - Revestimento`, reconhecidos pelo título do bloco (ex.: "ROIC",
  "Margem EBITDA") seguido dos anos. Novas colunas de ano e novas empresas no fim de um bloco são
  lidas automaticamente.

### O que não está no repositório

- **As planilhas** (ficam no SharePoint).
- **A senha do site** (é passada só na hora de gerar e deve ser combinada com o time).
- **O link publicado** pertence à conta de quem o criou. Outra pessoa consegue gerar os arquivos,
  mas para republicar **no mesmo link** precisa ter acesso de edição ao artifact (menu **Share**)
  ou pedir à dona do link.

## Estrutura

| Pasta / arquivo | O que é |
|---|---|
| `pipeline/build_site.py` | Comando único: lê as duas planilhas e gera o site (e o PowerPoint, com `--deck`). |
| `pipeline/protect.py` | Proteção com senha: criptografa cada página e coloca a tela de acesso na frente. |
| `pipeline/radar/` | Extração da Grande Base (`extract.py`, `build_data.py`) e modelo da página do Radar. `static_apresentacao.json` guarda o Brent e o impacto fiscal, que só existem na apresentação Modelo 2. |
| `pipeline/peers/` | Leitura das abas "Gráficos" de cada negócio (`parse_blocks.py`, `build.py`), tratamento dos dados (`clean.py`) e modelo da página de Peers. |
| `deck/` | Gerador do PowerPoint (`build_deck.py`) e o template do Offsite. |
| `site/` | Última versão gerada do site. |

## Regras de tratamento dos dados de Peers

As regras ficam em `pipeline/peers/build.py` e `pipeline/peers/clean.py`, e são aplicadas sempre que o site é gerado:

- **Fonte**: blocos das abas "Gráficos - Louças", "Gráficos - Metais", "Gráficos Madeira 1" e "Gráficos - Revestimento".
- **Células zeradas** nas séries financeiras são lidas como "sem dado".
- **Erros de origem** são retirados: lucro líquido igual à receita (e as margens de 100% que vêm disso) e prazos médios negativos.
- **Outliers** são retirados: |dívida líquida/EBITDA| > 10x; índices sobre EBITDA acima de 500% em módulo; prazos acima de 365 dias; giro acima de 6x; margens e ROIC abaixo de −50%; capex acima de 50% da receita.
- Cada tópico abre com a sua tabela-resumo (não há resumo geral no topo). Valores retirados aparecem como **n.s.**; cada coluna usa o ano da receita mais recente da empresa.
- **Alavancagem da Dexco é sempre a consolidada**, em todos os negócios.
- **Madeira**: EBITDA calculado como receita × margem EBITDA (a aba não traz o valor absoluto).
- **Metais**: nos blocos "Mkt Competitivo/Médio/Luxo", a participação é calculada como empresa ÷ total do segmento (as primeiras linhas desses blocos estão em R$ milhões).

## Projeções da consultoria (Radar)

Opcional: passe o arquivo da Tendências com `--proj "102026_Dexco.xlsx"`. Cada indicador vai para o
tópico do Radar definido em `pipeline/radar/projecoes_map.json` e aparece num cartão com realizado
e projeção (botão Mostrar/Ocultar no topo). A projeção começa no ano da edição indicada no nome do
arquivo (102026 = out/2026 → projeção a partir de 2026); para mudar, use `--proj-desde 2027`.
Indicadores novos que não estão no mapa aparecem em `projecoes.sem_topico` no resumo do comando.

## Recursos interativos

- **Peers**: faixa dos peers (1º–3º quartil e mediana, sem a Dexco) nos gráficos de índices; matriz de posicionamento com eixos e tamanho da bolha à escolha; botões Todos/Brasil/Internacionais selecionam todas as empresas do grupo.
- **Radar**: "Monte seu gráfico" para comparar até 3 séries de qualquer tópico, em nível, variação em 12 meses ou base 100.
- Nos dois painéis, clicar na legenda esconde ou mostra a série.

## Formato do Radar

- Sem textos de comentário: os títulos dos tópicos são neutros; comentários entram só na apresentação.
- Cada tópico abre com um resumo dos seus indicadores (último dado observado, sem projeções).
- Gráficos que têm dois eixos na apresentação mantêm os dois eixos; a legenda indica "esq." ou "dir.".

## Decisões em aberto (próxima fase)

- Madeira usa os números da planilha atual, que divergem do deck de junho/2026 (ex.: capital investido da Berneck, EBITDA da Floraplac).
- Docol 2022–23 (receita de R$ 1.100 e R$ 1.200 milhões, números redondos) mantida por enquanto.
