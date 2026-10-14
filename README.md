# Painel Estratégico Dexco

Gera o site do time de estratégia, que junta o **Radar macro** (planilha *Grande Base*) e a
**Análise de Peers** (planilha *Análise de Competidores*), e o PowerPoint de peers no template
do Offsite.

- Site publicado: https://claude.ai/artifact/N8CHHxXKwp8fjJsHYg15bb (privado; o acesso do time é liberado pelo menu **Share** da página)

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
  --deck Analise_de_Peers_Dashboard.pptx   # opcional: também gera o PowerPoint
```

Isso recria `site/index.html` (página inicial), `site/radar.html` e `site/peers.html`. Para colocar
no ar, publique `site/index.html` como página do artifact acima, com `radar.html` e `peers.html`
como arquivos de apoio (é o que o Claude faz quando você pede para republicar).

## Estrutura

| Pasta / arquivo | O que é |
|---|---|
| `pipeline/build_site.py` | Comando único: lê as duas planilhas e gera o site (e o PowerPoint, com `--deck`). |
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
- Na tabela-resumo, valores retirados aparecem como **n.s.**; cada coluna usa o ano da receita mais recente da empresa.
- **Alavancagem da Dexco é sempre a consolidada**, em todos os negócios.
- **Madeira**: EBITDA calculado como receita × margem EBITDA (a aba não traz o valor absoluto).
- **Metais**: nos blocos "Mkt Competitivo/Médio/Luxo", a participação é calculada como empresa ÷ total do segmento (as primeiras linhas desses blocos estão em R$ milhões).

## Decisões em aberto (próxima fase)

- Madeira usa os números da planilha atual, que divergem do deck de junho/2026 (ex.: capital investido da Berneck, EBITDA da Floraplac).
- Docol 2022–23 (receita de R$ 1.100 e R$ 1.200 milhões, números redondos) mantida por enquanto.
