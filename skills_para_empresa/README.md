# Skills do time de estratégia

Seis skills para o Claude, criadas a partir do trabalho no Painel Estratégico. Cada arquivo
`.skill` é um pacote (ZIP) pronto para instalar; o código-fonte de cada uma está em
`.claude/skills/<nome>/` neste projeto.

| Skill | Para que serve | Depende do projeto? |
|---|---|---|
| `atualizar-painel-estrategico` | Gerar, conferir e publicar o painel (Radar + Peers) com senha | Sim (scripts em `pipeline/`) |
| `conferir-peers` | Conferir os números da Análise de Competidores e gerar a planilha de conferência (com "o que mudou" desde a versão anterior) | Sim (usa os leitores de `pipeline/peers/`) |
| `incluir-peer-novo` | Levar os números de um release/DFP de um concorrente para o formato da planilha, com fonte e página | Não |
| `comentarios-radar` | Escrever títulos e textos do Radar no estilo do time, apoiados nos números | Não |
| `atualizar-deck-radar` | Atualizar o PowerPoint do Radar (Modelo 2) com os dados do painel e a leitura do time, sem quebrar o vínculo com a Grande Base | Sim (usa as faixas de `pipeline/radar/`) |
| `padrao-apresentacao-dexco` | Padrão visual do Radar e do template do Offsite/Peers (traz o template .pptx) | Não |

## Como instalar

**No Claude Code ou no Claude conectado à pasta do projeto:** não precisa instalar. As skills
estão em `.claude/skills/` e são carregadas automaticamente quando a pasta do projeto é aberta.

**No Claude da empresa (claude.ai), para uso individual:** em Configurações, na área de Skills
(Capacidades), faça o upload do arquivo `.skill` (ou renomeie para `.zip`, se o site pedir ZIP).
Repita para cada skill que quiser usar.

**Para o time todo:** peça ao administrador do Claude da Dexco que publique os arquivos `.skill`
para a organização; assim elas aparecem habilitadas para todos.

Os passos e nomes das telas podem mudar com as versões do Claude; se não encontrar a opção, peça
ao próprio Claude "como instalo uma skill a partir de um arquivo .skill".

## Cuidados

- As skills levam o **processo**, não os dados: nenhuma planilha, número ou senha está dentro delas.
- `padrao-apresentacao-dexco` inclui o template do Offsite (`template_offsite_peers.pptx`):
  compartilhe só dentro da Dexco.
- Ao mudar uma skill, edite a pasta em `.claude/skills/<nome>/` e gere o pacote de novo
  (peça ao Claude: "empacote a skill <nome>"), para que todos recebam a mesma versão.
