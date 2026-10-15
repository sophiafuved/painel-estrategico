"""Gera o Painel Estratégico (Radar + Peers) a partir das duas planilhas.

Uso:
  python build_site.py --radar "Grande Base.xlsx" --peers "Analise de Competidores.xlsx" --out ../site [--deck ../Analise_de_Peers.pptx]
                      [--proj "102026_Dexco.xlsx" [--proj-desde 2026]]   # projeções da consultoria no Radar

Saída (pasta --out): index.html (página inicial), radar.html, peers.html.
Publicar: index.html como página do artifact e os outros dois como arquivos de apoio.
"""
import argparse, os, sys, json, shutil, subprocess, tempfile, datetime, re

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
PROTEGIDO = False
PEERS_SHEETS = ['Gráficos - Louças', 'Gráficos - Metais', 'Gráficos Madeira 1', 'Gráficos - Revestimento']
MES = ['jan','fev','mar','abr','mai','jun','jul','ago','set','out','nov','dez']

def run(script, cwd, *args, env=None):
    r = subprocess.run([sys.executable, '-I', os.path.join(HERE, script), *args], cwd=cwd,
                       env={**os.environ, **(env or {})}, capture_output=True, text=True)
    if r.returncode != 0:
        sys.exit(f'Falha em {script}:\n{r.stderr[-3000:]}')
    return r.stdout

def page_doc(content, title_html):
    """Arquivo de apoio: precisa ser um documento completo."""
    return ('<!doctype html><html lang="pt-BR"><head><meta charset="utf-8">'
            '<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover"><style>body{margin:0}</style></head><body>'
            + content + '</body></html>')

def sitebar(active, sair=None):
    sair = PROTEGIDO if sair is None else sair
    links = [('index.html', 'Início', 'home'), ('radar.html', 'Radar macro', 'radar'), ('peers.html', 'Peers', 'peers')]
    items = ''.join(f'<a href="{h}" class="{"on" if k == active else ""}"{" aria-current=page" if k == active else ""}>{t}</a>' for h, t, k in links)
    if sair:
        items += ('<button type="button" class="pe-sair" onclick="try{for(const st of [localStorage,sessionStorage])'
                  'for(const k of Object.keys(st))if(k.startsWith(\'painel-chave-\'))st.removeItem(k)}catch(e){}location.reload()">Sair</button>')
    return f'''<style>
.pe-sair{{font:inherit;color:var(--pe-mut);background:none;border:1px solid #444;border-radius:6px;padding:4px 10px;cursor:pointer;margin-left:6px}}
.pe-sair:hover{{color:var(--pe-fg)}}
.pe-bar{{--pe-bg:#111;--pe-fg:#f1f0ed;--pe-mut:#a8a6a1;display:flex;flex-wrap:wrap;align-items:center;gap:6px 18px;background:var(--pe-bg);color:var(--pe-fg);
 margin:0 calc(-1*max(16px,env(safe-area-inset-right,0px))) 0 calc(-1*max(16px,env(safe-area-inset-left,0px)));
 padding:10px max(16px,env(safe-area-inset-right,0px)) 10px max(16px,env(safe-area-inset-left,0px));font:500 14px "Red Hat Text",system-ui,sans-serif}}
.pe-bar b{{font:800 15px "Red Hat Display",system-ui,sans-serif;letter-spacing:-.01em;margin-right:auto}}
.pe-bar nav{{display:flex;gap:4px;flex-wrap:wrap}}
.pe-bar a{{color:var(--pe-mut);text-decoration:none;padding:5px 10px;border-radius:6px}}
.pe-bar a:hover{{color:var(--pe-fg)}}
.pe-bar a.on{{background:var(--pe-fg);color:var(--pe-bg);font-weight:600}}
.pe-bar a:focus-visible{{outline:2px solid #4e86da;outline-offset:2px}}
</style><div class="pe-bar"><b>Painel Estratégico Dexco</b><nav aria-label="Seções do painel">{items}</nav></div>'''

def build_radar(xlsx, work, proj=None):
    run('radar/extract.py', work, os.path.abspath(xlsx))
    run('radar/build_data.py', work)
    data = open(os.path.join(work, 'data.json')).read()
    t = open(os.path.join(HERE, 'radar/template.html')).read()
    # tipografia comum ao painel
    t = t.replace('family=Archivo:wdth,wght@62..125,500..800&family=Source+Sans+3:wght@400;600&', 'family=Red+Hat+Display:wght@500;700;800;900&family=Red+Hat+Text:wght@400;500;600&')
    t = t.replace('--display:"Archivo","Arial Narrow",system-ui,sans-serif;', '--display:"Red Hat Display",system-ui,sans-serif;')
    t = t.replace('--body:"Source Sans 3",system-ui,-apple-system,"Segoe UI",sans-serif;', '--body:"Red Hat Text",system-ui,-apple-system,"Segoe UI",sans-serif;')
    t = t.replace('<title>Radar Estratégia</title>', '<title>Radar Estratégia</title>' + sitebar('radar'), 1)
    D = json.loads(data)
    last = D['iie']['labels'][-1]
    if proj:  # projeções da consultoria entram como D.tend
        D['tend'] = proj
        data = json.dumps(D, ensure_ascii=False, separators=(',', ':'))
    return t.replace('__DATA__', data.replace('</', '<\\/')), f"{MES[int(last[5:7])-1]}/{last[2:4]}"

def build_peers(xlsx, work):
    run('peers/parse_blocks.py', work, *PEERS_SHEETS, env={'PEERS_XLSX': os.path.abspath(xlsx)})
    run('peers/build.py', work)
    run('peers/clean.py', work)
    D = json.load(open(os.path.join(work, 'peers_clean.json')))
    for s in D.values(): s.pop('log', None)
    t = open(os.path.join(HERE, 'peers/template.html')).read()
    t = t.replace('<title>Peers Dexco</title>', '<title>Peers Dexco</title>' + sitebar('peers'), 1)
    # abre direto num negócio pelo link (ex.: peers.html#metais)
    t = t.replace("renderAll();\nconst rer", "if(D[location.hash.slice(1)])ST.seg=location.hash.slice(1);\naddEventListener('hashchange',()=>{if(D[location.hash.slice(1)]){ST.seg=location.hash.slice(1);renderAll()}});\nrenderAll();\nconst rer", 1)
    years = sorted({y for s in D.values() for m in s['metrics'].values() for y in m['years'] if 'LTM' not in y})
    log = json.load(open(os.path.join(work, 'ajustes.json')))
    # base para o PowerPoint: inclui a lista de ajustes (slides "Ajustes na base")
    full = json.load(open(os.path.join(work, 'peers_clean.json')))
    for seg in full:
        full[seg]['log'] = [{k: l[k] for k in ('kind','metric','mk','ent','year','v','fmt','why','cell')} for l in log if l['seg'] == seg]
    json.dump(full, open(os.path.join(work, 'peers_deck.json'), 'w'), ensure_ascii=False)
    return t.replace('__DATA__', json.dumps(D, ensure_ascii=False, separators=(',', ':')).replace('</', '<\\/')), years[-1], {s: D[s]['name'] for s in D}, len(log)

def hub(radar_ate, peers_ate, segs, built):
    seg_links = ''.join(f'<a href="peers.html#{k}">{v}</a>' for k, v in segs.items())
    return f'''<title>Painel Estratégico</title>
<meta name="claude:viewport-fit-sides" content="cover">
<link rel="preconnect" href="https://fonts.googleapis.com"><link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Red+Hat+Display:wght@700;900&family=Red+Hat+Text:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap">
<style>
/* Página inicial: duas portas de entrada (Radar macro e Peers), com a data dos dados de cada uma. */
:root{{--bg:#f4f4f2;--panel:#fff;--line:#e1e0dc;--ink:#111;--ink-2:#4d4c48;--ink-3:#7a7974;--acc:#e0661a;--acc-2:#2a78d6;
 --display:"Red Hat Display",system-ui,sans-serif;--body:"Red Hat Text",system-ui,-apple-system,sans-serif;--mono:"IBM Plex Mono",ui-monospace,Menlo,monospace}}
@media (prefers-color-scheme: dark){{:root:not([data-theme="light"]){{--bg:#141517;--panel:#1c1d20;--line:#33353a;--ink:#f1f0ed;--ink-2:#bdbbb5;--ink-3:#8c8a85;--acc:#e8732a;--acc-2:#4e86da;color-scheme:dark}}}}
:root[data-theme="dark"]{{--bg:#141517;--panel:#1c1d20;--line:#33353a;--ink:#f1f0ed;--ink-2:#bdbbb5;--ink-3:#8c8a85;--acc:#e8732a;--acc-2:#4e86da;color-scheme:dark}}
*{{box-sizing:border-box}}
body{{background:var(--bg);color:var(--ink);font:15px/1.5 var(--body);padding-left:max(16px,env(safe-area-inset-left,0px));padding-right:max(16px,env(safe-area-inset-right,0px));padding-block:0 48px}}
.wrap{{max-width:1080px;margin:0 auto}}
header{{padding-block:44px 28px}}
.eyebrow{{font:500 12px var(--mono);letter-spacing:.08em;text-transform:uppercase;color:var(--ink-3)}}
h1{{font:900 clamp(34px,5.5vw,56px)/1.02 var(--display);letter-spacing:-.02em;margin:8px 0 10px;text-wrap:balance}}
header p{{color:var(--ink-2);max-width:62ch;margin:0}}
.doors{{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:18px}}
.door{{display:flex;flex-direction:column;gap:12px;background:var(--panel);border:1px solid var(--line);border-radius:12px;padding:22px;text-decoration:none;color:inherit;min-width:0;transition:border-color .15s}}
.door:hover{{border-color:var(--ink-3)}}
.door:focus-visible{{outline:2px solid var(--acc-2);outline-offset:3px}}
.door .tag{{font:500 12px var(--mono);letter-spacing:.06em;text-transform:uppercase}}
.door.radar .tag{{color:var(--acc)}} .door.peers .tag{{color:var(--acc-2)}}
.door h2{{font:700 26px/1.15 var(--display);margin:0}}
.door p{{margin:0;color:var(--ink-2)}}
.door ul{{margin:0;padding-left:18px;color:var(--ink-2);font-size:14px}}
.meta{{display:flex;flex-wrap:wrap;gap:6px 16px;font:12.5px var(--mono);color:var(--ink-3);border-top:1px solid var(--line);padding-top:12px;margin-top:auto}}
.go{{font-weight:600;color:var(--ink)}}
.segs{{display:flex;flex-wrap:wrap;gap:8px;margin-top:18px}}
.segs span{{font:500 12px var(--mono);letter-spacing:.06em;text-transform:uppercase;color:var(--ink-3);align-self:center;margin-right:4px}}
.segs a{{font:600 13px var(--body);color:var(--ink-2);text-decoration:none;border:1px solid var(--line);background:var(--panel);border-radius:999px;padding:5px 12px}}
.segs a:hover{{color:var(--ink);border-color:var(--ink-3)}}
footer{{margin-top:36px;font-size:13px;color:var(--ink-3);max-width:80ch}}
@media (max-width:760px){{.doors{{grid-template-columns:minmax(0,1fr)}}}}
</style>
{sitebar('home')}
<div class="wrap">
<header>
  <div class="eyebrow">Planejamento estratégico · Offsite 2026</div>
  <h1>Painel Estratégico</h1>
  <p>Cenário macro do Radar de Estratégia e benchmarking da Dexco contra os peers de cada negócio, num só lugar.</p>
</header>
<div class="doors">
  <a class="door radar" href="radar.html">
    <span class="tag">Radar macro</span>
    <h2>Radar Estratégia</h2>
    <p>Indicadores de conjuntura que alimentam o Radar: global, atividade, juros, renda, crédito, custos da construção, imobiliário e varejo.</p>
    <ul><li>27 gráficos com filtro de período</li><li>Indicadores-chave com variação em 12 meses</li></ul>
    <div class="meta"><span>Dados até {radar_ate}</span><span>Fonte: Grande Base</span><span class="go">Abrir →</span></div>
  </a>
  <a class="door peers" href="peers.html">
    <span class="tag">Benchmarking</span>
    <h2>Análise de Peers</h2>
    <p>Escala, rentabilidade, custos, capital de giro, fluxo de caixa e dados setoriais da Dexco frente aos concorrentes.</p>
    <ul><li>Louças, Metais, Madeira e Revestimentos</li><li>Escolha das empresas e do período, com Foco Dexco</li></ul>
    <div class="meta"><span>Demonstrações até {peers_ate}</span><span>Fonte: Análise de Competidores</span><span class="go">Abrir →</span></div>
  </a>
</div>
<div class="segs"><span>Ir direto para</span>{seg_links}</div>
<footer>Atualizado em {built} a partir das planilhas Grande Base e Análise de Competidores. As regras de tratamento dos dados estão descritas no rodapé de cada página.</footer>
</div>'''

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--radar', required=True); ap.add_argument('--peers', required=True); ap.add_argument('--out', required=True)
    ap.add_argument('--proj', help='planilha de projeções da consultoria (ex.: Tendências) para o Radar')
    ap.add_argument('--proj-desde', type=int, help='primeiro ano tratado como projeção (padrão: ano da edição no nome do arquivo)')
    ap.add_argument('--deck', help='também gera o PowerPoint de peers neste caminho (.pptx)')
    ap.add_argument('--senha', default=os.environ.get('PAINEL_SENHA'), help='protege o site com senha (ou variável PAINEL_SENHA). Nunca salve a senha no repositório.')
    a = ap.parse_args()
    global PROTEGIDO
    PROTEGIDO = bool(a.senha)
    os.makedirs(a.out, exist_ok=True)
    with tempfile.TemporaryDirectory() as w1, tempfile.TemporaryDirectory() as w2:
        proj, sem_topico = (None, [])
        if a.proj:
            from radar.projecoes import ler
            proj, sem_topico = ler(a.proj, a.proj_desde)
        radar_html, radar_ate = build_radar(a.radar, w1, proj)
        peers_html, peers_ate, segs, nlog = build_peers(a.peers, w2)
        if a.deck:
            root = os.path.dirname(HERE)
            r = subprocess.run([sys.executable, '-I', os.path.join(root, 'deck', 'build_deck.py'), os.path.join(w2, 'peers_deck.json'),
                                os.path.join(root, 'deck', 'template_dashboard.pptx'), os.path.abspath(a.deck)], capture_output=True, text=True)
            if r.returncode != 0: sys.exit('Falha no PowerPoint:\n' + r.stderr[-3000:])
    built = datetime.datetime.now().strftime('%d/%m/%Y')
    pages = {'radar.html': page_doc(radar_html, 'Radar'), 'peers.html': page_doc(peers_html, 'Peers'),
             'index.html': hub(radar_ate, peers_ate, segs, built)}
    if a.senha:
        from protect import new_salt, derive_key, protect
        pages = {k: v.replace('</nav></div>', '</nav></div>', 1) for k, v in pages.items()}
        salt = new_salt(); key = derive_key(a.senha, salt)
        titles = {'index.html': 'Painel Estratégico', 'radar.html': 'Radar Estratégia', 'peers.html': 'Peers Dexco'}
        pages = {k: protect(v, key, salt, titles[k], standalone=(k != 'index.html')) for k, v in pages.items()}
    for k, v in pages.items():
        open(os.path.join(a.out, k), 'w').write(v)
    res = {'radar_ate': radar_ate, 'peers_ate': peers_ate, 'ajustes_peers': nlog, 'out': os.path.abspath(a.out)}
    if proj:
        res['projecoes'] = {'arquivo': proj['arquivo'], 'edicao': proj['edicao'], 'projecao_desde': proj['desde'],
                            'indicadores': len(proj['itens']), 'sem_topico': sem_topico}
    print(json.dumps(res, ensure_ascii=False))

if __name__ == '__main__':
    main()
