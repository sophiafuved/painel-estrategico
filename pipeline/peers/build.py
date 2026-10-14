import json,re,unicodedata
B=json.load(open('blocks.json'))
def norm(s):
    s=unicodedata.normalize('NFKD',s).encode('ascii','ignore').decode().lower()
    return re.sub(r'[^a-z0-9]+',' ',s).strip()
SEG={
 'loucas':dict(name='Louças',sheet='Gráficos - Louças',dex=['deca loucas','dexco','deca'],dexname='Deca Louças'),
 'metais':dict(name='Metais',sheet='Gráficos - Metais',dex=['deca metais','dexco'],dexname='Deca Metais'),
 'madeira':dict(name='Madeira',sheet='Gráficos Madeira 1',dex=['madeira total','dexco'],dexname='Dexco Madeira'),
 'revest':dict(name='Revestimentos',sheet='Gráficos - Revestimento',dex=['dx revestimento','dx revestimentos','dexco'],dexname='DX Revestimentos'),
}
COUNTRY={'BR':'BR','JP':'JP','ES':'ES','US':'US','IT':'IT','SUI':'CH','AL':'DE','MX':'MX','In':'BR','Out':'BR'}
M=[ # key,regex,label,fmt,group
 ('rol',r'^(ROL \(R\$ M\)|Receita L[ií]quida)$','Receita líquida','mi','escala'),
 ('ebitda',r'^EBITDA \(R\$ M\)$','EBITDA','mi','escala'),
 ('ll',r'^Lucro L[ií]quido \(R\$ M\)$','Lucro líquido','mi','escala'),
 ('ci',r'^Capital Investido','Capital investido','mi','escala'),
 ('roic',r'^ROIC( Te[oó]rico)?$','ROIC','pct','rent'),
 ('me',r'^Margem EBITDA$','Margem EBITDA','pct','rent'),
 ('mb',r'^Margem Bruta$','Margem bruta','pct','rent'),
 ('ml',r'^Margem L[ií]quida$','Margem líquida','pct','rent'),
 ('giro',r'^Giro de Ativos \(ROL/(IC|Capital Investido)\)$','Giro de ativos (ROL/CI)','x','rent'),
 ('alav',r'^Alavancagem Financeira$','Alavancagem (DL/EBITDA)','x','rent'),
 ('cpv',r'^CPV Caixa ?/ ?ROL$','CPV caixa / ROL','pct','custos'),
 ('dga',r'^(DGA|G&A)/ ?ROL$','DGA / ROL','pct','custos'),
 ('dv',r'^Desp','Despesas c/ vendas / ROL','pct','custos'),
 ('wk',r'^WK/ROL$','Capital de giro / ROL','pct','giro'),
 ('pmr',r'^(PMR|Dias de Receb[ií]veis)$','Dias de recebíveis','d','giro'),
 ('pme',r'^(PME|Dias de Estoque)$','Dias de estoque','d','giro'),
 ('pmp',r'^(PMP|Dias de Fornecedores)$','Dias de fornecedores','d','giro'),
 ('varwk',r'^Var\. WK/ ?EBITDA$','Variação do capital de giro / EBITDA','pct','caixa'),
 ('capexe',r'^Capex/ ?EBITDA$','Capex / EBITDA','pct','caixa'),
 ('outras',r'^Outras Desp\./ ?EBITDA$','Outras despesas / EBITDA','pct','caixa'),
 ('fcl',r'^FCL/ ?EBITDA$','FCL / EBITDA','pct','caixa'),
 ('capexr',r'^Capex/ROL$','Capex / ROL','pct','caixa'),
]
out={}
for sk,cfg in SEG.items():
    blocks=B[cfg['sheet']]
    ents={}; dexkey='dexco'
    def ekey(name,country=None):
        n=norm(name)
        if n in cfg['dex'] or n.startswith('dx revest') or n=='madeira total': 
            ents.setdefault(dexkey,dict(name=cfg['dexname'],country='BR',dex=1)); return dexkey
        k=n.replace(' ','_')
        if k not in ents: ents[k]=dict(name=re.sub(r'\s+',' ',name).strip(),country=COUNTRY.get(country or '', 'BR'))
        return k
    mets={}
    for key,rx,label,fmt,grp in M:
        b=next((b for b in blocks if re.search(rx,b['title'].strip())),None)
        if not b: continue
        ser={};cells={}
        for c in b['comp']:
            if norm(c['name']).startswith('portobello aju') or norm(c['name']).startswith('portobello a'): 
                if norm(c['name'])!='portobello': continue
            k=ekey(c['name'],c['country'])
            vals=[None if (v is None or v==0) else round(v,6) for v in c['vals']]
            if k in ser and any(x is not None for x in ser[k]): continue
            ser[k]=vals;cells[k]=c['row']
        mets[key]=dict(label=label,fmt=fmt,grp=grp,years=b['years'],series=ser,ref=f"'{cfg['sheet']}'!{b['col']}{b['row']}",sheet=cfg['sheet'],col=b['col'],rows=cells)
    # EBITDA Madeira = ROL x Margem
    if 'ebitda' not in mets and 'rol' in mets and 'me' in mets:
        r,m=mets['rol'],mets['me']; ser={}
        for k,v in r['series'].items():
            mm=m['series'].get(k)
            if mm: ser[k]=[(a*b if a is not None and b is not None else None) for a,b in zip(v,mm)]
        mets['ebitda']=dict(label='EBITDA',fmt='mi',grp='escala',years=r['years'],series=ser,ref='calculado: ROL × Margem EBITDA',calc=1)
    out[sk]=dict(name=cfg['name'],ents=ents,metrics=mets,sector=[])
# setoriais
def blk(sheet,title,nth=0):
    bs=[b for b in B[sheet] if b['title'].strip()==title]; return bs[nth] if len(bs)>nth else None
def sect(sk,sheet,title,label,fmt,skip=(),refs=(),share_of=None,unit=''):
    b=blk(sheet,title)
    if not b: print('MISSING',sk,title);return
    seg=out[sk]; ser=[]
    tot=None
    if share_of: tot=next(c['vals'] for c in b['comp'] if norm(c['name'])==share_of)
    for c in b['comp']:
        n=norm(c['name'])
        if n in skip or n=='-' : continue
        vals=c['vals']
        if tot: vals=[(v/t if v is not None and t else None) for v,t in zip(vals,tot)]
        isdex=n.startswith('deca') or n.startswith('dx ') or n=='dexco' or n=='madeira total'
        ser.append(dict(name=('Dexco' if n=='dexco' else re.sub(r'\s+',' ',c['name']).strip()),dex=1 if isdex else 0,ref=1 if n in refs else 0,vals=[None if v is None else round(v,6) for v in vals]))
    seg['sector'].append(dict(title=label,fmt=fmt,unit=unit,years=b['years'],series=ser,ref=f"'{sheet}'!{b['col']}{b['row']}"))
S='Gráficos - Louças'
for t,l in [('Mkt Share - Competitivo','Market share – Competitivo'),('Mkt Share - Médio','Market share – Médio'),('Mkt Share - Luxo','Market share – Luxo'),('Mkt Share - Médio/Luxo','Market share – Médio/Luxo')]:
    sect('loucas',S,t,l,'pct',skip=('mercado',),unit='% do valor do segmento')
sect('loucas',S,'Preços Médios - Competitivo','Preço médio – Competitivo','brl',refs=('mercado',),unit='R$ por peça')
sect('loucas',S,'Preços Médios - Médio/Luxo','Preço médio – Médio/Luxo','brl',refs=('mercado',),unit='R$ por peça')
S='Gráficos - Metais'
for t,l in [('Mkt Competitivo','Market share – Competitivo'),('Mkt Médio','Market share – Médio'),('Mkt Luxo','Market share – Luxo'),('Mkt Médio/Luxo','Market share – Médio/Luxo')]:
    # Competitivo/Médio/Luxo: as primeiras linhas do bloco estão em R$ mi; a participação = empresa / total do segmento
    sect('metais',S,t,l,'pct',skip=('total produtos acabados','total produtos'),unit='% do valor do segmento',share_of=None if t=='Mkt Médio/Luxo' else 'total produtos acabados')
sect('metais',S,'Preço Médio - Competitivo','Preço médio – Competitivo','brl',refs=('mercado',),unit='R$ por peça')
sect('metais',S,'Preço Médio - Médio/Luxo','Preço médio – Médio/Luxo','brl',refs=('mercado',),unit='R$ por peça')
S='Gráficos Madeira 1'
sect('madeira',S,'Dexco - ROIC vs WACC','Dexco: ROIC por negócio vs WACC','pct',refs=('wacc dexco',),unit='% a.a.')
sect('madeira',S,'Players Internacionais - ROIC','ROIC: players internacionais vs Dexco','pct',unit='%')
sect('madeira',S,'Taxa de Ocupação','Taxa de ocupação: Dexco vs mercado','pct',refs=('mercado mdf','mercado mdp'),unit='% da capacidade')
S='Gráficos - Revestimento'
sect('revest',S,'Value Share','Value share (participação em valor)','pct',skip=('mercado',),share_of='mercado',unit='% do mercado nacional · 2026 parcial')
sect('revest',S,'Preço Médio 90x90','Preço médio do porcelanato 90x90','brl',refs=('mercado',),unit='R$ por m²')
json.dump(out,open('peers.json','w'),ensure_ascii=False,separators=(',',':'))
import os;print(os.path.getsize('peers.json'))
for sk,s in out.items():
    print(sk,list(s['ents'].keys()),'| metrics',len(s['metrics']),[k for k in s['metrics']],'| sector',[x['title'] for x in s['sector']])
