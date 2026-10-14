import json
from openpyxl.utils import column_index_from_string as ci, get_column_letter as cl
D=json.load(open('peers.json'))
LOG=[]
# Alavancagem da Dexco é sempre a consolidada (linha 'Dexco' dos blocos de Louças/Metais/Revestimentos)
cons={}
for sg in ('loucas','metais','revest'):
    m=D[sg]['metrics']['alav']
    for y,v in zip(m['years'],m['series']['dexco']):
        if v is not None and 'LTM' not in y: cons.setdefault(y,v)
for sg,S in D.items():
    m=S['metrics'].get('alav')
    if not m: continue
    m['series']['dexco']=[cons.get(y) for y in m['years']]
    m['dexCons']=1
def cell(m,k,i):
    return f"'{m['sheet']}'!{cl(ci(m['col'])+1+i)}{m['rows'][k]}" if 'rows' in m and k in m['rows'] else m['ref']
def excl(seg,mk,k,i,kind,why):
    m=D[seg]['metrics'][mk];v=m['series'][k][i]
    if v is None: return
    m.setdefault('excl',{}).setdefault(k,{})[m['years'][i]]=dict(v=v,kind=kind,why=why)
    m['series'][k][i]=None
    LOG.append(dict(seg=seg,metric=m['label'],mk=mk,ent=D[seg]['ents'][k]['name'],year=m['years'][i],v=v,fmt=m['fmt'],kind=kind,why=why,cell=cell(m,k,i)))
OUT={'alav':(lambda v:abs(v)>10,'dívida líquida/EBITDA acima de 10x em módulo: EBITDA próximo de zero, índice sem significado'),
     'fcl':(lambda v:abs(v)>5,'acima de 500% em módulo: EBITDA próximo de zero ou negativo'),
     'varwk':(lambda v:abs(v)>5,'acima de 500% em módulo: EBITDA próximo de zero'),
     'capexe':(lambda v:abs(v)>5,'acima de 500% em módulo: EBITDA próximo de zero'),
     'outras':(lambda v:abs(v)>5,'acima de 500% em módulo: EBITDA próximo de zero'),
     'pmr':(lambda v:v>365,'prazo acima de 365 dias'),'pme':(lambda v:v>365,'prazo acima de 365 dias'),'pmp':(lambda v:v>365,'prazo acima de 365 dias'),
     'giro':(lambda v:v>6,'giro acima de 6x: capital investido muito baixo ou inconsistente no ano'),
     'me':(lambda v:v<-.5,'margem abaixo de −50%'),'ml':(lambda v:v<-.5,'margem abaixo de −50%'),'mb':(lambda v:v<-.5,'margem abaixo de −50%'),
     'roic':(lambda v:v<-.5,'ROIC abaixo de −50%'),'wk':(lambda v:v<-.5,'capital de giro abaixo de −50% da receita'),
     'capexr':(lambda v:v>.5,'capex acima de 50% da receita (ano de implantação)')}
for seg,S in D.items():
    M=S['metrics']
    # 1) erros: lucro igual à receita e margens de 100% derivadas
    if 'll' in M and 'rol' in M:
        for k,v in list(M['ll']['series'].items()):
            r=M['rol']['series'].get(k) or []
            for i,(a,b) in enumerate(zip(v,r)):
                if a is not None and b is not None and abs(a-b)<1e-6:
                    excl(seg,'ll',k,i,'erro','lucro líquido igual à receita líquida (célula puxando a linha errada)')
                    for mk in ['ml','mb']:
                        if mk in M and k in M[mk]['series'] and M[mk]['series'][k][i] is not None and abs(M[mk]['series'][k][i]-1)<1e-6:
                            excl(seg,mk,k,i,'erro','margem de 100% decorrente do mesmo erro (lucro = receita)')
    # margens = 100% exatos restantes
    for mk in ['mb','me','ml']:
        if mk not in M: continue
        for k,v in M[mk]['series'].items():
            for i,a in enumerate(v):
                if a is not None and abs(a-1)<1e-9: excl(seg,mk,k,i,'erro','margem de exatamente 100% (valor de preenchimento)')
    # dias negativos
    for mk in ['pmr','pme','pmp']:
        if mk not in M: continue
        for k,v in M[mk]['series'].items():
            for i,a in enumerate(v):
                if a is not None and a<0: excl(seg,mk,k,i,'erro','prazo médio negativo (impossível)')
    # 2) outliers
    for mk,(f,why) in OUT.items():
        if mk not in M: continue
        for k,v in M[mk]['series'].items():
            for i,a in enumerate(v):
                if a is not None and f(a): excl(seg,mk,k,i,'outlier',why)
json.dump(D,open('peers_clean.json','w'),ensure_ascii=False,separators=(',',':'))
json.dump(LOG,open('ajustes.json','w'),ensure_ascii=False,indent=1)
def f(v,fm):
    return f'{v*100:.1f}%' if fm=='pct' else f'{v:.1f}x' if fm=='x' else f'{v:,.1f}'
for l in LOG: print(f"{l['kind']:7} {l['seg']:7} {l['metric'][:26]:26} {l['ent'][:22]:22} {l['year']:8} {f(l['v'],l['fmt']):>12}  {l['cell']}")
print(len(LOG))
