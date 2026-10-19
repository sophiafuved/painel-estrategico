import openpyxl,warnings,json,datetime,re,sys,os
from lxml import etree
warnings.filterwarnings('ignore')
wb=openpyxl.load_workbook(sys.argv[1],data_only=True)
def rng(sheet,col,r1,r2):
    ws=wb[sheet]; return [ws[f'{col}{r}'].value for r in range(r1,r2+1)]
def norm_cat(v):
    if isinstance(v,datetime.datetime): return v.strftime('%Y-%m-%d')
    if isinstance(v,float) and v.is_integer(): v=int(v)
    return str(v).strip() if v is not None else None
def num(v):
    if isinstance(v,(int,float)) and not isinstance(v,bool): return float(v)
    return None
sys.path.insert(0,os.path.dirname(os.path.abspath(__file__)))
from faixas import SPECS as specs
out={}
for k,ser in specs.items():
    out[k]=[]
    for s in ser:
        vals=rng(s['sheet'],s['col'],s['r1'],s['r2'])
        cats=rng(s['cs'],s['cc'],s['cr'],s['cr']+len(vals)-1)
        pts=[[norm_cat(c),num(v)] for c,v in zip(cats,vals) if c is not None]
        out[k].append(dict(label=s['label'],ref=f"{s['sheet']}!{s['col']}{s['r1']}:{s['col']}{s['r2']}",data=pts))
# from pptx caches (not in Grande Base)
ns={'c':'http://schemas.openxmlformats.org/drawingml/2006/chart'}
def pcache(path):
    t=etree.parse(path); res=[]
    for s in t.iter('{%s}ser'%ns['c']):
        def arr(n):
            if n is None: return []
            c=int(n.find('.//c:ptCount',ns).get('val')); a=[None]*c
            for p in n.findall('.//c:pt',ns): a[int(p.get('idx'))]=p.find('c:v',ns).text
            return a
        name=arr(s.find('c:tx',ns)); cat=arr(s.find('c:cat',ns)); val=arr(s.find('c:val',ns))
        res.append(dict(name=name[0] if name else None,cat=cat,val=val))
    return res
# Brent e impacto fiscal não estão na Grande Base: vêm da apresentação Modelo 2 (dados fixos)
ST=json.load(open(os.path.join(os.path.dirname(os.path.abspath(__file__)),'static_apresentacao.json')))
out['_brent']=ST['_brent'];out['_fiscal']=ST['_fiscal']
json.dump(out,open('data_raw.json','w'),ensure_ascii=False)
for k,v in out.items():
    if k.startswith('_'):
        for s in v: print(k,s['name'],len(s['cat']),s['cat'][:3],s['cat'][-2:],s['val'][:3],s['val'][-3:])
        continue
    for s in v:
        d=[p for p in s['data'] if p[1] is not None]
        print(k,'|',s['label'],'| n',len(s['data']),'nn',len(d),'|',d[:2],d[-2:])
