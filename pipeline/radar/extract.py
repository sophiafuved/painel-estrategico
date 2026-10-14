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
# series: (label, sheet, col, r1, r2, catSheet, catCol, catR1)
S=lambda label,sheet,col,r1,r2,cs=None,cc='B',cr=None:dict(label=label,sheet=sheet,col=col,r1=r1,r2=r2,cs=cs or sheet,cc=cc,cr=cr or r1)
specs={
 'wui':[S('Índice de Incerteza Global (WUI)','Dados_Extra','W',9,131,cc='P')],
 'gpr':[S('Índice mensal','Dados_Extra','G',9,509),S('mm12m','Dados_Extra','H',9,509)],
 'trade':[S('Global Trade (var. % a/a)','Apoio_Radar','BK',35,52,cc='BI'),S('Share EUA (%)','Apoio_Radar','BL',35,52,cc='BI'),S('Mediana','Apoio_Radar','BM',35,52,cc='BI')],
 'infl_glob':[S('Inflação EUA','Dados_ITAÚ','G',9,27),S('Inflação Zona do Euro','Dados_ITAÚ','N',9,27),S('Média EUA (10-19)','Dados_ITAÚ','H',9,27),S('Média Zona do Euro (10-19)','Dados_ITAÚ','O',9,27)],
 'pib':[S('Serviços','Apoio_Radar','AN',82,88,cc='AK'),S('Indústria','Apoio_Radar','AM',82,88,cc='AK'),S('Agro','Apoio_Radar','AL',82,88,cc='AK'),S('PIB','Apoio_Radar','AQ',82,88,cc='AK')],
 'hiato':[S('Hiato do Produto','Dados_BACEN','D',35,101),S('IPCA','Dados_BACEN','E',35,101)],
 'focus_selic_ipca':[S('SELIC 2026','Dados_Focus','M',1023,1711),S('SELIC 2027','Dados_Focus','Q',1023,1711),S('IPCA 2026','Dados_Focus','L',1023,1711),S('IPCA 2027','Dados_Focus','P',1023,1711)],
 'focus_cambio':[S('Câmbio 2026','Dados_Focus','N',1023,1711),S('Câmbio 2027','Dados_Focus','R',1023,1711)],
 'neutra':[S('Taxa de Juros Neutra Real','Dados_BACEN','G',35,108),S('SELIC Real (ex-ante)','Dados_BACEN','H',35,108)],
 'iie':[S('Índice mensal','Dados_FGV','H',128,328),S('mm12m','Dados_FGV','I',128,328)],
 'conf':[S('Empresário¹','Dados_FGV','C',164,328),S('Consumidor','Dados_FGV','G',164,328)],
 'massa':[S('Trabalho','Dados_Tendências','G',129,328),S('Auxílios','Apoio_Radar','C',128,327,cs='Apoio_Radar'),S('Precatórios','Dados_Tendências','H',129,328)],
 'desemprego':[S('Taxa de Desocupação','Dados_Tendências','N',155,328),S('Renda Média do Trabalho','Dados_Tendências','L',155,328)],
 'juros_endiv':[S('Taxa de Juros² (% a.a.)','Dados_Tendências','AF',141,328),S('Endividamento das Famílias (% Renda)','Dados_BACEN','R',129,315,cc='P')],
 'renda_inad':[S('Renda comprometida com dívida¹ (% Renda)','Dados_Tendências','D',167,327),S('Inadimplência³ (% Saldo)','Dados_Tendências','AC',165,328)],
 'renda_disp':[S('Trabalho','Apoio_Radar','D',128,326),S('Auxílios','Apoio_Radar','E',128,326),S('Precatórios','Apoio_Radar','F',128,326)],
 'poupanca':[S('Taxa de Poupança das Famílias','Dados_Tendências','BD',49,114,cc='AZ')],
 'incc_comp':[S('Materiais','Apoio_Radar','N',164,327),S('Mão de Obra','Apoio_Radar','M',164,327)],
 'incc_ipca':[S('INCC','Apoio_Radar','H',164,328),S('IPCA','Dados_Tendências','O',165,329)],
 'financ':[S('Juros Longos (10y)','Dados_CIQ_Ipea_FED','C',44,208),S('Taxa de Financiamento Imobiliário','Dados_BACEN','T',153,316,cc='P')],
 'igmi':[S('Nominal','Dados_Extra','D',369,508),S('Real','Apoio_Radar','Q',188,327)],
 'vso':[S('Médio/Alto Padrão','Apoio_Radar','X',200,325),S('Minha Casa Minha Vida','Apoio_Radar','Z',200,325)],
 'lanc':[S('Minha Casa Minha Vida','Apoio_Radar','T',199,325),S('Médio/Alto Padrão','Apoio_Radar','S',199,325)],
 'varejo_var':[S('Var. a/a (mm3m)','Apoio_RCA','G',166,327,cc='F')],
 'varejo_proj':[S('Material de Construção','Apoio_Radar','AV',9,10,cc='AU'),S('Varejo Ampliado','Apoio_Radar','AW',9,10,cc='AU')],
 'varejo_idx':[S('Varejo Ampliado','Dados_Tendências','AL',165,327),S('Varejo - Material de Construção','Dados_Tendências','AM',165,327)],
 'varejo_idx_proj':[S('Varejo Ampliado','Apoio_Radar','AV',23,24,cc='AU'),S('Material de Construção','Apoio_Radar','AW',23,24,cc='AU')],
}
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
