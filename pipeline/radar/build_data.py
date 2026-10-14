import json,datetime,collections
d=json.load(open('data_raw.json'))
def r(v,sig=5):
    if v is None: return None
    return float(f'{v:.{sig}g}')
def quarterize(pts):
    cnt=collections.Counter(); out=[]
    for c,v in pts:
        y=c.strip(); cnt[y]+=1; out.append([f'{y}-T{cnt[y]}',v])
    return out
Q={'wui','hiato','neutra','poupanca'}
def serial(s): return (datetime.date(1899,12,30)+datetime.timedelta(days=int(float(s)))).isoformat()
charts={}
for k,ser in d.items():
    if k.startswith('_'): continue
    allpts=[]
    for s in ser:
        pts=s['data']
        if k in Q: pts=quarterize(pts)
        if k in ('massa','renda_disp'):
            pts=[[c,(v/1000 if v else None)] for c,v in pts if c>='2013']
        allpts.append(pts)
    labels=[]
    seen=set()
    for pts in allpts:
        for c,_ in pts:
            if c not in seen: seen.add(c); labels.append(c)
    if k not in ('trade','pib','infl_glob','varejo_proj','varejo_idx_proj'): labels.sort()
    idx={c:i for i,c in enumerate(labels)}
    series=[]
    for s,pts in zip(ser,allpts):
        arr=[None]*len(labels)
        for c,v in pts: arr[idx[c]]=r(v)
        series.append(dict(label=s['label'],ref=s['ref'],data=arr))
    # drop trailing labels where all null
    while labels and all(x['data'][-1] is None for x in series):
        labels.pop(); [x['data'].pop() for x in series]
    charts[k]=dict(labels=labels,series=series)
b=d['_brent']
labels=[serial(c) for c in b[0]['cat']]
charts['brent']=dict(labels=labels,series=[dict(label=['Brent (US$/barril)','Média pré-conflito','Curva futura'][i],ref='Apresentação Modelo 2 (aba Daily)',data=[r(float(v)) if v else None for v in s['val']]) for i,s in enumerate(b)])
json.dump(charts,open('data.json','w'),ensure_ascii=False,separators=(',',':'))
import os;print(os.path.getsize('data.json'))
for k,v in charts.items(): print(k,len(v['labels']),v['labels'][:1],v['labels'][-1:],[x['data'][-1] for x in v['series']])
