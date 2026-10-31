import openpyxl, json, sys
src=sys.argv[1]; out=sys.argv[2]
wb=openpyxl.load_workbook(src,data_only=True)
num=lambda x: x if isinstance(x,(int,float)) else None
months=[]
for ws in wb.worksheets:
    hdr=[c.value for c in ws[1]]
    if not hdr or hdr[0]!='房间号': continue
    def ix(name,which=0):
        L=[i for i,h in enumerate(hdr) if h==name]
        return L[which] if L else None
    rents=[i for i,h in enumerate(hdr) if h=='房租']
    tot_i=ix('合计') if '合计' in hdr else rents[-1]
    rg_i=rents[1] if len(rents)>=3 else None
    pay_i=ix('收款')
    wp=num(ws['Q2'].value) if pay_i is not None else None
    ep=num(ws['Q3'].value) if pay_i is not None else None
    rooms={}
    for r in ws.iter_rows(min_row=2,values_only=True):
        if not isinstance(r[0],(int,float)) or r[0]<100: continue
        tot=num(r[tot_i])
        if tot is None: continue
        d={'rent':num(r[rents[0]]),'w0':num(r[ix('上月水')]),'w1':num(r[ix('此月水')]),'wf':num(r[ix('水费')]),
           'e0':num(r[ix('上月电')]),'e1':num(r[ix('此月电')]),'ef':num(r[ix('电费')]),'net':num(r[ix('网络费')]),
           'gar':num(r[ix('垃圾费')]),'tot':tot}
        if rg_i is not None: d['rg']=num(r[rg_i])
        if pay_i is not None: d['paid']=bool(r[pay_i])
        rooms[str(int(r[0]))]=d
    if rooms: months.append({'m':ws.title,'wp':wp or 4,'ep':ep or 1.2,'pay':pay_i is not None,'rooms':rooms})
months=months[-1:]  # 只放本月，往月数据不进页面
data=json.dumps(months,ensure_ascii=False,separators=(',',':'))
html=open('template.html',encoding='utf-8').read().replace('/*DATA*/[]',data)
open(out,'w',encoding='utf-8').write(html)
print(len(months),'months', [m['m'] for m in months])
