import openpyxl, sys, json
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
src,out,PREVN,NEWN=sys.argv[1],sys.argv[2],sys.argv[3],sys.argv[4]
wb=openpyxl.load_workbook(src)
p=wb[PREVN]
rows=[]
for r in range(2,48):
    g=lambda c: p[f'{c}{r}'].value
    rows.append(dict(room=g('A'),rent=g('B'),w0=g('D'),w1=None,e0=g('G'),e1=None,net=g('I'),gar=g('J'),pay=None,note=None))
orig={d['room']:(d['rent'],d['net'],d['gar']) for d in rows}
OVR=json.loads(sys.argv[5]) if len(sys.argv)>5 else {}
for d in rows:
    for k,v in OVR.get(str(d['room']),{}).items(): d[k]=v
price_w,price_e=p['Q2'].value,p['Q3'].value
ws=wb.create_sheet(NEWN)
F=lambda **k: Font(name='Arial',**k)
yellow=PatternFill('solid',fgColor='FFF2CC'); grey=PatternFill('solid',fgColor='EEEEEE')
hdrfill=PatternFill('solid',fgColor='D9E1F2'); thin=Side(style='thin',color='BBBBBB')
bd=Border(left=thin,right=thin,top=thin,bottom=thin); ctr=Alignment(horizontal='center',vertical='center')
hdr=['房间号','房租','上月水','此月水','水费','上月电','此月电','电费','网络费','垃圾费','合计','收款','备注']
for i,h in enumerate(hdr,1):
    c=ws.cell(1,i,h); c.font=F(bold=True); c.fill=hdrfill; c.alignment=ctr; c.border=bd
last=1+len(rows)
for i,d in enumerate(rows):
    r=i+2
    vals={'A':d['room'],'B':d['rent'],'C':d['w0'],'D':d['w1'],'E':f'=IF(D{r}="","",(D{r}-C{r})*$Q$2)',
          'F':d['e0'],'G':d['e1'],'H':f'=IF(G{r}="","",ROUND((G{r}-F{r})*$Q$3,0))','I':d['net'],'J':d['gar'],
          'K':f'=IF(OR(D{r}="",G{r}=""),"",ROUND(B{r}+E{r}+H{r}+I{r}+J{r},0))','L':d['pay'],'M':d['note']}
    for col,v in vals.items():
        c=ws[f'{col}{r}']; c.value=v; c.font=F(); c.border=bd; c.alignment=ctr
        if col in 'DG': c.fill=yellow
    ws[f'K{r}'].font=F(bold=True)
    o=orig[d['room']]
    for hc,v in zip('TWX',o): ws[f'{hc}{r}']=v
t=last+1; ws[f'A{t}']='合计'
for col in 'BEHIJK': ws[f'{col}{t}']=f'=SUM({col}2:{col}{last})'
for col in 'ABCDEFGHIJKLM':
    c=ws[f'{col}{t}']; c.font=F(bold=True); c.fill=grey; c.border=bd; c.alignment=ctr
S=[('上月表',PREVN,'上月读数已从这个表抄好'),('水价(元/吨)',price_w,''),('电价(元/度)',price_e,''),(None,None,None),
   ('应收合计',f'=SUM(K2:K{last})',''),('已收金额',f'=SUMIF(L2:L{last},"?*",K2:K{last})',''),
   ('未付间数',f'=SUMPRODUCT((K2:K{last}<>"")*(L2:L{last}=""))',''),('未付金额','=Q5-Q6',''),
   (None,None,None),('本月没用电',f'=SUMPRODUCT((G2:G{last}<>"")*(G2:G{last}=F2:F{last}))','间（房号灰色，请确认是否空房）')]
for i,(a,b,c) in enumerate(S,1):
    if a is None: continue
    ws[f'P{i}']=a; ws[f'Q{i}']=b; ws[f'R{i}']=c
    ws[f'P{i}'].font=F(bold=True); ws[f'Q{i}'].font=F(bold=True,color='0000FF' if i<=3 else '000000')
    ws[f'R{i}'].font=F(color='888888',italic=True); ws[f'Q{i}'].border=bd; ws[f'P{i}'].border=bd
    if i<=3: ws[f'Q{i}'].fill=yellow
ws['Q1'].number_format='@'
for i in (7,8): ws[f'Q{i}'].font=F(bold=True,color='C00000')
dv=DataValidation(type='list',formula1='"微信,支付宝,现金"',allow_blank=True); ws.add_data_validation(dv); dv.add(f'L2:L{last}')
red=PatternFill('solid',fgColor='F8CBAD'); green=PatternFill('solid',fgColor='C6EFCE'); blue=PatternFill('solid',fgColor='BDD7EE')
vac=PatternFill('solid',fgColor='BFBFBF')
ws.conditional_formatting.add(f'A2:A{last}',FormulaRule(formula=['AND($G2<>"",$G2=$F2)'],fill=vac,font=Font(color='595959',italic=True),stopIfTrue=True))
ws.conditional_formatting.add(f'A2:A{last}',FormulaRule(formula=['AND($K2<>"",$L2="")'],fill=red))
ws.conditional_formatting.add(f'K2:K{last}',FormulaRule(formula=['AND($K2<>"",$L2="")'],fill=red))
ws.conditional_formatting.add(f'L2:L{last}',FormulaRule(formula=['$L2<>""'],fill=green))
ws.conditional_formatting.add(f'D2:D{last}',FormulaRule(formula=['AND($D2<>"",$D2<$C2)'],font=Font(color='FF0000',bold=True),fill=red))
ws.conditional_formatting.add(f'G2:G{last}',FormulaRule(formula=['AND($G2<>"",$G2<$F2)'],font=Font(color='FF0000',bold=True),fill=red))
for col,hc in zip('BIJ','TWX'):
    ws.conditional_formatting.add(f'{col}2:{col}{last}',FormulaRule(formula=[f'{col}2<>{hc}2'],fill=blue))
for hc in 'TWX': ws.column_dimensions[hc].hidden=True
for k,v in {'A':8,'B':7,'C':8,'D':8,'E':7,'F':8,'G':8,'H':8,'I':7,'J':7,'K':9,'L':8,'M':22,'N':3,'O':3,'P':12,'Q':10,'R':30}.items():
    ws.column_dimensions[k].width=v
ws.freeze_panes='B2'
wb.active=wb.sheetnames.index(NEWN)
wb.save(out)
