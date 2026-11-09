import json
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as CL
import scen
RC=scen.RC; SC=scen.SC; res=json.load(open('results.json'))
N=len(RC); NS=len(SC)
F=lambda **k: Font(name='Arial',size=10,**k)
BLUE=F(color='0000FF'); BLK=F(); BOLD=F(bold=True); GRN=F(color='008000'); WH=F(bold=True,color='FFFFFF')
HDR=PatternFill('solid',fgColor='1F3864'); SEC=PatternFill('solid',fgColor='D9E2F3'); YEL=PatternFill('solid',fgColor='FFFF00')
thin=Side(style='thin',color='BFBFBF')
CR='#,##0;(#,##0);-'; VOL='#,##0.##;(#,##0.##);-'
wb=Workbook()

def put(ws,ref,val,font=BLK,fmt=None,fill=None,al=None):
    c=ws[ref]; c.value=val; c.font=font
    if fmt: c.number_format=fmt
    if fill: c.fill=fill
    if al: c.alignment=al
    return c
def header(ws,row,vals,start=1):
    for i,v in enumerate(vals):
        put(ws,f'{CL(start+i)}{row}',v,WH,fill=HDR,al=Alignment(wrap_text=True,vertical='center'))

# ---------- Reference ----------
ref=wb.active; ref.title='Reference'
put(ref,'A1','Reference tables (Workday Flex Credits and Platform Entitlement Policy, 5 Nov 2026)',BOLD)
header(ref,3,['Min employees','Band','Complimentary credits / yr','Baseline API calls / yr'])
bands=[(0,'Under 3.5k',15000,2500000),(3500,'3.5k to 9,999',30000,3500000),(10000,'10k to 29,999',60000,4500000),(30000,'30k to 99,999',120000,6000000),(100000,'100k+',200000,6500000)]
for i,b in enumerate(bands):
    for j,v in enumerate(b): put(ref,f'{CL(1+j)}{4+i}',v,BLUE,'#,##0' if j!=1 else None)
put(ref,'A9','Source: Policy p1 (complimentary credits) and p2 (baseline API entitlement). HCM or FIN customers; Planning-only is set separately below.',F(italic=True))
header(ref,10,['SKU (raises API entitlement)','Uplift %'])
skus=[('Student',150),('Accounting Center',100),('Procurement',100),('Extend',50),('Prism or Data Cloud',15),('Financials',15)]
for i,(n,p) in enumerate(skus): put(ref,f'A{11+i}',n); put(ref,f'B{11+i}',p,BLUE)
put(ref,'A17','Source: Policy p3. Uplifts are summed and each applies to the baseline entitlement.',F(italic=True))
header(ref,19,['Ramp preset']+[f'M{i}' for i in range(1,13)]+['Sum (months at full use)'])
ramps={'Gradual':[10,15,25,40,55,70,85,100,100,100,100,100],'Stepped':[5,5,5,20,20,20,50,50,50,100,100,100],'Slow':[5,10,15,20,25,30,40,50,60,70,80,90],'Immediate':[100]*12}
for i,(n,v) in enumerate(ramps.items()):
    r=20+i; put(ref,f'A{r}',n)
    for j,x in enumerate(v): put(ref,f'{CL(2+j)}{r}',x,BLUE)
    put(ref,f'N{r}',f'=SUM(B{r}:M{r})/100',BLK,'0.00')
put(ref,'A24','Ramp percentages are modelling assumptions (not Workday data).',F(italic=True))
header(ref,26,['Constant','Value','Source'])
consts=[('API credits per block',60,'Rate Card p4: 60 credits per 10k API calls'),('API calls per block',10000,'Rate Card p4'),('Complimentary credits, Planning without HCM or FIN',10000,'Policy p1'),('Months in a contract year',12,'')]
for i,(n,v,s) in enumerate(consts): put(ref,f'A{27+i}',n); put(ref,f'B{27+i}',v,BLUE,'#,##0'); put(ref,f'C{27+i}',s)
ref.column_dimensions['A'].width=44
for c in 'BCDEFGHIJKLMN': ref.column_dimensions[c].width=14
ref.column_dimensions['C'].width=26; ref.column_dimensions['D'].width=24

# ---------- RateCard ----------
rcs=wb.create_sheet('RateCard')
put(rcs,'A1','Rate card as transcribed independently from the PDF (Workday Flex Credits Rate Card, last updated 5 Nov 2026)',BOLD)
put(rcs,'A2','Mode: mo = volume per month x ramp; yr = volume per year x ramp/12; yr_noramp = volume per year, no ramp; run = records per run, billed every run.',F(italic=True))
header(rcs,4,['Id','Offer','Skill','Usage meter','Block size (units)','Credits per block','Free units / yr','Mode','Source'])
for k,r in enumerate(RC):
    row=5+k
    for j,v in enumerate(r):
        put(rcs,f'{CL(1+j)}{row}',v,BLUE if j in (4,5,6,7) else BLK,'#,##0' if j in (4,5,6) else None)
for c,w in zip('ABCDEFGHI',[11,34,40,52,14,14,14,12,18]): rcs.column_dimensions[c].width=w
rcs.freeze_panes='A5'

# ---------- Model ----------
m=wb.create_sheet('Model')
put(m,'A1','Independent Excel model of Flex Credit consumption (formulas only; inputs in blue)',BOLD)
put(m,'A2','Each column is one test scenario. Change blue cells to explore. Page results for the same inputs are on "Page engine"; differences are on "Diff".',F(italic=True))
header(m,3,['Parameter','Notes']+[s['name'] for s in SC]); m.row_dimensions[3].height=42
P={'emp':4,'sub':5,'sku0':6,'upl':12,'ramp':13,'rs':14,'mult':15,'runs':16,'api':17,'grace':18,'compOn':19,'compM':20,'buf':21,'price':22,'cur':23,'fx':24}
labels=[(4,'Employees',''),(5,'Subscription','HCM or FIN / Planning'),(12,'API entitlement uplift %','Sum of flagged SKU uplifts'),(13,'Ramp preset','Name from Reference'),(14,'Ramp sum (months at full use)','From Reference'),(15,'Scenario multiplier',''),(16,'Payroll runs per year','Used by mode = run'),(17,'Metered API calls per year',''),(18,'Months inside API grace period','0 to 12'),(19,'Count complimentary credits (1/0)',''),(20,'Months of complimentary grant','12 = full year'),(21,'Buffer on purchase %',''),(22,'Price per credit (USD)',''),(23,'Display currency',''),(24,'AUD per USD','Used only if AUD')]
for r,l,n in labels: put(m,f'A{r}',l); put(m,f'B{r}',n,F(italic=True))
for i,(n,p) in enumerate(skus): put(m,f'A{6+i}',f'Owns {n} (1/0)')
put(m,'A26','VOLUMES in the unit of each RateCard row (see Mode)',BOLD,fill=SEC)
V0=27; C0=V0+N+2
put(m,f'A{C0-1}','CREDITS PER LINE (formulas)',BOLD,fill=SEC)
T0=C0+N+1
put(m,f'A{T0-1}','TOTALS',BOLD,fill=SEC)
for k,r in enumerate(RC):
    put(m,f'A{V0+k}',r[0]); put(m,f'B{V0+k}',f'=RateCard!C{5+k}',GRN)
    put(m,f'A{C0+k}',r[0]); put(m,f'B{C0+k}',f'=RateCard!C{5+k}',GRN)
tot=['Credits from agents and platform offers','API entitlement (calls)','Metered API calls','Calls over entitlement','API overage credits (after grace)','TOTAL expected credits','Complimentary credits for band','Complimentary credits counted','Net credits to cover','Recommended purchase (credits)','Cost of recommended purchase (display currency)']
for i,t in enumerate(tot): put(m,f'A{T0+i}',t,BOLD if i in (5,9,10) else BLK)
for si,s in enumerate(SC):
    c=CL(3+si)
    put(m,f'{c}{P["emp"]}',s['emp'],BLUE,'#,##0')
    put(m,f'{c}{P["sub"]}','Planning' if s['base']=='planning' else 'HCM or FIN',BLUE)
    for i,(n,p) in enumerate(skus):
        key=['student','acct','proc','extend','prism','fin'][i]
        put(m,f'{c}{6+i}',1 if key in s['skus'] else 0,BLUE)
    put(m,f'{c}12',f'=SUMPRODUCT({c}6:{c}11,Reference!$B$11:$B$16)',BLK,'0')
    put(m,f'{c}13',scen.NAMES[s['preset']],BLUE)
    put(m,f'{c}14',f'=INDEX(Reference!$N$20:$N$23,MATCH({c}13,Reference!$A$20:$A$23,0))',BLK,'0.00')
    put(m,f'{c}15',s['m'],BLUE,'0.0#'); put(m,f'{c}16',s['runs'],BLUE); put(m,f'{c}17',s['api'],BLUE,'#,##0'); put(m,f'{c}18',s['grace'],BLUE)
    put(m,f'{c}19',s['compOn'],BLUE); put(m,f'{c}20',s['compM'],BLUE); put(m,f'{c}21',s['buf'],BLUE); put(m,f'{c}22',s['price'],BLUE,'0.000'); put(m,f'{c}23',s['cur'],BLUE); put(m,f'{c}24',s['fx'],BLUE,'0.00')
    for k,r in enumerate(RC):
        v=s['vols'].get(r[0])
        if v is not None: put(m,f'{c}{V0+k}',v,BLUE,VOL)
        vr=f'{c}{V0+k}'; rr=5+k
        f=(f'=IF({vr}="",0,IF(RateCard!$H${rr}="run",CEILING(ROUND({vr}*{c}$15/RateCard!$E${rr},6),1)*RateCard!$F${rr}*{c}$16*{c}$14/Reference!$B$30,'
           f'CEILING(ROUND(MAX(0,IF(RateCard!$H${rr}="mo",{vr}*{c}$14*{c}$15,IF(RateCard!$H${rr}="yr",{vr}*{c}$14/Reference!$B$30*{c}$15,{vr}*{c}$15))-RateCard!$G${rr})/RateCard!$E${rr},6),1)*RateCard!$F${rr}))')
        put(m,f'{c}{C0+k}',f,BLK,CR)
    t=lambda i:f'{c}{T0+i}'
    put(m,t(0),f'=SUM({c}{C0}:{c}{C0+N-1})',BLK,CR)
    put(m,t(1),f'=LOOKUP({c}4,Reference!$A$4:$A$8,Reference!$D$4:$D$8)*(1+{c}12/100)',BLK,'#,##0')
    put(m,t(2),f'={c}17',BLK,'#,##0')
    put(m,t(3),f'=MAX(0,{t(2)}-{t(1)})',BLK,'#,##0')
    put(m,t(4),f'=CEILING(ROUND({t(3)}/Reference!$B$28,6),1)*Reference!$B$27*(Reference!$B$30-{c}18)/Reference!$B$30',BLK,CR)
    put(m,t(5),f'={t(0)}+{t(4)}',BOLD,CR)
    put(m,t(6),f'=IF({c}5="Planning",Reference!$B$29,LOOKUP({c}4,Reference!$A$4:$A$8,Reference!$C$4:$C$8))',BLK,CR)
    put(m,t(7),f'=IF({c}19=1,{t(6)}*MIN(Reference!$B$30,{c}20)/Reference!$B$30,0)',BLK,CR)
    put(m,t(8),f'=MAX(0,{t(5)}-{t(7)})',BLK,CR)
    put(m,t(9),f'=IF({t(8)}>0,CEILING({t(8)}*(1+{c}21/100)/1000,1)*1000,0)',BOLD,CR)
    put(m,t(10),f'={t(9)}*{c}22*IF({c}23="AUD",{c}24,1)',BOLD,'#,##0;(#,##0);-')
m.column_dimensions['A'].width=46; m.column_dimensions['B'].width=34
for si in range(NS): m.column_dimensions[CL(3+si)].width=15
m.freeze_panes='C4'

# ---------- Page engine (values produced by the page's own code) ----------
pe=wb.create_sheet('Page engine')
put(pe,'A1','Results produced by the calculation code inside the published page for the same inputs (pasted values, generated by running that code)',BOLD)
header(pe,3,['Line / total','']+[s['name'] for s in SC]); pe.row_dimensions[3].height=42
ids=[r[0] for r in RC]
for k,i in enumerate(ids): put(pe,f'A{C0+k}',i)
for i,t in enumerate(tot): put(pe,f'A{T0+i}',t)
keys=['linesTotal','ent','sum','over','api','total','compAnnual','comp','net','rec','cost']
for si,R in enumerate(res['res']):
    c=CL(3+si)
    for k,i in enumerate(ids): put(pe,f'{c}{C0+k}',R['lines'][i],BLK,CR)
    for i,key in enumerate(keys): put(pe,f'{c}{T0+i}',R[key],BLK,'#,##0.##;(#,##0.##);-')
pe.column_dimensions['A'].width=46
for si in range(NS): pe.column_dimensions[CL(3+si)].width=15
pe.freeze_panes='C4'

# ---------- Diff ----------
d=wb.create_sheet('Diff')
put(d,'A1','Excel model minus page engine. Every cell should be zero.',BOLD)
header(d,3,['Line / total','']+[s['name'] for s in SC]); d.row_dimensions[3].height=42
rows=list(range(C0,C0+N))+list(range(T0,T0+len(tot)))
for r in rows:
    put(d,f'A{r}',f'=Model!A{r}')
    for si in range(NS):
        c=CL(3+si); put(d,f'{c}{r}',f"=Model!{c}{r}-'Page engine'!{c}{r}",BLK,'#,##0.00;(#,##0.00);-')
d.column_dimensions['A'].width=46
for si in range(NS): d.column_dimensions[CL(3+si)].width=15
d.freeze_panes='C4'

# ---------- Data check ----------
dc=wb.create_sheet('Data check')
put(dc,'A1','Rate transcription check: this workbook (typed from the PDF) against the rate data inside the page',BOLD)
header(dc,3,['Id','Block (RateCard)','Credits (RateCard)','Free (RateCard)','Mode (RateCard)','Block (page)','Credits (page)','Free (page)','Mode (page)','Result'])
pl={x['id']:x for x in res['lines']}
for k,r in enumerate(RC):
    row=4+k; p=pl[r[0]]
    pmode='run' if p['kind']=='run' else ('yr_noramp' if p['noramp'] else ('yr' if r[0]=='rc' else p['per']))
    put(dc,f'A{row}',r[0]); put(dc,f'B{row}',f'=RateCard!E{5+k}',GRN); put(dc,f'C{row}',f'=RateCard!F{5+k}',GRN); put(dc,f'D{row}',f'=RateCard!G{5+k}',GRN); put(dc,f'E{row}',f'=RateCard!H{5+k}',GRN)
    put(dc,f'F{row}',p['b'],BLUE); put(dc,f'G{row}',p['r'],BLUE); put(dc,f'H{row}',p['free'],BLUE); put(dc,f'I{row}',pmode,BLUE)
    put(dc,f'J{row}',f'=IF(AND(B{row}=F{row},C{row}=G{row},D{row}=H{row},E{row}=I{row}),"Match","DIFFERENT")')
put(dc,f'A{4+N+1}','Lines in page but not in this workbook:',BOLD)
missing=[i for i in pl if i not in ids]
put(dc,f'B{4+N+1}',', '.join(missing) if missing else 'none')
for c,w in zip('ABCDEFGHIJ',[12,14,16,14,14,14,14,12,14,12]): dc.column_dimensions[c].width=w
dc.freeze_panes='A4'

# ---------- Summary ----------
sm=wb.create_sheet('Summary')
put(sm,'A1','Validation summary',BOLD)
put(sm,'A2','Overall result',BOLD)
put(sm,'B2','=IF(AND(COUNTIF(J6:J'+str(5+NS)+',"FAIL")=0,COUNTIF(\'Data check\'!J4:J'+str(3+N)+',"DIFFERENT")=0),"ALL CHECKS PASS","CHECK FAILED - see rows marked FAIL / DIFFERENT")',BOLD)
put(sm,'A3','Rate lines matching page data',BOLD)
put(sm,'B3',f'=COUNTIF(\'Data check\'!J4:J{3+N},"Match")&" of {N}"',BLK)
header(sm,5,['Scenario','Employees','Excel total credits','Page total credits','Difference','Excel purchase','Page purchase','Difference','Sum of absolute differences, all lines and totals','Result'])
sm.row_dimensions[5].height=54
for si,s in enumerate(SC):
    r=6+si; c=CL(3+si)
    put(sm,f'A{r}',s['name']); put(sm,f'B{r}',f'=Model!{c}4',GRN,'#,##0')
    put(sm,f'C{r}',f'=Model!{c}{T0+5}',GRN,CR); put(sm,f'D{r}',f"='Page engine'!{c}{T0+5}",GRN,CR); put(sm,f'E{r}',f'=C{r}-D{r}',BLK,'#,##0.00;(#,##0.00);-')
    put(sm,f'F{r}',f'=Model!{c}{T0+9}',GRN,CR); put(sm,f'G{r}',f"='Page engine'!{c}{T0+9}",GRN,CR); put(sm,f'H{r}',f'=F{r}-G{r}',BLK,'#,##0.00;(#,##0.00);-')
    put(sm,f'I{r}',f'=SUMPRODUCT(ABS(Diff!{c}{C0}:{c}{T0+len(tot)-1}))',BLK,'#,##0.00')
    put(sm,f'J{r}',f'=IF(I{r}<0.01,"PASS","FAIL")',BOLD)
for c,w in zip('ABCDEFGHIJ',[34,12,16,16,14,16,16,14,24,10]): sm.column_dimensions[c].width=w
sm.freeze_panes='A6'

# ---------- Read me ----------
rd=wb.create_sheet('Read me')
lines=[
('Flex Credit Modeller: independent validation workbook',BOLD),
('',BLK),
('Purpose',BOLD),
('An independent formula model of Flex Credit consumption, built only from the Workday documents dated 5 Nov 2026. It re-computes 15 test scenarios and compares them with the calculation code in the Flex Credit Modeller page.',BLK),
('',BLK),
('Sheets',BOLD),
('Summary: pass / fail for every scenario and for the rate transcription check.',BLK),
('Model: the independent calculation. Blue cells are inputs; edit them to try your own scenario.',BLK),
('RateCard: rates typed from the PDF, each with its page reference. Reference: bands, SKU uplifts and ramp assumptions.',BLK),
('Page engine: values produced by the page for the same inputs (pasted, not live). Diff: Model minus Page engine.',BLK),
('Data check: compares the typed rates with the rates stored in the page.',BLK),
('',BLK),
('What this proves',BOLD),
('That the page applies the rate card, block rounding, free allowances, payroll per-run billing, API overage, grace months and complimentary credits the same way as a separate spreadsheet implementation. It does not prove the interpretation of Workday terms is right.',BLK),
('',BLK),
('Still to do before wider release',BOLD),
('1. Have a second person check the RateCard sheet against the PDF line by line.',BLK),
('2. Run the same inputs through Workday\'s Flex Credit Estimator (release 2026R2) for several real tenants and record the variance.',BLK),
('3. Get Workday to confirm the open interpretation points: behaviour when credits run out, Sana Learn, Sana base entitlement, complimentary pro-rating, grace period treatment.',BLK),
('4. Back-test against real Platform Consumption Console usage.',BLK),
('',BLK),
('Modelling assumptions (not Workday data)',BOLD),
('Ramp presets, scenario multipliers and all default volumes are planning assumptions. Credits are billed in whole blocks on annual units, which can differ from Workday billing if usage is bursty. API grace is applied pro rata by months.',BLK),
('',BLK),
('Legend: blue text = input or typed constant; black = formula; green = link to another sheet.',BLK),
]
for i,(t,f) in enumerate(lines):
    put(rd,f'A{i+1}',t,f,al=Alignment(wrap_text=True,vertical='top'))
rd.column_dimensions['A'].width=120
wb._sheets=[wb['Read me'],wb['Summary'],wb['Model'],wb['RateCard'],wb['Reference'],wb['Page engine'],wb['Diff'],wb['Data check']]
wb.save('Flex_Credit_Validation_Model.xlsx')
print('saved',N,'lines',NS,'scenarios', 'volume rows',V0,'credit rows',C0,'totals',T0)
