#!/usr/bin/env python3
"""
Generador del Dashboard KUKUXUMUSU v10
Uso: python3 generar_dashboard.py
Requiere tener los ficheros en:
  ../Downloads/Ventas del 15 al 26 julio.xlsx
  ../Downloads/Reposicion del 15 al 26 julio.xlsx
"""

import pandas as pd
import json, os, sys
from datetime import datetime

# Paths - detect if running from repo root or Analisis Tiendas
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if os.path.basename(SCRIPT_DIR) == 'scripts':
    BASE = os.path.dirname(SCRIPT_DIR)  # scripts/ -> repo root
else:
    BASE = SCRIPT_DIR  # Use current dir

DOWNLOADS = os.path.join(os.path.expanduser('~'), 'Downloads')
OUTPUT = os.path.join(BASE, 'dashboard_ejecutivo.html')

print(f"Directorio base: {BASE}")
print(f"Downloads: {DOWNLOADS}")
print(f"Output: {OUTPUT}")

# ==== LOAD DATA ====
ventas_path = os.path.join(DOWNLOADS, 'Ventas del 15 al 26 julio.xlsx')
repo_path = os.path.join(DOWNLOADS, 'Reposicion del 15 al 26 julio.xlsx')

if not os.path.exists(ventas_path):
    print(f"ERROR: No se encuentra {ventas_path}")
    sys.exit(1)
if not os.path.exists(repo_path):
    print(f"ERROR: No se encuentra {repo_path}")
    sys.exit(1)

ventas = pd.read_excel(ventas_path)
repo = pd.read_excel(repo_path)
sales = ventas[ventas['Operaci\u00f3n'] == 'Venta'].copy()
ret = ventas[ventas['Operaci\u00f3n'] == 'Devoluci\u00f3n'].copy()
sales['Tienda'] = sales['Tienda'].replace({'Estafeta':'Pamplona'})
ret['Tienda'] = ret['Tienda'].replace({'Estafeta':'Pamplona'})

cc_v = next(c for c in ventas.columns if c.startswith('C') and 'digo' in c)
cc_r = next(c for c in repo.columns if c.startswith('C') and 'digo' in c)
fc = 'Precio final con impuestos'

df_kpi = pd.read_excel(os.path.join(BASE, 'DASHBOARD_GLOBAL_TIENDAS.xlsx'), sheet_name='1-RESUMEN KPIs')
kfc = [c for c in df_kpi.columns if 'Factura' in str(c)][0]
df_top = pd.read_excel(os.path.join(BASE, 'DASHBOARD_GLOBAL_TIENDAS.xlsx'), sheet_name='2-TOP 100 GLOBAL TOTAL')
tec = [c for c in df_top.columns if 'Total' in c and 'Unid' not in c][0]

scn = {c.lower().replace('\u00f3','o').replace('\u00ed','i').replace('\u00e9','e').replace('\u00e1','a'):c for c in sales.columns}
cat_col = scn.get('categoria','Categoria')

# [REST OF THE GENERATOR CODE - same as v10]
# I'll continue with the data processing...

# ==== GLOBAL ====
tf = round(float(df_kpi[kfc].sum()),2); tu = int(df_kpi['Ventas Totales (U)'].sum())
ju = int(df_kpi['Ventas Julio (U)'].sum()); st = int(repo['Total existencias disponibles'].sum())
pu = int(sales['Unidades'].sum()); pi = round(float(sales[fc].sum()),2)
pt = sales['Pedido'].nunique(); ru = int(abs(ret['Unidades'].sum()))
ri = round(float(abs(ret[fc].sum())),2); td = round(ri/(pi+ri)*100,2) if pi>0 else 0
pp = int(sales[cc_v].nunique())

# ==== TIENDAS ACUMULADO ====
tlist = []
for _, r in df_kpi.iterrows():
    t = str(r['Tienda']); fu = round(float(r[kfc]),2); tu_st = int(r['Ventas Totales (U)'])
    prods = int(r['Productos']); ju_st = int(r['Ventas Julio (U)'])
    growth = round((ju_st-int(r['Ventas Junio (U)']))/max(1,int(r['Ventas Junio (U)']))*100,1)
    store_col = f'Vendido {t}'
    if t == 'Pamplona': store_col = 'Vendido Estafeta'
    top_n = '-'; top_u_v = 0
    if store_col in repo.columns:
        tmp = repo[[cc_r, 'Nombre del Producto', store_col]].copy()
        tmp.columns = ['c','n','v']; tmp = tmp.sort_values('v', ascending=False).head(1)
        if len(tmp) > 0:
            top_n = str(tmp['n'].values[0]); top_u_v = int(tmp['v'].values[0])
    tlist.append({'t':t,'fu':fu,'tu':tu_st,'prods':prods,'ju':ju_st,'growth':growth,'top_n':top_n,'top_u':top_u_v})

# ==== TIENDAS PERIODO ====
sr = sales.groupby('Tienda').agg(U=('Unidades','sum'),I=(fc,'sum'),T=('Pedido','nunique')).sort_values('I',ascending=False).reset_index()
sr['%I'] = (sr['I']/sr['I'].sum()*100).round(1)
rets = ret.groupby('Tienda').agg(RU=('Unidades','sum'),RI=(fc,'sum')).reset_index()
sr = sr.merge(rets,on='Tienda',how='left').fillna(0)
sr['RU'] = sr['RU'].astype(int); sr['TM'] = (sr['I']/sr['T']).round(2)
top_pn = []; top_pu_v = []; top_pi_v = []
for _, row in sr.iterrows():
    sdd = sales[sales['Tienda']==row['Tienda']]
    sd_t = sdd.groupby([cc_v,'Nombre']).agg(U=('Unidades','sum'),I=(fc,'sum')).sort_values('U',ascending=False).head(1)
    if len(sd_t) > 0:
        top_pn.append(str(sd_t.index.get_level_values('Nombre')[0]))
        top_pu_v.append(int(sd_t.values[0][0])); top_pi_v.append(round(float(sd_t.values[0][1]),2))
    else: top_pn.append('-'); top_pu_v.append(0); top_pi_v.append(0)
sr['top_pn']=top_pn; sr['top_pu']=top_pu_v; sr['top_pi']=top_pi_v

# ==== STORE FILES MAP ====
store_file_map = {
    'Bilbao': 'analisis_rotacion_BILBAO_COMPLETO.xlsx',
    'Donostia': 'analisis_rotacion_DONOSTIA_COMPLETO.xlsx',
    'Lleida': 'analisis_rotacion_LLEIDA_ACTUALIZADO_ENERO.xlsx',
    'Pamplona': 'analisis_rotacion_PAMPLONA_COMPLETO.xlsx',
    'Segovia': 'analisis_rotacion_SEGOVIA_COMPLETO.xlsx',
    'Toros': 'analisis_rotacion_TOROS_COMPLETO.xlsx',
    'Zaragoza': 'analisis_rotacion_ZARAGOZA_COMPLETO.xlsx',
}

# ==== STORE DETAIL ====
store_detail = {}
code_caches = {}  # Cache column names per store file
for t in sales['Tienda'].unique():
    sd = sales[sales['Tienda']==t]; sret = ret[ret['Tienda']==t]
    sd_u = sd.groupby([cc_v,'Nombre']).agg(U=('Unidades','sum'),I=(fc,'sum')).sort_values('U',ascending=False).head(10).reset_index()
    sd_i = sd.groupby([cc_v,'Nombre']).agg(U=('Unidades','sum'),I=(fc,'sum')).sort_values('I',ascending=False).head(10).reset_index()
    sd_cats = sd.groupby(cat_col).agg(U=('Unidades','sum'),I=(fc,'sum')).sort_values('I',ascending=False).reset_index()
    
    hist_data = []
    fname = store_file_map.get(t)
    if fname:
        fpath = os.path.join(BASE, fname)
        if os.path.exists(fpath):
            sf = pd.read_excel(fpath, sheet_name='1-RANKING COMPLETO')
            code_col_sf = next((c for c in sf.columns if c.startswith('C') and ('digo' in c or 'od' in c)), None)
            total_col_sf = next((c for c in sf.columns if 'Total' in c and 'Unid' in c), None)
            name_col_sf = next((c for c in sf.columns if c.startswith('N')), None)
            if code_col_sf and total_col_sf and name_col_sf:
                top10 = sf.nlargest(10, total_col_sf)[[code_col_sf, name_col_sf, total_col_sf]]
                for _, r2 in top10.iterrows():
                    hist_data.append({'c':str(r2[code_col_sf]),'n':str(r2[name_col_sf]),'u':int(r2[total_col_sf])})
    
    t_acum = next(({'fu':round(float(r[kfc]),2),'tu':int(r['Ventas Totales (U)'])} for _, r in df_kpi.iterrows() if str(r['Tienda'])==t), {'fu':0,'tu':0})
    
    store_detail[t] = {
        'u':int(sd['Unidades'].sum()),'i':round(float(sd[fc].sum()),2),
        't':int(sd['Pedido'].nunique()),'tm':round(float(sd[fc].mean()),2),
        'ru':int(abs(sret['Unidades'].sum())) if len(sret)>0 else 0,
        'fu':t_acum['fu'],'tu_acum':t_acum['tu'],
        'top_u':json.loads(sd_u.to_json(orient='records')),
        'top_i':json.loads(sd_i.to_json(orient='records')),
        'hist_u':hist_data,
        'cats':json.loads(sd_cats.to_json(orient='records'))
    }

# ==== TOP 100 ANUAL ====
au = []
for _, r in df_top.iterrows():
    au.append({'n':str(r['Nombre']),'tu':int(r['Total_Unid']),'te':round(float(r[tec]),2),'ju':int(r['Jul_Unid'])})
ai = sorted(au, key=lambda x:x['te'], reverse=True)

top_u = sales.groupby([cc_v,'Nombre']).agg(U=('Unidades','sum'),I=(fc,'sum'),T=('Pedido','nunique')).sort_values('U',ascending=False).reset_index().head(50)
top_u['PM'] = (top_u['I']/top_u['U']).round(2)
top_i = sales.groupby([cc_v,'Nombre']).agg(U=('Unidades','sum'),I=(fc,'sum'),T=('Pedido','nunique')).sort_values('I',ascending=False).reset_index().head(50)
top_i['PM'] = (top_i['I']/top_i['U']).round(2)

cats = sales.groupby(cat_col).agg(U=('Unidades','sum'),I=(fc,'sum')).sort_values('I',ascending=False).reset_index()
cats['%I'] = (cats['I']/cats['I'].sum()*100).round(1); cn = list(cats.columns)[0]

sales['Fecha'] = pd.to_datetime(sales['Fecha'])
daily = sales.groupby('Fecha').agg(U=('Unidades','sum'),I=(fc,'sum'),T=('Pedido','nunique')).reset_index().sort_values('Fecha')
mc = ['Ventas Enero (U)','Ventas Febrero (U)','Ventas Marzo (U)','Ventas Abril (U)','Ventas Mayo (U)','Ventas Junio (U)','Ventas Julio (U)']
trend = []
for _, r in df_kpi.iterrows():
    trend.append({'t':str(r['Tienda']),'d':[int(r[c]) for c in mc]})
mn = ['Ene','Feb','Mar','Abr','May','Jun','Jul']

D = {
    'act': datetime.now().strftime('%d/%m/%Y %H:%M'),
    'tf':tf,'tu':tu,'ju':ju,'st':st,'pu':pu,'pi':pi,'pt':pt,'pp':pp,
    'ru':ru,'ri':ri,'td':td,
    'tl':tlist,'sr':json.loads(sr.to_json(orient='records')),
    'sd':store_detail,
    'au':au,'ai':ai,
    'top_u':json.loads(top_u.to_json(orient='records')),
    'top_i':json.loads(top_i.to_json(orient='records')),
    'cats':json.loads(cats.to_json(orient='records')),'cn':cn,
    'day':json.loads(daily.to_json(orient='records')),
    'trend':trend,'mn':mn,
}
JSON = json.dumps(D, ensure_ascii=False)
print(f"JSON OK: {len(JSON):,} chars")

# ==== HTML ==== (same template as v10)
T = r"""<!DOCTYPE html>
<html lang="es"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0"><title>KUKUXUMUSU - Dashboard</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js" defer></script>
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:'Segoe UI',Arial,sans-serif;background:#f0f2f5;color:#1a1a2e;padding:14px}
.container{max-width:1440px;margin:0 auto}
.header{background:linear-gradient(135deg,#0f0c29,#302b63,#24243e);color:#fff;padding:16px 20px;border-radius:10px;margin-bottom:12px;display:flex;justify-content:space-between;flex-wrap:wrap;gap:8px;align-items:center}
.header h1{font-size:18px;font-weight:800;letter-spacing:-.5px}
.header h1 span{color:#e94560}.header .sub{font-size:10px;opacity:.65;margin-top:1px}
.hdr-actions button{padding:3px 8px;border:1px solid rgba(255,255,255,.25);border-radius:4px;background:rgba(255,255,255,.08);color:#fff;cursor:pointer;font-size:9px}
.kpi-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(110px,1fr));gap:5px;margin-bottom:12px}
.kpi{background:#fff;border-radius:6px;padding:7px 9px;box-shadow:0 1px 2px rgba(0,0,0,.04)}
.kpi .l{font-size:7px;font-weight:600;text-transform:uppercase;letter-spacing:.3px;color:#6b7280}
.kpi .v{font-size:15px;font-weight:800;margin:1px 0}.kpi .s{font-size:8px;color:#9ca3af}
h2.sec{font-size:12px;font-weight:700;margin:12px 0 6px;color:#1f2937;display:flex;align-items:center;gap:6px}
h2.sec:before{content:'';width:3px;height:13px;background:#e94560;border-radius:2px}
h2.sec .tag{font-size:8px;font-weight:400;color:#9ca3af;background:#f3f4f6;padding:1px 6px;border-radius:4px;margin-left:auto}
.sg{display:grid;grid-template-columns:repeat(auto-fit,minmax(240px,1fr));gap:5px;margin-bottom:10px}
.sc{background:#fff;border-radius:7px;padding:7px 9px;box-shadow:0 1px 2px rgba(0,0,0,.04);cursor:pointer;transition:all .12s}
.sc:hover{transform:translateY(-1px);box-shadow:0 3px 8px rgba(0,0,0,.08)}
.sc h4{font-size:10px;font-weight:700;display:flex;justify-content:space-between;margin-bottom:3px}
.sc .dg{display:grid;grid-template-columns:1fr 1fr;gap:2px;margin-top:3px}
.sc .it .lb{font-size:7px;color:#6b7280;font-weight:600;text-transform:uppercase}
.sc .it .vl{font-size:11px;font-weight:700}
.sc .tp{font-size:8px;background:#fef3c7;border-radius:3px;padding:2px 5px;margin-top:3px;display:flex;justify-content:space-between;align-items:center}
.sc .tp .tl{overflow:hidden;text-overflow:ellipsis;white-space:nowrap;max-width:130px;font-weight:600}
.sc .ft{font-size:8px;color:#9ca3af;margin-top:3px}
.tw{background:#fff;border-radius:7px;padding:10px 12px;box-shadow:0 1px 2px rgba(0,0,0,.04);margin-bottom:10px;overflow-x:auto}
.tw h3{font-size:11px;font-weight:600;margin-bottom:4px;color:#374151}
.tc{display:flex;gap:3px;margin:3px 0;flex-wrap:wrap;align-items:center}
.tc button{padding:2px 6px;border:1px solid #d1d5db;border-radius:3px;background:#fff;cursor:pointer;font-size:8px;font-weight:500;color:#4b5563}
.tc button.act{background:#1a1a2e;color:#fff;border-color:#1a1a2e}
.tc input{padding:2px 5px;border:1px solid #d1d5db;border-radius:3px;font-size:8px;max-width:100px}
.cg{display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-bottom:10px}
.cf{grid-column:1/-1}.cb{background:#fff;border-radius:7px;padding:10px;box-shadow:0 1px 2px rgba(0,0,0,.04)}
.cb h3{font-size:10px;font-weight:600;margin-bottom:5px;color:#374151}
.cb canvas{max-height:220px;width:100%!important}
table{width:100%;border-collapse:collapse;font-size:10px}
th{padding:3px 4px;background:#f8fafc;color:#6b7280;font-weight:600;font-size:8px;text-transform:uppercase;letter-spacing:.2px;border-bottom:2px solid #e5e7eb;white-space:nowrap;text-align:left}
td{padding:2px 4px;border-bottom:1px solid #f3f4f6;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:160px}
tr:hover td{background:#f9fafb}
.num{text-align:right}.rk{color:#9ca3af;width:18px;text-align:center}
.modal{display:none;position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,.5);z-index:1000;justify-content:center;align-items:center}
.modal.show{display:flex}
.modal-inner{background:#fff;border-radius:12px;padding:16px;max-width:640px;width:90%;max-height:85vh;overflow-y:auto;box-shadow:0 12px 30px rgba(0,0,0,.25)}
.modal-inner h2{font-size:14px;font-weight:700;margin-bottom:6px}
.modal-close{float:right;background:none;border:none;font-size:20px;cursor:pointer;color:#6b7280;line-height:1}
.mg{display:grid;grid-template-columns:1fr 1fr;gap:5px;margin:6px 0}
.mi{background:#f8fafc;border-radius:5px;padding:7px;text-align:center}
.mi .l{font-size:7px;color:#6b7280;text-transform:uppercase}
.mi .v{font-size:13px;font-weight:700;margin-top:1px}
@media(max-width:768px){.cg{grid-template-columns:1fr}.kpi-grid{grid-template-columns:1fr 1fr}.sg{grid-template-columns:1fr}.header{flex-direction:column}}
</style></head><body>
<div class="container"><div id="app">Cargando...</div></div>
<div class="modal" id="modal"><div class="modal-inner"><button class="modal-close" onclick="closeModal()">&times;</button><div id="modalBody"></div></div></div>
<script>
var D = __DATA__;
function fmt(v){return v.toLocaleString('es-ES')}
function eur(v){return v.toLocaleString('es-ES',{style:'currency',currency:'EUR'})}
var cols=['#7c3aed','#ea580c','#16a34a','#0d9488','#f59e0b','#dc2626','#2563eb','#6b7280'];
function render(){
try{
var h='';h+='<div class="header"><div><h1>KUKU<span>XUMUSU</span></h1><div class="sub">Dashboard &bull; '+D.act+'</div></div><div class="hdr-actions"><button onclick="window.print()">Imprimir</button><button onclick="location.reload()">Actualizar</button></div></div>';
h+='<div class="kpi-grid">';
[{l:'Fact.Total',v:eur(D.tf),s:fmt(D.tu)+' u'},{l:'Periodo',v:eur(D.pi),s:fmt(D.pu)+'u/'+fmt(D.pt)+'t'},{l:'Julio',v:fmt(D.ju)+'u',s:'Todas'},{l:'Tiendas',v:D.tl.length,s:'Activas'},{l:'Stock',v:fmt(D.st),s:'Unidades'},{l:'Dev.',v:fmt(D.ru)+'u',s:D.td+'%'},{l:'Ticket.Med',v:eur(D.pi/D.pt),s:'Periodo'},{l:'Prods',v:fmt(D.pp),s:'vendidos'}]
.forEach(function(k){h+='<div class="kpi"><div class="l">'+k.l+'</div><div class="v">'+k.v+'</div><div class="s">'+k.s+'</div></div>'});h+='</div>';
h+='<h2 class="sec"><span>Acumulado Total por Tienda</span><span class="tag">ene-jul 2026</span></h2><div class="sg">';
for(var i=0;i<D.tl.length;i++){var t=D.tl[i];var p=(t.fu/D.tf*100).toFixed(1);
h+='<div class="sc" style="border-top:3px solid '+cols[i]+'" onclick="openModal(\''+t.t+'\')"><h4><span>'+(i+1)+'. '+t.t+'</span><span style="font-size:8px;color:'+cols[i]+'">'+p+'%</span></h4>'+
'<div class="dg"><div class="it"><div class="lb">Facturacion</div><div class="vl" style="color:#2563eb">'+eur(t.fu)+'</div></div><div class="it"><div class="lb">Unidades</div><div class="vl" style="color:#16a34a">'+fmt(t.tu)+'</div></div></div>'+
'<div class="tp"><span class="tl">&#9733; '+t.top_n+'</span><span>'+fmt(t.top_u)+' unids</span></div>'+
'<div class="ft">Crec.Jun-Jul: <span style="color:'+(t.growth>=0?'#16a34a':'#dc2626')+';font-weight:600">'+(t.growth>=0?'+':'')+t.growth+'%</span> | '+fmt(t.prods)+' prod.</div></div>';}h+='</div>';
h+='<h2 class="sec"><span>Mes en Curso por Tienda</span><span class="tag">15-26 jul 2026</span></h2><div class="sg">';
for(var i=0;i<D.sr.length;i++){var s=D.sr[i];var p=(s.I/D.pi*100).toFixed(1);
h+='<div class="sc" style="border-top:3px solid '+cols[i]+'" onclick="openModal(\''+s.Tienda+'\')"><h4><span>'+(i+1)+'. '+s.Tienda+'</span><span style="font-size:8px;color:'+cols[i]+'">'+p+'%</span></h4>'+
'<div class="dg"><div class="it"><div class="lb">Ingresos</div><div class="vl" style="color:#2563eb">'+eur(s.I)+'</div></div><div class="it"><div class="lb">Unidades</div><div class="vl" style="color:#16a34a">'+fmt(s.U)+'</div></div></div>'+
'<div class="tp"><span class="tl">&#9733; '+s.top_pn+'</span><span>'+fmt(s.top_pu)+'u / '+eur(s.top_pi)+'</span></div>'+
'<div class="ft">Tickets: '+fmt(s.T)+' | TM: '+eur(s.TM)+' | Dev: '+fmt(Math.abs(s.RU))+'u</div></div>';}h+='</div>';
h+='<h2 class="sec"><span>Ranking Productos del Periodo</span></h2><div class="tw"><div class="tc">';
h+='<button class="act" onclick="rp(\'U\',this)">Unidades</button><button onclick="rp(\'I\',this)">Ingresos</button>';
h+='<input type="text" placeholder="Buscar..." oninput="fp(this.value)" style="margin-left:auto">';
h+='</div><table><thead><tr><th>#</th><th>Cod</th><th>Producto</th><th class="num">Unids</th><th class="num">Ingresos</th><th class="num">P.Med</th><th class="num">Tickets</th></tr></thead><tbody id="pb"></tbody></table></div>';
h+='<h2 class="sec"><span>Top Productos Acumulado Anual</span></h2><div class="tw"><div class="tc">';
h+='<button class="act" onclick="ra(\'U\',this)">Unidades</button><button onclick="ra(\'I\',this)">Ingresos</button>';
h+='<input type="text" placeholder="Buscar..." oninput="fa(this.value)" style="margin-left:auto">';
h+='</div><table><thead><tr><th>#</th><th>Producto</th><th class="num">Total Unids</th><th class="num">Total Ingresos</th><th class="num">Julio</th></tr></thead><tbody id="ab"></tbody></table></div>';
h+='<div class="cg"><div class="cb"><h3>Facturacion por Tienda</h3><canvas id="c1"></canvas></div><div class="cb"><h3>Unidades Julio</h3><canvas id="c2"></canvas></div></div>';
h+='<div class="cg cf"><div class="cb"><h3>Tendencia Mensual</h3><canvas id="c3"></canvas></div></div>';
h+='<div class="cg"><div class="cb"><h3>Categorias</h3><canvas id="c4"></canvas></div><div class="cb"><h3>Ventas Diarias</h3><canvas id="c6"></canvas></div></div>';
h+='<footer style="text-align:center;color:#9ca3af;font-size:8px;padding:6px 0;border-top:1px solid #eee;margin-top:4px">Kukuxumusu &bull; '+D.act+'</footer>';
document.getElementById('app').innerHTML=h;renderP();renderA();
if(typeof Chart!=='undefined'){
new Chart(c1,{type:'bar',data:{labels:D.tl.map(function(t){return t.t}),datasets:[{data:D.tl.map(function(t){return t.fu}),backgroundColor:cols,borderRadius:6}]},options:{responsive:true,plugins:{legend:{display:false}},scales:{y:{beginAtZero:true,ticks:{callback:function(v){return eur(v)}}}}}});
new Chart(c2,{type:'bar',data:{labels:D.tl.map(function(t){return t.t}),datasets:[{data:D.tl.map(function(t){return t.ju}),backgroundColor:cols,borderRadius:6}]},options:{responsive:true,plugins:{legend:{display:false}},scales:{y:{beginAtZero:true}}}});
new Chart(c3,{type:'line',data:{labels:D.mn,datasets:D.trend.map(function(t,i){return{label:t.t,data:t.d,borderColor:cols[i],backgroundColor:cols[i]+'18',fill:true,tension:.3,pointRadius:2}})},options:{responsive:true,plugins:{legend:{position:'bottom',labels:{boxWidth:8,padding:6}}},scales:{y:{beginAtZero:true}}}});
new Chart(c4,{type:'doughnut',data:{labels:D.cats.map(function(c){return c[D.cn]}),datasets:[{data:D.cats.map(function(c){return c.I}),backgroundColor:['#7c3aed','#2563eb','#16a34a','#ea580c','#f59e0b','#0d9488','#dc2626','#db2777','#6b7280','#b8860b'],borderWidth:0}]},options:{responsive:true,plugins:{legend:{position:'right',labels:{boxWidth:7,font:{size:8}}}},cutout:'58%'}});
new Chart(c6,{type:'bar',data:{labels:D.day.map(function(d){var f=new Date(d.Fecha);return f.getDate()+' '+['Dom','Lun','Mar','Mie','Jue','Vie','Sab'][f.getDay()]}),datasets:[{label:'Unids',data:D.day.map(function(d){return d.U}),backgroundColor:'#2563eb',borderRadius:4,yAxisID:'y'},{label:'Ingresos',data:D.day.map(function(d){return d.I}),backgroundColor:'#16a34a',borderRadius:4,yAxisID:'y1',type:'line',borderColor:'#16a34a',borderWidth:2}]},options:{responsive:true,plugins:{legend:{position:'top'}},scales:{y:{beginAtZero:true,position:'left',grid:{drawOnChartArea:false}},y1:{beginAtZero:true,position:'right',grid:{drawOnChartArea:false},ticks:{callback:function(v){return eur(v)}}}}}});}
}catch(e){document.getElementById('app').innerHTML='<div style="padding:30px;color:red"><h2>Error</h2><pre>'+e.message+'</pre></div>';}}
var modalTab='U';
function openModal(tienda){
var d=D.sd[tienda];if(!d)return;
var h='<h2>'+tienda+'</h2><div class="mg">'+
'<div class="mi"><div class="l">Ingresos Periodo</div><div class="v" style="color:#2563eb">'+eur(d.i)+'</div></div>'+
'<div class="mi"><div class="l">Unidades Periodo</div><div class="v" style="color:#16a34a">'+fmt(d.u)+'</div></div>'+
'<div class="mi"><div class="l">Fact.Acumulada</div><div class="v" style="color:#2563eb">'+eur(d.fu)+'</div></div>'+
'<div class="mi"><div class="l">Unids.Acumuladas</div><div class="v" style="color:#16a34a">'+fmt(d.tu_acum)+'</div></div>'+
'<div class="mi"><div class="l">Tickets</div><div class="v">'+fmt(d.t)+'</div></div>'+
'<div class="mi"><div class="l">Ticket Medio</div><div class="v">'+eur(d.tm)+'</div></div></div>';
h+='<div class="tc" style="margin:4px 0"><button class="act" onclick="modalTab=\'U\';renderModal(\''+tienda+'\')">Top Periodo (Unids)</button><button onclick="modalTab=\'I\';renderModal(\''+tienda+'\')">Top Periodo (Ingresos)</button>';
if(d.hist_u&&d.hist_u.length>0)h+='<button onclick="modalTab=\'H\';renderModal(\''+tienda+'\')">Top Hist\u00f3rico</button>';
h+='</div><div id="modalTop">';
if(modalTab==='U'){h+='<table><thead><tr><th>#</th><th>Cod</th><th>Producto</th><th class="num">Unids</th><th class="num">Ingresos</th></tr></thead><tbody>';for(var i=0;i<d.top_u.length;i++){var p=d.top_u[i];var cd=p['Codigo']||p['C\u00f3digo']||'';h+='<tr><td>'+(i+1)+'</td><td>'+cd+'</td><td>'+p.Nombre+'</td><td class="num">'+fmt(p.U)+'</td><td class="num">'+eur(p.I)+'</td></tr>';}h+='</tbody></table>';}
else if(modalTab==='I'){h+='<table><thead><tr><th>#</th><th>Cod</th><th>Producto</th><th class="num">Unids</th><th class="num">Ingresos</th></tr></thead><tbody>';for(var i=0;i<d.top_i.length;i++){var p=d.top_i[i];var cd=p['Codigo']||p['C\u00f3digo']||'';h+='<tr><td>'+(i+1)+'</td><td>'+cd+'</td><td>'+p.Nombre+'</td><td class="num">'+fmt(p.U)+'</td><td class="num">'+eur(p.I)+'</td></tr>';}h+='</tbody></table>';}
else if(modalTab==='H'){h+='<table><thead><tr><th>#</th><th>Cod</th><th>Producto</th><th class="num">Total Vendido</th></tr></thead><tbody>';for(var i=0;i<d.hist_u.length;i++){var p=d.hist_u[i];h+='<tr><td>'+(i+1)+'</td><td>'+p.c+'</td><td>'+p.n+'</td><td class="num">'+fmt(p.u)+' unids</td></tr>';}h+='</tbody></table>';}
h+='</div><h3 style="font-size:10px;font-weight:600;margin:6px 0 3px;color:#374151">Categorias</h3>';
h+='<table><thead><tr><th>Categoria</th><th class="num">Unids</th><th class="num">Ingresos</th></tr></thead><tbody>';
for(var i=0;i<d.cats.length;i++){var c=d.cats[i];h+='<tr><td>'+(c[D.cn]||'')+'</td><td class="num">'+fmt(c.U)+'</td><td class="num">'+eur(c.I)+'</td></tr>';}h+='</tbody></table>';
document.getElementById('modalBody').innerHTML=h;document.getElementById('modal').classList.add('show');}
function renderModal(tienda){openModal(tienda);}
function closeModal(){document.getElementById('modal').classList.remove('show');}
document.addEventListener('keydown',function(e){if(e.key==='Escape')closeModal();});
var pM='U',pF=null;
function renderP(){var d=pF||(pM==='U'?D.top_u:D.top_i);var s=d.slice();if(pM==='U')s.sort(function(a,b){return b.U-a.U});else s.sort(function(a,b){return b.I-a.I});var h='';for(var i=0;i<s.length;i++){var p=s[i];var cd=p['Codigo']||p['C\u00f3digo']||'';h+='<tr><td class="rk">'+(i+1)+'</td><td>'+cd+'</td><td>'+p.Nombre+'</td><td class="num">'+fmt(p.U)+'</td><td class="num">'+eur(p.I)+'</td><td class="num">'+eur(p.PM)+'</td><td class="num">'+fmt(p.T)+'</td></tr>';}document.getElementById('pb').innerHTML=h;}
function rp(m,btn){pM=m;var b=btn.parentNode.querySelectorAll('button');for(var i=0;i<b.length;i++)b[i].classList.remove('act');btn.classList.add('act');renderP();}
function fp(v){var q=v.toLowerCase().trim();if(!q){pF=null;}else{pF=[];var src=pM==='U'?D.top_u:D.top_i;for(var i=0;i<src.length;i++){var p=src[i];if(p.Nombre.toLowerCase().includes(q)||(p['Codigo']||p['C\u00f3digo']||'').toLowerCase().includes(q))pF.push(p);}}renderP();}
var aM='U',aF=null;
function renderA(){var d=aF||(aM==='U'?D.au:D.ai);var s=d.slice();if(aM==='U')s.sort(function(a,b){return b.tu-a.tu});else s.sort(function(a,b){return b.te-a.te});var h='';for(var i=0;i<s.length;i++){var p=s[i];h+='<tr><td class="rk">'+(i+1)+'</td><td>'+p.n+'</td><td class="num">'+fmt(p.tu)+'</td><td class="num">'+eur(p.te)+'</td><td class="num">'+fmt(p.ju)+'</td></tr>';}document.getElementById('ab').innerHTML=h;}
function ra(m,btn){aM=m;var b=btn.parentNode.querySelectorAll('button');for(var i=0;i<b.length;i++)b[i].classList.remove('act');btn.classList.add('act');renderA();}
function fa(v){var q=v.toLowerCase().trim();if(!q){aF=null;}else{aF=[];var src=aM==='U'?D.au:D.ai;for(var i=0;i<src.length;i++){if(src[i].n.toLowerCase().includes(q))aF.push(src[i]);}}renderA();}
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',render);else render();
</script></body></html>"""

html = T.replace('__DATA__', JSON)

with open(OUTPUT, 'w', encoding='utf-8') as f:
    f.write(html)

print(f"Dashboard generado: {OUTPUT} ({os.path.getsize(OUTPUT):,} bytes)")
print(f"\nListo para publicar en GitHub + Render!")
