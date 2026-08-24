#!/usr/bin/env python3
"""
DASHBOARD KUKUXUMUSU v11 - DATOS COMPLETOS 2026 (1 Ene a 18 Ago)
Fuentes:
  - pyexcel: ventas linea a linea ENE-AGO
  - Pivots por tienda: ranking oficial productos/categorias con Pareto
  - Stock sheets: margenes (coste vs venta) y stock actual
"""
import pandas as pd
import json, os, sys
from datetime import datetime

BASE = 'C:/Users/hiato/Desktop/Analisis Tiendas'
DOWNLOADS = os.path.join(os.path.expanduser('~'), 'Downloads')
SRC = f'{DOWNLOADS}/20260819_Analisis Ranking.xlsx'
OUTPUT = f'{BASE}/dashboard_ejecutivo.html'

# ============ 1. VENTAS LINEA A LINEA ============
print("Leyendo ventas 2026 (ENE-AGO)...")
ventas = pd.read_excel(SRC, sheet_name='pyexcel sheet 1ENE a 18AGO 2026')
cc_v = next(c for c in ventas.columns if c.startswith('C') and 'digo' in c)
fc = 'Subtotal'

# Archivo semanal extra (18-23 ago) - evitar duplicar el dia 18
SEM = f'{DOWNLOADS}/retail_transactions_part_1 (8)/Productos vendidos , 18 al 23 agosto 2026.xlsx'
if os.path.exists(SEM):
    print("Leyendo ventas semana extra (19-23 ago)...")
    wv = pd.read_excel(SEM)
    wv['Fecha'] = pd.to_datetime(wv['Fecha'])
    wv_extra = wv[wv['Fecha'] > '2026-08-18']  # solo 19-23, el 18 ya esta en principal
    ventas = pd.concat([ventas, wv_extra], ignore_index=True)
    print(f"  Anadidas {len(wv_extra):,} filas nuevas")
else:
    print("WARN: no se encontro archivo semanal extra")

sales = ventas[ventas['Operaci\u00f3n'] == 'Venta'].copy()
returns = ventas[ventas['Operaci\u00f3n'] == 'Devoluci\u00f3n'].copy()
sales['Tienda'] = sales['Tienda'].replace({'Estafeta':'Pamplona'})
returns['Tienda'] = returns['Tienda'].replace({'Estafeta':'Pamplona'})
# Agregar mes para el periodo
sales['Fecha'] = pd.to_datetime(sales['Fecha'])
sales['Mes'] = sales['Fecha'].dt.month
returns['Fecha'] = pd.to_datetime(returns['Fecha'])
returns['Mes'] = returns['Fecha'].dt.month

# Periodo = mes en curso (agosto); resto = acumulado 2026
mes_act = 8
sales_year = sales[sales['Mes'] <= mes_act]
sales_act = sales[sales['Mes'] == mes_act]  # mes en curso (1-23 ago)

dias_agosto = sales_act['Fecha'].dt.day.max() if len(sales_act) > 0 else 23

print(f"Ventas ano: {len(sales_year)}, unidades {sales_year['Unidades'].sum():,.0f}")
print(f"Ventas agosto (1-{dias_agosto}): {len(sales_act)}, unidades {sales_act['Unidades'].sum():,.0f}")

# ============ 2. KPIS ============
K = {
    'fact_ano': round(float(sales_year[fc].sum()), 2),
    'unid_ano': int(sales_year['Unidades'].sum()),
    'tickets_ano': int(sales_year['Pedido'].nunique()),
    'fact_mes': round(float(sales_act[fc].sum()), 2),
    'unid_mes': int(sales_act['Unidades'].sum()),
    'tickets_mes': int(sales_act['Pedido'].nunique()),
    'dev_importe': round(float(abs(returns[fc].sum())), 2),
    'dev_unid': int(abs(returns['Unidades'].sum())),
    'prods_vendidos': int(sales_year[cc_v].nunique()),
    'dias_mes': dias_agosto,
}

# ============ 3. TIENDAS: ACOMO AÑO + PERIODO ============
def store_agg(df, fact_col=fc):
    return df.groupby('Tienda').agg(U=('Unidades','sum'), I=(fact_col,'sum'), T=('Pedido','nunique')).reset_index()

sr_year = store_agg(sales_year).sort_values('I', ascending=False).reset_index(drop=True)
sr_year['%I'] = (sr_year['I']/sr_year['I'].sum()*100).round(1)
sr_year['TM'] = (sr_year['I']/sr_year['T']).round(2)

sr_mes = store_agg(sales_act).sort_values('I', ascending=False).reset_index(drop=True)
sr_mes['%I'] = (sr_mes['I']/sr_mes['I'].sum()*100).round(1)
sr_mes['TM'] = (sr_mes['I']/sr_mes['T']).round(2)

# Devoluciones por tienda
ret_s = returns.groupby('Tienda').agg(RU=('Unidades','sum'), RI=(fc,'sum')).reset_index()
sr_year = sr_year.merge(ret_s, on='Tienda', how='left').fillna(0)
sr_mes = sr_mes.merge(ret_s, on='Tienda', how='left').fillna(0)
sr_year['RU'] = sr_year['RU'].astype(int); sr_mes['RU'] = sr_mes['RU'].astype(int)

# ============ 4. PIVOTS: RANKING OFICIAL POR TIENDA ============
def parse_products(df):
    """Extrae ranking de productos (nombre, total, %, %acum) del pivot"""
    header_row = None
    for i in range(min(8, len(df))):
        if any(str(v) == 'Nombre' for v in df.iloc[i]):
            header_row = i; break
    if header_row is None: return []
    prods = []
    for _, row in df.iloc[header_row+1:].iterrows():
        name = row.iloc[9]; total = row.iloc[10]
        if pd.notna(name) and pd.notna(total):
            prods.append({
                'n': str(name),
                't': round(float(total), 2),
                'p': round(float(row.iloc[11])*100, 2) if pd.notna(row.iloc[11]) else None,
                'c': round(float(row.iloc[12])*100, 2) if pd.notna(row.iloc[12]) else None
            })
    return prods

def parse_cats(df):
    """Extrae categorias del pivot (col S=18,T=19,U=20,V=21)"""
    header_row = None
    for i in range(min(8, len(df))):
        if any(str(v) == 'Nombre' for v in df.iloc[i]):
            header_row = i; break
    if header_row is None: return []
    cats = []
    for _, row in df.iloc[header_row+1:].iterrows():
        cat = row.iloc[18]; total = row.iloc[19]
        if pd.notna(cat) and pd.notna(total):
            cats.append({'n': str(cat), 't': round(float(total), 2),
                         'p': round(float(row.iloc[20])*100, 2) if pd.notna(row.iloc[20]) else None})
    return cats

def parse_notes(df):
    nota = None
    for _, row in df.head(3).iterrows():
        for v in row.values:
            if pd.notna(v) and '80%' in str(v):
                nota = str(v); break
        if nota: break
    return nota

# Map sheet names -> store name
sheet_store = {'Toros':'Toros','Donosti':'Donostia','Bilbo':'Bilbao','Estafeta':'Pamplona',
               'Zaragoza':'Zaragoza','Segovia':'Segovia','Lleida':'Lleida'}

store_rank = {}   # tienda -> ranking oficial (parcial autentico)
store_notes = {}  # tienda -> nota 80/20

# Hoja Total (global)
df_tot = pd.read_excel(SRC, sheet_name='Total', header=None)
store_rank['TOTAL'] = parse_products(df_tot)
store_notes['TOTAL'] = parse_notes(df_tot)

for sh, st in sheet_store.items():
    df_s = pd.read_excel(SRC, sheet_name=sh, header=None)
    store_rank[st] = parse_products(df_s)
    store_notes[st] = parse_notes(df_s)

print(f"Rankings oficiales: {[ (k, len(v)) for k,v in store_rank.items() ]}")
print(f"Notas: {store_notes}")

# ============ 5. MARGENES Y STOCK (hojas Stock) ============
stock_sheets = {'Zaragoza':'Stock Zaragoza','Segovia':'Stock Segovia','Lleida':'Stock Lleida'}
margenes = {}  # codigo -> {coste, venta, margen_pct}

def parse_stock(df):
    header = {i: str(v) for i, v in enumerate(df.iloc[0])}
    cc = nc = coste = venta = None
    for i, h in header.items():
        if 'digo' in h: cc = i
        if h == 'Nombre': nc = i
        if 'DDP' in h: coste = i
        if h == 'Tienda (InclTax)': venta = i
    if cc is None: return []
    items = []
    for _, row in df.iloc[1:].iterrows():
        if pd.isna(row.iloc[cc]): continue
        cod = str(row.iloc[cc]).strip()
        if cod == 'nan': continue
        it = {'c': cod, 'n': str(row.iloc[nc]) if nc is not None and pd.notna(row.iloc[nc]) else ''}
        if coste is not None and pd.notna(row.iloc[coste]): it['coste'] = float(row.iloc[coste])
        if venta is not None and pd.notna(row.iloc[venta]): it['venta'] = float(row.iloc[venta])
        items.append(it)
    return items

for st, sh in stock_sheets.items():
    try:
        df_s = pd.read_excel(SRC, sheet_name=sh, header=None)
        for it in parse_stock(df_s):
            if 'coste' in it and 'venta' in it and it['venta'] > 0 and it['coste'] > 0:
                margenes[it['c']] = {
                    'n': it['n'], 'coste': round(it['coste'], 3), 'venta': round(it['venta'], 2),
                    'mg': round((it['venta']-it['coste'])/it['venta']*100, 1)
                }
    except Exception as e:
        print(f"  Error stock {sh}: {e}")
print(f"Margenes calculados: {len(margenes)} productos")

# ============ 6. RANKINGS PRODUCTOS (unidades e ingresos, real) ============
# Ingresos reales por producto del ano
prod_ing = sales_year.groupby([cc_v,'Nombre']).agg(U=('Unidades','sum'), I=(fc,'sum'), T=('Pedido','nunique')).reset_index()
prod_ing['PM'] = (prod_ing['I']/prod_ing['U']).round(2)
top_u = prod_ing.sort_values('U', ascending=False).head(60).copy()
top_i = prod_ing.sort_values('I', ascending=False).head(60).copy()

# Añadir margen cuando exista
top_u['MG'] = top_u[cc_v].map(lambda c: margenes.get(str(c), {}).get('mg'))
top_i['MG'] = top_i[cc_v].map(lambda c: margenes.get(str(c), {}).get('mg'))

# ============ 7. CATEGORIAS / SUBCATS REALES ============
cats = sales_year.groupby('Categor\u00eda').agg(U=('Unidades','sum'), I=(fc,'sum')).sort_values('I',ascending=False).reset_index()
cats['%I'] = (cats['I']/cats['I'].sum()*100).round(1)
cn = list(cats.columns)[0]
subcats = sales_year.groupby('Subcategor\u00eda').agg(U=('Unidades','sum'), I=(fc,'sum')).sort_values('I',ascending=False).reset_index().head(12)
sn = list(subcats.columns)[0]

# ============ 8. MENSUAL + DIARIO ============
monthly = sales_year.groupby('Mes').agg(U=('Unidades','sum'), I=(fc,'sum')).reset_index()
mn = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago']

# Por tienda por mes (para tendencia)
trend = []
for tienda in sales_year['Tienda'].unique():
    td = sales_year[sales_year['Tienda']==tienda].groupby('Mes').agg(U=('Unidades','sum')).reset_index()
    vals = []
    for m in range(1, mes_act+1):
        r = td[td['Mes']==m]
        vals.append(int(r['U'].sum()) if len(r)>0 else 0)
    trend.append({'t': tienda, 'd': vals})

# Diario agosto
daily = sales_act.groupby('Fecha').agg(U=('Unidades','sum'), I=(fc,'sum'), T=('Pedido','nunique')).reset_index().sort_values('Fecha')

# ============ 9. TOP POR TIENDA (para tarjetas y modal) ============
store_detail = {}
for t in sales_year['Tienda'].unique():
    sd = sales_year[sales_year['Tienda']==t]
    sd_mes = sales_act[sales_act['Tienda']==t]
    # top por unidades e ingresos (ano)
    sd_u = sd.groupby([cc_v,'Nombre']).agg(U=('Unidades','sum'),I=(fc,'sum')).sort_values('U',ascending=False).head(10).reset_index()
    sd_i = sd.groupby([cc_v,'Nombre']).agg(U=('Unidades','sum'),I=(fc,'sum')).sort_values('I',ascending=False).head(10).reset_index()
    # top mes
    sd_mu = sd_mes.groupby([cc_v,'Nombre']).agg(U=('Unidades','sum'),I=(fc,'sum')).sort_values('U',ascending=False).head(5).reset_index()
    # ranking oficial de la tienda (top 15 por ingresos)
    oficial = store_rank.get(t, [])[:15]
    # categorias
    sd_cats = sd.groupby('Categor\u00eda').agg(U=('Unidades','sum'),I=(fc,'sum')).sort_values('I',ascending=False).reset_index()
    store_detail[t] = {
        'u': int(sd['Unidades'].sum()), 'i': round(float(sd[fc].sum()),2),
        't': int(sd['Pedido'].nunique()), 'tm': round(float(sd[fc].mean()),2),
        'u_mes': int(sd_mes['Unidades'].sum()), 'i_mes': round(float(sd_mes[fc].sum()),2), 's_mes': len(sd_mes),
        'top_u': json.loads(sd_u.to_json(orient='records')),
        'top_i': json.loads(sd_i.to_json(orient='records')),
        'top_mes': json.loads(sd_mu.to_json(orient='records')),
        'oficial': oficial,
        'cats': json.loads(sd_cats.to_json(orient='records')),
        'nota': store_notes.get(t)
    }

# ============ 10. D ============
D = {
    'act': datetime.now().strftime('%d/%m/%Y %H:%M'),
    'periodo': 'Agosto 2026 (1-23)',
    'K': K,
    'sr_year': json.loads(sr_year.to_json(orient='records')),
    'sr_mes': json.loads(sr_mes.to_json(orient='records')),
    'top_u': json.loads(top_u.to_json(orient='records')),
    'top_i': json.loads(top_i.to_json(orient='records')),
    'global_rank': store_rank.get('TOTAL', [])[:100],
    'cats': json.loads(cats.to_json(orient='records')), 'cn': cn,
    'scats': json.loads(subcats.to_json(orient='records')), 'sn': sn,
    'monthly': json.loads(monthly.to_json(orient='records')),
    'mn': mn[:mes_act],
    'trend': trend,
    'daily': json.loads(daily.to_json(orient='records')),
    'sd': store_detail,
    'num_tiendas': len(sr_year),
}
def _clean(o):
    if isinstance(o, dict): return {k: _clean(v) for k, v in o.items()}
    if isinstance(o, list): return [_clean(x) for x in o]
    if hasattr(o, 'item'): return o.item()
    return o

D = _clean(D)
JSON = json.dumps(D, ensure_ascii=False)
print(f"JSON: {len(JSON):,} chars")

# ============ HTML ============
T = r"""<!DOCTYPE html>
<html lang="es"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0"><title>KUKUXUMUSU - Dashboard 2026</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js" defer></script>
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:'Segoe UI',Arial,sans-serif;background:#f0f2f5;color:#1a1a2e;padding:14px}
.container{max-width:1440px;margin:0 auto}
.header{background:linear-gradient(135deg,#0f0c29,#302b63,#24243e);color:#fff;padding:16px 20px;border-radius:10px;margin-bottom:12px;display:flex;justify-content:space-between;flex-wrap:wrap;gap:8px;align-items:center}
.header h1{font-size:18px;font-weight:800;letter-spacing:-.5px}
.header h1 span{color:#e94560}.header .sub{font-size:10px;opacity:.65;margin-top:1px}
.hdr-actions button{padding:3px 8px;border:1px solid rgba(255,255,255,.25);border-radius:4px;background:rgba(255,255,255,.08);color:#fff;cursor:pointer;font-size:9px}
.kpi-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(120px,1fr));gap:5px;margin-bottom:12px}
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
.sc .nt{font-size:8px;color:#92400e;background:#fffbeb;border-radius:3px;padding:2px 5px;margin-top:3px}
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
.cb canvas{max-height:240px;width:100%!important}
table{width:100%;border-collapse:collapse;font-size:10px}
th{padding:3px 4px;background:#f8fafc;color:#6b7280;font-weight:600;font-size:8px;text-transform:uppercase;letter-spacing:.2px;border-bottom:2px solid #e5e7eb;white-space:nowrap;text-align:left}
td{padding:2px 4px;border-bottom:1px solid #f3f4f6;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:150px}
tr:hover td{background:#f9fafb}
.num{text-align:right}.rk{color:#9ca3af;width:18px;text-align:center}
.modal{display:none;position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,.5);z-index:1000;justify-content:center;align-items:center}
.modal.show{display:flex}
.modal-inner{background:#fff;border-radius:12px;padding:16px;max-width:660px;width:90%;max-height:85vh;overflow-y:auto;box-shadow:0 12px 30px rgba(0,0,0,.25)}
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
var h='';
h+='<div class="header"><div><h1>KUKU<span>XUMUSU</span></h1><div class="sub">Dashboard 2026 &bull; '+D.act+' &bull; 1 Ene - 18 Ago &bull; <strong style="color:#fbbf24">Cifras SIN IVA</strong></div></div><div class="hdr-actions"><button onclick="window.print()">Imprimir</button><button onclick="location.reload()">Actualizar</button></div></div>';
h+='<div class="kpi-grid">';
var K=D.K;
[{l:'Fact. 2026',v:eur(K.fact_ano),s:fmt(K.unid_ano)+' unids / '+fmt(K.tickets_ano)+' tickets'},{l:'Mes en Curso',v:eur(K.fact_mes),s:'Agosto (1-'+K.dias_mes+') '+fmt(K.unid_mes)+' unids'},
{l:'Media/Dia Ago',v:eur(K.fact_mes/K.dias_mes),s:'sobre '+K.dias_mes+' dias'},{l:'Tiendas',v:D.num_tiendas,s:'activas'},{l:'Devoluciones',v:fmt(K.dev_unid)+' u',s:eur(K.dev_importe)+' importe'},{l:'Prods.Vendidos',v:fmt(K.prods_vendidos),s:'en 2026'},{l:'Ticket Medio',v:eur(K.fact_ano/K.tickets_ano),s:'2026'}]
.forEach(function(k){h+='<div class="kpi"><div class="l">'+k.l+'</div><div class="v">'+k.v+'</div><div class="s">'+k.s+'</div></div>'});h+='</div>';

// ===== ACUMULADO 2026 =====
h+='<h2 class="sec"><span>Tiendas &bull; Acumulado 2026</span><span class="tag">Ene-Ago</span></h2><div class="sg">';
for(var i=0;i<D.sr_year.length;i++){var s=D.sr_year[i];var p=(s.I/D.sr_year[0].I*100).toFixed(1);
h+='<div class="sc" style="border-top:3px solid '+cols[i]+'" onclick="openModal(\''+s.Tienda+'\')"><h4><span>'+(i+1)+'. '+s.Tienda+'</span><span style="font-size:8px;color:'+cols[i]+'">'+p+'%</span></h4>'+
'<div class="dg"><div class="it"><div class="lb">Ingresos</div><div class="vl" style="color:#2563eb">'+eur(s.I)+'</div></div><div class="it"><div class="lb">Unidades</div><div class="vl" style="color:#16a34a">'+fmt(s.U)+'</div></div></div>'+
'<div class="ft">Tickets: '+fmt(s.T)+' | TM: '+eur(s.TM)+' | Dev: '+fmt(Math.abs(s.RU))+'</div></div>';}h+='</div>';

// ===== MES EN CURSO =====
h+='<h2 class="sec"><span>Tiendas &bull; Mes en Curso</span><span class="tag">Agosto 2026 (1-23)</span></h2><div class="sg">';
for(var i=0;i<D.sr_mes.length;i++){var s=D.sr_mes[i];var p=(s.I/D.sr_mes[0].I*100).toFixed(1);
h+='<div class="sc" style="border-top:3px solid '+cols[i]+'" onclick="openModal(\''+s.Tienda+'\')"><h4><span>'+(i+1)+'. '+s.Tienda+'</span><span style="font-size:8px;color:'+cols[i]+'">'+p+'%</span></h4>'+
'<div class="dg"><div class="it"><div class="lb">Ingresos</div><div class="vl" style="color:#2563eb">'+eur(s.I)+'</div></div><div class="it"><div class="lb">Unidades</div><div class="vl" style="color:#16a34a">'+fmt(s.U)+'</div></div></div>'+
'<div class="ft">Tickets: '+fmt(s.T)+' | TM: '+eur(s.TM)+'</div></div>';}h+='</div>';

// ===== RANKING PRODUCTOS =====
h+='<h2 class="sec"><span>Ranking Productos 2026</span></h2><div class="tw"><div class="tc">';
h+='<button class="act" onclick="rp(\'U\',this)">Unidades</button><button onclick="rp(\'I\',this)">Ingresos</button>';
h+='<input type="text" placeholder="Buscar..." oninput="fp(this.value)" style="margin-left:auto">';
h+='</div><table><thead><tr><th>#</th><th>Cod</th><th>Producto</th><th class="num">Unids</th><th class="num">Ingresos</th><th class="num">P.Med</th><th class="num">Margen</th></tr></thead><tbody id="pb"></tbody></table></div>';

// ===== PARETO GLOBAL =====
h+='<h2 class="sec"><span>Pareto Global &bull; Top 100 por Ingresos</span><span class="tag">analisis oficial</span></h2><div class="tw"><div class="tc">';
h+='<input type="text" placeholder="Buscar..." oninput="fg(this.value)" style="margin-left:auto">';
h+='</div><table><thead><tr><th>#</th><th>Producto</th><th class="num">Ingresos</th><th class="num">%</th><th class="num">% Acum</th></tr></thead><tbody id="gb"></tbody></table></div>';

// ===== CHARTS =====
h+='<div class="cg"><div class="cb"><h3>Ingresos por Tienda (2026)</h3><canvas id="c1"></canvas></div><div class="cb"><h3>Ventas Mensuales 2026</h3><canvas id="c2"></canvas></div></div>';
h+='<div class="cg cf"><div class="cb"><h3>Tendencia por Tienda</h3><canvas id="c3"></canvas></div></div>';
h+='<div class="cg"><div class="cb"><h3>Categorias</h3><canvas id="c4"></canvas></div><div class="cb"><h3>Subcategorias</h3><canvas id="c5"></canvas></div></div>';
h+='<div class="cg cf"><div class="cb"><h3>Ventas Diarias Agosto</h3><canvas id="c6"></canvas></div></div>';
document.getElementById('app').innerHTML=h;
renderP();renderG();
if(typeof Chart!=='undefined'){
new Chart(c1,{type:'bar',data:{labels:D.sr_year.map(function(s){return s.Tienda}),datasets:[{data:D.sr_year.map(function(s){return s.I}),backgroundColor:cols,borderRadius:6}]},options:{responsive:true,plugins:{legend:{display:false}},scales:{y:{beginAtZero:true,ticks:{callback:function(v){return eur(v)}}}}}});
new Chart(c2,{type:'bar',data:{labels:D.mn,datasets:[{label:'Ingresos',data:D.monthly.map(function(m){return m.I}),backgroundColor:'#2563eb',borderRadius:6,yAxisID:'y'},{label:'Unidades',data:D.monthly.map(function(m){return m.U}),backgroundColor:'#16a34a',borderRadius:6,yAxisID:'y1'}]},options:{responsive:true,plugins:{legend:{position:'top',labels:{boxWidth:10}}},scales:{y:{beginAtZero:true,position:'left'},y1:{beginAtZero:true,position:'right',grid:{drawOnChartArea:false}}}}});
new Chart(c3,{type:'line',data:{labels:D.mn,datasets:D.trend.map(function(t,i){return{label:t.t,data:t.d,borderColor:cols[i],backgroundColor:cols[i]+'18',fill:true,tension:.3,pointRadius:2}})},options:{responsive:true,plugins:{legend:{position:'bottom',labels:{boxWidth:8,padding:6}}},scales:{y:{beginAtZero:true}}}});
new Chart(c4,{type:'doughnut',data:{labels:D.cats.map(function(c){return c[D.cn]}),datasets:[{data:D.cats.map(function(c){return c.I}),backgroundColor:['#7c3aed','#2563eb','#16a34a','#ea580c','#f59e0b','#0d9488','#dc2626','#db2777','#6b7280','#b8860b'],borderWidth:0}]},options:{responsive:true,plugins:{legend:{position:'right',labels:{boxWidth:7,font:{size:8}}}},cutout:'58%'}});
new Chart(c5,{type:'bar',data:{labels:D.scats.map(function(s){return s[D.sn]}),datasets:[{label:'Ingresos',data:D.scats.map(function(s){return s.I}),backgroundColor:'#7c3aed',borderRadius:4}]},options:{indexAxis:'y',responsive:true,plugins:{legend:{display:false}},scales:{x:{beginAtZero:true,ticks:{callback:function(v){return eur(v)}}}}}});
new Chart(c6,{type:'bar',data:{labels:D.daily.map(function(d){var f=new Date(d.Fecha);return f.getDate()+' '+['Dom','Lun','Mar','Mie','Jue','Vie','Sab'][f.getDay()]}),datasets:[{label:'Unids',data:D.daily.map(function(d){return d.U}),backgroundColor:'#2563eb',borderRadius:4,yAxisID:'y'},{label:'Ingresos',data:D.daily.map(function(d){return d.I}),backgroundColor:'#16a34a',borderRadius:4,yAxisID:'y1',type:'line',borderColor:'#16a34a',borderWidth:2}]},options:{responsive:true,plugins:{legend:{position:'top'}},scales:{y:{beginAtZero:true,position:'left',grid:{drawOnChartArea:false}},y1:{beginAtZero:true,position:'right',grid:{drawOnChartArea:false},ticks:{callback:function(v){return eur(v)}}}}}});
}
}catch(e){document.getElementById('app').innerHTML='<div style="padding:30px;color:red"><h2>Error</h2><pre>'+e.message+'</pre></div>';}}

// MODAL TIENDA
function openModal(tienda){
var d=D.sd[tienda];if(!d)return;
var h='<h2>'+tienda+'</h2><div class="mg">'+
'<div class="mi"><div class="l">Ingresos 2026</div><div class="v" style="color:#2563eb">'+eur(d.i)+'</div></div>'+
'<div class="mi"><div class="l">Unids 2026</div><div class="v" style="color:#16a34a">'+fmt(d.u)+'</div></div>'+
'<div class="mi"><div class="l">Agosto</div><div class="v">'+eur(d.i_mes)+'</div></div>'+
'<div class="mi"><div class="l">Unids Agosto</div><div class="v">'+fmt(d.u_mes)+'</div></div>'+
'<div class="mi"><div class="l">Tickets</div><div class="v">'+fmt(d.t)+'</div></div>'+
'<div class="mi"><div class="l">Ticket Medio</div><div class="v">'+eur(d.tm)+'</div></div></div>';
if(d.nota){h+='<div class="nt" style="font-size:9px;color:#92400e;background:#fffbeb;padding:4px 8px;border-radius:5px;margin-bottom:6px">&#128161; '+d.nota+'</div>';}
h+='<h3 style="font-size:10px;font-weight:600;margin:6px 0 3px;color:#374151">Top 10 Productos 2026</h3>';
h+='<table><thead><tr><th>#</th><th>Cod</th><th>Producto</th><th class="num">Unids</th><th class="num">Ingresos</th></tr></thead><tbody>';
for(var i=0;i<d.top_i.length;i++){var p=d.top_i[i];var cd=p['Codigo']||p['C\u00f3digo']||'';h+='<tr><td>'+(i+1)+'</td><td>'+cd+'</td><td>'+p.Nombre+'</td><td class="num">'+fmt(p.U)+'</td><td class="num">'+eur(p.I)+'</td></tr>';}
h+='</tbody></table>';
if(d.oficial&&d.oficial.length>0){h+='<h3 style="font-size:10px;font-weight:600;margin:6px 0 3px;color:#374151">Ranking Oficial (Pareto)</h3><table><thead><tr><th>#</th><th>Producto</th><th class="num">Ingresos</th><th class="num">%</th><th class="num">%Acum</th></tr></thead><tbody>';
for(var i=0;i<d.oficial.length;i++){var p=d.oficial[i];h+='<tr><td>'+(i+1)+'</td><td>'+p.n+'</td><td class="num">'+eur(p.t)+'</td><td class="num">'+p.p+'%</td><td class="num">'+p.c+'%</td></tr>';}
h+='</tbody></table>';}
h+='<h3 style="font-size:10px;font-weight:600;margin:6px 0 3px;color:#374151">Categorias</h3><table><thead><tr><th>Categoria</th><th class="num">Unids</th><th class="num">Ingresos</th></tr></thead><tbody>';
for(var i=0;i<d.cats.length;i++){var c=d.cats[i];h+='<tr><td>'+(c['Categor\u00eda']||'')+'</td><td class="num">'+fmt(c.U)+'</td><td class="num">'+eur(c.I)+'</td></tr>';}
h+='</tbody></table>';
document.getElementById('modalBody').innerHTML=h;document.getElementById('modal').classList.add('show');}
function closeModal(){document.getElementById('modal').classList.remove('show');}
document.addEventListener('keydown',function(e){if(e.key==='Escape')closeModal();});

// RANKING PRODUCTOS
var pM='U',pF=null;
function renderP(){var d=pF||(pM==='U'?D.top_u:D.top_i);var s=d.slice();if(pM==='U')s.sort(function(a,b){return b.U-a.U});else s.sort(function(a,b){return b.I-a.I});
var h='';for(var i=0;i<s.length;i++){var p=s[i];var cd=p['Codigo']||p['C\u00f3digo']||'';var mg=p.MG!=null?p.MG+'%':'---';
h+='<tr><td class="rk">'+(i+1)+'</td><td>'+cd+'</td><td>'+p.Nombre+'</td><td class="num">'+fmt(p.U)+'</td><td class="num">'+eur(p.I)+'</td><td class="num">'+eur(p.PM)+'</td><td class="num">'+mg+'</td></tr>';}
document.getElementById('pb').innerHTML=h;}
function rp(m,btn){pM=m;var b=btn.parentNode.querySelectorAll('button');for(var i=0;i<b.length;i++)b[i].classList.remove('act');btn.classList.add('act');renderP();}
function fp(v){var q=v.toLowerCase().trim();if(!q){pF=null;}else{pF=[];var src=pM==='U'?D.top_u:D.top_i;for(var i=0;i<src.length;i++){var p=src[i];if(p.Nombre.toLowerCase().includes(q)||(p['Codigo']||p['C\u00f3digo']||'').toLowerCase().includes(q))pF.push(p);}}renderP();}

// PARETO GLOBAL
var gF=null;
function renderG(){var d=gF||D.global_rank;var h='';for(var i=0;i<Math.min(100,d.length);i++){var p=d[i];h+='<tr><td class="rk">'+(i+1)+'</td><td>'+p.n+'</td><td class="num">'+eur(p.t)+'</td><td class="num">'+p.p+'%</td><td class="num">'+p.c+'%</td></tr>';}document.getElementById('gb').innerHTML=h;}
function fg(v){var q=v.toLowerCase().trim();if(!q){gF=null;}else{gF=[];for(var i=0;i<D.global_rank.length;i++){if(D.global_rank[i].n.toLowerCase().includes(q))gF.push(D.global_rank[i]);}}renderG();}

if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',render);else render();
</script></body></html>"""

html = T.replace('__DATA__', JSON)
with open(OUTPUT, 'w', encoding='utf-8') as f:
    f.write(html)
print(f"Dashboard v11: {OUTPUT} ({os.path.getsize(OUTPUT):,} bytes)")