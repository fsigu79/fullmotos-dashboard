import os, json, base64, urllib.request, urllib.parse, io
from datetime import datetime, timezone, timedelta
import pandas as pd
import numpy as np

CLIENT_ID = os.environ["MS_CLIENT_ID"]
CLIENT_SECRET = os.environ["MS_CLIENT_SECRET"]
TENANT_ID = os.environ["MS_TENANT_ID"]
REFRESH_TOKEN = os.environ["MS_REFRESH_TOKEN"]
GITHUB_TOKEN = os.environ["GITHUB_TOKEN"]
GITHUB_USER = "fsigu79"
GITHUB_REPO = "fullmotos-dashboard"
ECUADOR = timezone(timedelta(hours=-5))

def get_access_token():
    data = urllib.parse.urlencode({
        "client_id": CLIENT_ID, "client_secret": CLIENT_SECRET,
        "refresh_token": REFRESH_TOKEN, "grant_type": "refresh_token",
        "scope": "https://graph.microsoft.com/Files.Read.All offline_access"
    }).encode()
    req = urllib.request.Request(f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token", data=data, method="POST")
    with urllib.request.urlopen(req) as r:
        res = json.loads(r.read())
        with open("new_refresh_token.txt", "w") as f:
            f.write(res.get("refresh_token", ""))
        return res["access_token"]

def download_excel(token):
    req = urllib.request.Request(
        "https://graph.microsoft.com/v1.0/me/drive/root:/Ventas/fullm2025.xlsx:/content",
        headers={"Authorization": f"Bearer {token}"}
    )
    with urllib.request.urlopen(req) as r:
        return r.read()

def fv(v):
    if v is None or (isinstance(v, float) and np.isnan(v)): return 0
    if isinstance(v, (np.integer,)): return int(v)
    if isinstance(v, (np.floating,)): return round(float(v), 2)
    return v

def analyze(excel_bytes):
    df = pd.read_excel(io.BytesIO(excel_bytes))
    dv = df[df['estadofac'] == 'FACTURAD'].copy()
    dv['fecha'] = pd.to_datetime(dv['fecha'], errors='coerce', dayfirst=True)
    dv = dv.dropna(subset=['fecha'])
    dv['mes'] = dv['fecha'].dt.strftime('%Y-%m')
    dv['año'] = dv['fecha'].dt.year.astype(int)

    tv = float(dv['vtaNeta'].sum())
    tc = float(dv['costotal'].sum())
    mg = round((tv-tc)/tv*100, 1)
    tf = int(dv['documento'].nunique())
    tcl = int(dv['codigocliente'].nunique())

    vm = dv.groupby('mes')['vtaNeta'].sum().reset_index()
    vm_labels = vm['mes'].tolist()
    vm_data = [round(float(v),2) for v in vm['vtaNeta'].tolist()]
    mejor = vm.loc[vm['vtaNeta'].idxmax(), 'mes']

    va = dv.groupby('año')['vtaNeta'].sum().reset_index()
    años = [int(a) for a in va['año'].tolist()]
    vpano = [round(float(v),2) for v in va['vtaNeta'].tolist()]

    cl = dv.groupby(['codigocliente','cliente','localidad']).agg(
        total=('vtaNeta','sum'), facturas=('documento','nunique'), costo=('costotal','sum')
    ).reset_index().sort_values('total', ascending=False)
    cl['mg'] = ((cl['total']-cl['costo'])/cl['total']*100).round(1)
    cl['ltv'] = (cl['total']*1.2).round(0)
    cl['ltv_r'] = ((cl['total']-cl['costo'])*1.2).round(0)
    cl['ltv_3'] = (cl['ltv_r']*3).round(0)
    cl['mgd'] = (cl['total']-cl['costo']).round(0)

    top = []
    for _, r in cl.head(10).iterrows():
        top.append({k: fv(v) for k,v in {
            'cliente': str(r['cliente']), 'localidad': str(r['localidad']),
            'total': r['total'], 'facturas': r['facturas'], 'costo': r['costo'],
            'margen': r['mg'], 'ltv': r['ltv'], 'ltv_real': r['ltv_r'],
            'ltv_3anos': r['ltv_3'], 'margen_d': r['mgd']
        }.items()})

    cl['pct'] = (cl['total']/tv*100).round(1)
    cl['cum'] = cl['pct'].cumsum().round(1)
    p80 = int(len(cl[cl['cum']<=80])+1)

    pr = dv.groupby(['codigo','articulo']).agg(
        ventas=('vtaNeta','sum'), costo=('costotal','sum')
    ).reset_index().sort_values('ventas', ascending=False).head(10)
    pr['mgp'] = ((pr['ventas']-pr['costo'])/pr['ventas']*100).round(1)
    prods = [{'articulo': str(r['articulo']), 'ventas': round(float(r['ventas']),2),
              'costo': round(float(r['costo']),2), 'margen_pct': float(r['mgp'])} for _,r in pr.iterrows()]

    locs = dv.groupby('localidad')['vtaNeta'].sum().reset_index().sort_values('vtaNeta', ascending=False)
    localidades = [{'localidad': str(r['localidad']), 'ventas': round(float(r['vtaNeta']),2)} for _,r in locs.iterrows()]

    vds = dv.groupby('vendedor')['vtaNeta'].sum().reset_index().sort_values('vtaNeta', ascending=False).head(5)
    vendedores = [{'vendedor': str(r['vendedor']), 'ventas': round(float(r['vtaNeta']),2)} for _,r in vds.iterrows()]

    now = datetime.now(ECUADOR).strftime("%d/%m/%Y %H:%M")
    return {
        "updated": now, "total_ventas": round(tv,0), "margen": mg,
        "margen_total": round(tv-tc,0), "total_facturas": tf, "total_clientes": tcl,
        "ticket_prom": round(tv/tf,0), "mejor_mes": mejor,
        "mes_labels": vm_labels, "mes_data": vm_data, "meses_count": len(vm_labels),
        "años": años, "ventas_por_año": vpano,
        "top_clientes": top, "pareto_clientes": p80, "pareto_pct": round(p80/len(cl)*100,1),
        "top_prods": prods, "localidades": localidades, "vendedores": vendedores,
        "cliente1_pct": round(top[0]['total']/tv*100,1) if top else 0,
        "ltv_real_total": round(sum(c['ltv_real'] for c in top),0),
    }

def build_html(d):
    av = d['avBg'] if 'avBg' in d else ['#DBEAFE','#D1FAE5','#EDE9FE','#FEE2E2','#FEF3C7','#DCFCE7','#FCE7F3','#DBEAFE','#D1FAE5','#EDE9FE']
    at = ['#1e40af','#065f46','#4c1d95','#991b1b','#92400e','#14532d','#831843','#1e40af','#065f46','#4c1d95']
    cl = d['top_clientes']
    pr = d['top_prods']
    lo = d['localidades']
    ve = d['vendedores']
    tv = d['total_ventas']

    clientes_js = json.dumps(cl)
    prods_js = json.dumps(pr)
    loc_js = json.dumps(lo)
    vend_js = json.dumps(ve)
    mes_labels_js = json.dumps(d['mes_labels'])
    mes_data_js = json.dumps(d['mes_data'])
    años_js = json.dumps(d['años'])
    vpano_js = json.dumps(d['ventas_por_año'])
    ltv_b_js = json.dumps([round(c['ltv'],0) for c in cl])
    ltv_r_js = json.dumps([round(c['ltv_real'],0) for c in cl])
    ltv_3_js = json.dumps([round(c['ltv_3anos'],0) for c in cl])
    mgd_js = json.dumps([round(c['margen_d'],0) for c in cl])
    mgp_js = json.dumps([c['margen'] for c in cl])
    names_js = json.dumps([c['cliente'][:18] for c in cl])
    mg_colors = json.dumps(['#1D9E75' if c['margen']>=30 else '#185FA5' if c['margen']>=25 else '#D97706' for c in cl])

    return f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Fullmotos — Dashboard BI</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.js"></script>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;background:#f4f5f7;color:#1a1a2e}}
header{{background:#fff;border-bottom:1px solid #e5e7eb;padding:1rem 1.5rem;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:100;flex-wrap:wrap;gap:8px}}
.logo{{display:flex;align-items:center;gap:10px}}
.logo-icon{{width:36px;height:36px;background:#185FA5;border-radius:8px;display:flex;align-items:center;justify-content:center;color:#fff;font-size:18px}}
.logo-text{{font-size:16px;font-weight:600}}.logo-sub{{font-size:11px;color:#6b7280}}
.hright{{display:flex;align-items:center;gap:8px;flex-wrap:wrap}}
.upd{{font-size:11px;color:#6b7280;background:#f3f4f6;padding:4px 10px;border-radius:20px}}
.btn-wa{{font-size:11px;background:#25D366;color:#fff;padding:5px 12px;border-radius:20px;text-decoration:none;font-weight:600}}
.main{{max-width:1100px;margin:0 auto;padding:1.5rem}}
.tabs{{display:flex;gap:6px;margin-bottom:1.5rem;flex-wrap:wrap;background:#fff;padding:.75rem 1rem;border-radius:12px;border:1px solid #e5e7eb}}
.tab{{font-size:12px;padding:6px 14px;border:1px solid #e5e7eb;border-radius:20px;cursor:pointer;background:#fff;color:#6b7280;font-weight:500}}
.tab.active{{background:#185FA5;color:#fff;border-color:#185FA5}}
.sec{{display:none}}.sec.active{{display:block}}
.sl{{font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:.08em;color:#9ca3af;margin:0 0 .75rem}}
.kg{{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:10px;margin-bottom:1.25rem}}
.k{{background:#fff;border:1px solid #e5e7eb;border-radius:10px;padding:.9rem 1rem}}
.kl{{font-size:11px;color:#6b7280;margin-bottom:6px}}
.kv{{font-size:22px;font-weight:700;line-height:1}}
.ks{{font-size:11px;color:#9ca3af;margin-top:4px}}
.k.bl{{border-top:3px solid #185FA5}}.k.bl .kv{{color:#185FA5}}
.k.gr{{border-top:3px solid #1D9E75}}.k.gr .kv{{color:#1D9E75}}
.k.am{{border-top:3px solid #D97706}}.k.am .kv{{color:#D97706}}
.k.re{{border-top:3px solid #DC2626}}.k.re .kv{{color:#DC2626}}
.k.pu{{border-top:3px solid #534AB7}}.k.pu .kv{{color:#534AB7}}
.card{{background:#fff;border:1px solid #e5e7eb;border-radius:12px;padding:1.25rem;margin-bottom:1rem}}
.ct{{font-size:13px;font-weight:600;margin-bottom:1rem}}
.tc{{display:grid;grid-template-columns:1fr 1fr;gap:10px}}
.leg{{display:flex;flex-wrap:wrap;gap:14px;margin-bottom:10px;font-size:11px;color:#6b7280}}
.ld{{width:10px;height:10px;border-radius:2px;display:inline-block;margin-right:4px;vertical-align:middle}}
.crow{{display:flex;align-items:center;gap:10px;padding:9px 0;border-bottom:1px solid #f3f4f6}}
.crow:last-child{{border:none}}
.crk{{font-size:11px;font-weight:600;color:#9ca3af;width:18px;text-align:center}}
.cav{{width:34px;height:34px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:700;flex-shrink:0}}
.ci{{flex:1;min-width:0}}
.cn{{font-size:12px;font-weight:600;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}}
.cm{{font-size:11px;color:#9ca3af}}
.cb{{background:#f3f4f6;border-radius:4px;height:5px;margin-top:4px}}
.cf{{height:5px;border-radius:4px;background:#185FA5}}
.cv{{text-align:right;flex-shrink:0}}
.ct2{{font-size:13px;font-weight:700}}
.clt{{font-size:11px;color:#9ca3af}}
.pr{{display:flex;align-items:center;gap:8px;margin-bottom:7px}}
.pl{{font-size:11px;width:130px;flex-shrink:0;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}}
.pbw{{flex:1;background:#f3f4f6;border-radius:4px;height:20px;overflow:hidden}}
.pbi{{height:100%;border-radius:4px;display:flex;align-items:center;padding-left:6px}}
.pp{{font-size:10px;font-weight:700;color:#fff}}
.pa{{font-size:11px;color:#6b7280;width:72px;text-align:right;flex-shrink:0}}
.rc{{border-radius:10px;padding:.9rem 1rem}}
.rc.g{{background:#f0fdf4;border-left:3px solid #16a34a}}
.rc.a{{background:#fffbeb;border-left:3px solid #d97706}}
.rc.g .rt{{color:#15803d}}.rc.a .rt{{color:#b45309}}
.rc.g .rb{{color:#166534}}.rc.a .rb{{color:#92400e}}
.rg{{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:1rem}}
.rt{{font-size:12px;font-weight:600;margin-bottom:4px}}
.rb{{font-size:11px;line-height:1.5}}
.alb{{background:#fffbeb;border-left:3px solid #d97706;border-radius:0 8px 8px 0;padding:.75rem 1rem;margin-bottom:1rem}}
.alt{{font-size:12px;font-weight:600;color:#92400e;margin-bottom:3px}}
.alb2{{font-size:11px;color:#b45309;line-height:1.5}}
.ins{{background:#f0fdf4;border-left:3px solid #16a34a;border-radius:0 8px 8px 0;padding:.75rem 1rem;margin-bottom:.6rem}}
.int{{font-size:12px;font-weight:600;color:#15803d;margin-bottom:3px}}
.inb{{font-size:11px;color:#166534;line-height:1.5}}
.footer{{text-align:center;padding:2rem 1rem;font-size:11px;color:#9ca3af}}
@media(max-width:600px){{.tc{{grid-template-columns:1fr}}.rg{{grid-template-columns:1fr}}.kv{{font-size:18px}}}}
</style>
</head>
<body>
<header>
  <div class="logo">
    <div class="logo-icon">🏍️</div>
    <div><div class="logo-text">Fullmotos Dashboard</div><div class="logo-sub">Inteligencia de negocios</div></div>
  </div>
  <div class="hright">
    <div class="upd">🕐 Actualizado: {d['updated']} (ECT)</div>
    <a class="btn-wa" href="https://wa.me/593998890863?text=Hola%2C%20actualiza%20el%20dashboard%20de%20Fullmotos" target="_blank">🔄 Solicitar actualización</a>
  </div>
</header>
<div class="main">
  <div class="tabs">
    <button class="tab active" onclick="st('res',this)">📊 Resumen</button>
    <button class="tab" onclick="st('cli',this)">👥 Clientes & LTV</button>
    <button class="tab" onclick="st('ltv',this)">💰 LTV Real</button>
    <button class="tab" onclick="st('par',this)">📐 Pareto</button>
    <button class="tab" onclick="st('pro',this)">🏍️ Productos</button>
    <button class="tab" onclick="st('rec',this)">💡 Recomendaciones</button>
  </div>

  <div id="res" class="sec active">
    <p class="sl">Indicadores clave — {d['meses_count']} meses de datos</p>
    <div class="kg">
      <div class="k bl"><div class="kl">Ventas netas totales</div><div class="kv">${d['total_ventas']/1e6:.2f}M</div><div class="ks">{d['meses_count']} meses</div></div>
      <div class="k gr"><div class="kl">Margen bruto</div><div class="kv">{d['margen']}%</div><div class="ks">${d['margen_total']/1e6:.2f}M ganancia</div></div>
      <div class="k"><div class="kl">Facturas válidas</div><div class="kv">{d['total_facturas']:,}</div><div class="ks">estado FACTURAD</div></div>
      <div class="k"><div class="kl">Clientes activos</div><div class="kv">{d['total_clientes']}</div><div class="ks">mayoristas B2B</div></div>
      <div class="k am"><div class="kl">Ticket promedio</div><div class="kv">${d['ticket_prom']:,.0f}</div><div class="ks">por factura</div></div>
      <div class="k re"><div class="kl">Mejor mes</div><div class="kv">{d['mejor_mes']}</div><div class="ks">pico del período</div></div>
    </div>
    <div class="card"><div class="ct">Ventas netas por mes (USD)</div>
      <div class="leg"><span><span class="ld" style="background:#185FA5"></span>Ventas</span><span><span class="ld" style="background:#1D9E75;border-radius:50%"></span>Tendencia</span></div>
      <div style="position:relative;height:240px"><canvas id="cMes"></canvas></div>
    </div>
    <div class="card"><div class="ct">Ventas por año</div>
      <div style="position:relative;height:200px"><canvas id="cAno"></canvas></div>
    </div>
    <div class="tc">
      <div class="card"><div class="ct">Ventas por localidad</div><div style="position:relative;height:210px"><canvas id="cLoc"></canvas></div></div>
      <div class="card"><div class="ct">Ventas por vendedor</div><div style="position:relative;height:210px"><canvas id="cVend"></canvas></div></div>
    </div>
  </div>

  <div id="cli" class="sec">
    <p class="sl">Top clientes por ventas y LTV</p>
    <div class="kg" style="grid-template-columns:repeat(3,1fr)">
      <div class="k bl"><div class="kl">Cliente #1 LTV anual</div><div class="kv">${d['top_clientes'][0]['ltv']/1e6:.2f}M</div><div class="ks">{d['top_clientes'][0]['cliente'][:20] if d['top_clientes'] else ''}</div></div>
      <div class="k gr"><div class="kl">Clientes clave Pareto</div><div class="kv">{d['pareto_clientes']} de {d['total_clientes']}</div><div class="ks">generan el 80% ventas</div></div>
      <div class="k am"><div class="kl">Cliente #1 participa</div><div class="kv">{d['cliente1_pct']}%</div><div class="ks">de ventas totales</div></div>
    </div>
    <div class="card"><div class="ct">Ranking clientes — ventas y LTV proyectado anual</div><div id="cliList"></div></div>
  </div>

  <div id="ltv" class="sec">
    <p class="sl">LTV real — lo que realmente te queda después de costos</p>
    <div class="kg">
      <div class="k bl"><div class="kl">LTV básico total/año</div><div class="kv">${sum(c['ltv'] for c in d['top_clientes'])/1e6:.2f}M</div><div class="ks">ventas proyectadas</div></div>
      <div class="k gr"><div class="kl">LTV real total/año</div><div class="kv">${d['ltv_real_total']/1e6:.2f}M</div><div class="ks">ganancia proyectada</div></div>
      <div class="k pu"><div class="kl">LTV real 3 años</div><div class="kv">${d['ltv_real_total']*3/1e6:.2f}M</div><div class="ks">valor total cartera</div></div>
      <div class="k re"><div class="kl">Se va en costos</div><div class="kv">{round((1-d['margen']/100)*100,1)}%</div><div class="ks">de cada venta</div></div>
    </div>
    <div class="alb"><div class="alt">⚠️ El dato que más importa</div>
      <div class="alb2">El LTV básico parece enorme pero incluye costos. Lo que realmente queda es el <strong>LTV real ({d['margen']}% de margen)</strong>. Siempre mira la ganancia, no solo las ventas.</div>
    </div>
    <div class="card"><div class="ct">LTV básico vs LTV real vs LTV 3 años por cliente</div>
      <div class="leg"><span><span class="ld" style="background:#B5D4F4"></span>LTV básico</span><span><span class="ld" style="background:#185FA5"></span>LTV real</span><span><span class="ld" style="background:#1D9E75"></span>LTV 3 años</span></div>
      <div style="position:relative;height:500px"><canvas id="cLTV"></canvas></div>
    </div>
    <div class="card"><div class="ct">Margen en dólares por cliente</div>
      <div style="position:relative;height:400px"><canvas id="cMgD"></canvas></div>
    </div>
    <div class="card"><div class="ct">Porcentaje de margen por cliente</div>
      <div style="position:relative;height:400px"><canvas id="cMgP"></canvas></div>
    </div>
  </div>

  <div id="par" class="sec">
    <p class="sl">Ley de Pareto — 80/20</p>
    <div class="kg" style="grid-template-columns:repeat(3,1fr)">
      <div class="k bl"><div class="kl">Clientes top 20%</div><div class="kv">{d['pareto_clientes']} de {d['total_clientes']}</div><div class="ks">{d['pareto_pct']}% de la base</div></div>
      <div class="k gr"><div class="kl">Ventas generadas</div><div class="kv">80%</div><div class="ks">por ese grupo</div></div>
      <div class="k re"><div class="kl">Cliente #1 solo</div><div class="kv">{d['cliente1_pct']}%</div><div class="ks">de todas las ventas</div></div>
    </div>
    <div class="card"><div class="ct">Participación por cliente en ventas totales</div><div id="parList"></div></div>
    <div class="card"><div class="ct">Curva de Pareto acumulada</div>
      <div style="position:relative;height:240px"><canvas id="cPar"></canvas></div>
    </div>
  </div>

  <div id="pro" class="sec">
    <p class="sl">Top 10 productos por ventas</p>
    <div class="kg" style="grid-template-columns:repeat(3,1fr)">
      <div class="k bl"><div class="kl">Producto #1</div><div class="kv">${d['top_prods'][0]['ventas']/1e6:.2f}M</div><div class="ks">{d['top_prods'][0]['articulo'][:25] if d['top_prods'] else ''}</div></div>
      <div class="k gr"><div class="kl">Mejor margen</div><div class="kv">{max(d['top_prods'], key=lambda x: x['margen_pct'])['margen_pct']}%</div><div class="ks">{max(d['top_prods'], key=lambda x: x['margen_pct'])['articulo'][:20]}</div></div>
      <div class="k"><div class="kl">Marca dominante</div><div class="kv">DAYTONA</div><div class="ks">100% catálogo</div></div>
    </div>
    <div class="card">
      <div class="ct">Ventas netas vs costo — top 10 productos</div>
      <div class="leg"><span><span class="ld" style="background:#185FA5"></span>Ventas netas</span><span><span class="ld" style="background:#85B7EB"></span>Costo</span></div>
      <div style="position:relative;height:500px"><canvas id="cProd"></canvas></div>
    </div>
  </div>

  <div id="rec" class="sec">
    <p class="sl">Lo que estás haciendo bien ✅</p>
    <div class="rg" style="margin-bottom:1rem">
      <div class="rc g"><div class="rt">📈 Crecimiento sostenido</div><div class="rb">Tus ventas muestran tendencia creciente a lo largo del período analizado. El negocio va en la dirección correcta.</div></div>
      <div class="rc g"><div class="rt">🤝 Clientes ancla fieles</div><div class="rb">Tus top 2 clientes compran cada semana. Eso es lo más valioso en un negocio B2B mayorista.</div></div>
      <div class="rc g"><div class="rt">💰 Margen saludable</div><div class="rb">{d['margen']}% de margen bruto en distribución de motos es excelente. Por cada $100 vendidos, ${d['margen']:.0f} quedan para el negocio.</div></div>
      <div class="rc g"><div class="rt">🏍️ Producto estrella claro</div><div class="rb">{d['top_prods'][0]['articulo'][:30] if d['top_prods'] else ''} lidera las ventas. Enfoca stock y marketing en los primeros 3 productos.</div></div>
    </div>
    <p class="sl">Lo que deberías mejorar ⚡</p>
    <div class="rg">
      <div class="rc a"><div class="rt">⚠️ Concentración de riesgo</div><div class="rb">{d['pareto_clientes']} clientes generan el 80% de ventas. Si uno se va, lo sientes inmediatamente. Diversifica ya.</div></div>
      <div class="rc a"><div class="rt">🗺️ Expansión geográfica</div><div class="rb">Cuenca domina. Guayaquil, Manta y Quito tienen potencial enorme. Una estrategia de expansión puede doblar el negocio.</div></div>
      <div class="rc a"><div class="rt">🧑‍💼 Dependencia de un vendedor</div><div class="rb">Un vendedor maneja más del 50% de ventas. Distribuir la cartera protege el negocio ante cualquier eventualidad.</div></div>
      <div class="rc a"><div class="rt">📅 Estacionalidad sin estrategia</div><div class="rb">Hay meses bajos predecibles. Diseña promociones anticipadas para nivelar ventas durante todo el año.</div></div>
    </div>
  </div>
</div>
<div class="footer">Fullmotos Dashboard · Actualizado: {d['updated']} ECT · Datos desde OneDrive · GitHub Actions + Claude AI</div>

<script>
const gC='rgba(0,0,0,.07)',tC='#6b7280';
const CLIENTES={clientes_js},PRODS={prods_js},LOC={loc_js},VEND={vend_js};
const MES_L={mes_labels_js},MES_D={mes_data_js};
const AÑOS={años_js},VPANO={vpano_js};
const LTV_B={ltv_b_js},LTV_R={ltv_r_js},LTV_3={ltv_3_js};
const MGD={mgd_js},MGP={mgp_js},NAMES={names_js},MGCOL={mg_colors};
const avBg=['{av[0]}','{av[1]}','{av[2]}','{av[3]}','{av[4]}','{av[5]}','{av[6]}','{av[7]}','{av[8]}','{av[9]}'];
const avTx=['{at[0]}','{at[1]}','{at[2]}','{at[3]}','{at[4]}','{at[5]}','{at[6]}','{at[7]}','{at[8]}','{at[9]}'];

function st(id,el){{document.querySelectorAll('.sec').forEach(s=>s.classList.remove('active'));document.querySelectorAll('.tab').forEach(t=>t.classList.remove('active'));document.getElementById(id).classList.add('active');el.classList.add('active');}}
function fx(v){{if(v===0)return '$0';if(v>=1e6)return '$'+(v/1e6).toFixed(1)+'M';if(v>=1000)return '$'+(v/1000).toFixed(0)+'K';return '$'+v;}}

// Clientes list
const maxT=CLIENTES[0]?.total||1;
const cl=document.getElementById('cliList');
CLIENTES.forEach((c,i)=>{{
  const ini=(c.cliente||'??').split(' ').slice(0,2).map(w=>w[0]).join('');
  cl.innerHTML+=`<div class="crow"><div class="crk">${{i+1}}</div><div class="cav" style="background:${{avBg[i]}};color:${{avTx[i]}}">${{ini}}</div><div class="ci"><div class="cn">${{c.cliente}}</div><div class="cm">${{c.localidad}} · ${{c.facturas}} facturas · ${{c.margen}}% margen</div><div class="cb"><div class="cf" style="width:${{(c.total/maxT*100).toFixed(0)}}%"></div></div></div><div class="cv"><div class="ct2">${{fx(c.total)}}</div><div class="clt">LTV ${{fx(c.ltv)}}/año</div></div></div>`;
}});

// Pareto list
const totV=CLIENTES.reduce((a,c)=>a+c.total,0);
let cum=0;const paretoAcum=[];
const pColors=['#185FA5','#185FA5','#185FA5','#185FA5','#185FA5','#378ADD','#378ADD','#85B7EB','#85B7EB','#B5D4F4'];
const pl=document.getElementById('parList');
CLIENTES.forEach((c,i)=>{{const pct=c.total/totV*100;cum+=pct;paretoAcum.push(parseFloat(cum.toFixed(1)));pl.innerHTML+=`<div class="pr"><div class="pl">${{(c.cliente||'').substring(0,18)}}</div><div class="pbw"><div class="pbi" style="width:${{(pct/(CLIENTES[0].total/totV*100)*100).toFixed(0)}}%;background:${{pColors[i]}}"><span class="pp">${{pct.toFixed(1)}}%</span></div></div><div class="pa">${{fx(c.total)}}</div></div>`;
}});

// Charts
new Chart(document.getElementById('cMes'),{{type:'bar',data:{{labels:MES_L,datasets:[{{label:'Ventas',data:MES_D,backgroundColor:'#185FA5',borderRadius:4}},{{type:'line',label:'Tendencia',data:MES_D,borderColor:'#1D9E75',pointBackgroundColor:'#1D9E75',pointRadius:3,fill:false,tension:.4,borderDash:[4,4]}}]}},options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:false}}}},scales:{{y:{{grid:{{color:gC}},ticks:{{color:tC,callback:v=>v===0?'$0':'$'+(v/1e6).toFixed(1)+'M'}}}},x:{{grid:{{display:false}},ticks:{{color:tC}}}}}}}}}});
new Chart(document.getElementById('cAno'),{{type:'bar',data:{{labels:AÑOS.map(String),datasets:[{{label:'Ventas por año',data:VPANO,backgroundColor:['#B5D4F4','#185FA5','#0C447C'],borderRadius:6}}]}},options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:false}}}},scales:{{y:{{grid:{{color:gC}},ticks:{{color:tC,callback:v=>v===0?'$0':'$'+(v/1e6).toFixed(1)+'M'}}}},x:{{grid:{{display:false}},ticks:{{color:tC}}}}}}}}}});
const locL=LOC.map(l=>l.localidad),locD=LOC.map(l=>l.ventas),locT=locD.reduce((a,b)=>a+b,0);
new Chart(document.getElementById('cLoc'),{{type:'doughnut',data:{{labels:locL.map((l,i)=>l+' '+((locD[i]/locT)*100).toFixed(0)+'%'),datasets:[{{data:locD,backgroundColor:['#185FA5','#1D9E75','#534AB7','#D85A30','#D1D5DB'],borderWidth:2,borderColor:'#fff'}}]}},options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{position:'bottom',labels:{{font:{{size:10}},color:tC,padding:8}}}}}}}}}});
const vendT=VEND.reduce((a,v)=>a+v.ventas,0);
new Chart(document.getElementById('cVend'),{{type:'doughnut',data:{{labels:VEND.map(v=>v.vendedor+' '+((v.ventas/vendT)*100).toFixed(0)+'%'),datasets:[{{data:VEND.map(v=>v.ventas),backgroundColor:['#185FA5','#D85A30','#1D9E75','#534AB7','#D97706'],borderWidth:2,borderColor:'#fff'}}]}},options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{position:'bottom',labels:{{font:{{size:10}},color:tC,padding:8}}}}}}}}}});
new Chart(document.getElementById('cLTV'),{{type:'bar',indexAxis:'y',data:{{labels:NAMES,datasets:[{{label:'LTV básico',data:LTV_B,backgroundColor:'#B5D4F4',borderRadius:3}},{{label:'LTV real',data:LTV_R,backgroundColor:'#185FA5',borderRadius:3}},{{label:'LTV 3 años',data:LTV_3,backgroundColor:'#1D9E75',borderRadius:3}}]}},options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:false}},tooltip:{{callbacks:{{label:ctx=>' '+fx(ctx.parsed.x)}}}}}},scales:{{x:{{type:'linear',min:0,grid:{{color:gC}},ticks:{{color:tC,callback:fx}}}},y:{{grid:{{display:false}},ticks:{{color:tC,font:{{size:10}}}}}}}}}}}});
new Chart(document.getElementById('cMgD'),{{type:'bar',indexAxis:'y',data:{{labels:NAMES,datasets:[{{label:'Margen $',data:MGD,backgroundColor:MGCOL,borderRadius:3}}]}},options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:false}},tooltip:{{callbacks:{{label:ctx=>' '+fx(ctx.parsed.x)+' ('+MGP[ctx.dataIndex]+'%)'}}}}}},scales:{{x:{{type:'linear',min:0,grid:{{color:gC}},ticks:{{color:tC,callback:fx}}}},y:{{grid:{{display:false}},ticks:{{color:tC,font:{{size:10}}}}}}}}}}}});
new Chart(document.getElementById('cMgP'),{{type:'bar',indexAxis:'y',data:{{labels:NAMES,datasets:[{{label:'% margen',data:MGP,backgroundColor:MGCOL,borderRadius:3}}]}},options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:false}},tooltip:{{callbacks:{{label:ctx=>` ${{ctx.parsed.x.toFixed(1)}}% margen`}}}}}},scales:{{x:{{type:'linear',min:0,max:40,grid:{{color:gC}},ticks:{{color:tC,callback:v=>v+'%'}}}},y:{{grid:{{display:false}},ticks:{{color:tC,font:{{size:10}}}}}}}}}}}});
new Chart(document.getElementById('cPar'),{{type:'line',data:{{labels:CLIENTES.map((_,i)=>(i+1)+''),datasets:[{{label:'% acumulado',data:paretoAcum,borderColor:'#185FA5',backgroundColor:'rgba(24,95,165,.08)',pointBackgroundColor:'#185FA5',fill:true,tension:.3}},{{label:'80%',data:CLIENTES.map(()=>80),borderColor:'#DC2626',borderDash:[6,4],pointRadius:0,fill:false}}]}},options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:false}}}},scales:{{y:{{min:0,max:100,grid:{{color:gC}},ticks:{{color:tC,callback:v=>v+'%'}}}},x:{{grid:{{display:false}},ticks:{{color:tC}}}}}}}}}});
new Chart(document.getElementById('cProd'),{{type:'bar',indexAxis:'y',data:{{labels:PRODS.map(p=>p.articulo.substring(0,25)),datasets:[{{label:'Ventas',data:PRODS.map(p=>p.ventas),backgroundColor:'#185FA5',borderRadius:3}},{{label:'Costo',data:PRODS.map(p=>p.costo),backgroundColor:'#85B7EB',borderRadius:3}}]}},options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:false}},tooltip:{{callbacks:{{label:ctx=>' '+fx(ctx.parsed.x)}}}}}},scales:{{x:{{type:'linear',min:0,grid:{{color:gC}},ticks:{{color:tC,callback:fx}}}},y:{{grid:{{display:false}},ticks:{{color:tC,font:{{size:11}}}}}}}}}}}});
</script>
</body>
</html>"""

def get_gh_sha(path):
    hdrs = {"Authorization": f"token {GITHUB_TOKEN}", "User-Agent": "Claude-Bot"}
    try:
        req = urllib.request.Request(f"https://api.github.com/repos/{GITHUB_USER}/{GITHUB_REPO}/contents/{path}", headers=hdrs)
        with urllib.request.urlopen(req) as r:
            return json.loads(r.read())['sha']
    except:
        return None

def upload_html(html):
    sha = get_gh_sha("index.html")
    hdrs = {"Authorization": f"token {GITHUB_TOKEN}", "Content-Type": "application/json", "User-Agent": "Claude-Bot"}
    ecuador = timezone(timedelta(hours=-5))
    now = datetime.now(ecuador).strftime('%Y-%m-%d %H:%M ECT')
    payload = {"message": f"Dashboard actualizado {now}", "content": base64.b64encode(html.encode()).decode()}
    if sha: payload["sha"] = sha
    req = urllib.request.Request(f"https://api.github.com/repos/{GITHUB_USER}/{GITHUB_REPO}/contents/index.html",
        data=json.dumps(payload).encode(), headers=hdrs, method="PUT")
    with urllib.request.urlopen(req) as r:
        print("GitHub OK:", json.loads(r.read())['content']['html_url'])

if __name__ == "__main__":
    print("1. Token...")
    token = get_access_token()
    print("2. Descargando Excel...")
    excel_bytes = download_excel(token)
    print(f"   {len(excel_bytes):,} bytes")
    print("3. Analizando...")
    d = analyze(excel_bytes)
    d['avBg'] = ['#DBEAFE','#D1FAE5','#EDE9FE','#FEE2E2','#FEF3C7','#DCFCE7','#FCE7F3','#DBEAFE','#D1FAE5','#EDE9FE']
    print(f"   Ventas: ${d['total_ventas']:,.0f} | Años: {d['años']}")
    print("4. Generando HTML...")
    html = build_html(d)
    print("5. Publicando...")
    upload_html(html)
    print("✅ Listo!")
