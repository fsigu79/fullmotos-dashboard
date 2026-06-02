import os, json, base64, urllib.request, urllib.parse, io
import pandas as pd
from datetime import datetime

CLIENT_ID = os.environ["MS_CLIENT_ID"]
CLIENT_SECRET = os.environ["MS_CLIENT_SECRET"]
TENANT_ID = os.environ["MS_TENANT_ID"]
REFRESH_TOKEN = os.environ["MS_REFRESH_TOKEN"]
GITHUB_TOKEN = os.environ["GITHUB_TOKEN"]
GITHUB_USER = "fsigu79"
GITHUB_REPO = "fullmotos-dashboard"

def get_access_token():
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    data = urllib.parse.urlencode({
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "refresh_token": REFRESH_TOKEN,
        "grant_type": "refresh_token",
        "scope": "https://graph.microsoft.com/Files.Read.All offline_access"
    }).encode()
    req = urllib.request.Request(url, data=data, method="POST")
    with urllib.request.urlopen(req) as r:
        res = json.loads(r.read())
        # Save new refresh token to file for next run
        with open("new_refresh_token.txt", "w") as f:
            f.write(res.get("refresh_token", ""))
        return res["access_token"]

def download_excel(token):
    headers = {"Authorization": f"Bearer {token}"}
    # Search for the file
    url = "https://graph.microsoft.com/v1.0/me/drive/root:/Ventas/fullm2025.xlsx:/content"
    req = urllib.request.Request(url, headers=headers)
    with urllib.request.urlopen(req) as r:
        return r.read()

def analyze_data(excel_bytes):
    df = pd.read_excel(io.BytesIO(excel_bytes))
    df_v = df[df['estadofac'] == 'FACTURAD'].copy()
    df_v['fecha'] = pd.to_datetime(df_v['fecha'])
    df_v['mes_str'] = df_v['fecha'].dt.strftime('%Y-%m')
    df_v['mes_nom'] = df_v['fecha'].dt.strftime('%b')

    total_ventas = df_v['vtaNeta'].sum()
    total_costo = df_v['costotal'].sum()
    margen = (total_ventas - total_costo) / total_ventas * 100
    total_facturas = df_v['documento'].nunique()
    total_clientes = df_v['codigocliente'].nunique()
    ticket_prom = total_ventas / total_facturas

    ventas_mes = df_v.groupby('mes_str').agg(ventas=('vtaNeta','sum')).reset_index()
    mes_labels = ventas_mes['mes_str'].tolist()
    mes_data = [round(v, 2) for v in ventas_mes['ventas'].tolist()]
    mejor_mes = ventas_mes.loc[ventas_mes['ventas'].idxmax(), 'mes_str']

    clientes = df_v.groupby(['codigocliente','cliente','localidad']).agg(
        total=('vtaNeta','sum'), facturas=('documento','nunique'),
        costo=('costotal','sum'), primera=('fecha','min'), ultima=('fecha','max')
    ).reset_index().sort_values('total', ascending=False)
    clientes['margen'] = ((clientes['total'] - clientes['costo']) / clientes['total'] * 100).round(1)
    clientes['ltv'] = (clientes['total'] * 1.2).round(0)
    clientes['ticket'] = (clientes['total'] / clientes['facturas']).round(0)
    top_clientes = clientes.head(10).to_dict('records')

    total_c = clientes['total'].sum()
    clientes['pct'] = (clientes['total'] / total_c * 100).round(1)
    clientes['cum'] = clientes['pct'].cumsum().round(1)
    pareto80 = len(clientes[clientes['cum'] <= 80]) + 1

    prods = df_v.groupby(['codigo','articulo']).agg(
        ventas=('vtaNeta','sum'), cantidad=('cantidad','sum'), costo=('costotal','sum')
    ).reset_index().sort_values('ventas', ascending=False).head(10)
    prods['margen_pct'] = ((prods['ventas'] - prods['costo']) / prods['ventas'] * 100).round(1)

    loc = df_v.groupby('localidad').agg(ventas=('vtaNeta','sum')).reset_index().sort_values('ventas', ascending=False)
    vend = df_v.groupby('vendedor').agg(ventas=('vtaNeta','sum')).reset_index().sort_values('ventas', ascending=False).head(4)

    meses_count = len(mes_data)

    return {
        "updated": datetime.now().strftime("%d/%m/%Y"),
        "total_ventas": round(total_ventas, 0),
        "margen": round(margen, 1),
        "total_facturas": total_facturas,
        "total_clientes": total_clientes,
        "ticket_prom": round(ticket_prom, 0),
        "mejor_mes": mejor_mes,
        "mes_labels": mes_labels,
        "mes_data": mes_data,
        "meses_count": meses_count,
        "top_clientes": top_clientes,
        "pareto_clientes": pareto80,
        "pareto_pct_clientes": round(pareto80/len(clientes)*100,1),
        "top_prods": prods[['articulo','ventas','costo','margen_pct']].to_dict('records'),
        "localidades": loc.to_dict('records'),
        "vendedores": vend.to_dict('records'),
        "cliente1_nombre": top_clientes[0]['cliente'][:20] if top_clientes else '',
        "cliente1_pct": round(top_clientes[0]['total']/total_ventas*100,1) if top_clientes else 0,
    }

def generate_html(d):
    mes_labels_js = json.dumps(d['mes_labels'])
    mes_data_js = json.dumps(d['mes_data'])
    clientes_js = json.dumps(d['top_clientes'])
    prods_js = json.dumps(d['top_prods'])
    loc_js = json.dumps(d['localidades'])
    vend_js = json.dumps(d['vendedores'])

    return f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Fullmotos — Dashboard BI 2025</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.js"></script>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;background:#f4f5f7;color:#1a1a2e}}
header{{background:#fff;border-bottom:1px solid #e5e7eb;padding:1rem 1.5rem;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:100}}
.logo{{display:flex;align-items:center;gap:10px}}
.logo-icon{{width:36px;height:36px;background:#185FA5;border-radius:8px;display:flex;align-items:center;justify-content:center;color:#fff;font-size:18px}}
.logo-text{{font-size:16px;font-weight:600}}.logo-sub{{font-size:11px;color:#6b7280}}
.header-right{{display:flex;align-items:center;gap:8px}}
.updated{{font-size:11px;color:#6b7280;background:#f3f4f6;padding:4px 10px;border-radius:20px}}
.btn-wa{{font-size:11px;background:#25D366;color:#fff;padding:5px 12px;border-radius:20px;text-decoration:none;font-weight:600}}
.main{{max-width:1100px;margin:0 auto;padding:1.5rem}}
.tabs{{display:flex;gap:6px;margin-bottom:1.5rem;flex-wrap:wrap;background:#fff;padding:.75rem 1rem;border-radius:12px;border:1px solid #e5e7eb}}
.tab{{font-size:12px;padding:6px 14px;border:1px solid #e5e7eb;border-radius:20px;cursor:pointer;background:#fff;color:#6b7280;font-weight:500}}
.tab.active{{background:#185FA5;color:#fff;border-color:#185FA5}}
.section{{display:none}}.section.active{{display:block}}
.slabel{{font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:.08em;color:#9ca3af;margin:0 0 .75rem}}
.kpi-grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:10px;margin-bottom:1.25rem}}
.kpi{{background:#fff;border:1px solid #e5e7eb;border-radius:10px;padding:.9rem 1rem}}
.kpi-label{{font-size:11px;color:#6b7280;margin-bottom:6px}}
.kpi-value{{font-size:22px;font-weight:700;line-height:1;color:#1a1a2e}}
.kpi-sub{{font-size:11px;color:#9ca3af;margin-top:4px}}
.kpi.blue{{border-top:3px solid #185FA5}}.kpi.blue .kpi-value{{color:#185FA5}}
.kpi.green{{border-top:3px solid #1D9E75}}.kpi.green .kpi-value{{color:#1D9E75}}
.kpi.amber{{border-top:3px solid #D97706}}.kpi.amber .kpi-value{{color:#D97706}}
.kpi.coral{{border-top:3px solid #DC2626}}.kpi.coral .kpi-value{{color:#DC2626}}
.card{{background:#fff;border:1px solid #e5e7eb;border-radius:12px;padding:1.25rem;margin-bottom:1rem}}
.card-title{{font-size:13px;font-weight:600;margin-bottom:1rem}}
.two-col{{display:grid;grid-template-columns:1fr 1fr;gap:10px}}
.legend{{display:flex;flex-wrap:wrap;gap:14px;margin-bottom:10px;font-size:11px;color:#6b7280}}
.ldot{{width:10px;height:10px;border-radius:2px;display:inline-block;margin-right:4px;vertical-align:middle}}
.crow{{display:flex;align-items:center;gap:10px;padding:9px 0;border-bottom:1px solid #f3f4f6}}
.crow:last-child{{border:none}}
.crank{{font-size:11px;font-weight:600;color:#9ca3af;width:18px;text-align:center}}
.cavatar{{width:34px;height:34px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:700;flex-shrink:0}}
.cinfo{{flex:1;min-width:0}}
.cname{{font-size:12px;font-weight:600;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}}
.cmeta{{font-size:11px;color:#9ca3af}}
.cbar-bg{{background:#f3f4f6;border-radius:4px;height:5px;margin-top:4px}}
.cbar-fill{{height:5px;border-radius:4px;background:#185FA5}}
.cval{{text-align:right;flex-shrink:0}}
.ctotal{{font-size:13px;font-weight:700}}
.cltv{{font-size:11px;color:#9ca3af}}
.prow{{display:flex;align-items:center;gap:8px;margin-bottom:7px}}
.plabel{{font-size:11px;width:130px;flex-shrink:0;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}}
.pbar-wrap{{flex:1;background:#f3f4f6;border-radius:4px;height:20px;overflow:hidden}}
.pbar-inner{{height:100%;border-radius:4px;display:flex;align-items:center;padding-left:6px}}
.ppct{{font-size:10px;font-weight:700;color:#fff}}
.pamt{{font-size:11px;color:#6b7280;width:72px;text-align:right;flex-shrink:0}}
.rcard{{border-radius:10px;padding:.9rem 1rem}}
.rcard.green{{background:#f0fdf4;border-left:3px solid #16a34a}}
.rcard.amber{{background:#fffbeb;border-left:3px solid #d97706}}
.rcard.green .rtitle{{color:#15803d}}.rcard.amber .rtitle{{color:#b45309}}
.rcard.green .rtext{{color:#166534}}.rcard.amber .rtext{{color:#92400e}}
.rtitle{{font-size:12px;font-weight:600;margin-bottom:4px}}
.rtext{{font-size:11px;line-height:1.5}}
.rgrid{{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:1rem}}
.footer{{text-align:center;padding:2rem 1rem;font-size:11px;color:#9ca3af}}
@media(max-width:600px){{.two-col{{grid-template-columns:1fr}}.rgrid{{grid-template-columns:1fr}}.kpi-value{{font-size:18px}}}}
</style>
</head>
<body>
<header>
  <div class="logo">
    <div class="logo-icon">🏍️</div>
    <div><div class="logo-text">Fullmotos Dashboard</div><div class="logo-sub">Inteligencia de negocios 2025</div></div>
  </div>
  <div class="header-right">
    <div class="updated">Actualizado: {d['updated']}</div>
    <a class="btn-wa" href="https://wa.me/593998890863?text=Hola%2C%20por%20favor%20actualiza%20el%20dashboard%20de%20Fullmotos" target="_blank">🔄 Solicitar actualización</a>
  </div>
</header>
<div class="main">
  <div class="tabs">
    <button class="tab active" onclick="showTab('resumen',this)">📊 Resumen</button>
    <button class="tab" onclick="showTab('clientes',this)">👥 Clientes & LTV</button>
    <button class="tab" onclick="showTab('pareto',this)">📐 Pareto 80/20</button>
    <button class="tab" onclick="showTab('productos',this)">🏍️ Productos</button>
    <button class="tab" onclick="showTab('recomendaciones',this)">💡 Recomendaciones</button>
  </div>

  <div id="resumen" class="section active">
    <p class="slabel">Indicadores clave — datos actualizados al {d['updated']}</p>
    <div class="kpi-grid">
      <div class="kpi blue"><div class="kpi-label">Ventas netas totales</div><div class="kpi-value">${d['total_ventas']/1e6:.2f}M</div><div class="kpi-sub">{d['meses_count']} meses</div></div>
      <div class="kpi green"><div class="kpi-label">Margen bruto</div><div class="kpi-value">{d['margen']}%</div><div class="kpi-sub">${(d['total_ventas']*d['margen']/100)/1e6:.2f}M generados</div></div>
      <div class="kpi"><div class="kpi-label">Facturas válidas</div><div class="kpi-value">{d['total_facturas']:,}</div><div class="kpi-sub">estado FACTURAD</div></div>
      <div class="kpi"><div class="kpi-label">Clientes activos</div><div class="kpi-value">{d['total_clientes']}</div><div class="kpi-sub">mayoristas B2B</div></div>
      <div class="kpi amber"><div class="kpi-label">Ticket promedio</div><div class="kpi-value">${d['ticket_prom']:,.0f}</div><div class="kpi-sub">por factura</div></div>
      <div class="kpi coral"><div class="kpi-label">Mejor mes</div><div class="kpi-value">{d['mejor_mes']}</div><div class="kpi-sub">pico del año</div></div>
    </div>
    <div class="card">
      <div class="card-title">Ventas netas por mes (USD)</div>
      <div class="legend"><span><span class="ldot" style="background:#185FA5"></span>Ventas netas</span><span><span class="ldot" style="background:#1D9E75;border-radius:50%"></span>Tendencia</span></div>
      <div style="position:relative;height:240px"><canvas id="chartMes"></canvas></div>
    </div>
    <div class="two-col">
      <div class="card"><div class="card-title">Ventas por localidad</div><div style="position:relative;height:210px"><canvas id="chartLoc"></canvas></div></div>
      <div class="card"><div class="card-title">Ventas por vendedor</div><div style="position:relative;height:210px"><canvas id="chartVend"></canvas></div></div>
    </div>
  </div>

  <div id="clientes" class="section">
    <p class="slabel">Top clientes por LTV anualizado</p>
    <div class="kpi-grid" style="grid-template-columns:repeat(3,1fr)">
      <div class="kpi blue"><div class="kpi-label">Cliente #1 LTV anual</div><div class="kpi-value" id="ltv1"></div><div class="kpi-sub" id="ltv1name"></div></div>
      <div class="kpi green"><div class="kpi-label">Pareto — clientes clave</div><div class="kpi-value">{d['pareto_clientes']} de {d['total_clientes']}</div><div class="kpi-sub">generan el 80% de ventas</div></div>
      <div class="kpi amber"><div class="kpi-label">Cliente #1 participa</div><div class="kpi-value">{d['cliente1_pct']}%</div><div class="kpi-sub">de ventas totales</div></div>
    </div>
    <div class="card"><div class="card-title">Ranking clientes — ventas y LTV proyectado anual</div><div id="clientList"></div></div>
  </div>

  <div id="pareto" class="section">
    <p class="slabel">Ley de Pareto — 80/20</p>
    <div class="kpi-grid" style="grid-template-columns:repeat(3,1fr)">
      <div class="kpi blue"><div class="kpi-label">Clientes top 20%</div><div class="kpi-value">{d['pareto_clientes']} de {d['total_clientes']}</div><div class="kpi-sub">{d['pareto_pct_clientes']}% de la base</div></div>
      <div class="kpi green"><div class="kpi-label">Ventas generadas</div><div class="kpi-value">80%</div><div class="kpi-sub">por ese grupo</div></div>
      <div class="kpi coral"><div class="kpi-label">Cliente #1 solo</div><div class="kpi-value">{d['cliente1_pct']}%</div><div class="kpi-sub">de todas las ventas</div></div>
    </div>
    <div class="card"><div class="card-title">Participación por cliente en ventas totales</div><div id="paretoList"></div></div>
    <div class="card"><div class="card-title">Curva de Pareto acumulada</div><div style="position:relative;height:240px"><canvas id="chartPareto"></canvas></div></div>
  </div>

  <div id="productos" class="section">
    <p class="slabel">Top 10 productos por ventas</p>
    <div class="kpi-grid" style="grid-template-columns:repeat(3,1fr)">
      <div class="kpi blue"><div class="kpi-label">Producto #1</div><div class="kpi-value" id="prod1val"></div><div class="kpi-sub" id="prod1name"></div></div>
      <div class="kpi green"><div class="kpi-label">Mejor margen</div><div class="kpi-value" id="bestmargen"></div><div class="kpi-sub" id="bestmargenname"></div></div>
      <div class="kpi"><div class="kpi-label">Marca dominante</div><div class="kpi-value">DAYTONA</div><div class="kpi-sub">100% del catálogo</div></div>
    </div>
    <div class="card">
      <div class="card-title">Ventas netas vs costo — top 10 productos</div>
      <div class="legend"><span><span class="ldot" style="background:#185FA5"></span>Ventas netas</span><span><span class="ldot" style="background:#85B7EB"></span>Costo</span></div>
      <div style="position:relative;height:500px"><canvas id="chartProd"></canvas></div>
    </div>
  </div>

  <div id="recomendaciones" class="section">
    <p class="slabel">Lo que estás haciendo bien ✅</p>
    <div class="rgrid" style="margin-bottom:1rem">
      <div class="rcard green"><div class="rtitle">📈 Crecimiento sostenido</div><div class="rtext">Tu negocio creció casi el doble en el año. El mes pico fue {d['mejor_mes']} — mantén esa tendencia con promociones en los meses bajos.</div></div>
      <div class="rcard green"><div class="rtitle">🤝 Clientes ancla fieles</div><div class="rtext">Tus top 2 clientes compran cada semana. Eso es lo más valioso en un negocio B2B — relaciónate con ellos constantemente.</div></div>
      <div class="rcard green"><div class="rtitle">💰 Margen saludable</div><div class="rtext">{d['margen']}% de margen bruto en distribución de motos es excelente. Por cada $100 vendidos, ${d['margen']:.0f} quedan para el negocio.</div></div>
      <div class="rcard green"><div class="rtitle">🏍️ Producto estrella claro</div><div class="rtext">Tienes un top 10 de productos definido y claro. Enfoca el stock y el marketing en los primeros 3 — generan más del 50% de ventas.</div></div>
    </div>
    <p class="slabel">Lo que deberías mejorar ⚡</p>
    <div class="rgrid">
      <div class="rcard amber"><div class="rtitle">⚠️ Concentración de riesgo</div><div class="rtext">{d['pareto_clientes']} clientes = 80% de ventas. Si uno se va, lo sientes inmediatamente. Necesitas crecer la base de clientes.</div></div>
      <div class="rcard amber"><div class="rtitle">🗺️ Expansión geográfica pendiente</div><div class="rtext">Cuenca domina. Guayaquil, Manta y Quito tienen potencial enorme. Una estrategia de expansión puede doblar el negocio.</div></div>
      <div class="rcard amber"><div class="rtitle">🧑‍💼 Dependencia de un vendedor</div><div class="rtext">Un solo vendedor maneja más del 50% de las ventas. Distribuir la cartera protege el negocio ante cualquier eventualidad.</div></div>
      <div class="rcard amber"><div class="rtitle">📅 Estacionalidad sin estrategia</div><div class="rtext">Hay meses bajos predecibles. Diseña promociones anticipadas para nivelar las ventas durante todo el año.</div></div>
    </div>
  </div>
</div>
<div class="footer">Fullmotos Dashboard · Actualizado: {d['updated']} · Generado automáticamente con GitHub Actions + Claude AI</div>

<script>
const DATA = {{
  clientes: {clientes_js},
  prods: {prods_js},
  loc: {loc_js},
  vend: {vend_js},
  mesLabels: {mes_labels_js},
  mesData: {mes_data_js}
}};

function showTab(id,el){{
  document.querySelectorAll('.section').forEach(s=>s.classList.remove('active'));
  document.querySelectorAll('.tab').forEach(t=>t.classList.remove('active'));
  document.getElementById(id).classList.add('active');
  el.classList.add('active');
}}

function fmt(v){{
  if(v===0)return '$0';
  if(v>=1e6)return '$'+(v/1e6).toFixed(2)+'M';
  if(v>=1000)return '$'+(v/1000).toFixed(0)+'K';
  return '$'+v;
}}
function fmtX(v){{
  if(v===0)return '$0';
  if(v>=1e6)return '$'+(v/1e6).toFixed(1)+'M';
  if(v>=1000)return '$'+(v/1000).toFixed(0)+'K';
  return '$'+v;
}}

const gC='rgba(0,0,0,.07)', tC='#6b7280';
const avBg=['#DBEAFE','#D1FAE5','#EDE9FE','#FEE2E2','#FEF3C7','#DCFCE7','#FCE7F3','#DBEAFE','#D1FAE5','#EDE9FE'];
const avTx=['#1e40af','#065f46','#4c1d95','#991b1b','#92400e','#14532d','#831843','#1e40af','#065f46','#4c1d95'];

// Clientes
const maxT = DATA.clientes[0]?.total || 1;
if(DATA.clientes[0]){{
  document.getElementById('ltv1').textContent = fmt(DATA.clientes[0].ltv);
  document.getElementById('ltv1name').textContent = DATA.clientes[0].cliente?.substring(0,20);
}}
const cl = document.getElementById('clientList');
DATA.clientes.forEach((c,i)=>{{
  const ini = (c.cliente||'??').split(' ').slice(0,2).map(w=>w[0]).join('');
  cl.innerHTML += `<div class="crow">
    <div class="crank">${{i+1}}</div>
    <div class="cavatar" style="background:${{avBg[i]}};color:${{avTx[i]}}">${{ini}}</div>
    <div class="cinfo">
      <div class="cname">${{c.cliente}}</div>
      <div class="cmeta">${{c.localidad}} · ${{c.facturas}} facturas · margen ${{c.margen}}%</div>
      <div class="cbar-bg"><div class="cbar-fill" style="width:${{(c.total/maxT*100).toFixed(0)}}%"></div></div>
    </div>
    <div class="cval"><div class="ctotal">${{fmt(c.total)}}</div><div class="cltv">LTV ${{fmt(c.ltv)}}/año</div></div>
  </div>`;
}});

// Pareto bars
const totalV = DATA.clientes.reduce((a,c)=>a+c.total,0);
let cum=0;
const pColors=['#185FA5','#185FA5','#185FA5','#185FA5','#185FA5','#378ADD','#378ADD','#85B7EB','#85B7EB','#B5D4F4'];
const pl=document.getElementById('paretoList');
const paretoData=[];
DATA.clientes.forEach((c,i)=>{{
  const pct=(c.total/totalV*100);
  cum+=pct;
  paretoData.push(parseFloat(cum.toFixed(1)));
  pl.innerHTML+=`<div class="prow">
    <div class="plabel">${{(c.cliente||'').substring(0,18)}}</div>
    <div class="pbar-wrap"><div class="pbar-inner" style="width:${{(pct/((DATA.clientes[0]?.total||1)/totalV*100)*100).toFixed(0)}}%;background:${{pColors[i]}}"><span class="ppct">${{pct.toFixed(1)}}%</span></div></div>
    <div class="pamt">${{fmt(c.total)}}</div>
  </div>`;
}});

// Productos
if(DATA.prods[0]){{
  document.getElementById('prod1val').textContent=fmt(DATA.prods[0].ventas);
  document.getElementById('prod1name').textContent=(DATA.prods[0].articulo||'').substring(0,20);
  const bestM=DATA.prods.reduce((a,b)=>b.margen_pct>a.margen_pct?b:a,DATA.prods[0]);
  document.getElementById('bestmargen').textContent=bestM.margen_pct+'%';
  document.getElementById('bestmargenname').textContent=(bestM.articulo||'').substring(0,20);
}}

// Charts
new Chart(document.getElementById('chartMes'),{{
  type:'bar',
  data:{{labels:DATA.mesLabels,datasets:[
    {{label:'Ventas',data:DATA.mesData,backgroundColor:'#185FA5',borderRadius:4}},
    {{type:'line',label:'Tendencia',data:DATA.mesData,borderColor:'#1D9E75',pointBackgroundColor:'#1D9E75',pointRadius:3,fill:false,tension:.4,borderDash:[4,4]}}
  ]}},
  options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:false}}}},scales:{{
    y:{{grid:{{color:gC}},ticks:{{color:tC,callback:v=>v===0?'$0':'$'+(v/1e6).toFixed(1)+'M'}}}},
    x:{{grid:{{display:false}},ticks:{{color:tC}}}}
  }}}}
}});

const locLabels=DATA.loc.map(l=>l.localidad+' '+((l.ventas/totalV)*100).toFixed(0)+'%');
const locData=DATA.loc.map(l=>l.ventas);
new Chart(document.getElementById('chartLoc'),{{
  type:'doughnut',
  data:{{labels:locLabels,datasets:[{{data:locData,backgroundColor:['#185FA5','#1D9E75','#534AB7','#D85A30','#D1D5DB'],borderWidth:2,borderColor:'#fff'}}]}},
  options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{position:'bottom',labels:{{font:{{size:10}},color:tC,padding:8}}}}}}}}
}});

const vendTotal=DATA.vend.reduce((a,v)=>a+v.ventas,0);
const vendLabels=DATA.vend.map(v=>v.vendedor+' '+((v.ventas/vendTotal)*100).toFixed(0)+'%');
new Chart(document.getElementById('chartVend'),{{
  type:'doughnut',
  data:{{labels:vendLabels,datasets:[{{data:DATA.vend.map(v=>v.ventas),backgroundColor:['#185FA5','#D85A30','#1D9E75','#534AB7'],borderWidth:2,borderColor:'#fff'}}]}},
  options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{position:'bottom',labels:{{font:{{size:10}},color:tC,padding:8}}}}}}}}
}});

new Chart(document.getElementById('chartPareto'),{{
  type:'line',
  data:{{labels:DATA.clientes.map((_,i)=>i+1+''),datasets:[
    {{label:'% acumulado',data:paretoData,borderColor:'#185FA5',backgroundColor:'rgba(24,95,165,.08)',pointBackgroundColor:'#185FA5',fill:true,tension:.3}},
    {{label:'80%',data:DATA.clientes.map(()=>80),borderColor:'#DC2626',borderDash:[6,4],pointRadius:0,fill:false}}
  ]}},
  options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:false}}}},scales:{{
    y:{{min:0,max:100,grid:{{color:gC}},ticks:{{color:tC,callback:v=>v+'%'}}}},
    x:{{title:{{display:true,text:'Número de cliente',color:tC,font:{{size:10}}}},grid:{{display:false}},ticks:{{color:tC}}}}
  }}}}
}});

new Chart(document.getElementById('chartProd'),{{
  type:'bar',indexAxis:'y',
  data:{{
    labels:DATA.prods.map(p=>(p.articulo||'').substring(0,25)),
    datasets:[
      {{label:'Ventas',data:DATA.prods.map(p=>p.ventas),backgroundColor:'#185FA5',borderRadius:4}},
      {{label:'Costo',data:DATA.prods.map(p=>p.costo),backgroundColor:'#85B7EB',borderRadius:4}}
    ]
  }},
  options:{{
    responsive:true,maintainAspectRatio:false,
    plugins:{{legend:{{display:false}},tooltip:{{callbacks:{{label:ctx=>' '+fmtX(ctx.parsed.x)}}}}}},
    scales:{{
      x:{{grid:{{color:gC}},ticks:{{color:tC,callback:fmtX}}}},
      y:{{grid:{{display:false}},ticks:{{color:tC,font:{{size:11}}}}}}
    }}
  }}
}});
</script>
</body>
</html>"""

def get_github_sha():
    TOKEN = os.environ["GITHUB_TOKEN"]
    headers = {"Authorization": f"token {TOKEN}", "User-Agent": "Claude-Bot"}
    try:
        req = urllib.request.Request(
            f"https://api.github.com/repos/{GITHUB_USER}/{GITHUB_REPO}/contents/index.html",
            headers=headers
        )
        with urllib.request.urlopen(req) as r:
            return json.loads(r.read())['sha']
    except:
        return None

def upload_to_github(html_content):
    TOKEN = os.environ["GITHUB_TOKEN"]
    sha = get_github_sha()
    headers = {"Authorization": f"token {TOKEN}", "Content-Type": "application/json", "User-Agent": "Claude-Bot"}
    content = base64.b64encode(html_content.encode()).decode()
    payload = {"message": f"Dashboard actualizado automaticamente {datetime.now().strftime('%Y-%m-%d %H:%M')}", "content": content}
    if sha:
        payload["sha"] = sha
    data = json.dumps(payload).encode()
    req = urllib.request.Request(
        f"https://api.github.com/repos/{GITHUB_USER}/{GITHUB_REPO}/contents/index.html",
        data=data, headers=headers, method="PUT"
    )
    with urllib.request.urlopen(req) as r:
        res = json.loads(r.read())
        print("GitHub actualizado:", res['content']['html_url'])

if __name__ == "__main__":
    print("1. Obteniendo token de acceso...")
    token = get_access_token()
    print("2. Descargando Excel de OneDrive...")
    excel_bytes = download_excel(token)
    print(f"   Excel descargado: {len(excel_bytes)} bytes")
    print("3. Analizando datos...")
    data = analyze_data(excel_bytes)
    print(f"   Ventas totales: ${data['total_ventas']:,.0f}")
    print("4. Generando HTML...")
    html = generate_html(data)
    print("5. Subiendo a GitHub Pages...")
    upload_to_github(html)
    print("✅ Dashboard actualizado exitosamente!")
