
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
        with open("new_refresh_token.txt", "w") as f:
            f.write(res.get("refresh_token", ""))
        return res["access_token"]

def download_excel(token):
    headers = {"Authorization": f"Bearer {token}"}
    url = "https://graph.microsoft.com/v1.0/me/drive/root:/Ventas/fullm2025.xlsx:/content"
    req = urllib.request.Request(url, headers=headers)
    with urllib.request.urlopen(req) as r:
        return r.read()

def safe_val(v):
    if pd.isna(v): return None
    if hasattr(v, 'isoformat'): return str(v)
    if isinstance(v, (np.integer,)): return int(v)
    if isinstance(v, (np.floating,)): return float(v)
    return v

def analyze_data(excel_bytes):
    df = pd.read_excel(io.BytesIO(excel_bytes))
    df_v = df[df['estadofac'] == 'FACTURAD'].copy()
    df_v['fecha'] = pd.to_datetime(df_v['fecha'], errors='coerce', dayfirst=True)
    df_v = df_v.dropna(subset=['fecha'])
    df_v['mes_str'] = df_v['fecha'].dt.strftime('%Y-%m')
    df_v['año'] = df_v['fecha'].dt.year

    total_ventas = float(df_v['vtaNeta'].sum())
    total_costo = float(df_v['costotal'].sum())
    margen = round((total_ventas - total_costo) / total_ventas * 100, 1)
    total_facturas = int(df_v['documento'].nunique())
    total_clientes = int(df_v['codigocliente'].nunique())
    ticket_prom = round(total_ventas / total_facturas, 0)

    ventas_mes = df_v.groupby('mes_str').agg(ventas=('vtaNeta','sum')).reset_index()
    mes_labels = ventas_mes['mes_str'].tolist()
    mes_data = [round(float(v), 2) for v in ventas_mes['ventas'].tolist()]
    mejor_mes = ventas_mes.loc[ventas_mes['ventas'].idxmax(), 'mes_str']

    # Ventas por año
    ventas_año = df_v.groupby('año').agg(ventas=('vtaNeta','sum')).reset_index()
    años = [int(a) for a in ventas_año['año'].tolist()]
    ventas_por_año = [round(float(v), 2) for v in ventas_año['ventas'].tolist()]

    clientes = df_v.groupby(['codigocliente','cliente','localidad']).agg(
        total=('vtaNeta','sum'), facturas=('documento','nunique'),
        costo=('costotal','sum')
    ).reset_index()
    clientes['margen'] = ((clientes['total'] - clientes['costo']) / clientes['total'] * 100).round(1)
    clientes['ltv'] = (clientes['total'] * 1.2).round(0)
    clientes['ltv_real'] = ((clientes['total'] - clientes['costo']) * 1.2).round(0)
    clientes['ltv_3anos'] = (clientes['ltv_real'] * 3).round(0)
    clientes['margen_d'] = (clientes['total'] - clientes['costo']).round(0)
    clientes = clientes.sort_values('total', ascending=False)

    top_clientes = []
    for _, r in clientes.head(10).iterrows():
        top_clientes.append({
            'codigocliente': str(r['codigocliente']),
            'cliente': str(r['cliente']),
            'localidad': str(r['localidad']),
            'total': round(float(r['total']), 2),
            'facturas': int(r['facturas']),
            'costo': round(float(r['costo']), 2),
            'margen': float(r['margen']),
            'ltv': round(float(r['ltv']), 2),
            'ltv_real': round(float(r['ltv_real']), 2),
            'ltv_3anos': round(float(r['ltv_3anos']), 2),
            'margen_d': round(float(r['margen_d']), 2),
        })

    total_c = float(clientes['total'].sum())
    clientes['pct'] = (clientes['total'] / total_c * 100).round(1)
    clientes['cum'] = clientes['pct'].cumsum().round(1)
    pareto80 = int(len(clientes[clientes['cum'] <= 80]) + 1)

    prods = df_v.groupby(['codigo','articulo']).agg(
        ventas=('vtaNeta','sum'), cantidad=('cantidad','sum'), costo=('costotal','sum')
    ).reset_index().sort_values('ventas', ascending=False).head(10)
    prods['margen_pct'] = ((prods['ventas'] - prods['costo']) / prods['ventas'] * 100).round(1)

    top_prods = []
    for _, r in prods.iterrows():
        top_prods.append({
            'articulo': str(r['articulo']),
            'ventas': round(float(r['ventas']), 2),
            'costo': round(float(r['costo']), 2),
            'margen_pct': float(r['margen_pct']),
        })

    loc = df_v.groupby('localidad').agg(ventas=('vtaNeta','sum'), clientes=('codigocliente','nunique')).reset_index().sort_values('ventas', ascending=False)
    localidades = [{"localidad": str(r['localidad']), "ventas": round(float(r['ventas']),2), "clientes": int(r['clientes'])} for _, r in loc.iterrows()]

    vend = df_v.groupby('vendedor').agg(ventas=('vtaNeta','sum')).reset_index().sort_values('ventas', ascending=False).head(5)
    vendedores = [{"vendedor": str(r['vendedor']), "ventas": round(float(r['ventas']),2)} for _, r in vend.iterrows()]

    ecuador = timezone(timedelta(hours=-5))
    now = datetime.now(ecuador).strftime("%d/%m/%Y %H:%M")
    meses_count = len(mes_labels)

    return {
        "updated": now,
        "total_ventas": round(total_ventas, 0),
        "margen": margen,
        "total_facturas": total_facturas,
        "total_clientes": total_clientes,
        "ticket_prom": ticket_prom,
        "mejor_mes": mejor_mes,
        "mes_labels": mes_labels,
        "mes_data": mes_data,
        "meses_count": meses_count,
        "años": años,
        "ventas_por_año": ventas_por_año,
        "top_clientes": top_clientes,
        "pareto_clientes": pareto80,
        "pareto_pct_clientes": round(pareto80/len(clientes)*100, 1),
        "top_prods": top_prods,
        "localidades": localidades,
        "vendedores": vendedores,
        "cliente1_nombre": top_clientes[0]['cliente'][:20] if top_clientes else '',
        "cliente1_pct": round(top_clientes[0]['total']/total_ventas*100, 1) if top_clientes else 0,
        "margen_total": round(total_ventas - total_costo, 0),
        "ltv_real_total": round(sum(c['ltv_real'] for c in top_clientes), 0),
    }

def get_github_sha():
    hdrs = {"Authorization": f"token {GITHUB_TOKEN}", "User-Agent": "Claude-Bot"}
    try:
        req = urllib.request.Request(f"https://api.github.com/repos/{GITHUB_USER}/{GITHUB_REPO}/contents/index.html", headers=hdrs)
        with urllib.request.urlopen(req) as r:
            return json.loads(r.read())['sha']
    except:
        return None

def generate_html(d):
    mes_labels_js = json.dumps(d['mes_labels'])
    mes_data_js = json.dumps(d['mes_data'])
    clientes_js = json.dumps(d['top_clientes'])
    prods_js = json.dumps(d['top_prods'])
    loc_js = json.dumps(d['localidades'])
    vend_js = json.dumps(d['vendedores'])
    años_js = json.dumps(d['años'])
    vpano_js = json.dumps(d['ventas_por_año'])
    ltv_basico_js = json.dumps([round(c['ltv'],0) for c in d['top_clientes']])
    ltv_real_js = json.dumps([round(c['ltv_real'],0) for c in d['top_clientes']])
    ltv_3anos_js = json.dumps([round(c['ltv_3anos'],0) for c in d['top_clientes']])
    margen_d_js = json.dumps([round(c['margen_d'],0) for c in d['top_clientes']])
    margen_p_js = json.dumps([c['margen'] for c in d['top_clientes']])
    client_names_js = json.dumps([c['cliente'][:18] for c in d['top_clientes']])

    return open("/home/runner/work/fullmotos-dashboard/fullmotos-dashboard/template.html").read().format(
        updated=d['updated'],
        total_ventas_m=f"{d['total_ventas']/1e6:.2f}",
        margen=d['margen'],
        margen_total_m=f"{d['margen_total']/1e6:.2f}",
        total_facturas=f"{d['total_facturas']:,}",
        total_clientes=d['total_clientes'],
        ticket_prom=f"{d['ticket_prom']:,.0f}",
        mejor_mes=d['mejor_mes'],
        meses_count=d['meses_count'],
        pareto_clientes=d['pareto_clientes'],
        pareto_pct=d['pareto_pct_clientes'],
        cliente1_pct=d['cliente1_pct'],
        cliente1_nombre=d['cliente1_nombre'],
        ltv1_m=f"{d['top_clientes'][0]['ltv']/1e6:.2f}" if d['top_clientes'] else "0",
        ltv_real_total_m=f"{d['ltv_real_total']/1e6:.2f}",
        mes_labels_js=mes_labels_js,
        mes_data_js=mes_data_js,
        clientes_js=clientes_js,
        prods_js=prods_js,
        loc_js=loc_js,
        vend_js=vend_js,
        años_js=años_js,
        vpano_js=vpano_js,
        ltv_basico_js=ltv_basico_js,
        ltv_real_js=ltv_real_js,
        ltv_3anos_js=ltv_3anos_js,
        margen_d_js=margen_d_js,
        margen_p_js=margen_p_js,
        client_names_js=client_names_js,
    )

def upload_to_github(html_content):
    sha = get_github_sha()
    hdrs = {"Authorization": f"token {GITHUB_TOKEN}", "Content-Type": "application/json", "User-Agent": "Claude-Bot"}
    ecuador = timezone(timedelta(hours=-5))
    now = datetime.now(ecuador).strftime('%Y-%m-%d %H:%M ECT')
    content = base64.b64encode(html_content.encode()).decode()
    payload = {"message": f"Dashboard actualizado {now}", "content": content}
    if sha: payload["sha"] = sha
    req = urllib.request.Request(
        f"https://api.github.com/repos/{GITHUB_USER}/{GITHUB_REPO}/contents/index.html",
        data=json.dumps(payload).encode(), headers=hdrs, method="PUT"
    )
    with urllib.request.urlopen(req) as r:
        print("✅ GitHub actualizado:", json.loads(r.read())['content']['html_url'])

if __name__ == "__main__":
    print("1. Obteniendo token de acceso...")
    token = get_access_token()
    print("2. Descargando Excel de OneDrive...")
    excel_bytes = download_excel(token)
    print(f"   Excel descargado: {len(excel_bytes):,} bytes")
    print("3. Analizando datos...")
    data = analyze_data(excel_bytes)
    print(f"   Ventas totales: ${data['total_ventas']:,.0f}")
    print(f"   Años: {data['años']}")
    print("4. Generando HTML...")
    html = generate_html(data)
    print("5. Subiendo a GitHub Pages...")
    upload_to_github(html)
    print("✅ Dashboard actualizado exitosamente!")
