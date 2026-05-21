"""
generar_dashboard.py
--------------------
Lee el Excel de TCO y genera el archivo docs/index.html
Ejecutado automáticamente por GitHub Actions al subir el Excel.
"""

import pandas as pd
import json
import os
from datetime import datetime

# ── 1. LEER EL EXCEL ────────────────────────────────────────────────────────
EXCEL_PATH = "data/Resultado_respuesta_ONNET_SN.xls"

# Detectar extensión para elegir motor correcto
ext = os.path.splitext(EXCEL_PATH)[1].lower()
engine = "xlrd" if ext == ".xls" else "openpyxl"

df = pd.read_excel(EXCEL_PATH, engine=engine)
print(f"✅ Excel leído: {len(df)} registros")

# ── 2. SELECCIONAR Y RENOMBRAR COLUMNAS ─────────────────────────────────────
cols_map = {
    "Número":                            "tco",
    "Creado":                            "fecha_creacion",
    "Estado":                            "estado",
    "Cerrado":                           "fecha_respuesta",
    "Notas de resolución":               "nota_resultado",
    "Ámbito [Tiquete Comercial]":        "ambito",
    "Localidad [Tiquete Comercial]":     "localidad",
    "Descripción breve":                 "descripcion",
    "Grupo de asignación":               "grupo",
}

df = df[[c for c in cols_map if c in df.columns]].rename(columns=cols_map)

# Convertir fechas a string legible
for col in ["fecha_creacion", "fecha_respuesta"]:
    if col in df.columns:
        df[col] = pd.to_datetime(df[col], errors="coerce").dt.strftime("%Y-%m-%d %H:%M")

df = df.fillna("")
data_json = json.dumps(df.to_dict(orient="records"), ensure_ascii=False, separators=(",", ":"))
print(f"✅ Datos preparados: {len(data_json)//1024} KB")

# ── 3. GENERAR EL HTML ───────────────────────────────────────────────────────
update_ts = datetime.now().strftime("%d/%m/%Y %H:%M")

html = f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Dashboard TCO — ONNET</title>
<link href="https://fonts.googleapis.com/css2?family=Space+Mono:wght@400;700&family=DM+Sans:wght@300;400;500;700&display=swap" rel="stylesheet">
<style>
  :root {{
    --bg:#0b0f1a; --surface:#131929; --surface2:#1a2235; --border:#243050;
    --accent:#00d4ff; --accent2:#ff6b35; --accent3:#7fff6b;
    --warn:#ffb547; --danger:#ff4d6d;
    --text:#e8edf7; --text2:#8a9bc0;
    --mono:'Space Mono',monospace; --body:'DM Sans',sans-serif;
  }}
  *{{box-sizing:border-box;margin:0;padding:0}}
  body{{background:var(--bg);color:var(--text);font-family:var(--body);min-height:100vh}}

  .header{{background:var(--surface);border-bottom:1px solid var(--border);padding:16px 28px;
    display:flex;align-items:center;gap:14px;position:sticky;top:0;z-index:100;
    box-shadow:0 4px 24px rgba(0,0,0,.4)}}
  .logo{{font-family:var(--mono);font-size:1.2rem;color:var(--accent);letter-spacing:2px;font-weight:700}}
  .logo span{{color:var(--text2);font-size:.7rem;display:block;letter-spacing:1px}}
  .hbadge{{background:rgba(0,212,255,.1);border:1px solid var(--accent);color:var(--accent);
    font-family:var(--mono);font-size:.65rem;padding:4px 10px;border-radius:4px;letter-spacing:1px}}
  .update-info{{margin-left:auto;font-family:var(--mono);font-size:.65rem;color:var(--text2)}}
  .update-info strong{{color:var(--accent3)}}

  .container{{max-width:1400px;margin:0 auto;padding:24px}}

  /* METRICS */
  .metrics{{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:14px;margin-bottom:24px}}
  .mc{{background:var(--surface);border:1px solid var(--border);border-radius:10px;padding:18px 16px;
    position:relative;overflow:hidden;transition:transform .2s}}
  .mc:hover{{transform:translateY(-2px)}}
  .mc::before{{content:'';position:absolute;top:0;left:0;right:0;height:3px}}
  .mc-total::before{{background:var(--accent)}}
  .mc-pend::before{{background:var(--accent2)}}
  .mc-abierto::before{{background:var(--danger)}}
  .mc-nuevo::before{{background:var(--warn)}}
  .mc-prog::before{{background:#a78bfa}}
  .mc-res::before{{background:var(--accent3)}}
  .mc-cerr::before{{background:var(--text2)}}
  .mc-label{{font-family:var(--mono);font-size:.65rem;color:var(--text2);letter-spacing:1px;
    text-transform:uppercase;margin-bottom:8px}}
  .mc-val{{font-family:var(--mono);font-size:2rem;font-weight:700;line-height:1}}
  .mc-total .mc-val{{color:var(--accent)}}
  .mc-pend  .mc-val{{color:var(--accent2)}}
  .mc-abierto .mc-val{{color:var(--danger)}}
  .mc-nuevo .mc-val{{color:var(--warn)}}
  .mc-prog  .mc-val{{color:#a78bfa}}
  .mc-res   .mc-val{{color:var(--accent3)}}
  .mc-cerr  .mc-val{{color:var(--text2)}}
  .mc-sub{{font-size:.68rem;color:var(--text2);margin-top:4px}}

  /* PANELS */
  .panels{{display:grid;grid-template-columns:360px 1fr;gap:18px;margin-bottom:24px}}
  @media(max-width:900px){{.panels{{grid-template-columns:1fr}}}}
  .panel{{background:var(--surface);border:1px solid var(--border);border-radius:10px;overflow:hidden}}
  .ph{{padding:12px 16px;border-bottom:1px solid var(--border);background:var(--surface2);
    display:flex;align-items:center;gap:10px}}
  .ph-dot{{width:8px;height:8px;border-radius:50%;background:var(--accent);
    box-shadow:0 0 8px var(--accent);animation:pulse 2s infinite}}
  .ph-dot.orange{{background:var(--accent2);box-shadow:0 0 8px var(--accent2)}}
  @keyframes pulse{{0%,100%{{opacity:1}}50%{{opacity:.4}}}}
  .ph-title{{font-family:var(--mono);font-size:.7rem;letter-spacing:1.5px;
    text-transform:uppercase;color:var(--text2)}}
  .ph-title strong{{color:var(--accent)}}

  /* COSTA */
  .costa-list{{padding:10px}}
  .citem{{display:flex;align-items:center;padding:9px 10px;border-radius:7px;
    margin-bottom:5px;cursor:pointer;transition:background .15s;border:1px solid transparent}}
  .citem:hover{{background:var(--surface2);border-color:var(--border)}}
  .citem.active{{background:rgba(0,212,255,.08);border-color:rgba(0,212,255,.3)}}
  .cc{{font-weight:600;font-size:.88rem;flex:1}}
  .cbar{{flex:2;margin:0 10px}}
  .cbar-bg{{background:var(--border);height:5px;border-radius:3px;overflow:hidden}}
  .cbar-fill{{height:100%;border-radius:3px;background:var(--accent);transition:width .5s}}
  .ccount{{font-family:var(--mono);font-size:.82rem;color:var(--accent);font-weight:700;min-width:32px;text-align:right}}
  .cpend{{font-family:var(--mono);font-size:.62rem;padding:2px 6px;border-radius:10px;
    background:rgba(255,77,109,.15);color:var(--danger);border:1px solid rgba(255,77,109,.3);margin-left:6px}}

  /* SEARCH */
  .sp{{padding:14px}}
  .srow{{display:flex;gap:8px;margin-bottom:12px;flex-wrap:wrap}}
  .siw{{position:relative;flex:1;min-width:180px}}
  .sicon{{position:absolute;left:10px;top:50%;transform:translateY(-50%);color:var(--text2);font-size:.9rem}}
  .sinput{{width:100%;background:var(--bg);border:1px solid var(--border);color:var(--text);
    font-family:var(--mono);font-size:.82rem;padding:9px 12px 9px 32px;border-radius:7px;outline:none;transition:border .2s}}
  .sinput:focus{{border-color:var(--accent)}}
  .sinput::placeholder{{color:var(--text2)}}
  select.fsel{{background:var(--bg);border:1px solid var(--border);color:var(--text);
    font-family:var(--body);font-size:.82rem;padding:9px 10px;border-radius:7px;outline:none;cursor:pointer}}
  select.fsel:focus{{border-color:var(--accent)}}
  .breset{{background:rgba(255,107,53,.12);border:1px solid rgba(255,107,53,.35);color:var(--accent2);
    font-family:var(--mono);font-size:.7rem;padding:9px 12px;border-radius:7px;cursor:pointer;letter-spacing:1px}}
  .breset:hover{{background:rgba(255,107,53,.22)}}
  .rinfo{{font-family:var(--mono);font-size:.68rem;color:var(--text2);margin-bottom:8px}}
  .rinfo strong{{color:var(--accent)}}

  /* TABLE */
  .tw{{overflow-x:auto}}
  table{{width:100%;border-collapse:collapse;font-size:.81rem}}
  thead th{{background:var(--surface2);padding:10px 12px;text-align:left;
    font-family:var(--mono);font-size:.67rem;letter-spacing:1px;color:var(--text2);
    text-transform:uppercase;border-bottom:1px solid var(--border);white-space:nowrap;
    cursor:pointer;user-select:none}}
  thead th:hover{{color:var(--accent)}}
  thead th.sorted{{color:var(--accent)}}
  tbody tr{{border-bottom:1px solid rgba(36,48,80,.5);transition:background .1s}}
  tbody tr:hover{{background:rgba(255,255,255,.03)}}
  tbody td{{padding:9px 12px;vertical-align:middle}}
  .tco-n{{font-family:var(--mono);font-size:.76rem;color:var(--accent);white-space:nowrap}}
  .dc{{font-family:var(--mono);font-size:.72rem;color:var(--text2);white-space:nowrap}}
  .desc-c,.nota-c{{max-width:220px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-size:.78rem;color:var(--text2)}}
  .nota-c{{font-style:italic}}

  .badge{{display:inline-flex;align-items:center;gap:4px;font-family:var(--mono);
    font-size:.65rem;padding:3px 8px;border-radius:20px;white-space:nowrap;font-weight:700}}
  .badge::before{{content:'●';font-size:.48rem}}
  .b-abierto{{background:rgba(255,77,109,.12);color:#ff4d6d;border:1px solid rgba(255,77,109,.3)}}
  .b-nuevo{{background:rgba(255,181,71,.12);color:#ffb547;border:1px solid rgba(255,181,71,.3)}}
  .b-prog{{background:rgba(167,139,250,.12);color:#a78bfa;border:1px solid rgba(167,139,250,.3)}}
  .b-res{{background:rgba(127,255,107,.1);color:#7fff6b;border:1px solid rgba(127,255,107,.3)}}
  .b-cerr{{background:rgba(138,155,192,.1);color:#8a9bc0;border:1px solid rgba(138,155,192,.3)}}
  .b-esc{{background:rgba(0,212,255,.1);color:#00d4ff;border:1px solid rgba(0,212,255,.3)}}
  .b-can{{background:rgba(50,50,60,.5);color:#555e7a;border:1px solid rgba(85,94,122,.3)}}

  /* TOOLTIP */
  .tip{{position:relative}}
  .tip:hover .tipbox{{display:block}}
  .tipbox{{display:none;position:absolute;bottom:110%;left:0;background:#1e2a40;
    border:1px solid var(--border);color:var(--text);font-size:.73rem;padding:8px 12px;
    border-radius:6px;max-width:320px;white-space:normal;z-index:200;
    box-shadow:0 8px 24px rgba(0,0,0,.5);line-height:1.5;min-width:140px}}

  /* PAGINATION */
  .pag{{display:flex;align-items:center;justify-content:center;gap:5px;
    padding:14px;border-top:1px solid var(--border);flex-wrap:wrap}}
  .pb{{background:var(--surface2);border:1px solid var(--border);color:var(--text2);
    font-family:var(--mono);font-size:.72rem;width:30px;height:30px;border-radius:6px;
    cursor:pointer;display:flex;align-items:center;justify-content:center;transition:all .15s}}
  .pb:hover{{border-color:var(--accent);color:var(--accent)}}
  .pb.active{{background:var(--accent);color:var(--bg);border-color:var(--accent);font-weight:700}}
  .pb:disabled{{opacity:.3;cursor:not-allowed}}
  .pi{{font-family:var(--mono);font-size:.68rem;color:var(--text2);margin:0 6px}}

  .empty{{text-align:center;padding:40px;color:var(--text2);font-family:var(--mono);font-size:.82rem}}
  .empty-icon{{font-size:2.2rem;margin-bottom:10px;opacity:.4}}
</style>
</head>
<body>

<div class="header">
  <div class="logo">TCO<span>ONNET FIBRA — DASHBOARD</span></div>
  <div class="hbadge">v2.0</div>
  <div class="update-info">Última actualización: <strong>{update_ts}</strong></div>
</div>

<div class="container">

  <div class="metrics" id="metricsGrid"></div>

  <div class="panels">
    <div class="panel">
      <div class="ph"><div class="ph-dot"></div>
        <span class="ph-title"><strong>COSTA ATLÁNTICA</strong> — CASOS POR CIUDAD</span>
      </div>
      <div class="costa-list" id="costaList"></div>
    </div>

    <div class="panel">
      <div class="ph"><div class="ph-dot orange"></div>
        <span class="ph-title"><strong>BÚSQUEDA</strong> — GESTIÓN DE TCO</span>
      </div>
      <div class="sp">
        <div class="srow">
          <div class="siw">
            <span class="sicon">🔍</span>
            <input type="text" class="sinput" id="q" placeholder="Buscar TCO, descripción, ciudad...">
          </div>
          <select class="fsel" id="fEst">
            <option value="">Todos los estados</option>
            <option>Abierto</option><option>Nuevo</option>
            <option>En Progreso</option><option>Resuelto</option>
            <option>Cerrado</option><option value="Escalado a fábrica">Escalado</option>
            <option>Cancelado</option>
          </select>
          <select class="fsel" id="fLoc"><option value="">Todas las ciudades</option></select>
          <button class="breset" onclick="reset()">↺ RESET</button>
        </div>
        <div class="rinfo" id="rinfo"></div>
      </div>
      <div class="tw">
        <table>
          <thead><tr>
            <th onclick="sort('tco')" id="th-tco">N° TCO <span id="ar-tco"></span></th>
            <th onclick="sort('fecha_creacion')" id="th-fecha_creacion">Fecha Creación <span id="ar-fecha_creacion"></span></th>
            <th onclick="sort('estado')" id="th-estado">Estado <span id="ar-estado"></span></th>
            <th onclick="sort('fecha_respuesta')" id="th-fecha_respuesta">Fecha Respuesta <span id="ar-fecha_respuesta"></span></th>
            <th>Nota / Resultado</th>
            <th>Descripción</th>
          </tr></thead>
          <tbody id="tbody"></tbody>
        </table>
        <div id="empty" class="empty" style="display:none">
          <div class="empty-icon">🔎</div><div>Sin resultados</div>
        </div>
      </div>
      <div class="pag" id="pag"></div>
    </div>
  </div>
</div>

<script>
const D={data_json};
const COSTA=['BARRANQUILLA','CARTAGENA','SOLEDAD','VALLEDUPAR','SANTA MARTA','MONTERIA','SINCELEJO','RIOHACHA','MAICAO'];
const PEND=['Abierto','Nuevo','En Progreso','Escalado a fábrica'];
const PAGE=25;
let fil=[...D],pg=1,sk='fecha_creacion',sd=-1,ac=null;

const BC={{
  'Abierto':'b-abierto','Nuevo':'b-nuevo','En Progreso':'b-prog',
  'Resuelto':'b-res','Cerrado':'b-cerr','Escalado a fábrica':'b-esc','Cancelado':'b-can'
}};

function buildMetrics(){{
  const total=D.length;
  const p=D.filter(r=>PEND.includes(r.estado)).length;
  const cnt=s=>D.filter(r=>r.estado===s).length;
  const cards=[
    {{c:'mc-total',l:'TOTAL TCO',v:total}},
    {{c:'mc-pend', l:'PENDIENTES',v:p,s:'Abierto+Nuevo+Progreso+Esc.'}},
    {{c:'mc-abierto',l:'ABIERTOS',v:cnt('Abierto')}},
    {{c:'mc-nuevo',l:'NUEVOS',v:cnt('Nuevo')}},
    {{c:'mc-prog',l:'EN PROGRESO',v:cnt('En Progreso')}},
    {{c:'mc-res',l:'RESUELTOS',v:cnt('Resuelto')}},
    {{c:'mc-cerr',l:'CERRADOS',v:cnt('Cerrado')}},
  ];
  document.getElementById('metricsGrid').innerHTML=cards.map(c=>
    `<div class="mc ${{c.c}}"><div class="mc-label">${{c.l}}</div>
     <div class="mc-val">${{c.v}}</div>
     ${{c.s?`<div class="mc-sub">${{c.s}}</div>`:''}}</div>`).join('');
}}

function buildCosta(){{
  const cities=COSTA.filter(c=>D.some(r=>r.localidad===c));
  const max=Math.max(...cities.map(c=>D.filter(r=>r.localidad===c).length));
  document.getElementById('costaList').innerHTML=cities.map(c=>{{
    const tot=D.filter(r=>r.localidad===c).length;
    const pen=D.filter(r=>r.localidad===c&&PEND.includes(r.estado)).length;
    const pct=Math.round(tot/max*100);
    return `<div class="citem" onclick="byCosta('${{c}}')" id="ci-${{c}}">
      <div class="cc">${{c}}</div>
      <div class="cbar"><div class="cbar-bg"><div class="cbar-fill" style="width:${{pct}}%"></div></div></div>
      <div class="ccount">${{tot}}</div>
      ${{pen>0?`<div class="cpend">${{pen}} pend.</div>`:''}}
    </div>`;
  }}).join('');
}}

function populateCities(){{
  const cities=[...new Set(D.map(r=>r.localidad).filter(Boolean))].sort();
  const sel=document.getElementById('fLoc');
  cities.forEach(c=>{{const o=document.createElement('option');o.value=o.textContent=c;sel.appendChild(o);}});
}}

function byCosta(c){{
  if(ac===c){{ac=null;document.querySelectorAll('.citem').forEach(e=>e.classList.remove('active'));document.getElementById('fLoc').value='';}}
  else{{ac=c;document.querySelectorAll('.citem').forEach(e=>e.classList.remove('active'));document.getElementById('ci-'+c).classList.add('active');document.getElementById('fLoc').value=c;}}
  applyFilters();
}}

function applyFilters(){{
  const q=document.getElementById('q').value.trim().toLowerCase();
  const est=document.getElementById('fEst').value;
  const loc=document.getElementById('fLoc').value;
  fil=D.filter(r=>{{
    if(est&&r.estado!==est)return false;
    if(loc&&r.localidad!==loc)return false;
    if(q&&!(r.tco+r.descripcion+r.localidad+r.nota_resultado+r.estado).toLowerCase().includes(q))return false;
    return true;
  }});
  sort(sk,true);
}}

let dt;
function di(){{clearTimeout(dt);dt=setTimeout(applyFilters,250);}}

function sort(k,keep){{
  if(!keep){{sk===k?sd*=-1:(sk=k,sd=-1);}}
  document.querySelectorAll('thead th').forEach(t=>{{t.classList.remove('sorted');}});
  const th=document.getElementById('th-'+k);
  if(th){{th.classList.add('sorted');document.getElementById('ar-'+k).textContent=sd===1?'▲':'▼';}}
  ['tco','fecha_creacion','estado','fecha_respuesta'].forEach(x=>{{if(x!==k)document.getElementById('ar-'+x).textContent='';}});
  fil.sort((a,b)=>{{let va=a[k]||'',vb=b[k]||'';return va<vb?-sd:va>vb?sd:0;}});
  pg=1;render();
}}

function render(){{
  const tot=fil.length,pages=Math.max(1,Math.ceil(tot/PAGE));
  pg=Math.min(pg,pages);
  const sl=fil.slice((pg-1)*PAGE,pg*PAGE);
  document.getElementById('rinfo').innerHTML=`Mostrando <strong>${{sl.length}}</strong> de <strong>${{tot}}</strong> registros`;
  const tb=document.getElementById('tbody');
  const em=document.getElementById('empty');
  if(!sl.length){{tb.innerHTML='';em.style.display='';document.getElementById('pag').innerHTML='';return;}}
  em.style.display='none';
  tb.innerHTML=sl.map(r=>`<tr>
    <td><span class="tco-n">${{r.tco}}</span></td>
    <td><span class="dc">${{r.fecha_creacion||'—'}}</span></td>
    <td><span class="badge ${{BC[r.estado]||'b-cerr'}}">${{r.estado}}</span></td>
    <td><span class="dc">${{r.fecha_respuesta||'—'}}</span></td>
    <td><div class="tip nota-c">${{r.nota_resultado||'—'}}<div class="tipbox">${{r.nota_resultado||'Sin nota'}}</div></div></td>
    <td><div class="tip desc-c">${{r.descripcion||'—'}}<div class="tipbox">${{r.descripcion||''}}</div></div></td>
  </tr>`).join('');
  renderPag(pages);
}}

function renderPag(pages){{
  if(pages<=1){{document.getElementById('pag').innerHTML='';return;}}
  let h=`<button class="pb" onclick="go(${{pg-1}})" ${{pg===1?'disabled':''}}>‹</button>`;
  const rng=[];
  for(let i=1;i<=pages;i++){{if(i===1||i===pages||Math.abs(i-pg)<=2)rng.push(i);else if(rng[rng.length-1]!=='…')rng.push('…');}}
  rng.forEach(p=>{{h+=p==='…'?`<span class="pi">…</span>`:`<button class="pb ${{p===pg?'active':''}}" onclick="go(${{p}})">${{p}}</button>`;}});
  h+=`<button class="pb" onclick="go(${{pg+1}})" ${{pg===pages?'disabled':''}}>›</button>`;
  h+=`<span class="pi">${{pg}}/${{pages}}</span>`;
  document.getElementById('pag').innerHTML=h;
}}

function go(p){{pg=p;render();}}

function reset(){{
  document.getElementById('q').value='';
  document.getElementById('fEst').value='';
  document.getElementById('fLoc').value='';
  ac=null;document.querySelectorAll('.citem').forEach(e=>e.classList.remove('active'));
  fil=[...D];sk='fecha_creacion';sd=-1;applyFilters();
}}

document.getElementById('q').addEventListener('input',di);
document.getElementById('fEst').addEventListener('change',applyFilters);
document.getElementById('fLoc').addEventListener('change',()=>{{ac=null;document.querySelectorAll('.citem').forEach(e=>e.classList.remove('active'));applyFilters();}});

buildMetrics();buildCosta();populateCities();applyFilters();
</script>
</body>
</html>"""

# ── 4. GUARDAR HTML ──────────────────────────────────────────────────────────
os.makedirs("docs", exist_ok=True)
out = "docs/index.html"
with open(out, "w", encoding="utf-8") as f:
    f.write(html)

size = os.path.getsize(out)
print(f"✅ HTML generado: {out} ({size//1024} KB)")
print(f"✅ Timestamp: {update_ts}")
