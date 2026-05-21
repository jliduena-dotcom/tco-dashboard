"""
generar_dashboard.py  v2.1
Lee el Excel de TCO y genera docs/index.html
Ejecutado automáticamente por GitHub Actions.
"""

import pandas as pd, json, os
from datetime import datetime

# ── 1. LEER EXCEL ─────────────────────────────────────────────────────────────
EXCEL_PATH = "data/Resultado_respuesta_ONNET_SN.xls"
ext = os.path.splitext(EXCEL_PATH)[1].lower()
engine = "xlrd" if ext == ".xls" else "openpyxl"
df = pd.read_excel(EXCEL_PATH, engine=engine)
print(f"✅ Excel leído: {len(df)} registros")

# ── 2. PREPARAR DATOS ─────────────────────────────────────────────────────────
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
    "Tipo de cierre [Tiquete Comercial]":"tipo_cierre",
    "Causa":                             "causa",
}
df2 = df[[c for c in cols_map if c in df.columns]].rename(columns=cols_map)
for col in ["fecha_creacion","fecha_respuesta"]:
    if col in df2.columns:
        df2[col] = pd.to_datetime(df2[col], errors="coerce").dt.strftime("%Y-%m-%d %H:%M")
df2 = df2.fillna("")
DATA_JSON = json.dumps(df2.to_dict(orient="records"), ensure_ascii=False, separators=(",",":"))
update_ts = datetime.now().strftime("%d/%m/%Y %H:%M")
print(f"✅ Datos: {len(DATA_JSON)//1024} KB")

# ── 3. GENERAR HTML ───────────────────────────────────────────────────────────
html = f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>Seguimiento Escalamiento Red Neutra - Costa</title>
<link href="https://fonts.googleapis.com/css2?family=Space+Mono:wght@400;700&family=DM+Sans:wght@300;400;500;700&display=swap" rel="stylesheet">
<style>
:root{{
  --bg:#0b0f1a;--surface:#131929;--surface2:#1a2235;--border:#243050;
  --accent:#00d4ff;--accent2:#ff6b35;--accent3:#7fff6b;
  --warn:#ffb547;--danger:#ff4d6d;--purple:#a78bfa;
  --text:#e8edf7;--text2:#8a9bc0;
  --mono:'Space Mono',monospace;--body:'DM Sans',sans-serif;
}}
*{{box-sizing:border-box;margin:0;padding:0}}
body{{background:var(--bg);color:var(--text);font-family:var(--body);min-height:100vh}}
.header{{background:var(--surface);border-bottom:2px solid var(--accent);padding:14px 28px;
  display:flex;align-items:center;gap:14px;position:sticky;top:0;z-index:200;
  box-shadow:0 4px 32px rgba(0,212,255,.15)}}
.logo-title{{font-family:var(--mono);font-size:.72rem;color:var(--accent);letter-spacing:2px;font-weight:700;text-transform:uppercase}}
.logo-sub{{font-size:1rem;font-weight:700;color:var(--text);margin-top:2px}}
.hbadge{{background:rgba(0,212,255,.1);border:1px solid var(--accent);color:var(--accent);
  font-family:var(--mono);font-size:.62rem;padding:4px 10px;border-radius:4px;letter-spacing:1px}}
.update-info{{margin-left:auto;font-family:var(--mono);font-size:.62rem;color:var(--text2);text-align:right}}
.update-info strong{{color:var(--accent3)}}
.seg-bar{{background:var(--surface);border-bottom:1px solid var(--border);padding:10px 28px;
  display:flex;align-items:center;gap:8px;flex-wrap:wrap}}
.seg-label{{font-family:var(--mono);font-size:.63rem;color:var(--text2);letter-spacing:1px;text-transform:uppercase;margin-right:2px}}
.seg-btn{{font-family:var(--mono);font-size:.7rem;padding:6px 14px;border-radius:20px;
  border:1px solid var(--border);background:transparent;color:var(--text2);cursor:pointer;
  transition:all .2s;white-space:nowrap}}
.seg-btn:hover{{border-color:var(--accent);color:var(--accent)}}
.seg-btn.active{{background:var(--accent);color:var(--bg);border-color:var(--accent);font-weight:700}}
.seg-btn.active-costa{{background:linear-gradient(90deg,#00d4ff,#0099bb);color:var(--bg);border-color:var(--accent)}}
.seg-sep{{width:1px;height:22px;background:var(--border);margin:0 4px}}
#seg-cities{{display:none;flex-wrap:wrap;gap:6px}}
.container{{max-width:1400px;margin:0 auto;padding:20px 24px}}
.metrics{{display:grid;grid-template-columns:repeat(auto-fit,minmax(145px,1fr));gap:12px;margin-bottom:20px}}
.mc{{background:var(--surface);border:1px solid var(--border);border-radius:10px;padding:15px 14px;
  position:relative;overflow:hidden;transition:transform .2s}}
.mc:hover{{transform:translateY(-2px)}}
.mc::before{{content:'';position:absolute;top:0;left:0;right:0;height:3px}}
.mc-total::before{{background:var(--accent)}}.mc-pend::before{{background:var(--accent2)}}
.mc-abierto::before{{background:var(--danger)}}.mc-nuevo::before{{background:var(--warn)}}
.mc-prog::before{{background:var(--purple)}}.mc-res::before{{background:var(--accent3)}}
.mc-cerr::before{{background:var(--text2)}}
.mc-label{{font-family:var(--mono);font-size:.6rem;color:var(--text2);letter-spacing:1px;text-transform:uppercase;margin-bottom:6px}}
.mc-val{{font-family:var(--mono);font-size:1.85rem;font-weight:700;line-height:1}}
.mc-total .mc-val{{color:var(--accent)}}.mc-pend .mc-val{{color:var(--accent2)}}
.mc-abierto .mc-val{{color:var(--danger)}}.mc-nuevo .mc-val{{color:var(--warn)}}
.mc-prog .mc-val{{color:var(--purple)}}.mc-res .mc-val{{color:var(--accent3)}}
.mc-cerr .mc-val{{color:var(--text2)}}
.mc-sub{{font-size:.62rem;color:var(--text2);margin-top:4px}}
.chart-row{{display:grid;grid-template-columns:1fr 1fr;gap:18px;margin-bottom:20px}}
@media(max-width:800px){{.chart-row{{grid-template-columns:1fr}}}}
.panels{{display:grid;grid-template-columns:300px 1fr;gap:18px;margin-bottom:20px}}
@media(max-width:900px){{.panels{{grid-template-columns:1fr}}}}
.panel,.chart-panel{{background:var(--surface);border:1px solid var(--border);border-radius:10px;overflow:hidden}}
.ph{{padding:11px 16px;border-bottom:1px solid var(--border);background:var(--surface2);display:flex;align-items:center;gap:8px}}
.ph-dot{{width:8px;height:8px;border-radius:50%;background:var(--accent);box-shadow:0 0 8px var(--accent);animation:pulse 2s infinite}}
.ph-dot.o{{background:var(--accent2);box-shadow:0 0 8px var(--accent2)}}
@keyframes pulse{{0%,100%{{opacity:1}}50%{{opacity:.4}}}}
.ph-title{{font-family:var(--mono);font-size:.67rem;letter-spacing:1.5px;text-transform:uppercase;color:var(--text2)}}
.ph-title strong{{color:var(--accent)}}
.chart-body{{padding:20px;display:flex;align-items:center;justify-content:center;min-height:190px}}
.donut-wrap{{display:flex;align-items:center;gap:24px;width:100%}}
.donut-legend{{display:flex;flex-direction:column;gap:12px}}
.dl-item{{display:flex;align-items:center;gap:10px}}
.dl-dot{{width:12px;height:12px;border-radius:50%;flex-shrink:0}}
.dl-label{{font-size:.8rem;color:var(--text2)}}
.dl-val{{font-family:var(--mono);font-size:1rem;font-weight:700;color:var(--text)}}
.dl-pct{{font-family:var(--mono);font-size:.68rem;color:var(--text2)}}
.causa-list{{width:100%;display:flex;flex-direction:column;gap:8px}}
.causa-name{{font-size:.73rem;color:var(--text2);margin-bottom:3px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}}
.causa-row{{display:flex;align-items:center;gap:8px}}
.causa-bar-bg{{flex:1;background:var(--border);height:6px;border-radius:3px;overflow:hidden}}
.causa-bar-fill{{height:100%;border-radius:3px;background:var(--accent);transition:width .6s}}
.causa-count{{font-family:var(--mono);font-size:.7rem;color:var(--accent);min-width:28px;text-align:right}}
.costa-list{{padding:10px}}
.citem{{display:flex;align-items:center;padding:9px 10px;border-radius:7px;margin-bottom:5px;cursor:pointer;transition:background .15s;border:1px solid transparent}}
.citem:hover{{background:var(--surface2);border-color:var(--border)}}
.citem.active{{background:rgba(0,212,255,.08);border-color:rgba(0,212,255,.3)}}
.cc{{font-weight:600;font-size:.85rem;flex:1}}
.cbar{{flex:2;margin:0 10px}}
.cbar-bg{{background:var(--border);height:5px;border-radius:3px;overflow:hidden}}
.cbar-fill{{height:100%;border-radius:3px;background:var(--accent)}}
.ccount{{font-family:var(--mono);font-size:.8rem;color:var(--accent);font-weight:700;min-width:28px;text-align:right}}
.cpend{{font-family:var(--mono);font-size:.6rem;padding:2px 6px;border-radius:10px;background:rgba(255,77,109,.15);color:var(--danger);border:1px solid rgba(255,77,109,.3);margin-left:5px}}
.sp{{padding:12px}}
.srow{{display:flex;gap:8px;margin-bottom:10px;flex-wrap:wrap}}
.siw{{position:relative;flex:1;min-width:170px}}
.sicon{{position:absolute;left:10px;top:50%;transform:translateY(-50%);color:var(--text2)}}
.sinput{{width:100%;background:var(--bg);border:1px solid var(--border);color:var(--text);font-family:var(--mono);font-size:.79rem;padding:9px 12px 9px 30px;border-radius:7px;outline:none;transition:border .2s}}
.sinput:focus{{border-color:var(--accent)}}
.sinput::placeholder{{color:var(--text2)}}
select.fsel{{background:var(--bg);border:1px solid var(--border);color:var(--text);font-family:var(--body);font-size:.79rem;padding:9px 10px;border-radius:7px;outline:none;cursor:pointer}}
select.fsel:focus{{border-color:var(--accent)}}
.breset{{background:rgba(255,107,53,.12);border:1px solid rgba(255,107,53,.35);color:var(--accent2);font-family:var(--mono);font-size:.67rem;padding:9px 12px;border-radius:7px;cursor:pointer;letter-spacing:1px}}
.breset:hover{{background:rgba(255,107,53,.22)}}
.rinfo{{font-family:var(--mono);font-size:.64rem;color:var(--text2);margin-bottom:8px}}
.rinfo strong{{color:var(--accent)}}
.tw{{overflow-x:auto}}
table{{width:100%;border-collapse:collapse;font-size:.79rem}}
thead th{{background:var(--surface2);padding:9px 12px;text-align:left;font-family:var(--mono);font-size:.62rem;letter-spacing:1px;color:var(--text2);text-transform:uppercase;border-bottom:1px solid var(--border);white-space:nowrap;cursor:pointer;user-select:none}}
thead th:hover{{color:var(--accent)}}
thead th.sorted{{color:var(--accent)}}
tbody tr{{border-bottom:1px solid rgba(36,48,80,.5);transition:background .1s}}
tbody tr:hover{{background:rgba(255,255,255,.03)}}
tbody td{{padding:9px 12px;vertical-align:top}}
.tco-n{{font-family:var(--mono);font-size:.72rem;color:var(--accent);white-space:nowrap}}
.dc{{font-family:var(--mono);font-size:.68rem;color:var(--text2);white-space:nowrap}}
.exp-cell{{font-size:.77rem;color:var(--text2);line-height:1.45;max-width:250px}}
.exp-cell.nota{{font-style:italic}}
.exp-full{{display:none;margin-top:4px;color:var(--text);white-space:pre-wrap;word-break:break-word}}
.exp-btn{{display:inline-block;margin-top:4px;font-family:var(--mono);font-size:.58rem;color:var(--accent);cursor:pointer;border:none;background:none;padding:0}}
.exp-btn:hover{{text-decoration:underline}}
.badge{{display:inline-flex;align-items:center;gap:4px;font-family:var(--mono);font-size:.62rem;padding:3px 8px;border-radius:20px;white-space:nowrap;font-weight:700}}
.badge::before{{content:'●';font-size:.44rem}}
.b-abierto{{background:rgba(255,77,109,.12);color:#ff4d6d;border:1px solid rgba(255,77,109,.3)}}
.b-nuevo{{background:rgba(255,181,71,.12);color:#ffb547;border:1px solid rgba(255,181,71,.3)}}
.b-prog{{background:rgba(167,139,250,.12);color:#a78bfa;border:1px solid rgba(167,139,250,.3)}}
.b-res{{background:rgba(127,255,107,.1);color:#7fff6b;border:1px solid rgba(127,255,107,.3)}}
.b-cerr{{background:rgba(138,155,192,.1);color:#8a9bc0;border:1px solid rgba(138,155,192,.3)}}
.b-esc{{background:rgba(0,212,255,.1);color:#00d4ff;border:1px solid rgba(0,212,255,.3)}}
.b-can{{background:rgba(50,50,60,.5);color:#555e7a;border:1px solid rgba(85,94,122,.3)}}
.b-pos{{background:rgba(127,255,107,.12);color:#7fff6b;border:1px solid rgba(127,255,107,.35)}}
.b-neg{{background:rgba(255,77,109,.12);color:#ff4d6d;border:1px solid rgba(255,77,109,.3)}}
.pag{{display:flex;align-items:center;justify-content:center;gap:5px;padding:12px;border-top:1px solid var(--border);flex-wrap:wrap}}
.pb{{background:var(--surface2);border:1px solid var(--border);color:var(--text2);font-family:var(--mono);font-size:.68rem;width:28px;height:28px;border-radius:6px;cursor:pointer;display:flex;align-items:center;justify-content:center;transition:all .15s}}
.pb:hover{{border-color:var(--accent);color:var(--accent)}}
.pb.active{{background:var(--accent);color:var(--bg);border-color:var(--accent);font-weight:700}}
.pb:disabled{{opacity:.3;cursor:not-allowed}}
.pi{{font-family:var(--mono);font-size:.65rem;color:var(--text2);margin:0 4px}}
.empty{{text-align:center;padding:36px;color:var(--text2);font-family:var(--mono);font-size:.8rem}}
.empty-icon{{font-size:2rem;margin-bottom:8px;opacity:.4}}
</style>
</head>
<body>

<div class="header">
  <div>
    <div class="logo-title">TCO · ONNET FIBRA</div>
    <div class="logo-sub">Seguimiento Escalamiento Red Neutra — Costa</div>
  </div>
  <div class="hbadge">v2.1</div>
  <div class="update-info">Última actualización:<br><strong>{update_ts}</strong></div>
</div>

<div class="seg-bar">
  <span class="seg-label">📍 Zona:</span>
  <button class="seg-btn active" id="seg-todas" onclick="segmentar('todas')">Todas las ciudades</button>
  <button class="seg-btn" id="seg-costa" onclick="segmentar('costa')">🌊 Costa Atlántica</button>
  <div class="seg-sep"></div>
  <span class="seg-label" id="seg-cities-label" style="display:none">Ciudad:</span>
  <div id="seg-cities"></div>
</div>

<div class="container">
  <div class="metrics" id="metricsGrid"></div>

  <div class="chart-row">
    <div class="chart-panel">
      <div class="ph"><div class="ph-dot"></div>
        <span class="ph-title"><strong>RESULTADO</strong> — Positivo vs Negativo</span>
      </div>
      <div class="chart-body"><div class="donut-wrap" id="donutWrap"></div></div>
    </div>
    <div class="chart-panel">
      <div class="ph"><div class="ph-dot o"></div>
        <span class="ph-title"><strong>TOP CAUSAS</strong> — Principales motivos</span>
      </div>
      <div class="chart-body" style="align-items:flex-start">
        <div class="causa-list" id="causaList"></div>
      </div>
    </div>
  </div>

  <div class="panels">
    <div class="panel">
      <div class="ph"><div class="ph-dot"></div>
        <span class="ph-title"><strong>COSTA ATLÁNTICA</strong> — Por ciudad</span>
      </div>
      <div class="costa-list" id="costaList"></div>
    </div>
    <div class="panel">
      <div class="ph"><div class="ph-dot o"></div>
        <span class="ph-title"><strong>BÚSQUEDA</strong> — Gestión de TCO</span>
      </div>
      <div class="sp">
        <div class="srow">
          <div class="siw">
            <span class="sicon">🔍</span>
            <input type="text" class="sinput" id="q" placeholder="Buscar TCO, descripción, ciudad...">
          </div>
          <select class="fsel" id="fEst">
            <option value="">Todos los estados</option>
            <option>Abierto</option><option>Nuevo</option><option>En Progreso</option>
            <option>Resuelto</option><option>Cerrado</option>
            <option value="Escalado a fábrica">Escalado</option><option>Cancelado</option>
          </select>
          <select class="fsel" id="fCierre">
            <option value="">Tipo cierre</option>
            <option value="Positivo">✅ Positivo</option>
            <option value="Negativo">❌ Negativo</option>
          </select>
          <button class="breset" onclick="resetAll()">↺ RESET</button>
        </div>
        <div class="rinfo" id="rinfo"></div>
      </div>
      <div class="tw">
        <table>
          <thead><tr>
            <th onclick="srt('tco')" id="th-tco">N° TCO <span id="ar-tco"></span></th>
            <th onclick="srt('fecha_creacion')" id="th-fecha_creacion">Creación <span id="ar-fecha_creacion"></span></th>
            <th onclick="srt('estado')" id="th-estado">Estado <span id="ar-estado"></span></th>
            <th>Cierre</th>
            <th onclick="srt('fecha_respuesta')" id="th-fecha_respuesta">F. Resp. <span id="ar-fecha_respuesta"></span></th>
            <th>Nota / Resultado</th>
            <th>Descripción</th>
          </tr></thead>
          <tbody id="tbody"></tbody>
        </table>
        <div id="empty" class="empty" style="display:none"><div class="empty-icon">🔎</div><div>Sin resultados</div></div>
      </div>
      <div class="pag" id="pag"></div>
    </div>
  </div>
</div>

<script>
const D={DATA_JSON};
const COSTA=['BARRANQUILLA','CARTAGENA','SOLEDAD','VALLEDUPAR','SANTA MARTA','MONTERIA','SINCELEJO','RIOHACHA','MAICAO'];
const PEND=['Abierto','Nuevo','En Progreso','Escalado a fábrica'];
const BC={{'Abierto':'b-abierto','Nuevo':'b-nuevo','En Progreso':'b-prog','Resuelto':'b-res','Cerrado':'b-cerr','Escalado a fábrica':'b-esc','Cancelado':'b-can'}};
const PAGE=25;
let zona='todas',fil=[...D],pg=1,sk='fecha_creacion',sd=-1;

function segmentar(z){{
  zona=z;
  document.querySelectorAll('.seg-btn').forEach(b=>b.classList.remove('active','active-costa'));
  if(z==='todas')document.getElementById('seg-todas').classList.add('active');
  else if(z==='costa')document.getElementById('seg-costa').classList.add('active','active-costa');
  else{{document.getElementById('seg-costa').classList.add('active-costa');
    document.querySelectorAll('.seg-city-btn').forEach(b=>b.classList.toggle('active',b.dataset.city===z));}}
  const showC=z==='costa'||COSTA.includes(z);
  document.getElementById('seg-cities-label').style.display=showC?'':'none';
  document.getElementById('seg-cities').style.display=showC?'flex':'none';
  applyFilters();
}}

function buildSegCities(){{
  const cities=COSTA.filter(c=>D.some(r=>r.localidad===c));
  const w=document.getElementById('seg-cities');
  w.style.flexWrap='wrap';w.style.gap='6px';
  w.innerHTML=cities.map(c=>`<button class="seg-btn seg-city-btn" data-city="${{c}}" onclick="segmentar('${{c}}')">${{c}}</button>`).join('');
}}

function getBase(){{
  if(zona==='todas')return D;
  if(zona==='costa')return D.filter(r=>COSTA.includes(r.localidad));
  return D.filter(r=>r.localidad===zona);
}}

function buildMetrics(data){{
  const cnt=s=>data.filter(r=>r.estado===s).length;
  const cards=[
    {{c:'mc-total',l:'TOTAL TCO',v:data.length}},
    {{c:'mc-pend',l:'PENDIENTES',v:data.filter(r=>PEND.includes(r.estado)).length,s:'Abierto+Nuevo+Progreso+Esc.'}},
    {{c:'mc-abierto',l:'ABIERTOS',v:cnt('Abierto')}},
    {{c:'mc-nuevo',l:'NUEVOS',v:cnt('Nuevo')}},
    {{c:'mc-prog',l:'EN PROGRESO',v:cnt('En Progreso')}},
    {{c:'mc-res',l:'RESUELTOS',v:cnt('Resuelto')}},
    {{c:'mc-cerr',l:'CERRADOS',v:cnt('Cerrado')}},
  ];
  document.getElementById('metricsGrid').innerHTML=cards.map(c=>
    `<div class="mc ${{c.c}}"><div class="mc-label">${{c.l}}</div><div class="mc-val">${{c.v}}</div>${{c.s?`<div class="mc-sub">${{c.s}}</div>`:''}}</div>`).join('');
}}

function buildDonut(data){{
  const pos=data.filter(r=>r.tipo_cierre==='Positivo').length;
  const neg=data.filter(r=>r.tipo_cierre==='Negativo').length;
  const tot=pos+neg;
  if(!tot){{document.getElementById('donutWrap').innerHTML='<span style="color:var(--text2);font-family:var(--mono);font-size:.8rem">Sin datos de cierre</span>';return;}}
  const R=68,cx=88,cy=88,circ=2*Math.PI*R;
  const dp=circ*(pos/tot),dn=circ*(neg/tot);
  document.getElementById('donutWrap').innerHTML=`
    <svg width="176" height="176" viewBox="0 0 176 176">
      <circle cx="${{cx}}" cy="${{cy}}" r="${{R}}" fill="none" stroke="var(--border)" stroke-width="20"/>
      <circle cx="${{cx}}" cy="${{cy}}" r="${{R}}" fill="none" stroke="#7fff6b" stroke-width="20"
        stroke-dasharray="${{dp}} ${{circ-dp}}" stroke-dashoffset="${{circ*.25}}" stroke-linecap="round"/>
      <circle cx="${{cx}}" cy="${{cy}}" r="${{R}}" fill="none" stroke="#ff4d6d" stroke-width="20"
        stroke-dasharray="${{dn}} ${{circ-dn}}" stroke-dashoffset="${{circ*.25-dp}}" stroke-linecap="round"/>
      <text x="${{cx}}" y="${{cy-6}}" text-anchor="middle" fill="var(--text)" font-family="Space Mono,monospace" font-size="20" font-weight="700">${{tot}}</text>
      <text x="${{cx}}" y="${{cy+12}}" text-anchor="middle" fill="var(--text2)" font-family="Space Mono,monospace" font-size="9">con cierre</text>
    </svg>
    <div class="donut-legend">
      <div class="dl-item"><div class="dl-dot" style="background:#7fff6b"></div>
        <div><div class="dl-label">Positivo</div><div class="dl-val">${{pos}}</div><div class="dl-pct">${{Math.round(pos/tot*100)}}%</div></div></div>
      <div class="dl-item"><div class="dl-dot" style="background:#ff4d6d"></div>
        <div><div class="dl-label">Negativo</div><div class="dl-val">${{neg}}</div><div class="dl-pct">${{Math.round(neg/tot*100)}}%</div></div></div>
    </div>`;
}}

function buildCausas(data){{
  const cnt={{}};
  data.forEach(r=>{{if(r.causa){{const k=r.causa.trim();cnt[k]=(cnt[k]||0)+1;}}}});
  const top=Object.entries(cnt).sort((a,b)=>b[1]-a[1]).slice(0,6);
  if(!top.length){{document.getElementById('causaList').innerHTML='<span style="color:var(--text2);font-size:.8rem">Sin datos</span>';return;}}
  const max=top[0][1];
  document.getElementById('causaList').innerHTML=top.map(([k,v])=>`
    <div><div class="causa-name" title="${{k}}">${{k}}</div>
    <div class="causa-row"><div class="causa-bar-bg"><div class="causa-bar-fill" style="width:${{Math.round(v/max*100)}}%"></div></div>
    <div class="causa-count">${{v}}</div></div></div>`).join('');
}}

function buildCosta(data){{
  const cities=COSTA.filter(c=>D.some(r=>r.localidad===c));
  const maxAll=Math.max(...cities.map(c=>D.filter(r=>r.localidad===c).length));
  document.getElementById('costaList').innerHTML=cities.map(c=>{{
    const tot=data.filter(r=>r.localidad===c).length;
    const pen=data.filter(r=>r.localidad===c&&PEND.includes(r.estado)).length;
    const pct=Math.round(D.filter(r=>r.localidad===c).length/maxAll*100);
    return `<div class="citem${{zona===c?' active':''}}" onclick="segmentar('${{c}}')" id="ci-${{c}}">
      <div class="cc">${{c}}</div>
      <div class="cbar"><div class="cbar-bg"><div class="cbar-fill" style="width:${{pct}}%"></div></div></div>
      <div class="ccount">${{tot}}</div>
      ${{pen>0?`<div class="cpend">${{pen}} pend.</div>`:''}}
    </div>`;
  }}).join('');
}}

function applyFilters(){{
  const q=document.getElementById('q').value.trim().toLowerCase();
  const est=document.getElementById('fEst').value;
  const cierre=document.getElementById('fCierre').value;
  const base=getBase();
  fil=base.filter(r=>{{
    if(est&&r.estado!==est)return false;
    if(cierre&&r.tipo_cierre!==cierre)return false;
    if(q&&!(r.tco+r.descripcion+r.localidad+r.nota_resultado+r.estado+r.causa).toLowerCase().includes(q))return false;
    return true;
  }});
  buildMetrics(base);buildDonut(base);buildCausas(base);buildCosta(base);
  srt(sk,true);
}}

let dt;
function di(){{clearTimeout(dt);dt=setTimeout(applyFilters,250);}}

function srt(k,keep){{
  if(!keep){{sk===k?sd*=-1:(sk=k,sd=-1);}}
  ['tco','fecha_creacion','estado','fecha_respuesta'].forEach(x=>{{
    const th=document.getElementById('th-'+x);
    if(th){{th.classList.toggle('sorted',x===k);document.getElementById('ar-'+x).textContent=x===k?(sd===1?'▲':'▼'):'';}}
  }});
  fil.sort((a,b)=>{{let va=a[k]||'',vb=b[k]||'';return va<vb?-sd:va>vb?sd:0;}});
  pg=1;render();
}}

function render(){{
  const tot=fil.length,pages=Math.max(1,Math.ceil(tot/PAGE));
  pg=Math.min(pg,pages);
  const sl=fil.slice((pg-1)*PAGE,pg*PAGE);
  document.getElementById('rinfo').innerHTML=`Mostrando <strong>${{sl.length}}</strong> de <strong>${{tot}}</strong> registros`;
  const tb=document.getElementById('tbody'),em=document.getElementById('empty');
  if(!sl.length){{tb.innerHTML='';em.style.display='';document.getElementById('pag').innerHTML='';return;}}
  em.style.display='none';
  tb.innerHTML=sl.map((r,i)=>{{
    const id=`r${{(pg-1)*PAGE+i}}`;
    const dShort=r.descripcion.length>55?r.descripcion.slice(0,55)+'…':r.descripcion;
    const nShort=r.nota_resultado.length>55?r.nota_resultado.slice(0,55)+'…':r.nota_resultado;
    const cb=r.tipo_cierre==='Positivo'?'<span class="badge b-pos">✅ Positivo</span>':r.tipo_cierre==='Negativo'?'<span class="badge b-neg">❌ Negativo</span>':'—';
    return `<tr>
      <td><span class="tco-n">${{r.tco}}</span></td>
      <td><span class="dc">${{r.fecha_creacion||'—'}}</span></td>
      <td><span class="badge ${{BC[r.estado]||'b-cerr'}}">${{r.estado}}</span></td>
      <td>${{cb}}</td>
      <td><span class="dc">${{r.fecha_respuesta||'—'}}</span></td>
      <td><div class="exp-cell nota">${{nShort||'—'}}
        ${{r.nota_resultado.length>55?`<span class="exp-full" id="nf-${{id}}">${{r.nota_resultado}}</span><button class="exp-btn" onclick="tog('nf-${{id}}',this)">[+ ver todo]</button>`:''}}
      </div></td>
      <td><div class="exp-cell">${{dShort||'—'}}
        ${{r.descripcion.length>55?`<span class="exp-full" id="df-${{id}}">${{r.descripcion}}</span><button class="exp-btn" onclick="tog('df-${{id}}',this)">[+ ver todo]</button>`:''}}
      </div></td>
    </tr>`;
  }}).join('');
  renderPag(pages);
}}

function tog(id,btn){{
  const el=document.getElementById(id);
  const open=el.style.display==='block';
  el.style.display=open?'none':'block';
  btn.textContent=open?'[+ ver todo]':'[− ocultar]';
}}

function renderPag(pages){{
  if(pages<=1){{document.getElementById('pag').innerHTML='';return;}}
  let h=`<button class="pb" onclick="go(${{pg-1}})" ${{pg===1?'disabled':''}}>‹</button>`;
  const rng=[];
  for(let i=1;i<=pages;i++){{if(i===1||i===pages||Math.abs(i-pg)<=2)rng.push(i);else if(rng[rng.length-1]!=='…')rng.push('…');}}
  rng.forEach(p=>{{h+=p==='…'?`<span class="pi">…</span>`:`<button class="pb ${{p===pg?'active':''}}" onclick="go(${{p}})">${{p}}</button>`;}} );
  h+=`<button class="pb" onclick="go(${{pg+1}})" ${{pg===pages?'disabled':''}}>›</button><span class="pi">${{pg}}/${{pages}}</span>`;
  document.getElementById('pag').innerHTML=h;
}}

function go(p){{pg=p;render();}}

function resetAll(){{
  zona='todas';
  document.querySelectorAll('.seg-btn').forEach(b=>b.classList.remove('active','active-costa'));
  document.getElementById('seg-todas').classList.add('active');
  document.getElementById('seg-cities-label').style.display='none';
  document.getElementById('seg-cities').style.display='none';
  document.getElementById('q').value='';
  document.getElementById('fEst').value='';
  document.getElementById('fCierre').value='';
  applyFilters();
}}

document.getElementById('q').addEventListener('input',di);
document.getElementById('fEst').addEventListener('change',applyFilters);
document.getElementById('fCierre').addEventListener('change',applyFilters);
buildSegCities();applyFilters();
</script>
</body>
</html>"""

os.makedirs("docs", exist_ok=True)
with open("docs/index.html","w",encoding="utf-8") as f:
    f.write(html)
print(f"✅ HTML guardado: docs/index.html ({os.path.getsize('docs/index.html')//1024} KB)")
