"""
build_final.py — Gera outputs/index.html completo com dados reais
"""
import json, openpyxl, collections, re, os

# ── Carregar D_final_v2.json ──────────────────────────────────────────────────
with open('inputs/D_final_v2.json', encoding='utf-8') as f:
    D = json.load(f)

# ── Valores de referência (aprovados) ────────────────────────────────────────
# C05 dist por UGN (do HTML de referência)
C05_DIST_UGN = {
    'LESTE':    6857,
    'NE':       8077,
    'APOLO RJ': 5769,
    'CO NO':    6923,
    'SPI OESTE':6923,
    'SUL':      6923,
    'APOLO SP': 9231,
}
C04_DIST_UGN = {
    'LESTE':    5000,
    'NE':       8000,
    'APOLO RJ': 4553,
    'CO NO':    7000,
    'SPI OESTE':7000,
    'SUL':      6898,
    'APOLO SP': 11000,
}
# conv por ciclo por UGN [nexus, vital, id]
C04_CONV_UGN = {
    'LESTE':    [258, 307, 0],
    'NE':       [213, 209, 0],
    'APOLO RJ': [134, 75,  0],
    'CO NO':    [205, 229, 0],
    'SPI OESTE':[99,  82,  0],
    'SUL':      [92,  60,  0],
    'APOLO SP': [22,  1,   0],
}
C05_CONV_UGN = {
    'LESTE':    [418, 415, 0],
    'NE':       [423, 533, 0],
    'APOLO RJ': [277, 289, 0],
    'CO NO':    [253, 376, 0],
    'SPI OESTE':[310, 296, 0],
    'SUL':      [194, 198, 0],
    'APOLO SP': [237, 180, 0],
}
# Nexus C05 dist por UGN (aprox 50/50 exceto ID)
NEXUS_C05_DIST = {k: (C05_DIST_UGN[k]-int(C05_DIST_UGN[k]*0.12 if k in ['APOLO SP'] else C05_DIST_UGN[k]*0.13 if k in ['LESTE'] else 0))//2 for k in C05_DIST_UGN}
# Simplify: use 3000 per linha for display in metrics

UGN_ORDER = ['SPI OESTE','LESTE','NE','CO NO','APOLO RJ','SUL','APOLO SP']
UGN_RANKING = [  # sorted by % C05 desc
    ('LESTE',    5000, 565, 11.3, 6857, 833, 12.1),
    ('NE',       8000, 422,  5.3, 8077, 956, 11.8),
    ('APOLO RJ', 4553, 209,  4.6, 5769, 566,  9.8),
    ('CO NO',    7000, 434,  6.2, 6923, 629,  9.1),
    ('SPI OESTE',7000, 181,  2.6, 6923, 606,  8.8),
    ('SUL',      6898, 152,  2.2, 6923, 392,  5.7),
    ('APOLO SP',11000,  23,  0.2, 9231, 417,  4.5),
]

# ── Ler C04 inputs para PDV/Redes ────────────────────────────────────────────
wb4 = openpyxl.load_workbook('inputs/Cupons C04.xlsx', read_only=True)
ws4 = wb4.active
rows4 = [r for r in ws4.iter_rows(values_only=True)]
# cols: Setor(0), Resp(1), SetorFull(2), Cartao(3), Apresentacao(4), Desconto(5), CNPJ(6), NomePDV(7), Data(8), Hora(9), QtdCx(10)
good4 = [r for r in rows4[1:] if r and len(r)>=11 and r[10] is not None]

# Redes: group by NomePDV
redes_data = collections.defaultdict(lambda: {'cx':0,'cnpjs':set(),'setores':set()})
setor_pdv = collections.defaultdict(lambda: collections.defaultdict(lambda: {'cx':0,'cnpj':''}))

for r in good4:
    setor = str(r[0]) if r[0] else ''
    nome = str(r[7]) if r[7] else ''
    cnpj_raw = str(r[6]) if r[6] else ''
    cx = int(r[10])
    if not nome or nome == 'None': continue
    redes_data[nome]['cx'] += cx
    redes_data[nome]['cnpjs'].add(cnpj_raw)
    redes_data[nome]['setores'].add(setor)
    setor_pdv[setor][nome]['cx'] += cx
    if not setor_pdv[setor][nome]['cnpj']:
        setor_pdv[setor][nome]['cnpj'] = cnpj_raw

# Top redes sorted by cx
top_redes = sorted(redes_data.items(), key=lambda x: -x[1]['cx'])[:20]

# Format helpers
def fmt_br(n):
    if n is None: return '—'
    return f'{int(n):,}'.replace(',', '.')

def fmt_pct(p):
    if p is None or p == 0: return '—'
    return f'{p:.1f}%'.replace('.', ',')

def pct_calc(conv, dist):
    if not dist: return 0.0
    return round(conv/dist*100, 1)

# ── Build sectors JS data ─────────────────────────────────────────────────────
ugn_js = {}
for ugn in UGN_ORDER:
    u = D['por_ugn'][ugn]
    setores = u['setores']
    nexus_s = [s for s in setores if s['linha']=='NEXUS']
    vital_s  = [s for s in setores if s['linha']=='VITAL']
    ident_s  = [s for s in setores if s['linha']=='IDENTIDADE']

    c04conv = C04_CONV_UGN[ugn]
    c05conv = C05_CONV_UGN[ugn]
    c05dist_total = C05_DIST_UGN[ugn]
    c04dist_total = C04_DIST_UGN[ugn]

    # Approximate nexus/vital split for C05 dist
    id_count = len(ident_s)
    # ID C05 dist: proportional from total
    id_c05dist = sum(s['dist_c05'] for s in ident_s)
    nv_c05dist = c05dist_total - id_c05dist
    nexus_c05dist = nv_c05dist // 2
    vital_c05dist = nv_c05dist - nexus_c05dist

    def build_setor_list(sl):
        out = []
        for s in sl:
            c04c = s['conv_por_mes'].get('2026-04', 0)
            c05c = s['conv_por_mes'].get('2026-05', 0)
            c04d = s['dist_c04']
            c05d = s['dist_c05']
            pdvs = s.get('pdvs', [])
            # supplement with C04 file data
            sp = setor_pdv.get(s['cd_setor'], {})
            pdv_list = [{'cnpj': v.get('cnpj',''), 'nome': nm, 'cx': v['cx']}
                        for nm, v in sorted(sp.items(), key=lambda x:-x[1]['cx'])[:8]]
            out.append({
                'cd': s['cd_setor'],
                'nome': s['nome_setor'],
                'c04d': c04d, 'c04c': c04c,
                'c05d': c05d, 'c05c': c05c,
                'pdvs': pdv_list
            })
        return out

    ugn_js[ugn] = {
        'c04dist': c04dist_total,
        'c05dist': c05dist_total,
        'c04conv': sum(c04conv),
        'c05conv': sum(c05conv),
        'nexus_c05dist': nexus_c05dist,
        'vital_c05dist': vital_c05dist,
        'id_c05dist': id_c05dist,
        'nexus_c05conv': c05conv[0],
        'vital_c05conv': c05conv[1],
        'nexus_c04conv': c04conv[0],
        'vital_c04conv': c04conv[1],
        'nexus': build_setor_list(nexus_s),
        'vital': build_setor_list(vital_s),
        'identidade': build_setor_list(ident_s),
    }

ugn_js_str = json.dumps(ugn_js, ensure_ascii=False)

# Top redes for Redes page
top_redes_js = [{'nome': nm, 'cx': d['cx'], 'cnpjs': len(d['cnpjs'])}
                for nm, d in top_redes]
top_redes_js_str = json.dumps(top_redes_js, ensure_ascii=False)

# Evolução data from reference
EVO_DATA = {
    'todas':   {'c04conv':[903,1083,0], 'c05conv':[1930,2469,0], 'c04dist':49451, 'c05dist':56000, 'label':'Todas UGNs'},
    'SPI OESTE': {'c04conv':[99,82,0],   'c05conv':[310,296,0],   'c04dist':7000,  'c05dist':6923,  'label':'SPI OESTE', 'mediaUGN':8.8,
                  'distritos':['11030100','11030200','11630100','11630200'],
                  'dLabels':['11030100','11030200','11630100','11630200'],
                  'd_c04':[6.0,5.4,3.3,4.1],'d_c05':[17.1,19.0,17.1,16.6]},
    'LESTE':   {'c04conv':[258,307,0],  'c05conv':[418,415,0],   'c04dist':5000,  'c05dist':6857,  'label':'LESTE',    'mediaUGN':12.1},
    'NE':      {'c04conv':[213,209,0],  'c05conv':[423,533,0],   'c04dist':8000,  'c05dist':8077,  'label':'NE',       'mediaUGN':11.8},
    'CO NO':   {'c04conv':[205,229,0],  'c05conv':[253,376,0],   'c04dist':7000,  'c05dist':6923,  'label':'CO NO',    'mediaUGN':9.1},
    'APOLO RJ':{'c04conv':[134,75,0],   'c05conv':[277,289,0],   'c04dist':4553,  'c05dist':5769,  'label':'APOLO RJ', 'mediaUGN':9.8},
    'SUL':     {'c04conv':[92,60,0],    'c05conv':[194,198,0],   'c04dist':6898,  'c05dist':6923,  'label':'SUL',      'mediaUGN':5.7},
    'APOLO SP':{'c04conv':[22,1,0],     'c05conv':[237,180,0],   'c04dist':11000, 'c05dist':9231,  'label':'APOLO SP', 'mediaUGN':4.5},
}

# ── Generate HTML ─────────────────────────────────────────────────────────────
html_parts = []

def h(s): html_parts.append(s)

h('''<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Olire & Lirux — Vida Mais Leve 2026</title>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/@tabler/icons-webfont@2.47.0/tabler-icons.min.css">
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
<style>
*{box-sizing:border-box;margin:0;padding:0}
:root{
  --ems-blue:#003087;--ems-blue-mid:#0055B3;--ems-blue-light:#E8F0FA;--ems-blue-border:#B8CEEE;
  --nexus:#1B5E20;--nexus-light:#E8F5E9;--nexus-mid:#388E3C;
  --vital:#B71C1C;--vital-light:#FFEBEE;--vital-mid:#D32F2F;
  --id:#7B5200;--id-light:#FFF8E1;--id-mid:#F9A825;--id-border:#FFE082;
  --txt:#212529;--txt2:#6c757d;--txt3:#adb5bd;
  --border:#dee2e6;--bg:#fff;--bg2:#f8f9fa;
  --border-radius-lg:10px;--border-radius-md:6px;--font-sans:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif
}
body{font-family:var(--font-sans);background:var(--bg2);color:var(--txt);line-height:1.5}
.container{max-width:1200px;margin:0 auto;padding:16px}
.dash-header{background:var(--ems-blue);border-radius:var(--border-radius-lg);padding:14px 20px;margin-bottom:1.25rem;display:flex;align-items:center;justify-content:space-between}
.dash-header-left{display:flex;align-items:center;gap:12px}
.dash-logo{width:36px;height:36px;background:#fff;border-radius:6px;display:flex;align-items:center;justify-content:center;font-size:12px;font-weight:700;color:var(--ems-blue)}
.dash-title{color:#fff;font-size:16px;font-weight:500}
.dash-sub{color:rgba(255,255,255,0.65);font-size:12px;margin-top:2px}
.dash-badge{background:rgba(255,255,255,0.15);color:#fff;font-size:11px;padding:4px 10px;border-radius:20px;border:0.5px solid rgba(255,255,255,0.25)}
.nav{display:flex;gap:4px;margin-bottom:1.25rem;flex-wrap:wrap;padding-bottom:1rem;border-bottom:0.5px solid var(--border)}
.nav-btn{padding:6px 13px;font-size:12px;border:0.5px solid var(--ems-blue-border);border-radius:var(--border-radius-md);background:transparent;color:var(--ems-blue-mid);cursor:pointer;display:flex;align-items:center;gap:5px}
.nav-btn:hover{background:var(--ems-blue-light)}
.nav-btn.active{background:var(--ems-blue);color:#fff;border-color:var(--ems-blue)}
.page{display:none}.page.active{display:block}
.metrics{display:grid;grid-template-columns:repeat(auto-fit,minmax(130px,1fr));gap:10px;margin-bottom:1.25rem}
.metric{background:var(--ems-blue-light);border:0.5px solid var(--ems-blue-border);border-radius:var(--border-radius-md);padding:12px 14px}
.metric-label{font-size:11px;color:var(--ems-blue-mid);margin-bottom:4px;font-weight:500}
.metric-value{font-size:22px;font-weight:500;color:var(--ems-blue)}
.metric-sub{font-size:11px;color:var(--txt3);margin-top:2px}
.card{background:var(--bg);border:0.5px solid var(--border);border-radius:var(--border-radius-lg);padding:1rem 1.25rem;margin-bottom:1rem}
.card-title{font-size:11px;font-weight:500;text-transform:uppercase;letter-spacing:0.07em;color:var(--ems-blue-mid);margin-bottom:12px;display:flex;align-items:center;gap:6px}
.badge{display:inline-flex;align-items:center;font-size:11px;font-weight:500;padding:2px 8px;border-radius:20px}
.b-nexus{background:var(--nexus-light);color:var(--nexus)}
.b-vital{background:var(--vital-light);color:var(--vital)}
.b-id{background:var(--id-light);color:var(--id);border:0.5px solid var(--id-border)}
.tw{overflow-x:auto}
table{width:100%;border-collapse:collapse;font-size:13px}
th{font-size:11px;font-weight:500;color:var(--txt2);text-align:left;padding:6px 8px;border-bottom:0.5px solid var(--border);white-space:nowrap;background:var(--bg2)}
td{padding:7px 8px;border-bottom:0.5px solid var(--border);color:var(--txt)}
tr:last-child td{border-bottom:none}
tr.row-nexus td:first-child{border-left:3px solid var(--nexus-mid)}
tr.row-vital td:first-child{border-left:3px solid var(--vital-mid)}
tr.row-id{background:var(--id-light)}
tr.row-id td:first-child{border-left:3px solid var(--id-mid)}
tr.row-total{background:var(--ems-blue-light);font-weight:500}
tr.row-subtotal-n{background:var(--nexus-light);font-weight:500}
tr.row-subtotal-v{background:var(--vital-light);font-weight:500}
tr.row-subtotal-id{background:var(--id-light);font-weight:500}
tr.linha-sep td{background:var(--bg2);font-size:11px;font-weight:500;color:var(--txt2);padding:5px 8px;border-top:1px solid var(--border);border-bottom:0.5px solid var(--border)}
tr.detail-row{display:none;background:var(--bg2)}
tr.detail-row.open{display:table-row}
tr.detail-row td{padding:5px 8px 5px 24px;font-size:12px;color:var(--txt2)}
tr.detail-hdr td{padding:5px 8px 5px 24px;font-size:11px;font-weight:500;background:var(--bg2)}
td.pct-n{color:var(--nexus);font-weight:500}
td.pct-v{color:var(--vital);font-weight:500}
td.pct-id{color:var(--id);font-weight:500}
td.pct-e{color:var(--ems-blue);font-weight:500}
td.zero{color:var(--txt3);font-style:italic}
.bar-wrap{display:flex;align-items:center;gap:7px}
.bar-bg{flex:1;height:6px;background:var(--bg2);border-radius:3px;overflow:hidden}
.bar-fill{height:100%;border-radius:3px}
.bar-n{background:var(--nexus-mid)}.bar-v{background:var(--vital-mid)}.bar-e{background:var(--ems-blue)}.bar-i{background:var(--id-mid)}
.bar-pct{font-size:11px;color:var(--txt2);min-width:36px;text-align:right}
.chips{display:flex;gap:6px;flex-wrap:wrap;margin-bottom:1rem}
.chip{padding:4px 12px;font-size:12px;border:0.5px solid var(--ems-blue-border);border-radius:20px;cursor:pointer;color:var(--ems-blue-mid);background:transparent}
.chip.sel{background:var(--ems-blue);color:#fff;border-color:var(--ems-blue)}
.setor-row{cursor:pointer}
.setor-row:hover td{background:var(--ems-blue-light)}
.setor-detail{display:none;background:var(--bg2)}
.setor-detail.open{display:table-row-group}
.setor-detail td{padding:5px 8px 5px 24px;font-size:12px;color:var(--txt2)}
.two{display:grid;grid-template-columns:1fr 1fr;gap:1rem}
@media(max-width:520px){.two{grid-template-columns:1fr}}
.rank-row{display:flex;align-items:center;gap:10px;padding:8px 0;border-bottom:0.5px solid var(--border)}
.rank-row:last-child{border-bottom:none}
.rank-num{font-size:11px;font-weight:500;color:var(--txt3);min-width:18px;text-align:center}
.rank-info{flex:1;min-width:0}
.rank-name{font-size:13px;color:var(--txt);font-weight:500;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.rank-detail{font-size:11px;color:var(--txt3)}
.rank-val{font-size:14px;font-weight:500;color:var(--ems-blue);text-align:right;min-width:50px}
.cycle-tag{font-size:10px;font-weight:500;padding:1px 7px;border-radius:20px;background:var(--ems-blue-light);color:var(--ems-blue-mid)}
.id-note{background:var(--id-light);border:0.5px solid var(--id-border);border-radius:var(--border-radius-md);padding:8px 12px;font-size:12px;color:var(--id);margin-bottom:1rem;display:flex;align-items:center;gap:8px}
.chart-wrap{position:relative;height:260px}
</style>
</head>
<body>
<div class="container">
''')

# ── HEADER ───────────────────────────────────────────────────────────────────
h('''
<div class="dash-header">
  <div class="dash-header-left">
    <div class="dash-logo">EMS</div>
    <div>
      <div class="dash-title">Olire &amp; Lirux — Vida Mais Leve 2026</div>
      <div class="dash-sub">Conversão de Cupons C04 / C05 · Nexus · Vital · Identidade</div>
    </div>
  </div>
  <div class="dash-badge"><i class="ti ti-refresh" style="font-size:12px"></i> Atualizado: 29/05/2026</div>
</div>

<div class="nav">
  <button class="nav-btn active" onclick="showPage('consolidado',this)"><i class="ti ti-layout-dashboard"></i> Consolidado</button>
  <button class="nav-btn" onclick="showPage('ugn',this)"><i class="ti ti-map-pin"></i> Por UGN</button>
  <button class="nav-btn" onclick="showPage('pdv-ugn',this)"><i class="ti ti-building-store"></i> PDV por Setor</button>
  <button class="nav-btn" onclick="showPage('redes',this)"><i class="ti ti-trophy"></i> Redes Brasil</button>
  <button class="nav-btn" onclick="showPage('evolucao',this)"><i class="ti ti-chart-line"></i> Evolução</button>
</div>
''')

# ── PAGE 1: CONSOLIDADO ───────────────────────────────────────────────────────
h('<div id="page-consolidado" class="page active">')
h('''
  <div class="id-note">
    <i class="ti ti-info-circle" style="font-size:16px"></i>
    <span><strong>Identidade</strong> recebeu 6.000 cupons no C05 · cupons somados ao total de cada UGN · conversão registrada como 0</span>
  </div>
  <div class="metrics">
    <div class="metric"><div class="metric-label"><i class="ti ti-ticket"></i> Cupons C04</div><div class="metric-value">49.451</div><div class="metric-sub">Nexus + Vital · Abril</div></div>
    <div class="metric"><div class="metric-label"><i class="ti ti-ticket"></i> Cupons C05</div><div class="metric-value">56.000</div><div class="metric-sub">Nexus + Vital + ID · Maio</div></div>
    <div class="metric"><div class="metric-label"><i class="ti ti-check"></i> Convertidos C04</div><div class="metric-value">1.986</div><div class="metric-sub">4,0% do racional</div></div>
    <div class="metric"><div class="metric-label"><i class="ti ti-check"></i> Convertidos C05</div><div class="metric-value">4.399</div><div class="metric-sub">7,9% do racional</div></div>
  </div>
  <div class="card">
    <div class="card-title"><i class="ti ti-table"></i> Conversão por ciclo e linha</div>
    <div class="tw"><table>
      <thead><tr><th>Ciclo</th><th>Linha</th><th>Cupons dist.</th><th>Convertidos</th><th>% Conv.</th><th style="min-width:110px">Progresso</th></tr></thead>
      <tbody>
        <tr class="row-nexus"><td><span class="cycle-tag">C04</span></td><td><span class="badge b-nexus">NEXUS</span></td><td>23.645</td><td>903</td><td class="pct-n">3,8%</td><td><div class="bar-wrap"><div class="bar-bg"><div class="bar-fill bar-n" style="width:14%"></div></div><span class="bar-pct">3,8%</span></div></td></tr>
        <tr class="row-vital"><td><span class="cycle-tag">C04</span></td><td><span class="badge b-vital">VITAL</span></td><td>25.806</td><td>1.083</td><td class="pct-v">4,2%</td><td><div class="bar-wrap"><div class="bar-bg"><div class="bar-fill bar-v" style="width:15%"></div></div><span class="bar-pct">4,2%</span></div></td></tr>
        <tr class="row-id"><td><span class="cycle-tag">C04</span></td><td><span class="badge b-id">IDENTIDADE</span></td><td class="zero">—</td><td class="zero">0</td><td class="zero">—</td><td><div class="bar-wrap"><div class="bar-bg"><div class="bar-fill bar-i" style="width:0%"></div></div><span class="bar-pct" style="color:var(--txt3)">—</span></div></td></tr>
        <tr class="row-total"><td><span class="cycle-tag">C04</span></td><td><strong>TOTAL</strong></td><td>49.451</td><td>1.986</td><td class="pct-e">4,0%</td><td><div class="bar-wrap"><div class="bar-bg"><div class="bar-fill bar-e" style="width:15%"></div></div><span class="bar-pct">4,0%</span></div></td></tr>
        <tr class="row-nexus"><td><span class="cycle-tag">C05</span></td><td><span class="badge b-nexus">NEXUS</span></td><td>25.000</td><td>1.930</td><td class="pct-n">7,7%</td><td><div class="bar-wrap"><div class="bar-bg"><div class="bar-fill bar-n" style="width:28%"></div></div><span class="bar-pct">7,7%</span></div></td></tr>
        <tr class="row-vital"><td><span class="cycle-tag">C05</span></td><td><span class="badge b-vital">VITAL</span></td><td>25.000</td><td>2.469</td><td class="pct-v">9,9%</td><td><div class="bar-wrap"><div class="bar-bg"><div class="bar-fill bar-v" style="width:35%"></div></div><span class="bar-pct">9,9%</span></div></td></tr>
        <tr class="row-id"><td><span class="cycle-tag">C05</span></td><td><span class="badge b-id">IDENTIDADE</span></td><td>6.000</td><td class="zero">0</td><td class="zero">0,0%</td><td><div class="bar-wrap"><div class="bar-bg"><div class="bar-fill bar-i" style="width:0%"></div></div><span class="bar-pct" style="color:var(--txt3)">0,0%</span></div></td></tr>
        <tr class="row-total"><td><span class="cycle-tag">C05</span></td><td><strong>TOTAL</strong></td><td>56.000</td><td>4.399</td><td class="pct-e">7,9%</td><td><div class="bar-wrap"><div class="bar-bg"><div class="bar-fill bar-e" style="width:29%"></div></div><span class="bar-pct">7,9%</span></div></td></tr>
      </tbody>
    </table></div>
  </div>
  <div class="card">
    <div class="card-title"><i class="ti ti-map"></i> Ranking UGNs — total por UGN (Nexus + Vital + Identidade)</div>
    <p style="font-size:11px;color:var(--txt3);margin-bottom:10px">C05 dist. inclui cupons Identidade · C05 conv. inclui apenas Nexus + Vital (Identidade = 0 conv.)</p>
    <div class="tw"><table>
      <thead><tr><th>#</th><th>UGN</th><th>C04 dist.</th><th>C04 conv.</th><th>% C04</th><th>C05 dist.</th><th>C05 conv.</th><th>% C05</th><th>Δ pp</th><th style="min-width:80px">C05</th></tr></thead>
      <tbody>
''')

max_pct_c05 = max(r[6] for r in UGN_RANKING)
bar_classes = ['bar-n','bar-n','bar-n','bar-v','bar-v','bar-v','bar-v']
for i, (ugn, c04d, c04c, pct4, c05d, c05c, pct5) in enumerate(UGN_RANKING):
    delta = round(pct5 - pct4, 1)
    delta_str = f'+{delta:.1f}pp'.replace('.',',') if delta >= 0 else f'{delta:.1f}pp'.replace('.',',')
    delta_color = 'var(--nexus)' if delta >= 0 else 'var(--vital)'
    bar_w = int(pct5 / max_pct_c05 * 100)
    bc = bar_classes[i]
    pct4_s = f'{pct4:.1f}%'.replace('.',',')
    pct5_s = f'{pct5:.1f}%'.replace('.',',')
    h(f'        <tr><td>{i+1}</td><td>{ugn}</td><td>{fmt_br(c04d)}</td><td>{fmt_br(c04c)}</td><td>{pct4_s}</td><td>{fmt_br(c05d)}</td><td>{fmt_br(c05c)}</td><td class="pct-n">{pct5_s}</td><td style="color:{delta_color};font-weight:500">{delta_str}</td><td><div class="bar-wrap"><div class="bar-bg"><div class="bar-fill {bc}" style="width:{bar_w}%"></div></div></div></td></tr>')

h('''        <tr class="row-total"><td></td><td><strong>TOTAL GERAL</strong></td><td>49.451</td><td>1.986</td><td class="pct-e">4,0%</td><td>56.000</td><td>4.399</td><td class="pct-e">7,9%</td><td style="color:var(--nexus);font-weight:500">+3,9pp</td><td></td></tr>
      </tbody>
    </table></div>
  </div>
</div>''')

# ── PAGE 2: POR UGN ───────────────────────────────────────────────────────────
h('<div id="page-ugn" class="page">')
h('''
  <div class="chips" id="ugn-chips">
    <button class="chip sel" onclick="selectUGN(this,'SPI OESTE')">SPI OESTE</button>
    <button class="chip" onclick="selectUGN(this,'LESTE')">LESTE</button>
    <button class="chip" onclick="selectUGN(this,'NE')">NE</button>
    <button class="chip" onclick="selectUGN(this,'CO NO')">CO NO</button>
    <button class="chip" onclick="selectUGN(this,'APOLO RJ')">APOLO RJ</button>
    <button class="chip" onclick="selectUGN(this,'SUL')">SUL</button>
    <button class="chip" onclick="selectUGN(this,'APOLO SP')">APOLO SP</button>
  </div>
  <div class="metrics" id="ugn-metrics"></div>
  <div class="card">
    <div class="card-title"><i class="ti ti-list"></i> Setores por linha — clique para expandir CNPJs</div>
    <div class="tw" id="ugn-table-wrap"></div>
  </div>
</div>''')

# ── PAGE 3: PDV POR SETOR ─────────────────────────────────────────────────────
h('<div id="page-pdv-ugn" class="page">')
h('''
  <p style="font-size:13px;color:var(--txt2);margin-bottom:1rem">Ranking de conversões dentro de cada setor (Cupons C04) · clique para expandir · filtre por UGN</p>
  <div class="chips" id="pdv-chips">
    <button class="chip sel" onclick="selectPDVUGN(this,'SPI OESTE')">SPI OESTE</button>
    <button class="chip" onclick="selectPDVUGN(this,'LESTE')">LESTE</button>
    <button class="chip" onclick="selectPDVUGN(this,'NE')">NE</button>
    <button class="chip" onclick="selectPDVUGN(this,'CO NO')">CO NO</button>
    <button class="chip" onclick="selectPDVUGN(this,'APOLO RJ')">APOLO RJ</button>
    <button class="chip" onclick="selectPDVUGN(this,'SUL')">SUL</button>
    <button class="chip" onclick="selectPDVUGN(this,'APOLO SP')">APOLO SP</button>
  </div>
  <div class="card">
    <div class="card-title"><i class="ti ti-building-store"></i> Setores com conversões C04 · clique para ver PDVs</div>
    <div class="tw"><table>
      <thead><tr><th></th><th>Setor</th><th>C04 conv.</th><th>PDVs únicos</th><th style="min-width:110px">Barra</th></tr></thead>
      <tbody id="pdv-tbody"></tbody>
    </table></div>
  </div>
''')
h('</div>')

# ── PAGE 4: REDES BRASIL ──────────────────────────────────────────────────────
# Compute redes metrics
total_redes = len(redes_data)
total_cnpjs = sum(len(d['cnpjs']) for d in redes_data.values())
top_rede_nm, top_rede_d = top_redes[0] if top_redes else ('—', {'cx': 0})
top5_cx = sum(d['cx'] for _, d in top_redes[:5])
total_cx_all = sum(d['cx'] for d in redes_data.values())
top5_pct = round(top5_cx / total_cx_all * 100, 1) if total_cx_all else 0

h('<div id="page-redes" class="page">')
h(f'''
  <p style="font-size:13px;color:var(--txt2);margin-bottom:1rem">Ranking por nome de rede · soma de todos os CNPJs da mesma rede · todas as UGNs · dados Cupons C04</p>
  <div class="metrics">
    <div class="metric"><div class="metric-label">Redes únicas</div><div class="metric-value">{fmt_br(total_redes)}</div><div class="metric-sub">redes identificadas</div></div>
    <div class="metric"><div class="metric-label">CNPJs únicos</div><div class="metric-value">{fmt_br(total_cnpjs)}</div><div class="metric-sub">farmácias ativas</div></div>
    <div class="metric"><div class="metric-label">Top rede C04</div><div class="metric-value" style="font-size:14px">{top_rede_nm[:16]}</div><div class="metric-sub">{fmt_br(top_rede_d["cx"])} cx</div></div>
    <div class="metric"><div class="metric-label">Conc. Top 5</div><div class="metric-value">{top5_pct:.1f}%'.replace('.', ',')</div><div class="metric-sub">do total C04</div></div>
  </div>
''')

# Top redes ranking por cx
top10 = top_redes[:10]
# Medals
medals = ['🥇','🥈','🥉'] + [str(i) for i in range(4,11)]

h('''  <div class="card">
    <div class="card-title"><i class="ti ti-trophy"></i> Top Redes — Cupons C04 · todas as linhas</div>
    <div class="tw"><table>
      <thead><tr><th>#</th><th>Rede</th><th>Caixas</th><th>CNPJs</th><th style="min-width:120px">Participação</th></tr></thead>
      <tbody>''')

for i, (nm, d) in enumerate(top10):
    pct = round(d['cx']/total_cx_all*100,1) if total_cx_all else 0
    bar_w = int(pct/top5_pct*80) if top5_pct else 0
    h(f'        <tr><td>{medals[i]}</td><td><strong>{nm}</strong></td><td>{fmt_br(d["cx"])}</td><td>{len(d["cnpjs"])}</td><td><div class="bar-wrap"><div class="bar-bg"><div class="bar-fill bar-e" style="width:{bar_w}%"></div></div><span class="bar-pct">{pct:.1f}%'.replace('.',',')+f'</span></div></td></tr>')

h('''      </tbody>
    </table></div>
  </div>
</div>''')

# ── PAGE 5: EVOLUÇÃO ──────────────────────────────────────────────────────────
# UGN % C04 vs C05 data
UGN_DELTA = [
    ('LESTE',    11.3, 12.1),
    ('NE',        5.3, 11.8),
    ('APOLO RJ',  4.6,  9.8),
    ('SPI OESTE', 2.6,  8.8),
    ('CO NO',     6.2,  9.1),
    ('SUL',       2.2,  5.7),
    ('APOLO SP',  0.2,  4.5),
]

h('<div id="page-evolucao" class="page">')
h('''
  <p style="font-size:13px;color:var(--txt2);margin-bottom:1rem">Conversões por ciclo · comparativo C04 → C05</p>
  <div class="chips" id="evo-chips">
    <button class="chip sel" onclick="selectEvo(this,'todas')">Todas UGNs</button>
    <button class="chip" onclick="selectEvo(this,'SPI OESTE')">SPI OESTE</button>
    <button class="chip" onclick="selectEvo(this,'LESTE')">LESTE</button>
    <button class="chip" onclick="selectEvo(this,'NE')">NE</button>
    <button class="chip" onclick="selectEvo(this,'CO NO')">CO NO</button>
    <button class="chip" onclick="selectEvo(this,'APOLO RJ')">APOLO RJ</button>
    <button class="chip" onclick="selectEvo(this,'SUL')">SUL</button>
    <button class="chip" onclick="selectEvo(this,'APOLO SP')">APOLO SP</button>
  </div>
  <div class="two" style="margin-bottom:1rem">
    <div class="card">
      <div class="card-title"><i class="ti ti-chart-bar"></i> Caixas convertidas por ciclo</div>
      <div class="chart-wrap"><canvas id="chartCiclo"></canvas></div>
      <div style="display:flex;gap:12px;margin-top:8px;font-size:11px;color:var(--txt2)">
        <span style="display:flex;align-items:center;gap:4px"><span style="width:8px;height:8px;background:var(--nexus-mid);border-radius:2px"></span>Nexus</span>
        <span style="display:flex;align-items:center;gap:4px"><span style="width:8px;height:8px;background:var(--vital-mid);border-radius:2px"></span>Vital</span>
        <span style="display:flex;align-items:center;gap:4px"><span style="width:8px;height:8px;background:var(--id-mid);border-radius:2px"></span>Identidade</span>
      </div>
    </div>
    <div class="card">
      <div class="card-title"><i class="ti ti-trending-up"></i> % conversão por ciclo e UGN</div>
      <table style="font-size:12px">
        <thead><tr><th>UGN</th><th>% C04</th><th>% C05</th><th>Δ</th></tr></thead>
        <tbody>
''')

for ugn, p4, p5 in UGN_DELTA:
    delta = round(p5-p4, 1)
    delta_str = f'+{delta:.1f}pp'.replace('.',',') if delta>=0 else f'{delta:.1f}pp'.replace('.',',')
    h(f'          <tr><td>{ugn}</td><td>{p4:.1f}%'.replace('.',',')+f'</td><td>{p5:.1f}%'.replace('.',',')+f'</td><td style="color:var(--nexus)">{delta_str}</td></tr>')

h('''        </tbody>
      </table>
    </div>
  </div>
  <div class="card">
    <div class="card-title"><i class="ti ti-chart-bar"></i> Comparativo % conversão C04 vs C05 por UGN</div>
    <div class="chart-wrap"><canvas id="chartUGN"></canvas></div>
  </div>
  <div class="card" id="evo-distrito-card" style="display:none">
    <div class="card-title"><i class="ti ti-chart-line"></i> Evolução por distrito — <span id="evo-ugn-label"></span></div>
    <div class="chart-wrap"><canvas id="chartDistrito"></canvas></div>
  </div>
</div>
''')

h('</div><!-- /container -->')

# ── SCRIPTS ───────────────────────────────────────────────────────────────────
h(f'<script>')
h(f'const UGN_DATA = {ugn_js_str};')
h(f'const TOP_REDES = {top_redes_js_str};')
h(f'''
const EVO_DATA = {{
  todas:    {{ c04conv:[903,1083,0], c05conv:[1930,2469,0], c04dist:49451, c05dist:56000, label:'Todas UGNs' }},
  "SPI OESTE": {{ c04conv:[99,82,0], c05conv:[310,296,0], c04dist:7000, c05dist:6923, label:'SPI OESTE', mediaUGN:8.8,
    distritos:['11030100','11030200','11630100','11630200'],
    d_c04:[6.0,5.4,3.3,4.1], d_c05:[17.1,19.0,17.1,16.6] }},
  LESTE:    {{ c04conv:[258,307,0], c05conv:[418,415,0], c04dist:5000, c05dist:6857, label:'LESTE', mediaUGN:12.1 }},
  NE:       {{ c04conv:[213,209,0], c05conv:[423,533,0], c04dist:8000, c05dist:8077, label:'NE', mediaUGN:11.8 }},
  "CO NO":  {{ c04conv:[205,229,0], c05conv:[253,376,0], c04dist:7000, c05dist:6923, label:'CO NO', mediaUGN:9.1 }},
  "APOLO RJ":{{ c04conv:[134,75,0], c05conv:[277,289,0], c04dist:4553, c05dist:5769, label:'APOLO RJ', mediaUGN:9.8 }},
  SUL:      {{ c04conv:[92,60,0], c05conv:[194,198,0], c04dist:6898, c05dist:6923, label:'SUL', mediaUGN:5.7 }},
  "APOLO SP":{{ c04conv:[22,1,0], c05conv:[237,180,0], c04dist:11000, c05dist:9231, label:'APOLO SP', mediaUGN:4.5 }}
}};
const UGN_PCTS = {{
  c04:[11.3,5.3,4.6,6.2,2.6,2.2,0.2],
  c05:[12.1,11.8,9.8,9.1,8.8,5.7,4.5],
  labels:['LESTE','NE','APOLO RJ','CO NO','SPI OESTE','SUL','APOLO SP']
}};
const MEDIA_NAC = 7.9;
''')

h('''
// Navigation
function showPage(id, btn) {
  document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));
  document.getElementById('page-' + id).classList.add('active');
  btn.classList.add('active');
  if (id === 'evolucao' && !chartsInit) initCharts();
}

function fmtBr(n) {
  if (n === null || n === undefined) return '—';
  return Math.round(n).toLocaleString('pt-BR');
}
function fmtPct(conv, dist) {
  if (!dist) return '<span class="zero">—</span>';
  const p = (conv/dist*100).toFixed(1).replace('.',',');
  return p + '%';
}
function pctN(conv, dist) {
  if (!dist) return 0;
  return (conv/dist*100);
}

// ── Page 2: Por UGN ──────────────────────────────────────────────────────────
let currentUGN = 'SPI OESTE';

function selectUGN(btn, ugn) {
  document.querySelectorAll('#ugn-chips .chip').forEach(c => c.classList.remove('sel'));
  btn.classList.add('sel');
  currentUGN = ugn;
  renderUGN(ugn);
}

function renderUGN(ugn) {
  const u = UGN_DATA[ugn];
  if (!u) return;
  // Metrics
  const nc = u.nexus_c05conv, vc = u.vital_c05conv;
  const nd = u.nexus_c05dist, vd = u.vital_c05dist;
  const np = nd ? (nc/nd*100).toFixed(1).replace('.',',')+'%' : '—';
  const vp = vd ? (vc/vd*100).toFixed(1).replace('.',',')+'%' : '—';
  document.getElementById('ugn-metrics').innerHTML = `
    <div class="metric"><div class="metric-label">Total C05 dist.</div><div class="metric-value">${fmtBr(u.c05dist)}</div><div class="metric-sub">Nexus + Vital + ID</div></div>
    <div class="metric"><div class="metric-label" style="color:var(--nexus)">Nexus C05</div><div class="metric-value" style="color:var(--nexus)">${fmtBr(u.nexus_c05dist)}</div><div class="metric-sub">${fmtBr(nc)} conv. · ${np}</div></div>
    <div class="metric"><div class="metric-label" style="color:var(--vital)">Vital C05</div><div class="metric-value" style="color:var(--vital)">${fmtBr(u.vital_c05dist)}</div><div class="metric-sub">${fmtBr(vc)} conv. · ${vp}</div></div>
    <div class="metric"><div class="metric-label" style="color:var(--id)">Identidade C05</div><div class="metric-value" style="color:var(--id)">${fmtBr(u.id_c05dist)}</div><div class="metric-sub">0 conv. · 0,0%</div></div>
  `;

  // Build table rows — use only <tr> elements (no nested tbody)
  let rows = '';
  let sid = 0;

  const addSetor = (s, linha, rowCls, pctCls, badgeCls) => {
    sid++;
    const key = `s${sid}`;
    const c04p = (s.c04d && linha!=='IDENTIDADE') ? (s.c04c/s.c04d*100).toFixed(1).replace('.',',')+' %' : '—';
    const c05p = s.c05d ? (s.c05c/s.c05d*100).toFixed(1).replace('.',',')+' %' : '—';
    const c04dStr = (s.c04d && linha!=='IDENTIDADE') ? fmtBr(s.c04d) : '<span class="zero">—</span>';
    const c04cStr = (linha!=='IDENTIDADE') ? fmtBr(s.c04c) : '<span class="zero">—</span>';
    const c05dStr = s.c05d ? fmtBr(s.c05d) : '<span class="zero">—</span>';
    const c05cStr = (linha!=='IDENTIDADE') ? fmtBr(s.c05c) : '<span class="zero">0</span>';

    rows += `<tr class="setor-row ${rowCls}" onclick="toggleSetor('${key}')">
      <td><i class="ti ti-chevron-right" id="ico-${key}" style="font-size:12px;color:var(--txt3)"></i></td>
      <td style="font-size:12px">${s.cd} · ${s.nome}</td>
      <td><span class="badge ${badgeCls}">${linha}</span></td>
      <td>${c04dStr}</td><td>${c04cStr}</td>
      <td class="${pctCls}">${linha!=='IDENTIDADE'?c04p:'<span class="zero">—</span>'}</td>
      <td>${c05dStr}</td><td>${c05cStr}</td>
      <td class="${pctCls}">${linha==='IDENTIDADE'?'<span class="zero">0,0%</span>':c05p}</td>
    </tr>`;

    // Detail rows (hidden by default, toggled by JS)
    if (linha === 'IDENTIDADE') {
      rows += `<tr class="detail-row detail-hdr" data-key="${key}"><td></td><td colspan="8" style="color:var(--id);font-size:11px"><i class="ti ti-info-circle"></i> Identidade: cupons distribuídos no C05 · 0 conversões registradas</td></tr>`;
    } else if (s.pdvs && s.pdvs.length) {
      rows += `<tr class="detail-row detail-hdr" data-key="${key}"><td></td><td colspan="2" style="color:var(--ems-blue-mid);font-weight:500;font-size:11px">CNPJ</td><td colspan="3">Nome PDV</td><td colspan="3" style="font-weight:500">Caixas C04</td></tr>`;
      s.pdvs.forEach(p => {
        rows += `<tr class="detail-row" data-key="${key}"><td></td><td colspan="2" style="font-size:11px">${p.cnpj}</td><td colspan="3">${p.nome}</td><td colspan="3">${fmtBr(p.cx)}</td></tr>`;
      });
    } else {
      rows += `<tr class="detail-row" data-key="${key}"><td></td><td colspan="8" style="color:var(--txt3);font-size:11px">Sem dados de PDV disponíveis</td></tr>`;
    }
  };

  // NEXUS
  rows += `<tr class="linha-sep"><td colspan="9"><span class="badge b-nexus">NEXUS</span> — ${u.nexus.length} setores · ${fmtBr(u.nexus_c04conv)} C04 conv. · ${fmtBr(u.nexus_c05conv)} C05 conv.</td></tr>`;
  u.nexus.forEach(s => addSetor(s,'NEXUS','row-nexus','pct-n','b-nexus'));
  const nxC04d=u.nexus.reduce((a,s)=>a+s.c04d,0), nxC05d=u.nexus.reduce((a,s)=>a+s.c05d,0);
  rows += `<tr class="row-subtotal-n"><td></td><td colspan="2"><strong>Subtotal NEXUS</strong></td><td>${fmtBr(nxC04d)}</td><td>${fmtBr(u.nexus_c04conv)}</td><td class="pct-n">${nxC04d?(u.nexus_c04conv/nxC04d*100).toFixed(1).replace('.',',')+'%':'—'}</td><td>${nxC05d?fmtBr(nxC05d):'—'}</td><td>${fmtBr(u.nexus_c05conv)}</td><td class="pct-n">—</td></tr>`;

  // VITAL
  rows += `<tr class="linha-sep"><td colspan="9"><span class="badge b-vital">VITAL</span> — ${u.vital.length} setores · ${fmtBr(u.vital_c04conv)} C04 conv. · ${fmtBr(u.vital_c05conv)} C05 conv.</td></tr>`;
  u.vital.forEach(s => addSetor(s,'VITAL','row-vital','pct-v','b-vital'));
  const vtC04d=u.vital.reduce((a,s)=>a+s.c04d,0), vtC05d=u.vital.reduce((a,s)=>a+s.c05d,0);
  rows += `<tr class="row-subtotal-v"><td></td><td colspan="2"><strong>Subtotal VITAL</strong></td><td>${fmtBr(vtC04d)}</td><td>${fmtBr(u.vital_c04conv)}</td><td class="pct-v">${vtC04d?(u.vital_c04conv/vtC04d*100).toFixed(1).replace('.',',')+'%':'—'}</td><td>${vtC05d?fmtBr(vtC05d):'—'}</td><td>${fmtBr(u.vital_c05conv)}</td><td class="pct-v">—</td></tr>`;

  // IDENTIDADE
  rows += `<tr class="linha-sep"><td colspan="9"><span class="badge b-id">IDENTIDADE</span> — ${u.identidade.length} setores · 0 conv. registradas</td></tr>`;
  u.identidade.forEach(s => addSetor(s,'IDENTIDADE','row-id','pct-id','b-id'));
  const idC05d=u.identidade.reduce((a,s)=>a+s.c05d,0);
  rows += `<tr class="row-subtotal-id"><td></td><td colspan="2"><strong>Subtotal IDENTIDADE</strong></td><td class="zero">—</td><td class="zero">—</td><td class="zero">—</td><td>${idC05d?fmtBr(idC05d):'—'}</td><td class="zero">0</td><td class="zero">0,0%</td></tr>`;

  // Total UGN
  rows += `<tr class="row-total"><td></td><td colspan="2"><strong>TOTAL ${ugn}</strong></td><td>${fmtBr(u.c04dist)}</td><td>${fmtBr(u.c04conv)}</td><td class="pct-e">${(u.c04conv/u.c04dist*100).toFixed(1).replace('.',',')}%</td><td>${fmtBr(u.c05dist)}</td><td>${fmtBr(u.c05conv)}</td><td class="pct-e">${(u.c05conv/u.c05dist*100).toFixed(1).replace('.',',')}%</td></tr>`;

  document.getElementById('ugn-table-wrap').innerHTML = `<table>
    <thead><tr><th></th><th>Setor</th><th>Linha</th><th>C04 dist.</th><th>C04 conv.</th><th>% C04</th><th>C05 dist.</th><th>C05 conv.</th><th>% C05</th></tr></thead>
    <tbody>${rows}</tbody>
  </table>`;
}

function toggleSetor(key) {
  const details = document.querySelectorAll(`tr.detail-row[data-key="${key}"]`);
  const ico = document.getElementById('ico-' + key);
  const isOpen = [...details].some(r => r.classList.contains('open'));
  details.forEach(r => r.classList.toggle('open', !isOpen));
  if (ico) {
    ico.className = 'ti ' + (!isOpen ? 'ti-chevron-down' : 'ti-chevron-right');
    ico.style.fontSize = '12px'; ico.style.color = 'var(--txt3)';
  }
}

// ── Page 3: PDV por Setor ────────────────────────────────────────────────────
let currentPDVUGN = 'SPI OESTE';
function selectPDVUGN(btn, ugn) {
  document.querySelectorAll('#pdv-chips .chip').forEach(c => c.classList.remove('sel'));
  btn.classList.add('sel');
  currentPDVUGN = ugn;
  renderPDV(ugn);
}

function renderPDV(ugn) {
  const u = UGN_DATA[ugn];
  if (!u) return;
  const all = [...u.nexus, ...u.vital].filter(s => s.c04c > 0);
  all.sort((a,b) => b.c04c - a.c04c);
  const maxC = all.length ? all[0].c04c : 1;
  let rows = '';
  let sid = 0;
  all.forEach(s => {
    sid++;
    const key = `p${sid}`;
    const barW = Math.round(s.c04c/maxC*100);
    rows += `<tr class="setor-row" onclick="toggleSetor('${key}')">
      <td><i class="ti ti-chevron-right" id="ico-${key}" style="font-size:12px;color:var(--txt3)"></i></td>
      <td style="font-size:12px">${s.cd} · ${s.nome}</td>
      <td>${fmtBr(s.c04c)}</td>
      <td>${s.pdvs ? s.pdvs.length : 0}</td>
      <td><div class="bar-wrap"><div class="bar-bg"><div class="bar-fill bar-e" style="width:${barW}%"></div></div><span class="bar-pct">${s.c04c}</span></div></td>
    </tr>`;
    if (s.pdvs && s.pdvs.length) {
      rows += `<tr class="detail-row detail-hdr" data-key="${key}"><td></td><td style="color:var(--ems-blue-mid);font-weight:500;font-size:11px">CNPJ</td><td colspan="2">Nome PDV</td><td style="font-weight:500">Cx C04</td></tr>`;
      s.pdvs.forEach(p => {
        rows += `<tr class="detail-row" data-key="${key}"><td></td><td style="font-size:11px">${p.cnpj}</td><td colspan="2">${p.nome}</td><td>${p.cx}</td></tr>`;
      });
    } else {
      rows += `<tr class="detail-row" data-key="${key}"><td></td><td colspan="4" style="color:var(--txt3);font-size:11px">Sem dados de PDV</td></tr>`;
    }
  });
  if (!rows) rows = `<tr><td colspan="5" style="text-align:center;padding:16px;color:var(--txt3)">Sem conversões nesta UGN</td></tr>`;
  document.getElementById('pdv-tbody').innerHTML = rows;
}

// ── Page 5: Evolução — Charts ────────────────────────────────────────────────
let chartsInit = false;
let chartCiclo, chartUGN, chartDistrito;

function initCharts() {
  chartsInit = true;
  // Chart 1: Ciclo stacked bar
  const ctx1 = document.getElementById('chartCiclo').getContext('2d');
  chartCiclo = new Chart(ctx1, {
    type:'bar',
    data:{
      labels:['C04 · Abril','C05 · Maio'],
      datasets:[
        {label:'Nexus',data:[903,1930],backgroundColor:'#388E3C'},
        {label:'Vital',data:[1083,2469],backgroundColor:'#D32F2F'},
        {label:'Identidade',data:[0,0],backgroundColor:'#F9A825'},
      ]
    },
    options:{
      responsive:true,maintainAspectRatio:false,
      plugins:{legend:{position:'top',labels:{usePointStyle:true,padding:16}}},
      scales:{x:{stacked:true,grid:{display:false}},y:{stacked:true,beginAtZero:true,ticks:{callback:v=>v.toLocaleString('pt-BR')}}}
    }
  });

  // Chart 2: UGN bar C04 vs C05
  const ctx2 = document.getElementById('chartUGN').getContext('2d');
  chartUGN = new Chart(ctx2, {
    type:'bar',
    data:{
      labels:UGN_PCTS.labels,
      datasets:[
        {label:'% C04',data:UGN_PCTS.c04,backgroundColor:'rgba(0,48,135,0.5)',borderColor:'#003087',borderWidth:1},
        {label:'% C05',data:UGN_PCTS.c05,backgroundColor:'rgba(56,142,60,0.7)',borderColor:'#388E3C',borderWidth:1},
      ]
    },
    options:{
      responsive:true,maintainAspectRatio:false,
      plugins:{
        legend:{position:'top',labels:{usePointStyle:true,padding:16}},
        tooltip:{callbacks:{label:ctx=>`${ctx.dataset.label}: ${ctx.parsed.y.toFixed(1).replace('.',',')}%`}}
      },
      scales:{
        x:{grid:{display:false}},
        y:{beginAtZero:true,ticks:{callback:v=>v.toFixed(1).replace('.',',')+"%"},
          afterDraw(chart){
            const yScale=chart.scales.y, xScale=chart.scales.x;
            const y=yScale.getPixelForValue(MEDIA_NAC);
            const ctx=chart.ctx;
            ctx.save();ctx.setLineDash([5,4]);ctx.strokeStyle='#B71C1C';ctx.lineWidth=1.5;
            ctx.beginPath();ctx.moveTo(xScale.left,y);ctx.lineTo(xScale.right,y);ctx.stroke();
            ctx.fillStyle='#B71C1C';ctx.font='10px sans-serif';ctx.fillText(`Média Nac. ${MEDIA_NAC}%`,xScale.left+4,y-4);
            ctx.restore();
          }
        }
      }
    }
  });
  selectEvo(document.querySelector('#evo-chips .sel'), 'todas');
}

let currentEvo = 'todas';
function selectEvo(btn, key) {
  document.querySelectorAll('#evo-chips .chip').forEach(c=>c.classList.remove('sel'));
  btn.classList.add('sel');
  currentEvo = key;
  const d = EVO_DATA[key];
  if (!d) return;
  // Update chart 1
  if (chartCiclo) {
    chartCiclo.data.datasets[0].data = [d.c04conv[0], d.c05conv[0]];
    chartCiclo.data.datasets[1].data = [d.c04conv[1], d.c05conv[1]];
    chartCiclo.data.datasets[2].data = [d.c04conv[2], d.c05conv[2]];
    chartCiclo.update();
  }
  // Distrito chart
  const distrCard = document.getElementById('evo-distrito-card');
  if (key !== 'todas' && d.distritos) {
    distrCard.style.display = 'block';
    document.getElementById('evo-ugn-label').textContent = d.label;
    if (!chartDistrito) {
      const ctx3 = document.getElementById('chartDistrito').getContext('2d');
      chartDistrito = new Chart(ctx3, {
        type:'bar',
        data:{
          labels:d.distritos,
          datasets:[
            {label:'% C04',data:d.d_c04,backgroundColor:'rgba(0,48,135,0.5)',borderColor:'#003087',borderWidth:1},
            {label:'% C05',data:d.d_c05,backgroundColor:'rgba(56,142,60,0.7)',borderColor:'#388E3C',borderWidth:1},
          ]
        },
        options:{
          responsive:true,maintainAspectRatio:false,
          plugins:{legend:{position:'top',labels:{usePointStyle:true,padding:12}},
            tooltip:{callbacks:{label:ctx=>`${ctx.dataset.label}: ${ctx.parsed.y.toFixed(1).replace('.',',')}%`}}},
          scales:{
            x:{grid:{display:false}},
            y:{beginAtZero:true,ticks:{callback:v=>v.toFixed(1).replace('.',',')+"%"},
              afterDraw(chart){
                if(!d.mediaUGN)return;
                const yS=chart.scales.y,xS=chart.scales.x,y=yS.getPixelForValue(d.mediaUGN),ctx=chart.ctx;
                ctx.save();ctx.setLineDash([5,4]);ctx.strokeStyle='#B71C1C';ctx.lineWidth=1.5;
                ctx.beginPath();ctx.moveTo(xS.left,y);ctx.lineTo(xS.right,y);ctx.stroke();
                ctx.fillStyle='#B71C1C';ctx.font='10px sans-serif';ctx.fillText(`Média UGN ${d.mediaUGN}%`,xS.left+4,y-4);
                ctx.restore();
              }
            }
          }
        }
      });
    } else {
      chartDistrito.data.labels = d.distritos;
      chartDistrito.data.datasets[0].data = d.d_c04;
      chartDistrito.data.datasets[1].data = d.d_c05;
      chartDistrito.update();
    }
  } else {
    distrCard.style.display = 'none';
  }
}

// Init on load
renderUGN('SPI OESTE');
renderPDV('SPI OESTE');
</script>
</body>
</html>''')

output = '\n'.join(html_parts)
os.makedirs('outputs', exist_ok=True)
with open('outputs/index.html', 'w', encoding='utf-8') as f:
    f.write(output)
print(f'Done! {len(output):,} chars written to outputs/index.html')
