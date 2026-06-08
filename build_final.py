"""
build_final.py v7 — Gera outputs/index.html
  - Abril  lido de C04.xlsx (mês 4 apenas)
  - Maio   lido de inputs/Cupons Maio.xlsx (fonte oficial)
  - Junho  lido de inputs/Cupons Jun.xlsx  (fonte oficial)
  - Mapeamento por setor e por distrito
"""
import json, openpyxl, zipfile, zlib, re, os, datetime
import pandas as pd
from collections import defaultdict

# ── Helpers ───────────────────────────────────────────────────────────────────
def fmt_br(n):
    if n is None: return '—'
    return f'{int(round(n)):,}'.replace(',', '.')

def fmt_pct(num, den):
    if not den: return '—'
    return f'{num/den*100:.1f}%'.replace('.', ',')

# ── Carregar D ────────────────────────────────────────────────────────────────
with open('inputs/D_final_v2.json', encoding='utf-8') as f:
    D = json.load(f)

# setor → {ugn, linha, cd_distrito}
setor_map = {}
for ugn, u in D['por_ugn'].items():
    for s in u['setores']:
        setor_map[s['cd_setor']] = {
            'ugn': ugn, 'linha': s['linha'], 'cd_distrito': s['cd_distrito']
        }

# distrito → {ugn, linha}  (fallback para codigos de distrito)
dist_map = {}
for ugn, u in D['por_ugn'].items():
    for s in u['setores']:
        d = s['cd_distrito']
        if d not in dist_map:
            dist_map[d] = {'ugn': ugn, 'linha': s['linha']}

def lookup_setor(cd):
    return setor_map.get(cd) or dist_map.get(cd)

# ── Ler C04.xlsx ──────────────────────────────────────────────────────────────
wb4 = openpyxl.load_workbook('inputs/Cupons C04.xlsx', read_only=True)
rows4 = [r for r in wb4.active.iter_rows(values_only=True)]
good4 = [r for r in rows4[1:] if r and len(r)>=11 and r[10] is not None and r[0] != 'Total']

# ── Ler C05.xlsx (CRC bypass) ─────────────────────────────────────────────────
class CRCBypassZip(zipfile.ZipFile):
    def read(self, name):
        info = self.getinfo(name)
        self.fp.seek(info.header_offset)
        fh = self.fp.read(30)
        fl = int.from_bytes(fh[26:28],'little')
        el = int.from_bytes(fh[28:30],'little')
        self.fp.seek(fl+el, 1)
        data = self.fp.read(info.compress_size)
        return zlib.decompress(data, -15) if info.compress_type==8 else data

zf5 = CRCBypassZip('inputs/Cupons C05.xlsx')
raw5 = zf5.read('xl/worksheets/sheet1.xml').decode('utf-8-sig', errors='replace')
rows5_raw = re.split(r'<x:row>', raw5)[1:]

def parse_c05_row(row_str):
    vals = []
    for cell in re.findall(r'<x:c[^>]*>(.*?)</x:c>', row_str, re.DOTALL):
        m = re.search(r'<x:is>.*?<x:t>(.*?)</x:t>', cell, re.DOTALL)
        if m: vals.append(m.group(1).strip()); continue
        m = re.search(r'<x:v>(.*?)</x:v>', cell, re.DOTALL)
        if m:
            try: vals.append(float(m.group(1)))
            except: vals.append(m.group(1))
            continue
        vals.append(None)
    return vals

good5 = []
for rstr in rows5_raw[1:]:
    r = parse_c05_row(rstr)
    if len(r) >= 11 and r[0] and r[10] and r[8]:
        good5.append(r)

def xl_month(v):
    try:
        return (datetime.date(1899,12,30) + datetime.timedelta(days=int(float(v)))).month
    except:
        return None

def xl_month_dt(v):
    if hasattr(v, 'month'): return v.month
    return None

# ── Ler Cupons Maio.xlsx (fonte oficial Maio) ─────────────────────────────────
df_maio = pd.read_excel('inputs/Cupons Maio.xlsx', header=0)
df_maio.columns = ['Setor','Resp','SetorFull','Cartao','Apresentacao','Desconto','CNPJ','NomePDV','Data','Hora','QtdCx']
df_maio['Setor'] = df_maio['Setor'].astype(str).str.strip()
df_maio['QtdCx'] = pd.to_numeric(df_maio['QtdCx'], errors='coerce')
df_maio = df_maio[df_maio['QtdCx'].notna() & (df_maio['QtdCx'] > 0)
                  & (~df_maio['Setor'].isin(['Setor','Total','nan','None']))]

# ── Ler Cupons Jun.xlsx (fonte oficial Junho) ─────────────────────────────────
df_jun = pd.read_excel('inputs/Cupons Jun.xlsx', header=0)
df_jun.columns = ['Setor','Resp','SetorFull','Cartao','Apresentacao','Desconto','CNPJ','NomePDV','Data','Hora','QtdCx']
df_jun['Setor'] = df_jun['Setor'].astype(str).str.strip()
df_jun['QtdCx'] = pd.to_numeric(df_jun['QtdCx'], errors='coerce')
df_jun = df_jun[df_jun['QtdCx'].notna() & (df_jun['QtdCx'] > 0)
                & (~df_jun['Setor'].isin(['Setor','Total','nan','None']))]

# ── Agregar conversoes por mes / setor ────────────────────────────────────────
conv_setor = defaultdict(lambda: defaultdict(int))   # conv[setor][mes] = cx
pdv_setor  = defaultdict(lambda: defaultdict(lambda: {'cx': 0, 'cnpj': ''}))

# C04.xlsx: Abril (4) apenas — Maio vem de maio.xlsx, Junho de jun.xlsx
for r in good4:
    setor = str(r[0])
    month = xl_month_dt(r[8])
    cx = int(r[10])
    nome = str(r[7]) if r[7] else ''
    cnpj = str(r[6]) if r[6] else ''
    if not month or month != 4: continue   # apenas Abril do C04
    conv_setor[setor][month] += cx
    if nome:
        pdv_setor[setor][nome]['cx'] += cx
        if not pdv_setor[setor][nome]['cnpj']:
            pdv_setor[setor][nome]['cnpj'] = cnpj

# maio.xlsx: Maio (5) — fonte oficial
redes_maio = defaultdict(lambda: {'cx': 0, 'cnpjs': set()})
for _, row in df_maio.iterrows():
    setor = str(row['Setor'])
    cx = int(row['QtdCx'])
    nome = str(row['NomePDV']) if pd.notna(row['NomePDV']) else ''
    cnpj = str(row['CNPJ']) if pd.notna(row['CNPJ']) else ''
    conv_setor[setor][5] += cx
    if nome:
        pdv_setor[setor][nome]['cx'] += cx
        if not pdv_setor[setor][nome]['cnpj']:
            pdv_setor[setor][nome]['cnpj'] = cnpj
        redes_maio[nome]['cx'] += cx
        redes_maio[nome]['cnpjs'].add(cnpj)

# jun.xlsx: Junho (6) — fonte oficial
redes_jun = defaultdict(lambda: {'cx': 0, 'cnpjs': set()})
for _, row in df_jun.iterrows():
    setor = str(row['Setor'])
    cx = int(row['QtdCx'])
    nome = str(row['NomePDV']) if pd.notna(row['NomePDV']) else ''
    cnpj = str(row['CNPJ']) if pd.notna(row['CNPJ']) else ''
    conv_setor[setor][6] += cx
    if nome:
        pdv_setor[setor][nome]['cx'] += cx
        if not pdv_setor[setor][nome]['cnpj']:
            pdv_setor[setor][nome]['cnpj'] = cnpj
        redes_jun[nome]['cx'] += cx
        redes_jun[nome]['cnpjs'].add(cnpj)

# ── Totais por UGN / linha / mes ──────────────────────────────────────────────
MESES = [4, 5, 6]
MES_LABEL = {4: 'Abril', 5: 'Maio', 6: 'Junho'}

ugn_conv  = defaultdict(lambda: defaultdict(int))
linha_conv = defaultdict(lambda: defaultdict(int))

for setor, mes_dict in conv_setor.items():
    info = lookup_setor(setor)
    if not info: continue
    ugn   = info.get('ugn')
    linha = info.get('linha')
    if not ugn: continue
    for mes, cx in mes_dict.items():
        ugn_conv[ugn][mes] += cx
        if linha:
            linha_conv[linha][mes] += cx

# ── Distribuicoes ─────────────────────────────────────────────────────────────
C04_DIST_UGN = {'LESTE':5000,'NE':8000,'APOLO RJ':4553,'CO NO':7000,'SPI OESTE':7000,'SUL':6898,'APOLO SP':11000}
C05_DIST_UGN = {'LESTE':6857,'NE':8077,'APOLO RJ':5769,'CO NO':6923,'SPI OESTE':6923,'SUL':6923,'APOLO SP':9231}
C04_DIST_LINHA = {'NEXUS':23645,'VITAL':25806,'IDENTIDADE':0}
C05_DIST_LINHA = {'NEXUS':25000,'VITAL':25000,'IDENTIDADE':6000}

UGN_ORDER = ['SPI OESTE','LESTE','NE','CO NO','APOLO RJ','SUL','APOLO SP']

# Per-setor PDV list
def get_pdvs(cd):
    sp = pdv_setor.get(cd, {})
    return sorted([{'nome': nm, 'cx': d['cx'], 'cnpj': d['cnpj']}
                   for nm, d in sp.items()], key=lambda x: -x['cx'])[:8]

# ── Build ugn_js dict ─────────────────────────────────────────────────────────
ugn_js = {}
for ugn in UGN_ORDER:
    u = D['por_ugn'][ugn]
    nexus_s = [s for s in u['setores'] if s['linha']=='NEXUS']
    vital_s  = [s for s in u['setores'] if s['linha']=='VITAL']
    ident_s  = [s for s in u['setores'] if s['linha']=='IDENTIDADE']

    def build_sl(sl):
        out = []
        for s in sl:
            cd = s['cd_setor']
            mc = conv_setor.get(cd, {})
            out.append({
                'cd': cd, 'nome': s['nome_setor'],
                'dist': s['cd_distrito'],
                'c04d': s['dist_c04'], 'c05d': s['dist_c05'],
                'abril': mc.get(4,0), 'maio': mc.get(5,0), 'junho': mc.get(6,0),
                'pdvs': get_pdvs(cd)
            })
        return out

    c04d = C04_DIST_UGN[ugn]
    c05d = C05_DIST_UGN[ugn]
    ugn_js[ugn] = {
        'c04dist': c04d, 'c05dist': c05d, 'totdist': c04d + c05d,
        'abril': ugn_conv[ugn].get(4,0),
        'maio':  ugn_conv[ugn].get(5,0),
        'junho': ugn_conv[ugn].get(6,0),
        'nexus': build_sl(nexus_s),
        'vital':  build_sl(vital_s),
        'identidade': build_sl(ident_s),
    }
    ugn_js[ugn]['totconv'] = ugn_js[ugn]['abril']+ugn_js[ugn]['maio']+ugn_js[ugn]['junho']

ugn_js_str = json.dumps(ugn_js, ensure_ascii=False)

# ── Totals for Page 1 ─────────────────────────────────────────────────────────
def linha_row(linha):
    c04d = C04_DIST_LINHA[linha]
    c05d = C05_DIST_LINHA[linha]
    tot_d = c04d + c05d
    abril = linha_conv[linha].get(4,0)
    maio  = linha_conv[linha].get(5,0)
    junho = linha_conv[linha].get(6,0)
    tot_c = abril+maio+junho
    pct   = fmt_pct(tot_c, tot_d) if tot_d else '—'
    return {'c04d':c04d,'c05d':c05d,'tot_d':tot_d,'abril':abril,'maio':maio,'junho':junho,'tot_c':tot_c,'pct':pct}

LN = {l: linha_row(l) for l in ['NEXUS','VITAL','IDENTIDADE']}
TOT_abril = sum(ugn_conv[u].get(4,0) for u in UGN_ORDER)
TOT_maio  = sum(ugn_conv[u].get(5,0) for u in UGN_ORDER)
TOT_junho = sum(ugn_conv[u].get(6,0) for u in UGN_ORDER)
TOT_conv  = TOT_abril+TOT_maio+TOT_junho
TOT_c04d  = sum(C04_DIST_UGN.values())
TOT_c05d  = sum(C05_DIST_UGN.values()) + 6000  # incluindo identidade
TOT_nv_d  = C04_DIST_LINHA['NEXUS']+C04_DIST_LINHA['VITAL']+C05_DIST_LINHA['NEXUS']+C05_DIST_LINHA['VITAL']  # 99451

# UGN Ranking sorted by %conv
ugn_ranking = []
for ugn in UGN_ORDER:
    u = ugn_js[ugn]
    pct = u['totconv']/u['totdist']*100 if u['totdist'] else 0
    ugn_ranking.append({'ugn':ugn,'c04d':u['c04dist'],'c05d':u['c05dist'],'totd':u['totdist'],
                        'abril':u['abril'],'maio':u['maio'],'junho':u['junho'],'totc':u['totconv'],'pct':pct})
ugn_ranking.sort(key=lambda x: -x['pct'])

# Redes — consolida Abril (C04) + Maio (maio.xlsx) + Junho (jun.xlsx)
redes_data = defaultdict(lambda: {'cx':0,'cnpjs':set()})
# Abril — C04
for r in good4:
    month = xl_month_dt(r[8])
    if month != 4: continue
    nome = str(r[7]) if r[7] else ''
    cnpj = str(r[6]) if r[6] else ''
    cx = int(r[10])
    if not nome: continue
    redes_data[nome]['cx'] += cx
    redes_data[nome]['cnpjs'].add(cnpj)
# Maio
for nm, d in redes_maio.items():
    redes_data[nm]['cx'] += d['cx']
    redes_data[nm]['cnpjs'].update(d['cnpjs'])
# Junho
for nm, d in redes_jun.items():
    redes_data[nm]['cx'] += d['cx']
    redes_data[nm]['cnpjs'].update(d['cnpjs'])

top_redes = sorted(redes_data.items(), key=lambda x: -x[1]['cx'])[:15]
top_redes_js = [{'nome':nm,'cx':d['cx'],'cnpjs':len(d['cnpjs'])} for nm,d in top_redes]
top_redes_js_str = json.dumps(top_redes_js, ensure_ascii=False)
total_cx_redes = sum(d['cx'] for d in redes_data.values())

# ── Gerar HTML ────────────────────────────────────────────────────────────────
parts = []
def h(s): parts.append(s)

h(f'''<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Olire &amp; Lirux — Vida Mais Leve 2026</title>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/@tabler/icons-webfont@2.47.0/tabler-icons.min.css">
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
:root{{
  --ems-blue:#003087;--ems-blue-mid:#0055B3;--ems-blue-light:#E8F0FA;--ems-blue-border:#B8CEEE;
  --nexus:#1B5E20;--nexus-light:#E8F5E9;--nexus-mid:#388E3C;
  --vital:#B71C1C;--vital-light:#FFEBEE;--vital-mid:#D32F2F;
  --id:#7B5200;--id-light:#FFF8E1;--id-mid:#F9A825;--id-border:#FFE082;
  --txt:#212529;--txt2:#6c757d;--txt3:#adb5bd;
  --border:#dee2e6;--bg:#fff;--bg2:#f8f9fa;
  --border-radius-lg:10px;--border-radius-md:6px
}}
body{{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;background:var(--bg2);color:var(--txt);line-height:1.5}}
.container{{max-width:1200px;margin:0 auto;padding:16px}}
.dash-header{{background:var(--ems-blue);border-radius:var(--border-radius-lg);padding:14px 20px;margin-bottom:1.25rem;display:flex;align-items:center;justify-content:space-between}}
.dash-header-left{{display:flex;align-items:center;gap:12px}}
.dash-logo{{width:36px;height:36px;background:#fff;border-radius:6px;display:flex;align-items:center;justify-content:center;font-size:12px;font-weight:700;color:var(--ems-blue)}}
.dash-title{{color:#fff;font-size:16px;font-weight:500}}
.dash-sub{{color:rgba(255,255,255,0.65);font-size:12px;margin-top:2px}}
.dash-badge{{background:rgba(255,255,255,0.15);color:#fff;font-size:11px;padding:4px 10px;border-radius:20px;border:0.5px solid rgba(255,255,255,0.25)}}
.nav{{display:flex;gap:4px;margin-bottom:1.25rem;flex-wrap:wrap;padding-bottom:1rem;border-bottom:0.5px solid var(--border)}}
.nav-btn{{padding:6px 13px;font-size:12px;border:0.5px solid var(--ems-blue-border);border-radius:var(--border-radius-md);background:transparent;color:var(--ems-blue-mid);cursor:pointer;display:flex;align-items:center;gap:5px}}
.nav-btn:hover{{background:var(--ems-blue-light)}}
.nav-btn.active{{background:var(--ems-blue);color:#fff;border-color:var(--ems-blue)}}
.page{{display:none}}.page.active{{display:block}}
.metrics{{display:grid;grid-template-columns:repeat(auto-fit,minmax(130px,1fr));gap:10px;margin-bottom:1.25rem}}
.metric{{background:var(--ems-blue-light);border:0.5px solid var(--ems-blue-border);border-radius:var(--border-radius-md);padding:12px 14px}}
.metric-label{{font-size:11px;color:var(--ems-blue-mid);margin-bottom:4px;font-weight:500}}
.metric-value{{font-size:22px;font-weight:500;color:var(--ems-blue)}}
.metric-sub{{font-size:11px;color:var(--txt3);margin-top:2px}}
.card{{background:var(--bg);border:0.5px solid var(--border);border-radius:var(--border-radius-lg);padding:1rem 1.25rem;margin-bottom:1rem}}
.card-title{{font-size:11px;font-weight:500;text-transform:uppercase;letter-spacing:0.07em;color:var(--ems-blue-mid);margin-bottom:12px;display:flex;align-items:center;gap:6px}}
.badge{{display:inline-flex;align-items:center;font-size:11px;font-weight:500;padding:2px 8px;border-radius:20px}}
.b-nexus{{background:var(--nexus-light);color:var(--nexus)}}
.b-vital{{background:var(--vital-light);color:var(--vital)}}
.b-id{{background:var(--id-light);color:var(--id);border:0.5px solid var(--id-border)}}
.tw{{overflow-x:auto}}
table{{width:100%;border-collapse:collapse;font-size:13px}}
th{{font-size:11px;font-weight:500;color:var(--txt2);text-align:left;padding:6px 8px;border-bottom:0.5px solid var(--border);white-space:nowrap;background:var(--bg2)}}
td{{padding:7px 8px;border-bottom:0.5px solid var(--border);color:var(--txt)}}
tr:last-child td{{border-bottom:none}}
tr.row-nexus td:first-child{{border-left:3px solid var(--nexus-mid)}}
tr.row-vital td:first-child{{border-left:3px solid var(--vital-mid)}}
tr.row-id{{background:var(--id-light)}}
tr.row-id td:first-child{{border-left:3px solid var(--id-mid)}}
tr.row-total{{background:var(--ems-blue-light);font-weight:500}}
tr.row-subtotal-n{{background:var(--nexus-light);font-weight:500}}
tr.row-subtotal-v{{background:var(--vital-light);font-weight:500}}
tr.row-subtotal-id{{background:var(--id-light);font-weight:500}}
tr.linha-sep td{{background:var(--bg2);font-size:11px;font-weight:500;color:var(--txt2);padding:5px 8px;border-top:1px solid var(--border)}}
tr.dist-sep td{{background:#f1f5ff;font-size:10px;font-weight:600;color:var(--ems-blue-mid);padding:4px 8px 4px 20px;border-top:0.5px solid var(--ems-blue-border);letter-spacing:0.05em}}
td.pct-n{{color:var(--nexus);font-weight:500}}
td.pct-v{{color:var(--vital);font-weight:500}}
td.pct-id{{color:var(--id);font-weight:500}}
td.pct-e{{color:var(--ems-blue);font-weight:500}}
td.zero{{color:var(--txt3);font-style:italic}}
.bar-wrap{{display:flex;align-items:center;gap:7px}}
.bar-bg{{flex:1;height:6px;background:var(--bg2);border-radius:3px;overflow:hidden}}
.bar-fill{{height:100%;border-radius:3px;background:var(--ems-blue)}}
.bar-pct{{font-size:11px;color:var(--txt2);min-width:40px;text-align:right}}
.chips{{display:flex;gap:6px;flex-wrap:wrap;margin-bottom:1rem}}
.chip{{padding:4px 12px;font-size:12px;border:0.5px solid var(--ems-blue-border);border-radius:20px;cursor:pointer;color:var(--ems-blue-mid);background:transparent}}
.chip.sel{{background:var(--ems-blue);color:#fff;border-color:var(--ems-blue)}}
.setor-row{{cursor:pointer}}
.setor-row:hover td{{background:var(--ems-blue-light)}}
tr.detail-row{{display:none;background:var(--bg2)}}
tr.detail-row.open{{display:table-row}}
tr.detail-row td{{padding:5px 8px 5px 28px;font-size:12px;color:var(--txt2)}}
tr.detail-hdr td{{padding:5px 8px 5px 28px;font-size:11px;font-weight:500;background:var(--bg2);color:var(--txt2)}}
.two{{display:grid;grid-template-columns:1fr 1fr;gap:1rem}}
@media(max-width:520px){{.two{{grid-template-columns:1fr}}}}
.rank-row{{display:flex;align-items:center;gap:10px;padding:8px 0;border-bottom:0.5px solid var(--border)}}
.rank-row:last-child{{border-bottom:none}}
.rank-num{{font-size:11px;font-weight:500;color:var(--txt3);min-width:18px;text-align:center}}
.rank-info{{flex:1;min-width:0}}
.rank-name{{font-size:13px;color:var(--txt);font-weight:500;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}}
.rank-detail{{font-size:11px;color:var(--txt3)}}
.rank-val{{font-size:14px;font-weight:500;color:var(--ems-blue);text-align:right;min-width:50px}}
.id-note{{background:var(--id-light);border:0.5px solid var(--id-border);border-radius:var(--border-radius-md);padding:8px 12px;font-size:12px;color:var(--id);margin-bottom:1rem;display:flex;align-items:center;gap:8px}}
.chart-wrap{{position:relative;height:260px}}
</style>
</head>
<body>
<div class="container">
''')

# ── HEADER ────────────────────────────────────────────────────────────────────
h('''<div class="dash-header">
  <div class="dash-header-left">
    <div class="dash-logo">EMS</div>
    <div>
      <div class="dash-title">Olire &amp; Lirux — Vida Mais Leve 2026</div>
      <div class="dash-sub">Conversão de Cupons C04 / C05 · Nexus · Vital · Identidade</div>
    </div>
  </div>
''')
h(f'  <div class="dash-badge"><i class="ti ti-refresh" style="font-size:12px"></i> Atualizado: {datetime.date.today().strftime("%d/%m/%Y")}</div>')
h('''</div>
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
h('''<div class="id-note">
  <i class="ti ti-info-circle" style="font-size:16px"></i>
  <span><strong>Identidade</strong> recebeu 6.000 cupons C05 · conversão registrada separadamente · % total calculado sobre Nexus+Vital</span>
</div>
''')

# Métricas
h(f'''<div class="metrics">
  <div class="metric"><div class="metric-label"><i class="ti ti-ticket"></i> Dist. C04</div><div class="metric-value">{fmt_br(TOT_c04d)}</div><div class="metric-sub">Nexus + Vital · Abril</div></div>
  <div class="metric"><div class="metric-label"><i class="ti ti-ticket"></i> Dist. C05</div><div class="metric-value">{fmt_br(56000)}</div><div class="metric-sub">Nexus + Vital + ID · Maio</div></div>
  <div class="metric"><div class="metric-label"><i class="ti ti-check"></i> Total Conv.</div><div class="metric-value">{fmt_br(TOT_conv)}</div><div class="metric-sub">Abril + Maio + Junho</div></div>
  <div class="metric"><div class="metric-label"><i class="ti ti-percent"></i> % Conversão</div><div class="metric-value">{fmt_pct(TOT_conv, TOT_nv_d)}</div><div class="metric-sub">Conv / Dist Nexus+Vital</div></div>
</div>
''')

# Tabela conversão por linha
h('''<div class="card">
  <div class="card-title"><i class="ti ti-table"></i> Conversão por linha</div>
  <p style="font-size:11px;color:var(--txt3);margin-bottom:10px">Dist. C04/C05 = racional de distribuição · colunas de mês = conversões por data · % = Total Conv. / Total Dist.</p>
  <div class="tw"><table>
    <thead><tr>
      <th>Linha</th><th>Dist.C04</th><th>Dist.C05</th><th>Tot.Dist.</th>
      <th>Abril</th><th>Maio</th><th>Junho</th><th>Tot.Conv.</th><th>%Conv.</th>
    </tr></thead>
    <tbody>''')

for linha, cls, pcls, badge in [
    ('NEXUS','row-nexus','pct-n','b-nexus'),
    ('VITAL','row-vital','pct-v','b-vital'),
    ('IDENTIDADE','row-id','pct-id','b-id'),
]:
    d = LN[linha]
    h(f'''      <tr class="{cls}">
        <td><span class="badge {badge}">{linha}</span></td>
        <td>{fmt_br(d["c04d"]) if d["c04d"] else '<span class="zero">—</span>'}</td>
        <td>{fmt_br(d["c05d"])}</td>
        <td>{fmt_br(d["tot_d"]) if d["tot_d"] else '<span class="zero">—</span>'}</td>
        <td>{fmt_br(d["abril"]) if d["abril"] else '<span class="zero">0</span>'}</td>
        <td>{fmt_br(d["maio"])}</td>
        <td>{fmt_br(d["junho"])}</td>
        <td>{fmt_br(d["tot_c"]) if d["tot_c"] else '<span class="zero">0</span>'}</td>
        <td class="{pcls}">{d["pct"]}</td>
      </tr>''')

tot_c04d = sum(C04_DIST_LINHA.values())
tot_c05d = sum(C05_DIST_LINHA.values())
tot_nv_d = C04_DIST_LINHA['NEXUS']+C04_DIST_LINHA['VITAL']+C05_DIST_LINHA['NEXUS']+C05_DIST_LINHA['VITAL']
h(f'''      <tr class="row-total">
        <td><strong>TOTAL</strong></td>
        <td>{fmt_br(tot_c04d)}</td><td>{fmt_br(tot_c05d)}</td><td>{fmt_br(tot_nv_d)}</td>
        <td>{fmt_br(TOT_abril)}</td><td>{fmt_br(TOT_maio)}</td><td>{fmt_br(TOT_junho)}</td>
        <td>{fmt_br(TOT_conv)}</td>
        <td class="pct-e">{fmt_pct(TOT_conv, tot_nv_d)}</td>
      </tr>
    </tbody>
  </table></div>
</div>
''')

# Ranking UGNs (cores azul)
h('''<div class="card">
  <div class="card-title"><i class="ti ti-map"></i> Ranking UGNs — Nexus + Vital + Identidade</div>
  <p style="font-size:11px;color:var(--txt3);margin-bottom:10px">Identidade somada ao total · barras em azul (visão consolidada)</p>
  <div class="tw"><table>
    <thead><tr>
      <th>#</th><th>UGN</th><th>Dist.C04</th><th>Dist.C05</th><th>Tot.Dist.</th>
      <th>Abril</th><th>Maio</th><th>Junho</th><th>Tot.Conv.</th><th>%Conv.</th>
      <th style="min-width:100px">Barra</th>
    </tr></thead>
    <tbody>''')

max_pct = max(r['pct'] for r in ugn_ranking) or 1
for i, r in enumerate(ugn_ranking):
    pct_s = f'{r["pct"]:.1f}%'.replace('.', ',')
    bar_w = int(r['pct']/max_pct*100)
    h(f'''      <tr>
        <td>{i+1}</td><td>{r["ugn"]}</td>
        <td>{fmt_br(r["c04d"])}</td><td>{fmt_br(r["c05d"])}</td><td>{fmt_br(r["totd"])}</td>
        <td>{fmt_br(r["abril"])}</td><td>{fmt_br(r["maio"])}</td><td>{fmt_br(r["junho"])}</td>
        <td>{fmt_br(r["totc"])}</td>
        <td class="pct-e">{pct_s}</td>
        <td><div class="bar-wrap"><div class="bar-bg"><div class="bar-fill" style="width:{bar_w}%"></div></div><span class="bar-pct">{pct_s}</span></div></td>
      </tr>''')

tot_ugn_c = sum(r['totc'] for r in ugn_ranking)
tot_ugn_d = sum(r['totd'] for r in ugn_ranking)
h(f'''      <tr class="row-total">
        <td></td><td><strong>TOTAL</strong></td>
        <td>{fmt_br(TOT_c04d)}</td><td>{fmt_br(56000)}</td><td>{fmt_br(tot_ugn_d)}</td>
        <td>{fmt_br(TOT_abril)}</td><td>{fmt_br(TOT_maio)}</td><td>{fmt_br(TOT_junho)}</td>
        <td>{fmt_br(TOT_conv)}</td>
        <td class="pct-e">{fmt_pct(TOT_conv, tot_nv_d)}</td><td></td>
      </tr>
    </tbody>
  </table></div>
</div>
</div>''')  # /page-consolidado

# ── PAGE 2: POR UGN ───────────────────────────────────────────────────────────
h('''<div id="page-ugn" class="page">
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
    <div class="card-title"><i class="ti ti-list"></i> Setores por linha e distrito — clique para expandir CNPJs</div>
    <div class="tw" id="ugn-table-wrap"></div>
  </div>
</div>''')

# ── PAGE 3: PDV POR SETOR ─────────────────────────────────────────────────────
h('''<div id="page-pdv-ugn" class="page">
  <p style="font-size:13px;color:var(--txt2);margin-bottom:1rem">Ranking de conversões dentro de cada setor (C04+C05) · clique para expandir PDVs · filtre por UGN</p>
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
    <div class="card-title"><i class="ti ti-building-store"></i> Setores com conversões · clique para ver PDVs</div>
    <div class="tw"><table>
      <thead><tr><th></th><th>Setor</th><th>Abril</th><th>Maio</th><th>Junho</th><th>Total</th><th>PDVs únicos</th><th style="min-width:110px">Barra</th></tr></thead>
      <tbody id="pdv-tbody"></tbody>
    </table></div>
  </div>
</div>''')

# ── PAGE 4: REDES BRASIL ──────────────────────────────────────────────────────
total_redes = len(redes_data)
total_cnpjs = sum(len(d['cnpjs']) for d in redes_data.values())
top1_nm, top1_d = top_redes[0] if top_redes else ('—', {'cx': 0})
top5_cx = sum(d['cx'] for _, d in top_redes[:5])
top5_pct = round(top5_cx/total_cx_redes*100,1) if total_cx_redes else 0

h(f'''<div id="page-redes" class="page">
  <p style="font-size:13px;color:var(--txt2);margin-bottom:1rem">Ranking por rede · soma de todos os CNPJs · C04 + C05 (dados de conversão)</p>
  <div class="metrics">
    <div class="metric"><div class="metric-label">Redes únicas</div><div class="metric-value">{fmt_br(total_redes)}</div><div class="metric-sub">redes identificadas</div></div>
    <div class="metric"><div class="metric-label">CNPJs únicos</div><div class="metric-value">{fmt_br(total_cnpjs)}</div><div class="metric-sub">farmácias ativas</div></div>
    <div class="metric"><div class="metric-label">Top rede</div><div class="metric-value" style="font-size:14px">{top1_nm[:16]}</div><div class="metric-sub">{fmt_br(top1_d["cx"])} cx</div></div>
    <div class="metric"><div class="metric-label">Conc. Top 5</div><div class="metric-value">{str(top5_pct).replace(".",",")}%</div><div class="metric-sub">do total de caixas</div></div>
  </div>
  <div class="card">
    <div class="card-title"><i class="ti ti-trophy"></i> Top Redes — todas as linhas</div>
    <div class="tw"><table>
      <thead><tr><th>#</th><th>Rede</th><th>Caixas</th><th>CNPJs</th><th style="min-width:120px">Participação</th></tr></thead>
      <tbody>''')

medals = ['🥇','🥈','🥉']+[str(i) for i in range(4,16)]
for i, (nm, d) in enumerate(top_redes[:10]):
    pct = round(d['cx']/total_cx_redes*100,1) if total_cx_redes else 0
    bw = int(pct/top5_pct*80) if top5_pct else 0
    pct_s = f'{pct:.1f}'.replace('.',',')
    h(f'        <tr><td>{medals[i]}</td><td><strong>{nm}</strong></td><td>{fmt_br(d["cx"])}</td><td>{len(d["cnpjs"])}</td><td><div class="bar-wrap"><div class="bar-bg"><div class="bar-fill" style="width:{bw}%"></div></div><span class="bar-pct">{pct_s}%</span></div></td></tr>')

h('''      </tbody>
    </table></div>
  </div>
</div>''')

# ── PAGE 5: EVOLUÇÃO ──────────────────────────────────────────────────────────
# % acumulada mês a mês por UGN (Abril, Abril+Maio, Abril+Maio+Junho)
evol_rows = []
for r in sorted(ugn_ranking, key=lambda x: -x['pct']):
    td = r['totd']
    p1 = round(r['abril']/td*100, 1) if td else 0
    p2 = round((r['abril']+r['maio'])/td*100, 1) if td else 0
    p3 = round(r['totc']/td*100, 1) if td else 0
    evol_rows.append({'ugn': r['ugn'], 'dist': td,
                      'abril': r['abril'], 'maio': r['maio'], 'junho': r['junho'], 'tot': r['totc'],
                      'p1': p1, 'p2': p2, 'p3': p3})

# totais nacionais acumulados (só Nexus+Vital para %)
tot_nv_d = TOT_nv_d
p_nac_1 = round(TOT_abril/tot_nv_d*100,1) if tot_nv_d else 0
p_nac_2 = round((TOT_abril+TOT_maio)/tot_nv_d*100,1) if tot_nv_d else 0
p_nac_3 = round((TOT_abril+TOT_maio+TOT_junho)/tot_nv_d*100,1) if tot_nv_d else 0

evol_js = json.dumps(evol_rows, ensure_ascii=False)

h('<div id="page-evolucao" class="page">')
h(f'''<p style="font-size:13px;color:var(--txt2);margin-bottom:1rem">Evolução mês a mês · caixas convertidas e % acumulada sobre total distribuído</p>
  <div class="two" style="margin-bottom:1rem">
    <div class="card">
      <div class="card-title"><i class="ti ti-chart-bar"></i> Caixas convertidas por mês</div>
      <div class="chart-wrap"><canvas id="chartMes"></canvas></div>
    </div>
    <div class="card">
      <div class="card-title"><i class="ti ti-trending-up"></i> % Conversão acumulada por mês · por UGN</div>
      <div class="chart-wrap"><canvas id="chartEvol"></canvas></div>
    </div>
  </div>
  <div class="card">
    <div class="card-title"><i class="ti ti-table"></i> Evolução mês a mês por UGN</div>
    <div class="tw"><table>
      <thead>
        <tr>
          <th>UGN</th>
          <th>Dist. Total</th>
          <th>Abril cx</th><th>% Abril</th>
          <th>Maio cx</th><th>% Ac. Maio</th>
          <th>Junho cx</th><th>% Ac. Junho</th>
          <th>Total cx</th>
        </tr>
      </thead>
      <tbody>''')

best_p3 = max(r['p3'] for r in evol_rows) if evol_rows else 1
for r in evol_rows:
    bw = int(r['p3']/best_p3*90) if best_p3 else 0
    p1s = f'{r["p1"]:.1f}'.replace('.',',')+'%'
    p2s = f'{r["p2"]:.1f}'.replace('.',',')+'%'
    p3s = f'{r["p3"]:.1f}'.replace('.',',')+'%'
    h(f'''        <tr>
          <td><strong>{r["ugn"]}</strong></td>
          <td>{fmt_br(r["dist"])}</td>
          <td>{fmt_br(r["abril"])}</td><td style="color:var(--ems-blue)">{p1s}</td>
          <td>{fmt_br(r["maio"])}</td><td style="color:var(--ems-blue)">{p2s}</td>
          <td>{fmt_br(r["junho"])}</td><td style="color:var(--ems-blue)">{p3s}</td>
          <td><div class="bar-wrap"><div class="bar-bg"><div class="bar-fill" style="width:{bw}%"></div></div><span class="bar-pct"><strong>{fmt_br(r["tot"])}</strong></span></div></td>
        </tr>''')

p_nac_1s = f'{p_nac_1:.1f}'.replace('.',',')+'%'
p_nac_2s = f'{p_nac_2:.1f}'.replace('.',',')+'%'
p_nac_3s = f'{p_nac_3:.1f}'.replace('.',',')+'%'
h(f'''        <tr style="font-weight:700;border-top:2px solid var(--border);background:var(--bg2)">
          <td>TOTAL NACIONAL</td>
          <td>{fmt_br(TOT_c04d+TOT_c05d)}</td>
          <td>{fmt_br(TOT_abril)}</td><td style="color:var(--ems-blue)">{p_nac_1s}</td>
          <td>{fmt_br(TOT_maio)}</td><td style="color:var(--ems-blue)">{p_nac_2s}</td>
          <td>{fmt_br(TOT_junho)}</td><td style="color:var(--ems-blue)">{p_nac_3s}</td>
          <td><strong>{fmt_br(TOT_conv)}</strong></td>
        </tr>''')

h(f'''      </tbody>
    </table></div>
  </div>
</div>
</div><!-- /container -->
<script>const EVOL_DATA = {evol_js};</script>''')

# ── SCRIPTS ───────────────────────────────────────────────────────────────────
# Compute chart data from aggregations
nexus_mes = [linha_conv['NEXUS'].get(m,0) for m in MESES]
vital_mes  = [linha_conv['VITAL'].get(m,0) for m in MESES]
id_mes     = [linha_conv['IDENTIDADE'].get(m,0) for m in MESES]

media_nac = round(TOT_conv/TOT_nv_d*100, 1) if TOT_nv_d else 0

h(f'<script>')
h(f'const UGN_DATA = {ugn_js_str};')
h(f'const TOP_REDES = {top_redes_js_str};')
h(f'''
const CHART_MES = {{
  labels: ['Abril','Maio','Junho'],
  nexus: {nexus_mes},
  vital: {vital_mes},
  id: {id_mes}
}};
const MEDIA_NAC = {media_nac};
''')

h('''
function showPage(id, btn) {
  document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));
  document.getElementById('page-' + id).classList.add('active');
  btn.classList.add('active');
  if (id === 'evolucao' && !chartsInit) initCharts();
}

function fmtBr(n) {
  if (n === null || n === undefined || n === 0) return '0';
  return Math.round(n).toLocaleString('pt-BR');
}

// ── Page 2: Por UGN ──────────────────────────────────────────────────────────
function selectUGN(btn, ugn) {
  document.querySelectorAll('#ugn-chips .chip').forEach(c => c.classList.remove('sel'));
  btn.classList.add('sel');
  renderUGN(ugn);
}

function renderUGN(ugn) {
  const u = UGN_DATA[ugn];
  if (!u) return;
  const totC = u.abril + u.maio + u.junho;
  const pctS = u.totdist ? (totC/u.totdist*100).toFixed(1).replace('.',',')+'%' : '—';
  document.getElementById('ugn-metrics').innerHTML = `
    <div class="metric"><div class="metric-label">Dist. C04</div><div class="metric-value">${u.c04dist.toLocaleString('pt-BR')}</div><div class="metric-sub">distribuição Abril</div></div>
    <div class="metric"><div class="metric-label">Dist. C05</div><div class="metric-value">${u.c05dist.toLocaleString('pt-BR')}</div><div class="metric-sub">distribuição Maio+ID</div></div>
    <div class="metric"><div class="metric-label">Total Conv.</div><div class="metric-value">${totC.toLocaleString('pt-BR')}</div><div class="metric-sub">Abr+Mai+Jun</div></div>
    <div class="metric"><div class="metric-label">% Conversão</div><div class="metric-value" style="color:var(--ems-blue)">${pctS}</div><div class="metric-sub">Conv / TotDist</div></div>
  `;

  let rows = '';
  let sid = 0;

  const addSetor = (s, linha, rowCls, pctCls, badgeCls) => {
    sid++;
    const key = `s${sid}`;
    const totS = s.abril + s.maio + s.junho;
    const totD = s.c04d + s.c05d;
    const pctT = totD ? (totS/totD*100).toFixed(1).replace('.',',')+' %' : '—';
    const c04dStr = (s.c04d && linha!=='IDENTIDADE') ? s.c04d.toLocaleString('pt-BR') : '<span class="zero">—</span>';
    const c05dStr = s.c05d ? s.c05d.toLocaleString('pt-BR') : '<span class="zero">—</span>';
    const abrilStr = s.abril ? s.abril : '<span class="zero">—</span>';
    const maioStr  = s.maio  ? s.maio  : '<span class="zero">—</span>';
    const junhoStr = s.junho ? s.junho : '<span class="zero">—</span>';
    const totSStr  = totS    ? totS    : '<span class="zero">—</span>';
    rows += `<tr class="setor-row ${rowCls}" onclick="toggleSetor('${key}')">
      <td><i class="ti ti-chevron-right" id="ico-${key}" style="font-size:12px;color:var(--txt3)"></i></td>
      <td style="font-size:12px">${s.cd} · ${s.nome}</td>
      <td><span class="badge ${badgeCls}">${linha}</span></td>
      <td>${c04dStr}</td><td>${c05dStr}</td>
      <td>${abrilStr}</td>
      <td>${maioStr}</td>
      <td>${junhoStr}</td>
      <td>${totSStr}</td>
      <td class="${pctCls}">${pctT}</td>
    </tr>`;
    if (linha === 'IDENTIDADE') {
      const identNote = totS > 0
        ? `<i class="ti ti-info-circle"></i> Identidade: ${totS} cx convertidas (% não calculada — excluída do denominador)`
        : `<i class="ti ti-info-circle"></i> Identidade: cupons distribuídos no C05 · sem conversões registradas`;
      rows += `<tr class="detail-row detail-hdr" data-key="${key}"><td></td><td colspan="9" style="color:var(--id);font-size:11px">${identNote}</td></tr>`;
    }
    if (s.pdvs && s.pdvs.length) {
      rows += `<tr class="detail-row detail-hdr" data-key="${key}"><td></td><td colspan="2" style="color:var(--ems-blue-mid);font-weight:500;font-size:11px">CNPJ</td><td colspan="4">Nome PDV</td><td colspan="3" style="font-weight:500">Caixas</td></tr>`;
      s.pdvs.forEach(p => {
        rows += `<tr class="detail-row" data-key="${key}"><td></td><td colspan="2" style="font-size:11px">${p.cnpj}</td><td colspan="4">${p.nome}</td><td colspan="3">${p.cx}</td></tr>`;
      });
    } else if (linha !== 'IDENTIDADE' || totS > 0) {
      rows += `<tr class="detail-row" data-key="${key}"><td></td><td colspan="9" style="color:var(--txt3);font-size:11px">Sem dados de PDV disponíveis</td></tr>`;
    }
  };

  const addBlocoLinha = (setores, linha, rowCls, pctCls, badgeCls) => {
    // Group by district
    const byDist = {};
    setores.forEach(s => {
      if (!byDist[s.dist]) byDist[s.dist] = [];
      byDist[s.dist].push(s);
    });
    const totConvBloco = setores.reduce((a,s) => a+s.abril+s.maio+s.junho, 0);
    rows += `<tr class="linha-sep"><td colspan="10"><span class="badge ${badgeCls}">${linha}</span> — ${setores.length} setores · ${totConvBloco.toLocaleString('pt-BR')} cx conv.</td></tr>`;
    Object.keys(byDist).sort().forEach(dist => {
      rows += `<tr class="dist-sep"><td colspan="10"><i class="ti ti-map-pin" style="font-size:10px;margin-right:4px"></i>DISTRITO ${dist}</td></tr>`;
      byDist[dist].forEach(s => addSetor(s, linha, rowCls, pctCls, badgeCls));
    });
  };

  addBlocoLinha(u.nexus, 'NEXUS', 'row-nexus', 'pct-n', 'b-nexus');
  addBlocoLinha(u.vital, 'VITAL', 'row-vital', 'pct-v', 'b-vital');
  addBlocoLinha(u.identidade, 'IDENTIDADE', 'row-id', 'pct-id', 'b-id');

  rows += `<tr class="row-total"><td colspan="2"><strong>TOTAL ${ugn}</strong></td><td></td><td>${u.c04dist.toLocaleString('pt-BR')}</td><td>${u.c05dist.toLocaleString('pt-BR')}</td><td>${u.abril.toLocaleString('pt-BR')}</td><td>${u.maio.toLocaleString('pt-BR')}</td><td>${u.junho.toLocaleString('pt-BR')}</td><td>${totC.toLocaleString('pt-BR')}</td><td class="pct-e">${pctS}</td></tr>`;

  document.getElementById('ugn-table-wrap').innerHTML = `<table>
    <thead><tr><th></th><th>Setor</th><th>Linha</th><th>Dist.C04</th><th>Dist.C05</th><th>Abril</th><th>Maio</th><th>Junho</th><th>Total</th><th>%Conv.</th></tr></thead>
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
function selectPDVUGN(btn, ugn) {
  document.querySelectorAll('#pdv-chips .chip').forEach(c => c.classList.remove('sel'));
  btn.classList.add('sel');
  renderPDV(ugn);
}

function renderPDV(ugn) {
  const u = UGN_DATA[ugn];
  if (!u) return;
  const all = [...u.nexus, ...u.vital].filter(s => (s.abril+s.maio+s.junho) > 0);
  all.sort((a,b) => (b.abril+b.maio+b.junho)-(a.abril+a.maio+a.junho));
  const maxC = all.length ? all[0].abril+all[0].maio+all[0].junho : 1;
  let rows = '';
  let sid = 0;
  all.forEach(s => {
    sid++;
    const key = `p${sid}`;
    const tot = s.abril+s.maio+s.junho;
    const barW = Math.round(tot/maxC*100);
    rows += `<tr class="setor-row" onclick="toggleSetor('${key}')">
      <td><i class="ti ti-chevron-right" id="ico-${key}" style="font-size:12px;color:var(--txt3)"></i></td>
      <td style="font-size:12px">${s.cd} · ${s.nome}</td>
      <td>${s.abril}</td><td>${s.maio}</td><td>${s.junho}</td><td><strong>${tot}</strong></td>
      <td>${s.pdvs ? s.pdvs.length : 0}</td>
      <td><div class="bar-wrap"><div class="bar-bg"><div class="bar-fill" style="width:${barW}%"></div></div><span class="bar-pct">${tot}</span></div></td>
    </tr>`;
    if (s.pdvs && s.pdvs.length) {
      rows += `<tr class="detail-row detail-hdr" data-key="${key}"><td></td><td style="color:var(--ems-blue-mid);font-weight:500;font-size:11px">CNPJ</td><td colspan="4">Nome PDV</td><td colspan="2" style="font-weight:500">Cx C04</td></tr>`;
      s.pdvs.forEach(p => {
        rows += `<tr class="detail-row" data-key="${key}"><td></td><td style="font-size:11px">${p.cnpj}</td><td colspan="4">${p.nome}</td><td colspan="2">${p.cx}</td></tr>`;
      });
    } else {
      rows += `<tr class="detail-row" data-key="${key}"><td></td><td colspan="7" style="color:var(--txt3);font-size:11px">Sem dados de PDV</td></tr>`;
    }
  });
  if (!rows) rows = `<tr><td colspan="8" style="text-align:center;padding:16px;color:var(--txt3)">Sem conversões nesta UGN</td></tr>`;
  document.getElementById('pdv-tbody').innerHTML = rows;
}

// ── Charts ────────────────────────────────────────────────────────────────────
let chartsInit = false;

function initCharts() {
  chartsInit = true;
  const ctx1 = document.getElementById('chartMes').getContext('2d');
  new Chart(ctx1, {
    type: 'bar',
    data: {
      labels: CHART_MES.labels,
      datasets: [
        {label:'Nexus', data:CHART_MES.nexus, backgroundColor:'#388E3C'},
        {label:'Vital',  data:CHART_MES.vital,  backgroundColor:'#D32F2F'},
        {label:'Identidade', data:CHART_MES.id, backgroundColor:'#F9A825'},
      ]
    },
    options: {
      responsive:true, maintainAspectRatio:false,
      plugins:{legend:{position:'top',labels:{usePointStyle:true,padding:16}}},
      scales:{x:{stacked:true,grid:{display:false}},y:{stacked:true,beginAtZero:true,ticks:{callback:v=>v.toLocaleString('pt-BR')}}}
    }
  });

  const ctx2 = document.getElementById('chartEvol').getContext('2d');
  const COLORS_UGN = ['#003087','#1B5E20','#B71C1C','#E65100','#6A1B9A','#00695C','#4E342E'];
  new Chart(ctx2, {
    type: 'line',
    data: {
      labels: ['Abril','Ac. Maio','Ac. Junho'],
      datasets: EVOL_DATA.map((r,i) => ({
        label: r.ugn,
        data: [r.p1, r.p2, r.p3],
        borderColor: COLORS_UGN[i % COLORS_UGN.length],
        backgroundColor: COLORS_UGN[i % COLORS_UGN.length] + '18',
        borderWidth: 2,
        pointRadius: 4,
        pointHoverRadius: 6,
        tension: 0.3,
        fill: false
      }))
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      interaction: {mode:'index', intersect:false},
      plugins:{
        legend:{position:'top', labels:{usePointStyle:true, padding:12, font:{size:11}}},
        tooltip:{callbacks:{label:ctx=>`${ctx.dataset.label}: ${ctx.parsed.y.toFixed(1).replace('.',',')}%`}}
      },
      scales:{
        x:{grid:{display:false}},
        y:{beginAtZero:true, ticks:{callback:v=>v.toFixed(1).replace('.',',')+"%"}}
      }
    }
  });
}

// Init
renderUGN('SPI OESTE');
renderPDV('SPI OESTE');
</script>
</body>
</html>''')

output = '\n'.join(parts)
os.makedirs('outputs', exist_ok=True)
with open('outputs/index.html', 'w', encoding='utf-8') as f:
    f.write(output)
print(f'Done! {len(output):,} chars written to outputs/index.html')
