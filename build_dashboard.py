"""
build_dashboard.py  ·  Planta Curicó · COMFRUT
═══════════════════════════════════════════════
Lee el Excel de la carpeta data/ y genera docs/dashboard.html

Uso local:   python build_dashboard.py
GitHub Actions: se ejecuta automáticamente al hacer push del Excel
"""
import os, sys, json, glob
import pandas as pd
from collections import defaultdict

# ── Locate Excel ──────────────────────────────────────────────────────────
BASE = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE, 'data')
DOCS_DIR = os.path.join(BASE, 'docs')
TMPL     = os.path.join(BASE, 'templates', 'dashboard_template.html')

os.makedirs(DOCS_DIR, exist_ok=True)

xlsx_files = glob.glob(os.path.join(DATA_DIR, '*.xlsx')) + \
             glob.glob(os.path.join(DATA_DIR, '*.xls'))
if not xlsx_files:
    print("ERROR: No Excel found in data/"); sys.exit(1)

EXCEL = sorted(xlsx_files)[-1]
print(f"Excel: {os.path.basename(EXCEL)}")
xl = pd.ExcelFile(EXCEL)
print(f"Sheets: {xl.sheet_names}")

# ── Helpers ───────────────────────────────────────────────────────────────
ZCIQ_L = ['S_L01','S_L03','S_L04','S_L05','S_LMANCU','S_MANUAL']
ZENV_L = ['S_ENV1']

def safe_str(v): return str(v) if pd.notna(v) else ''
def safe_num(v): return float(v) if pd.notna(v) and str(v).replace('.','').replace('-','').isdigit() else 0.0
def to_date(v):
    if pd.isna(v): return None
    if isinstance(v, (int, float)):
        try: return pd.Timestamp('1899-12-30') + pd.Timedelta(days=int(v))
        except: return None
    try: return pd.Timestamp(v)
    except: return None

def build_idx(records, filter_lines=None):
    MS, MD, SD = defaultdict(set), defaultdict(set), defaultdict(set)
    for r in records:
        if filter_lines and r.get('linea') not in filter_lines: continue
        ms, ss = str(r['mes']), str(r['semana'])
        MS[ms].add(r['semana']); MD[ms].add(r['dia']); SD[ss].add(r['dia'])
    return ({k:sorted(v) for k,v in MS.items()},
            {k:sorted(v) for k,v in MD.items()},
            {k:sorted(v) for k,v in SD.items()})

# ── Parse Tiempos Perdidos ────────────────────────────────────────────────
shTP = next((s for s in xl.sheet_names if 'perdid' in s.lower() or 'tiempos' in s.lower()), None)
PD = []
if shTP:
    df = pd.read_excel(xl, shTP)
    df['Fecha']     = pd.to_datetime(df.get('Fecha'), errors='coerce')
    df['T.Minutos'] = pd.to_numeric(df.get('T.Minutos'), errors='coerce').fillna(0)
    df['Semana']    = pd.to_numeric(df.get('Semana'), errors='coerce').fillna(0).astype(int)
    df['Turno']     = pd.to_numeric(df.get('Turno'), errors='coerce').fillna(1).astype(int)
    for _, r in df.iterrows():
        f = r['Fecha']
        if pd.isna(f): continue
        PD.append({
            'fecha':  f.strftime('%Y-%m-%d'),
            'año': int(f.year), 'mes': int(f.month), 'dia': int(f.day),
            'semana': int(r['Semana']), 'turno': int(r['Turno']),
            'linea':    safe_str(r.get('Pto. Trabajo')),
            'area':     safe_str(r.get('Clase de Orden')) or 'ZCIQ',
            'tipo_paro':safe_str(r.get('Tipo de Paro')),
            'categoria':safe_str(r.get('Desc.Clasifi. del Paro')) or 'Producción',
            'falla':    safe_str(r.get('Desc.Falla')) or 'SIN DESCRIPCIÓN',
            'obs':      safe_str(r.get('Observaciones')),
            'minutos':  float(r['T.Minutos']),
        })
print(f"Tiempos Perdidos: {len(PD)} registros")

# ── Parse Asistencia + Produccion ─────────────────────────────────────────
shAP = next((s for s in xl.sheet_names if 'asistencia' in s.lower() or 'produccion' in s.lower()), xl.sheet_names[0])
df_ap = pd.read_excel(xl, shAP)
df_ap['Inic.tratamiento'] = pd.to_datetime(df_ap.get('Inic.tratamiento'), errors='coerce')
df_ap['Semana'] = pd.to_numeric(df_ap.get('Semana'), errors='coerce').fillna(0).astype(int)
df_ap['Turno']  = pd.to_numeric(df_ap.get('Turno'), errors='coerce').fillna(0).astype(int)
NUM_COLS = ['T.Minutos','Tiempo Efec.Min.','Paros Plan Min.','Paros No Plan Min.',
            'Ton.Real','Cajas Produc.','BPM Total','BPM Estandar','BPM sin PP',
            'Cant.Personas','Produc.(Kg/H/Personas)','Kilos Ingresados','IQF Aprobado',
            'Kilos Aprobadas','Kilos Pure','Kilos Jugo','Kilos Crumble',
            'Teorico Cajas','Consumo Cajas','Teorico Bolsas','Consumo Bolsas']
for c in NUM_COLS:
    if c in df_ap.columns:
        df_ap[c] = pd.to_numeric(df_ap[c], errors='coerce').fillna(0)

PROD, TD = [], []
for _, r in df_ap.iterrows():
    f = r['Inic.tratamiento']
    if pd.isna(f): continue
    n = lambda k: float(r.get(k, 0) or 0)
    linea = safe_str(r.get('Pto. Trabajo'))
    row = {
        'fecha': f.strftime('%Y-%m-%d'),
        'año': int(f.year), 'mes': int(f.month), 'dia': int(f.day),
        'semana': int(r['Semana']), 'turno': int(r['Turno']),
        'linea': linea, 'area': 'ZENV' if linea == 'S_ENV1' else 'ZCIQ',
        'especie': safe_str(r.get('Especie')), 'sku': safe_str(r.get('Desc.Material')),
        'teo_min': n('T.Minutos'), 'efec_min': n('Tiempo Efec.Min.'),
        'plan_min': n('Paros Plan Min.'), 'nopl_min': n('Paros No Plan Min.'),
        'kg_ingresados': n('Kilos Ingresados'), 'iqf_aprobado': n('IQF Aprobado'),
        'kg_puro': n('Kilos Pure'), 'kg_jugo': n('Kilos Jugo'), 'kg_crumble': n('Kilos Crumble'),
        'ton_real': n('Ton.Real'), 'cajas': n('Cajas Produc.'), 'kg_aprobadas': n('Kilos Aprobadas'),
        'teo_cajas': n('Teorico Cajas'), 'con_cajas': n('Consumo Cajas'),
        'teo_bolsas': n('Teorico Bolsas'), 'con_bolsas': n('Consumo Bolsas'),
        'bpm_total': n('BPM Total'), 'bpm_std': n('BPM Estandar'), 'bpm_sinpp': n('BPM sin PP'),
        'personas': n('Cant.Personas'), 'kg_h_pers': n('Produc.(Kg/H/Personas)'),
    }
    PROD.append(row)
    TD.append({'fecha': row['fecha'], 'año': row['año'], 'mes': row['mes'],
               'dia': row['dia'], 'semana': row['semana'], 'linea': linea,
               'minutos': n('T.Minutos'), 'efec_min': n('Tiempo Efec.Min.'),
               'plan_min': n('Paros Plan Min.'), 'nopl_min': n('Paros No Plan Min.')})
print(f"Producción: {len(PROD)} registros")

# ── Build Indices ─────────────────────────────────────────────────────────
PIDX = {}
for area, lines in [('ZCIQ', ZCIQ_L), ('ZENV', ZENV_L)]:
    ms, md, sd = build_idx(PROD, lines)
    PIDX[f'MS_{area}'] = ms; PIDX[f'MD_{area}'] = md; PIDX[f'SD_{area}'] = sd

ms_all, md_all, sd_all = build_idx(PD, None)
ms_z,   md_z,   sd_z   = build_idx(PD, ZCIQ_L)
ms_e,   md_e,   sd_e   = build_idx(PD, ZENV_L)
IDX = {'MS': ms_all, 'MD': md_all, 'SD': sd_all,
       'MS_ZCIQ': ms_z, 'MD_ZCIQ': md_z, 'SD_ZCIQ': sd_z,
       'MS_ZENV': ms_e, 'MD_ZENV': md_e, 'SD_ZENV': sd_e}

# ── Line Availability ─────────────────────────────────────────────────────
all_combos  = sorted(set((r['año'], r['mes']) for r in TD))
combo_labels = [f"{m:02d}/{y}" for y, m in all_combos]
line_disp = {}
for l in ZCIQ_L + ZENV_L:
    pts = []
    for (año, mes) in all_combos:
        sub  = [r for r in TD if r['linea'] == l and r['año'] == año and r['mes'] == mes]
        teo  = sum(r['minutos'] for r in sub)
        efec = sum(r['efec_min'] for r in sub)
        pts.append(round(efec / teo * 100, 1) if teo > 0 else None)
    line_disp[l] = pts
LD = {'combos': combo_labels, 'data': line_disp}

# ── Programa Envasado ─────────────────────────────────────────────────────
shProg = next((s for s in xl.sheet_names if 'programa' in s.lower()), None)
PROG = []
if shProg:
    raw = pd.read_excel(xl, shProg, header=None)
    if len(raw) >= 3:
        turno_row = raw.iloc[0]
        date_row  = raw.iloc[1]
        for ri in range(2, len(raw)):
            row = raw.iloc[ri]
            if pd.isna(row.iloc[0]): continue
            cod = str(row.iloc[0]); sku = str(row.iloc[1]) if pd.notna(row.iloc[1]) else ''
            if not sku or sku == 'nan': continue
            for ci in range(3, len(date_row)):
                val = row.iloc[ci]
                if pd.isna(val) or float(str(val).replace(',','') if val else 0) <= 0: continue
                date_val = date_row.iloc[ci]
                if pd.isna(date_val): continue
                try: fecha = pd.Timestamp(date_val).strftime('%Y-%m-%d')
                except: continue
                tv = turno_row.iloc[ci]
                turno = int(tv) if pd.notna(tv) and str(tv).strip().isdigit() else 0
                PROG.append({'fecha': fecha, 'sku': sku, 'cod': cod,
                             'turno': turno, 'cajas_prog': float(val)})
print(f"Programa Envasado: {len(PROG)} registros")

# ── Generate dashboard.html ───────────────────────────────────────────────
if not os.path.exists(TMPL):
    print(f"ERROR: Template not found: {TMPL}"); sys.exit(1)

html = open(TMPL, encoding='utf-8').read()
html = html.replace('/*PROD_DATA*/', json.dumps(PROD))
html = html.replace('/*PIDX_DATA*/', json.dumps(PIDX))
html = html.replace('/*PROG_DATA*/', json.dumps(PROG))
html = html.replace('/*PD_DATA*/',   json.dumps(PD))
html = html.replace('/*TD_DATA*/',   json.dumps(TD))
html = html.replace('/*IDX_DATA*/',  json.dumps(IDX))
html = html.replace('/*LD_DATA*/',   json.dumps(LD))

out = os.path.join(DOCS_DIR, 'dashboard.html')
open(out, 'w', encoding='utf-8').write(html)
print(f"\n✅ dashboard.html generado → {len(html)//1024} KB")
print(f"   PROD:{len(PROD)} · PD:{len(PD)} · PROG:{len(PROG)}")
