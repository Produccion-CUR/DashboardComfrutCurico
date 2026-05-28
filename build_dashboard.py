#!/usr/bin/env python3
"""
build_dashboard.py — Lee el Excel de data/, genera los JSONs, embebe en el template
y escribe docs/index.html
"""
import sys, json, os, datetime
from pathlib import Path
from collections import defaultdict

# ── Rutas ────────────────────────────────────────────────────────────────────
ROOT     = Path(__file__).parent
DATA_DIR = ROOT / 'data'
TMPL_DIR = ROOT / 'templates'
OUT_DIR  = ROOT / 'docs'
OUT_DIR.mkdir(exist_ok=True)

excels = sorted(DATA_DIR.glob('*.xlsx'))
if not excels:
    print("ERROR: No se encontró ningún .xlsx en data/"); sys.exit(1)
EXCEL = str(excels[-1])
print(f"Excel: {EXCEL}")

templates = list(TMPL_DIR.glob('*.html'))
if not templates:
    print("ERROR: No se encontró template en templates/"); sys.exit(1)
TEMPLATE = str(templates[0])
print(f"Template: {TEMPLATE}")

try:
    import pandas as pd
except ImportError:
    os.system('pip install pandas openpyxl --quiet')
    import pandas as pd

# ── Helpers ───────────────────────────────────────────────────────────────────
def n(x):
    try: return float(x) if pd.notna(x) else 0.0
    except: return 0.0

def s(x):
    v = str(x).strip() if pd.notna(x) else ''
    return '' if v == 'nan' else v

def fechastr(x):
    try:
        t = pd.Timestamp(x)
        return t.strftime('%Y-%m-%d') if pd.notna(t) else ''
    except: return ''

ZCIQ_L = ['S_L01','S_L03','S_L04','S_L05','S_LMANCU','S_MANUAL']
ZENV_L = ['S_ENV1']

xl = pd.ExcelFile(EXCEL)
print(f"Hojas: {xl.sheet_names}")

# ── 1. PRODUCCION ─────────────────────────────────────────────────────────────
print("1. Asistencia + Produccion...")
df_ap = pd.read_excel(xl, 'Asistencia + Produccion', dtype={'Material': str})
df_ap['fecha'] = df_ap['Inic.tratamiento'].apply(fechastr)
df_ap = df_ap[df_ap['fecha'].str.len() == 10].copy()

# Build authoritative fecha→semana from AP FIRST (before processing any other sheet)
FECHA_SEM = df_ap.groupby('fecha')['Semana'].first().astype(int).to_dict()

def get_sem(fecha, fallback):
    """AP semana is authoritative. Fallback for dates not in AP."""
    if fecha in FECHA_SEM:
        return FECHA_SEM[fecha]
    try: return int(float(fallback)) if fallback and str(fallback) != 'nan' else 0
    except: return 0

prod = []
for _, r in df_ap.iterrows():
    fecha = r['fecha']; p = fecha.split('-')
    linea = s(r['Pto. Trabajo'])
    area  = 'ZENV' if linea == 'S_ENV1' else 'ZCIQ'
    prod.append({
        'fecha': fecha,
        'año': int(p[0]), 'mes': int(p[1]), 'dia': int(p[2]),
        'semana': int(n(r['Semana'])),
        'turno':  int(n(r['Turno'])),
        'linea': linea, 'area': area,
        'especie': s(r['Especie']),
        'sku':     s(r['Desc.Material']),
        'cod':     s(r['Material']),
        'teo_min':   n(r['T.Minutos']),
        'efec_min':  n(r['Tiempo Efec.Min.']),
        'plan_min':  n(r['Paros Plan Min.']),
        'nopl_min':  n(r['Paros No Plan Min.']),
        'kg_ingresados': n(r['Kilos Ingresados']),
        'iqf_aprobado':  n(r['IQF Aprobado']),
        'kg_puro':    n(r['Kilos Pure']),
        'kg_jugo':    n(r['Kilos Jugo']),
        'kg_crumble': n(r['Kilos Crumble']),
        'ton_real':   n(r['Ton.Real']),
        'cajas':      n(r['Cajas Produc.']),
        'kg_aprobadas': n(r['Kilos Aprobadas']),
        'teo_cajas':  n(r['Teorico Cajas']),
        'con_cajas':  n(r['Consumo Cajas']),
        'teo_bolsas': n(r['Teorico Bolsas']),
        'con_bolsas': n(r['Consumo Bolsas']),
        'bpm_total':  n(r['BPM Total']),
        'bpm_std':    n(r['BPM Estandar']),
        'bpm_sinpp':  n(r['BPM sin PP']),
        'personas':   n(r['Cant.Personas']),
        'kg_h_pers':  n(r['Produc.(Kg/H/Personas)']),
    })
print(f"   {len(prod)} registros")

# ── 2. PROGRAMA ENVASADO ──────────────────────────────────────────────────────
print("2. Programa Envasado...")
df_prog = pd.read_excel(xl, 'Programa Envasado', header=None)
turno_row = df_prog.iloc[0].tolist()
date_row  = df_prog.iloc[1].tolist()
prog = []
for ri in range(2, len(df_prog)):
    row = df_prog.iloc[ri]
    cod = s(row.iloc[0]); sku = s(row.iloc[1])
    if not cod or not sku: continue
    for ci in range(3, len(date_row)):
        val = row.iloc[ci]
        if pd.isna(val): continue
        try: val_f = float(val)
        except: continue
        if val_f <= 0: continue
        fecha = fechastr(date_row[ci])
        if not fecha or len(fecha) != 10: continue
        try: turno = int(float(turno_row[ci])) if pd.notna(turno_row[ci]) else 0
        except: turno = 0
        if turno not in [1, 2]: continue
        prog.append({'fecha': fecha, 'cod': cod, 'sku': sku,
                     'turno': turno, 'cajas_prog': val_f})
print(f"   {len(prog)} registros")

# Validación cumplimiento
prod_zenv     = [r for r in prod if r['area'] == 'ZENV' and r['cod']]
prog_keys_set = {(p['fecha'], p['cod'], p['turno']) for p in prog}
matched       = [p for p in prog if any(
    r['fecha'] == p['fecha'] and r['cod'] == p['cod'] and r['turno'] == p['turno']
    for r in prod_zenv)]
con_prog = sum(r['cajas'] for r in prod_zenv
               if (r['fecha'], r['cod'], r['turno']) in prog_keys_set)
tot_prog = sum(p['cajas_prog'] for p in matched)
if tot_prog:
    print(f"   Validación cumpl: {con_prog/tot_prog*100:.1f}%")

# ── 3. TIEMPOS PERDIDOS ───────────────────────────────────────────────────────
print("3. Tiempos Perdidos...")
df_tp = pd.read_excel(xl, 'Tiempos Perdidos', dtype=str)
df_tp.columns = [c.strip() for c in df_tp.columns]
perdidas = []
for _, r in df_tp.iterrows():
    linea = s(r.get('Pto. Trabajo', ''))
    if not linea: continue
    area = 'ZCIQ' if linea in ZCIQ_L else ('ZENV' if linea in ZENV_L else '')
    if not area: continue
    fecha = fechastr(r.get('Fecha', ''))
    if not fecha or len(fecha) != 10: continue
    minutos = n(r.get('T.Minutos', '0'))
    if minutos <= 0: continue
    p = fecha.split('-')
    sem = get_sem(fecha, r.get('Semana', '0'))
    if not sem:
        try: sem = datetime.date(int(p[0]), int(p[1]), int(p[2])).isocalendar()[1]
        except: sem = 0
    perdidas.append({
        'fecha': fecha,
        'año': int(p[0]), 'mes': int(p[1]), 'dia': int(p[2]),
        'semana': sem, 'linea': linea, 'area': area,
        'falla':     s(r.get('Desc.Falla', '')),
        'categoria': s(r.get('Desc.Clasifi. del Paro', '')),
        'tipo_paro': s(r.get('Tipo de Paro', '')),
        'obs':       s(r.get('Observaciones', '')),
        'minutos':   minutos,
        'turno':     int(n(r.get('Turno', '0'))),
    })
print(f"   {len(perdidas)} registros")

# ── 4. SEGURIDAD ──────────────────────────────────────────────────────────────
seg = []
if 'Seguridad' in xl.sheet_names:
    print("4. Seguridad...")
    df_s = pd.read_excel(xl, 'Seguridad')
    df_s['fecha'] = df_s['Fecha'].apply(fechastr)
    df_s = df_s[df_s['fecha'].str.len() == 10].copy()
    for _, r in df_s.iterrows():
        fecha = r['fecha']; p = fecha.split('-')
        sem = get_sem(fecha, str(r['Semana']))
        seg.append({
            'fecha': fecha, 'año': int(p[0]), 'mes': int(p[1]), 'dia': int(p[2]),
            'semana': sem, 'rut': s(r['Rut']), 'turno': s(r['Turno']),
            'supervisor': s(r['Supervisor']), 'linea': s(r['Linea']),
            'cantidad': int(n(r['Cantidad'])), 'nombre': s(r['Nombre']),
            'lesion':   s(r['Lesión o situación ']),
            'accion':   s(r['Acción / Condición Sub Estándar']),
            'medida':   s(r['Medida inmediata adoptada por jefatura']),
            'diat':     s(r['DIAT']),
        })
    print(f"   {len(seg)} registros")
else:
    print("4. Seguridad: hoja no encontrada, omitiendo")

# ── 5. TEORICO (desde AP agrupado por fecha+linea) ────────────────────────────
print("5. Teorico...")
grp_td = df_ap.groupby(['fecha', 'Pto. Trabajo']).agg({
    'T.Minutos': 'sum', 'Tiempo Efec.Min.': 'sum',
    'Paros Plan Min.': 'sum', 'Paros No Plan Min.': 'sum', 'Semana': 'first',
}).reset_index()
teorico = []
for _, r in grp_td.iterrows():
    fecha = r['fecha']; p = fecha.split('-')
    sem = get_sem(fecha, r['Semana'])
    if n(r['T.Minutos']) <= 0: continue
    teorico.append({
        'fecha': fecha, 'año': int(p[0]), 'mes': int(p[1]), 'dia': int(p[2]),
        'semana': sem, 'linea': s(r['Pto. Trabajo']),
        'minutos':  n(r['T.Minutos']),  'efec_min': n(r['Tiempo Efec.Min.']),
        'plan_min': n(r['Paros Plan Min.']), 'nopl_min': n(r['Paros No Plan Min.']),
    })
print(f"   {len(teorico)} registros")

# ── 6. PIDX (PROD + TP combinado para filtros consistentes) ──────────────────
print("6. PIDX...")
def build_pidx(records):
    MS = defaultdict(set); MD = defaultdict(set); SD = defaultdict(set)
    for r in records:
        mes = r['mes']; sem = r['semana']; dia = r['dia']
        if mes and sem: MS[str(mes)].add(sem)
        if mes and dia: MD[str(mes)].add(dia)
        if sem and dia: SD[str(sem)].add(dia)
    return {
        'MS': {k: sorted(v) for k, v in MS.items()},
        'MD': {k: sorted(v) for k, v in MD.items()},
        'SD': {k: sorted(v) for k, v in SD.items()},
    }

zciq_r  = [r for r in prod     if r['area'] == 'ZCIQ']
zenv_r  = [r for r in prod     if r['area'] == 'ZENV']
zciq_tp = [r for r in perdidas if r['area'] == 'ZCIQ']
zenv_tp = [r for r in perdidas if r['area'] == 'ZENV']

all_idx  = build_pidx(prod + perdidas)
zciq_idx = build_pidx(zciq_r + zciq_tp)
zenv_idx = build_pidx(zenv_r + zenv_tp)
pidx = {
    'MS': all_idx['MS'],   'MD': all_idx['MD'],   'SD': all_idx['SD'],
    'MS_ZCIQ': zciq_idx['MS'], 'MD_ZCIQ': zciq_idx['MD'], 'SD_ZCIQ': zciq_idx['SD'],
    'MS_ZENV': zenv_idx['MS'], 'MD_ZENV': zenv_idx['MD'], 'SD_ZENV': zenv_idx['SD'],
}

# ── 7. LINE_DISP ──────────────────────────────────────────────────────────────
ld = defaultdict(lambda: {'min': 0.0, 'efec': 0.0, 'plan': 0.0, 'nopl': 0.0})
for r in teorico:
    key = f"{r['linea']}|{r['año']}-{str(r['mes']).zfill(2)}"
    ld[key]['min']  += r['minutos'];  ld[key]['efec'] += r['efec_min']
    ld[key]['plan'] += r['plan_min']; ld[key]['nopl'] += r['nopl_min']
line_disp = [{'key': k, 'linea': k.split('|')[0], 'periodo': k.split('|')[1], **v}
             for k, v in ld.items()]

# ── 8. EMBEBER EN TEMPLATE ────────────────────────────────────────────────────
print("8. Generando docs/index.html...")
with open(TEMPLATE, encoding='utf-8') as f:
    html = f.read()

for placeholder, data in [
    ('/*PROD_DATA*/', json.dumps(prod)),
    ('/*PIDX_DATA*/', json.dumps(pidx)),
    ('/*PROG_DATA*/', json.dumps(prog)),
    ('/*PD_DATA*/',   json.dumps(perdidas)),
    ('/*TD_DATA*/',   json.dumps(teorico)),
    ('/*IDX_DATA*/',  json.dumps(pidx)),
    ('/*LD_DATA*/',   json.dumps(line_disp)),
    ('/*SEG_DATA*/',  json.dumps(seg)),
]:
    html = html.replace(placeholder, data)

out_path = OUT_DIR / 'index.html'
with open(out_path, 'w', encoding='utf-8') as f:
    f.write(html)

size_kb = len(html) // 1024
print(f"   ✅ docs/index.html — {size_kb} KB")
print()
print("=== RESUMEN ===")
print(f"  PROD:    {len(prod)} registros")
print(f"  PROG:    {len(prog)} registros")
print(f"  TP:      {len(perdidas)} registros")
print(f"  SEG:     {len(seg)} registros")
print(f"  TEORICO: {len(teorico)} registros")
print(f"  Output:  {size_kb} KB")
