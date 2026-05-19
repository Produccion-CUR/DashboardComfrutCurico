#!/usr/bin/env python3
"""
generate_data.py — Genera todos los JSON desde el Excel de Comfrut
Uso: python3 generate_data.py [ruta_excel]
"""
import sys, json, datetime
import pandas as pd
from collections import defaultdict
from pathlib import Path

# ─── Rutas ──────────────────────────────────────────────────────────────────
EXCEL = sys.argv[1] if len(sys.argv)>1 else \
        next(Path('/mnt/user-data/uploads').glob('*.xlsx'), None)
if not EXCEL:
    print("ERROR: no se encontró Excel"); sys.exit(1)
EXCEL = str(EXCEL)
OUT   = Path('/home/claude')
print(f"Leyendo: {EXCEL}")

xl  = pd.ExcelFile(EXCEL)

def n(x):
    try: return float(x) if pd.notna(x) else 0.0
    except: return 0.0

def s(x):
    return str(x).strip() if pd.notna(x) and str(x).strip()!='nan' else ''

def fechastr(x):
    try:
        t = pd.Timestamp(x)
        return t.strftime('%Y-%m-%d') if pd.notna(t) else ''
    except: return ''

ZCIQ_L = ['S_L01','S_L03','S_L04','S_L05','S_LMANCU','S_MANUAL']
ZENV_L = ['S_ENV1']

# ─── 1. PRODUCCION ───────────────────────────────────────────────────────────
print("1. Asistencia + Produccion...")
df_ap = pd.read_excel(xl, 'Asistencia + Produccion', dtype={'Material':str})
df_ap['fecha'] = df_ap['Inic.tratamiento'].apply(fechastr)
df_ap = df_ap[df_ap['fecha'].str.len()==10].copy()

prod = []
for _, r in df_ap.iterrows():
    fecha = r['fecha']; p = fecha.split('-')
    linea = s(r['Pto. Trabajo'])
    area  = 'ZENV' if linea=='S_ENV1' else 'ZCIQ'
    prod.append({
        'fecha': fecha,
        'año': int(p[0]), 'mes': int(p[1]), 'dia': int(p[2]),
        'semana': int(n(r['Semana'])),
        'turno':  int(n(r['Turno'])),
        'linea':  linea, 'area': area,
        'especie': s(r['Especie']),
        'sku':     s(r['Desc.Material']),
        'cod':     s(r['Material']),
        'teo_min':   n(r['T.Minutos']),
        'efec_min':  n(r['Tiempo Efec.Min.']),
        'plan_min':  n(r['Paros Plan Min.']),
        'nopl_min':  n(r['Paros No Plan Min.']),
        'kg_ingresados': n(r['Kilos Ingresados']),
        'iqf_aprobado':  n(r['IQF Aprobado']),
        'kg_puro':   n(r['Kilos Pure']),
        'kg_jugo':   n(r['Kilos Jugo']),
        'kg_crumble': n(r['Kilos Crumble']),
        'ton_real':  n(r['Ton.Real']),
        'cajas':     n(r['Cajas Produc.']),
        'kg_aprobadas': n(r['Kilos Aprobadas']),
        'teo_cajas': n(r['Teorico Cajas']),
        'con_cajas': n(r['Consumo Cajas']),
        'teo_bolsas': n(r['Teorico Bolsas']),
        'con_bolsas': n(r['Consumo Bolsas']),
        'bpm_total': n(r['BPM Total']),
        'bpm_std':   n(r['BPM Estandar']),
        'bpm_sinpp': n(r['BPM sin PP']),
        'personas':  n(r['Cant.Personas']),
        'kg_h_pers': n(r['Produc.(Kg/H/Personas)']),
    })
print(f"   {len(prod)} registros")
with open(OUT/'produccion_full.json','w') as f: json.dump(prod, f)

# ─── 2. PROGRAMA ENVASADO ────────────────────────────────────────────────────
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
        if not fecha or len(fecha)!=10: continue
        try: turno = int(float(turno_row[ci])) if pd.notna(turno_row[ci]) else 0
        except: turno = 0
        if turno not in [1, 2]: continue  # Turno 3 = totales, no matchea PROD
        prog.append({'fecha':fecha,'cod':cod,'sku':sku,'turno':turno,'cajas_prog':val_f})

print(f"   {len(prog)} registros (turno 1/2 only)")
with open(OUT/'prog_env_full.json','w') as f: json.dump(prog, f)

# Validación
prod_zenv = [r for r in prod if r['area']=='ZENV' and r['cod']]
prog_keys_set = {(p['fecha'],p['cod'],p['turno']) for p in prog}
matched = [p for p in prog if any(r['fecha']==p['fecha'] and r['cod']==p['cod']
           and r['turno']==p['turno'] for r in prod_zenv)]
con_prog = sum(r['cajas'] for r in prod_zenv
               if (r['fecha'],r['cod'],r['turno']) in prog_keys_set)
tot_prog = sum(p['cajas_prog'] for p in matched)
print(f"   Validación: {len(matched)} matches, cumpl={con_prog/tot_prog*100:.1f}%" if tot_prog else "   Sin matches")

# ─── 3. TIEMPOS PERDIDOS ─────────────────────────────────────────────────────
print("3. Tiempos Perdidos...")
df_tp = pd.read_excel(xl, 'Tiempos Perdidos', dtype=str)
df_tp.columns = [c.strip() for c in df_tp.columns]
perdidas = []
for _, r in df_tp.iterrows():
    linea = s(r.get('Pto. Trabajo',''))
    if not linea: continue
    area = 'ZCIQ' if linea in ZCIQ_L else ('ZENV' if linea in ZENV_L else '')
    if not area: continue
    fecha = fechastr(r.get('Fecha',''))
    if not fecha or len(fecha)!=10: continue
    minutos = n(r.get('T.Minutos','0'))
    if minutos <= 0: continue
    p = fecha.split('-')
    sem = int(n(r.get('Semana','0')))
    if not sem:
        try:
            dt = datetime.date(int(p[0]),int(p[1]),int(p[2]))
            sem = dt.isocalendar()[1]
        except: sem = 0
    perdidas.append({
        'fecha': fecha,
        'año': int(p[0]), 'mes': int(p[1]), 'dia': int(p[2]),
        'semana': sem, 'linea': linea, 'area': area,
        'falla':     s(r.get('Desc.Falla','')),
        'categoria': s(r.get('Desc.Clasifi. del Paro','')),
        'tipo_paro': s(r.get('Tipo de Paro','')),
        'obs':       s(r.get('Observaciones','')),
        'minutos':   minutos,
        'turno':     int(n(r.get('Turno','0'))),
    })
print(f"   {len(perdidas)} registros")
with open(OUT/'perdidas_full.json','w') as f: json.dump(perdidas, f)

# ─── 4. TEORICO (desde Asistencia+Produccion, agrupado por fecha+linea) ──────
print("4. Teorico (desde AP, agrupado por fecha+linea)...")
grp_td = df_ap.groupby(['fecha','Pto. Trabajo']).agg({
    'T.Minutos':        'sum',
    'Tiempo Efec.Min.': 'sum',
    'Paros Plan Min.':  'sum',
    'Paros No Plan Min.': 'sum',
    'Semana':           'first',
}).reset_index()

teorico = []
for _, r in grp_td.iterrows():
    fecha = r['fecha']; p = fecha.split('-')
    try:
        dt  = datetime.date(int(p[0]),int(p[1]),int(p[2]))
        sem = dt.isocalendar()[1]
    except: sem = int(n(r['Semana']))
    min_val = n(r['T.Minutos'])
    if min_val <= 0: continue
    teorico.append({
        'fecha': fecha,
        'año': int(p[0]), 'mes': int(p[1]), 'dia': int(p[2]),
        'semana': sem,
        'linea':    s(r['Pto. Trabajo']),
        'minutos':  min_val,
        'efec_min': n(r['Tiempo Efec.Min.']),
        'plan_min': n(r['Paros Plan Min.']),
        'nopl_min': n(r['Paros No Plan Min.']),
    })
print(f"   {len(teorico)} registros")
with open(OUT/'teorico_full.json','w') as f: json.dump(teorico, f)

# ─── 5. PROD_IDX + INDEX_DATA (PIDX format: mes→sems, sem→dias) ─────────────
print("5. prod_idx / index_data (PIDX format)...")
def build_pidx(records):
    MS = defaultdict(set)  # mes→semanas
    MD = defaultdict(set)  # mes→dias
    SD = defaultdict(set)  # sem→dias
    for r in records:
        mes=r['mes']; sem=r['semana']; dia=r['dia']
        if mes and sem: MS[str(mes)].add(sem)
        if mes and dia: MD[str(mes)].add(dia)
        if sem and dia: SD[str(sem)].add(dia)
    return {'MS':{k:sorted(v) for k,v in MS.items()},
            'MD':{k:sorted(v) for k,v in MD.items()},
            'SD':{k:sorted(v) for k,v in SD.items()}}

zciq_r = [r for r in prod if r['area']=='ZCIQ']
zenv_r = [r for r in prod if r['area']=='ZENV']
all_idx = build_pidx(prod)
zciq_idx = build_pidx(zciq_r)
zenv_idx = build_pidx(zenv_r)

pidx = {
    'MS': all_idx['MS'],  'MD': all_idx['MD'],  'SD': all_idx['SD'],
    'MS_ZCIQ': zciq_idx['MS'], 'MD_ZCIQ': zciq_idx['MD'], 'SD_ZCIQ': zciq_idx['SD'],
    'MS_ZENV': zenv_idx['MS'], 'MD_ZENV': zenv_idx['MD'], 'SD_ZENV': zenv_idx['SD'],
}
# Both prod_idx and index_data use the same PIDX format (for filter dropdowns)
with open(OUT/'prod_idx.json','w') as f: json.dump(pidx, f)
with open(OUT/'index_data.json','w') as f: json.dump(pidx, f)
print(f"   MS_ZENV keys: {sorted(pidx['MS_ZENV'].keys())}")

# ─── 7. LINE_DISP ────────────────────────────────────────────────────────────
print("7. line_disp...")
ld = defaultdict(lambda:{'min':0.0,'efec':0.0,'plan':0.0,'nopl':0.0})
for r in teorico:
    key = f"{r['linea']}|{r['año']}-{str(r['mes']).zfill(2)}"
    ld[key]['min']  += r['minutos']
    ld[key]['efec'] += r['efec_min']
    ld[key]['plan'] += r['plan_min']
    ld[key]['nopl'] += r['nopl_min']
line_disp = [{'key':k,'linea':k.split('|')[0],'periodo':k.split('|')[1],**v}
             for k,v in ld.items()]
with open(OUT/'line_disp.json','w') as f: json.dump(line_disp, f)
print(f"   {len(line_disp)} entradas")

print()
print("=== RESUMEN FINAL ===")
for name, cnt in [('PROD',len(prod)),('PROG',len(prog)),
                   ('TP',len(perdidas)),('TD',len(teorico)),
                   ('prod_idx',len(pidx.get('MS',{}))),('line_disp',len(line_disp))]:
    print(f"  {name:12s}: {cnt} registros")
print()
print("✅ Todos los JSONs generados. Ejecutar ahora:")
print("   python3 /home/claude/build_combined.py")
