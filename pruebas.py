"""
DIAGNÓSTICO DE TRAZABILIDAD DE SERIALES
========================================
Ejecuta este script y pega el resultado aquí para identificar
exactamente por qué algunos seriales no cruzan entre hojas.

Uso:
    python diagnostico_seriales.py
"""

import pandas as pd
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ruta_sistem = os.path.join(BASE_DIR, 'datos', 'SISTEM.xlsx')

print("=" * 60)
print("CARGANDO HOJAS DE SISTEM.xlsx...")
print("=" * 60)

doc_envios       = pd.read_excel(ruta_sistem, sheet_name='ENVIOS')
doc_entradas     = pd.read_excel(ruta_sistem, sheet_name='ENTRADAS')
doc_devoluciones = pd.read_excel(ruta_sistem, sheet_name='DEVOLUCIONES')
doc_salidas      = pd.read_excel(ruta_sistem, sheet_name='SALIDAS')
doc_entregas     = pd.read_excel(ruta_sistem, sheet_name='ENTREGAS')

# Normalizar columnas
for df in [doc_entradas, doc_devoluciones, doc_salidas, doc_entregas, doc_envios]:
    df.columns = df.columns.str.strip()

# ── 1. TIPOS DE DATO REALES ───────────────────────────────────────────────────
print("\n[1] TIPO DE DATO DEL SERIAL POR HOJA")
print("-" * 40)
print(f"  ENVIOS       → NºSerieFab : {doc_envios['NºSerieFab'].dtype}")
print(f"  ENTRADAS     → Serial     : {doc_entradas['Serial'].dtype}")
print(f"  DEVOLUCIONES → Serial     : {doc_devoluciones['Serial'].dtype}")
print(f"  SALIDAS      → Serial     : {doc_salidas['Serial'].dtype}")
print(f"  ENTREGAS     → Serial     : {doc_entregas['Serial'].dtype}")

# ── 2. LIMPIAR Y CONSTRUIR SETS ───────────────────────────────────────────────
def limpiar(serie):
    return serie.astype(str).str.strip().str.replace(r'\s+', ' ', regex=True)

envios_seriales    = set(limpiar(doc_envios['NºSerieFab']))
entradas_seriales  = set(limpiar(doc_entradas['Serial']))
dev_seriales       = set(limpiar(doc_devoluciones['Serial']))
salidas_seriales   = set(limpiar(doc_salidas['Serial']))
entregas_seriales  = set(limpiar(doc_entregas['Serial']))

todos_movimientos  = entradas_seriales | dev_seriales | salidas_seriales | entregas_seriales

# ── 3. SERIALES QUE DEBERÍAN CRUZAR PERO NO CRUZAN ───────────────────────────
sin_cruce = envios_seriales - todos_movimientos
sin_cruce = {s for s in sin_cruce if s not in ('nan', 'NaN', '', '#N/D', '#N/A')}

print(f"\n[2] SERIALES EN ENVIOS SIN NINGÚN MOVIMIENTO: {len(sin_cruce)}")
print("-" * 40)

# Mostrar los primeros 10 con análisis detallado
muestra = list(sin_cruce)[:10]
for serial in muestra:
    print(f"\n  Serial problemático : {repr(serial)}")
    print(f"  Longitud            : {len(serial)}")
    print(f"  Bytes (utf-8)       : {serial.encode('utf-8')}")

    # Buscar si existe algo parecido en movimientos (fuzzy manual)
    candidatos = [s for s in todos_movimientos if serial.strip() in s or s in serial.strip()]
    if candidatos:
        print(f"  ⚠️  Posibles matches cercanos en movimientos:")
        for c in candidatos[:3]:
            print(f"      → {repr(c)}  (longitud: {len(c)})")
    else:
        print(f"  ✗  No existe ningún serial parecido en movimientos")

# ── 4. MUESTRA DE SERIALES QUE SÍ CRUZAN (control) ──────────────────────────
si_cruzan = envios_seriales & todos_movimientos
si_cruzan = {s for s in si_cruzan if s not in ('nan', 'NaN', '')}

print(f"\n[3] SERIALES EN ENVIOS QUE SÍ CRUZAN CON MOVIMIENTOS: {len(si_cruzan)}")
print("-" * 40)
muestra_ok = list(si_cruzan)[:3]
for serial in muestra_ok:
    print(f"  ✓ {repr(serial)}")

# ── 5. RESUMEN ────────────────────────────────────────────────────────────────
print("\n" + "=" * 60)
print("RESUMEN")
print("=" * 60)
print(f"  Total seriales en ENVIOS       : {len(envios_seriales)}")
print(f"  Con al menos 1 movimiento      : {len(si_cruzan)}")
print(f"  Sin ningún movimiento (falla)  : {len(sin_cruce)}")
print(f"  % de falla                     : {len(sin_cruce)/len(envios_seriales)*100:.1f}%")
print("=" * 60)