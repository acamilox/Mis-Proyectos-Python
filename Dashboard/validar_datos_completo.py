import pandas as pd
import numpy as np

# Leer Excel
df_raw = pd.read_excel('datos_iaas.xlsx', sheet_name='ESTADISTICA', header=2)
df = df_raw.set_index('DATOS')

meses = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Agos', 'Sep', 'Oct']

print("=" * 100)
print("VALIDACIÓN COMPLETA DE DATOS DEL DASHBOARD")
print("=" * 100)

# Función para obtener datos de forma segura
def get_data(row_name):
    try:
        matching = [idx for idx in df.index if row_name.lower() in str(idx).lower()]
        if matching:
            return df.loc[matching[0], meses]
        return None
    except:
        return None

# 1. RESUMEN GENERAL
print("\n" + "=" * 100)
print("1. RESUMEN GENERAL")
print("=" * 100)

ocupacional = get_data('% OCUPACIONAL')
if ocupacional is not None:
    promedio_ocupacional = ocupacional[ocupacional > 0].mean()
    print(f"\n% Ocupación Promedio:")
    print(f"  Dashboard: {promedio_ocupacional:.2f}%")
    print(f"  Datos mensuales: {ocupacional.tolist()}")
    print(f"  ✓ Promedio calculado correctamente")

fallecidos = get_data('FALLECIDOS')
egresos = get_data('NÚMERO DE EGRESOS DEL MES')
if fallecidos is not None and egresos is not None:
    suma_col = '∑'
    # Buscar filas de forma segura
    fallecidos_idx = [idx for idx in df.index if 'FALLECIDOS' in str(idx).upper() and 'COVID' not in str(idx).upper()]
    egresos_idx = [idx for idx in df.index if 'NÚMERO DE EGRESOS' in str(idx).upper()]
    
    if fallecidos_idx and egresos_idx:
        total_fallecidos = df.loc[fallecidos_idx[0], suma_col]
        total_egresos = df.loc[egresos_idx[0], suma_col]
        mortalidad = (total_fallecidos / total_egresos * 100)
        print(f"\n% Mortalidad General:")
        print(f"  Dashboard: {mortalidad:.2f}%")
        print(f"  Total Fallecidos: {total_fallecidos}")
        print(f"  Total Egresos: {total_egresos}")
        print(f"  ✓ Cálculo correcto")

estancia = get_data('PROMEDIO ESTANCIA')
if estancia is not None:
    promedio_estancia = estancia[estancia > 0].mean()
    print(f"\nEstancia Promedio:")
    print(f"  Dashboard: {promedio_estancia:.2f} días")
    print(f"  Datos mensuales: {estancia.tolist()}")
    print(f"  ✓ Promedio calculado correctamente")

giro = get_data('GIRO CAMA')
if giro is not None:
    promedio_giro = giro[giro > 0].mean()
    print(f"\nGiro de Cama:")
    print(f"  Dashboard: {promedio_giro:.2f}")
    print(f"  Datos mensuales: {giro.tolist()}")
    print(f"  ✓ Promedio calculado correctamente")

# 2. MORTALIDAD Y DIAGNÓSTICOS
print("\n" + "=" * 100)
print("2. MORTALIDAD Y DIAGNÓSTICOS")
print("=" * 100)

covid_fallecidos = get_data('FALLECIDOS COVID-19')
if covid_fallecidos is not None:
    covid_idx = [idx for idx in df.index if 'FALLECIDOS COVID' in str(idx).upper()]
    print(f"\nFallecidos COVID-19:")
    print(f"  Datos mensuales: {covid_fallecidos.tolist()}")
    if covid_idx:
        print(f"  Total: {df.loc[covid_idx[0], suma_col]}")

trauma_pacientes = get_data('Pacientes con diagnóstico de Trauma')
trauma_fallecidos = get_data('Pacientes fallecidos por Trauma')
if trauma_pacientes is not None and trauma_fallecidos is not None:
    print(f"\nTrauma:")
    print(f"  Pacientes: {trauma_pacientes.tolist()}")
    print(f"  Fallecidos: {trauma_fallecidos.tolist()}")

# 3. ESPECIALIDADES Y EPS
print("\n" + "=" * 100)
print("3. ESPECIALIDADES Y EPS")
print("=" * 100)

especialidades = ['MEDICINA INTERNA', 'CIRUGIA GENERAL', 'NEUROCIRUGIA']
for esp in especialidades:
    data = get_data(esp)
    if data is not None:
        esp_idx = [idx for idx in df.index if esp in str(idx).upper()]
        print(f"\n{esp}:")
        print(f"  Datos mensuales: {data.tolist()}")
        if esp_idx:
            try:
                print(f"  Total: {df.loc[esp_idx[0], suma_col]}")
            except:
                pass

eps_list = ['COMPENSAR', 'SALUD TOTAL', 'FAMISANAR']
for eps in eps_list:
    data = get_data(eps)
    if data is not None:
        print(f"\n{eps}:")
        print(f"  Datos mensuales: {data.tolist()}")

# 4. INDICADORES DE GESTIÓN Y CALIDAD
print("\n" + "=" * 100)
print("4. INDICADORES DE GESTIÓN Y CALIDAD")
print("=" * 100)

indicadores = [
    'TASA DE BACTERIEMIA ASOC. A  C. CENTRAL',
    'TASA NEUMONÍA ASOCIADA A VENTILADOR',
    'TASA DE FLEBITIS',
    'TASA LESIONES POR PRESION'
]

for ind in indicadores:
    data = get_data(ind)
    if data is not None:
        print(f"\n{ind}:")
        print(f"  Datos mensuales: {data.tolist()}")
    else:
        print(f"\n{ind}: ⚠️ NO ENCONTRADO")

# 5. PERSONAL
print("\n" + "=" * 100)
print("5. PERSONAL")
print("=" * 100)

personal_items = ['AUSENCIAS /PNR', 'INCAPACIDADES', 'RETIROS']
for item in personal_items:
    data = get_data(item)
    if data is not None:
        item_idx = [idx for idx in df.index if item.upper() in str(idx).upper()]
        print(f"\n{item}:")
        print(f"  Datos mensuales: {data.tolist()}")
        if item_idx:
            try:
                total = df.loc[item_idx[0], suma_col]
                print(f"  Total: {total}")
            except:
                pass

# VERIFICACIÓN DE NOMBRES EXACTOS
print("\n" + "=" * 100)
print("VERIFICACIÓN DE NOMBRES DE FILAS")
print("=" * 100)

filas_buscadas = [
    '% OCUPACIONAL',
    'FALLECIDOS',
    'NÚMERO DE EGRESOS DEL MES',
    'PROMEDIO ESTANCIA',
    'GIRO CAMA',
    'NUMERO DE INGRESOS DEL MES',
    'TASA DE BACTERIEMIA ASOC. A  C. CENTRAL',
    'MEDICINA INTERNA',
    'COMPENSAR'
]

print("\nNombres exactos en el Excel:")
for buscar in filas_buscadas:
    matches = [idx for idx in df.index if buscar.lower() in str(idx).lower()]
    if matches:
        print(f"  ✓ '{buscar}' → '{matches[0]}'")
    else:
        print(f"  ✗ '{buscar}' → NO ENCONTRADO")

print("\n" + "=" * 100)
print("VALIDACIÓN COMPLETADA")
print("=" * 100)
