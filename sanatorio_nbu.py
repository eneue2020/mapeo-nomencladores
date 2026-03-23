import pandas as pd
from difflib import SequenceMatcher

def leer_csv(path):
    df = pd.read_csv(path, sep=';', header=0, dtype=str, encoding='latin-1')
    df.columns = df.columns.str.strip()
    df = df.dropna(how='all')
    return df

def limpiar_numero(valor):
    if pd.isna(valor):
        return 0.0
    v = str(valor).strip().replace('.', '').replace(',', '.')
    try:
        return float(v)
    except ValueError:
        return 0.0

def similitud(a, b):
    return SequenceMatcher(None, a.upper().strip(), b.upper().strip()).ratio()

def buscar_nbu(practica, dict_nbu):
    mejor_cod = ''
    mejor_desc = ''
    mejor_score = 0.0
    for cod, desc in dict_nbu.items():
        score = similitud(practica, desc)
        if score > mejor_score:
            mejor_score = score
            mejor_cod   = cod
            mejor_desc  = desc
    return mejor_cod, mejor_desc, mejor_score

# --- Cargar ---
print('Cargando archivos...')
sanatorio = leer_csv('sanatorio_del_oeste.csv')
nbu       = leer_csv('NBU.csv')

nbu['CODIGO']          = nbu['CODIGO'].str.strip()
nbu['DETERMINACIONES'] = nbu['DETERMINACIONES'].str.strip()
dict_nbu = dict(zip(nbu['CODIGO'], nbu['DETERMINACIONES']))

print('  sanatorio : %d filas' % len(sanatorio))
print('  NBU       : %d entradas' % len(dict_nbu))

# --- Agrupar solo filas NBU (nomenclador == 3) ---
print('Agrupando filas NBU...')
grupos = []
fila_principal   = None
suma_adicionales = 0.0

for _, row in sanatorio.iterrows():
    cod  = str(row.get('codigo', '')).strip()
    nom  = str(row.get('nomenclador', '')).strip()
    copa = limpiar_numero(row.get('copago_especial', 0))
    es_adicional = (cod == '0' and nom in ['0', '4'])
    if es_adicional:
        suma_adicionales += copa
    else:
        if fila_principal is not None:
            nom_fp = str(fila_principal.get('nomenclador', '')).strip()
            if nom_fp == '3':
                grupos.append((fila_principal, suma_adicionales))
        fila_principal   = row
        suma_adicionales = 0.0

if fila_principal is not None:
    nom_fp = str(fila_principal.get('nomenclador', '')).strip()
    if nom_fp == '3':
        grupos.append((fila_principal, suma_adicionales))

print('  Grupos NBU: %d' % len(grupos))

# --- Procesar ---
print('Procesando grupos NBU...')
UMBRAL = 0.4
filas_resultado = []
total = len(grupos)

for i, (row, suma_adic) in enumerate(grupos):
    if (i+1) % 10 == 0 or (i+1) == total:
        print('  Grupo %d / %d' % (i+1, total))

    codigo       = str(row.get('codigo', '')).strip()
    nom_col_ext  = str(row.get('NOMENCLADORES', '')).strip()
    copa_base    = limpiar_numero(row.get('copago_especial', 0))
    copago_final = copa_base + suma_adic
    cuit         = str(row.get('cuit', '')).strip()
    nombre       = str(row.get('nombre', '')).strip()
    practicas    = str(row.get('Practicas', '')).strip()

    if codigo != '' and codigo != '0':
        if codigo in dict_nbu:
            cod_nom = codigo; desc_nom = dict_nbu[codigo]; estado = 'Encontrado'
        else:
            cod_nom, desc_nom, score = buscar_nbu(practicas, dict_nbu)
            estado = 'Encontrado' if score >= UMBRAL else 'No Encontrado'
            if score < UMBRAL:
                cod_nom = desc_nom = ''
    else:
        cod_nom, desc_nom, score = buscar_nbu(practicas, dict_nbu)
        estado = 'Encontrado' if score >= UMBRAL else 'No Encontrado'
        if score < UMBRAL:
            cod_nom = desc_nom = ''

    filas_resultado.append({
        'CUIT':                    cuit,
        'Nombre':                  nombre,
        'Practica':                practicas,
        'Codigo':                  codigo if codigo != '0' else '',
        'Nomenclador':             nom_col_ext,
        'Copago Especial':         round(copago_final, 2),
        'Codigo Nomenclador':      cod_nom,
        'Descripcion Nomenclador': desc_nom,
        'Estado':                  estado,
    })

df = pd.DataFrame(filas_resultado)
df.to_csv('resultado_sanatorio_nbu.csv', sep=';', index=False, encoding='latin-1')
print('CSV guardado: resultado_sanatorio_nbu.csv')

encontrados    = (df['Estado'] == 'Encontrado').sum()
no_encontrados = (df['Estado'] == 'No Encontrado').sum()
print('\n=== RESUMEN NBU ===')
print('  Total filas    : %d' % len(df))
print('  Encontrados    : %d' % encontrados)
print('  No encontrados : %d' % no_encontrados)
