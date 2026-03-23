import pandas as pd

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

def expandir_codigos(codigo, copago, cuit, nombre, nom_col, nom_col_ext, practicas):
    filas = []
    if ' al ' in str(codigo):
        partes  = str(codigo).split(' al ')
        cod_ini = partes[0].strip()
        cod_fin = partes[1].strip()
        try:
            ini   = int(cod_ini)
            fin   = int(cod_fin)
            largo = len(cod_ini)
            for cod in range(ini, fin + 1):
                filas.append({'codigo': str(cod).zfill(largo), 'copago_especial': copago,
                              'cuit': cuit, 'nombre': nombre, 'nomenclador': nom_col,
                              'NOMENCLADORES': nom_col_ext, 'Practicas': practicas})
        except ValueError:
            filas.append({'codigo': codigo, 'copago_especial': copago,
                          'cuit': cuit, 'nombre': nombre, 'nomenclador': nom_col,
                          'NOMENCLADORES': nom_col_ext, 'Practicas': practicas})
    else:
        filas.append({'codigo': codigo, 'copago_especial': copago,
                      'cuit': cuit, 'nombre': nombre, 'nomenclador': nom_col,
                      'NOMENCLADORES': nom_col_ext, 'Practicas': practicas})
    return filas

# --- Cargar ---
print('Cargando archivos...')
sanatorio = leer_csv('sanatorio_del_oeste.csv')
nn        = leer_csv('NN.csv')
nn_osmiss = leer_csv('NN_OSMISS.csv')

nn['Codigo Nomenclador'] = nn['C\xf3digo Nomenclador'].str.strip()
nn['Descripcion']        = nn['Descripci\xf3n'].str.strip()
nn_osmiss['Codigo']      = nn_osmiss['C\xf3digo'].str.strip()
nn_osmiss['Descripcion'] = nn_osmiss['Descripci\xf3n'].str.strip()

dict_nn     = dict(zip(nn['Codigo Nomenclador'], nn['Descripcion']))
dict_osmiss = dict(zip(nn_osmiss['Codigo'], nn_osmiss['Descripcion']))

print('  sanatorio : %d filas' % len(sanatorio))
print('  NN        : %d entradas' % len(dict_nn))
print('  NN_OSMISS : %d entradas' % len(dict_osmiss))

# --- Agrupar ---
print('Agrupando filas...')
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
            if nom_fp in ['2', '4']:
                grupos.append((fila_principal, suma_adicionales))
        fila_principal   = row
        suma_adicionales = 0.0

if fila_principal is not None:
    nom_fp = str(fila_principal.get('nomenclador', '')).strip()
    if nom_fp in ['2', '4']:
        grupos.append((fila_principal, suma_adicionales))

print('  Grupos NN + N OSMISS: %d' % len(grupos))

# --- Expandir y buscar ---
print('Procesando grupos...')
filas_resultado = []
total = len(grupos)

for i, (row, suma_adic) in enumerate(grupos):
    if (i+1) % 25 == 0 or (i+1) == total:
        print('  Grupo %d / %d' % (i+1, total))

    codigo       = str(row.get('codigo', '')).strip()
    nom_col      = str(row.get('nomenclador', '')).strip()
    nom_col_ext  = str(row.get('NOMENCLADORES', '')).strip()
    copa_base    = limpiar_numero(row.get('copago_especial', 0))
    copago_final = copa_base + suma_adic
    cuit         = str(row.get('cuit', '')).strip()
    nombre       = str(row.get('nombre', '')).strip()
    practicas    = str(row.get('Practicas', '')).strip()

    if codigo == '' or codigo == '0':
        continue

    for f in expandir_codigos(codigo, copago_final, cuit, nombre, nom_col, nom_col_ext, practicas):
        cod_exp = str(f['codigo']).strip()
        nom_exp = str(f['nomenclador']).strip()

        if nom_exp == '2' and cod_exp in dict_nn:
            cod_nom = cod_exp; desc_nom = dict_nn[cod_exp]; estado = 'Encontrado'
        elif nom_exp == '4' and cod_exp in dict_osmiss:
            cod_nom = cod_exp; desc_nom = dict_osmiss[cod_exp]; estado = 'Encontrado'
        elif cod_exp in dict_nn:
            cod_nom = cod_exp; desc_nom = dict_nn[cod_exp]; estado = 'Encontrado'
        elif cod_exp in dict_osmiss:
            cod_nom = cod_exp; desc_nom = dict_osmiss[cod_exp]; estado = 'Encontrado'
        else:
            cod_nom = desc_nom = ''; estado = 'No Encontrado'

        filas_resultado.append({
            'CUIT':                    f['cuit'],
            'Nombre':                  f['nombre'],
            'Practica':                f['Practicas'],
            'Codigo':                  cod_exp,
            'Nomenclador':             f['NOMENCLADORES'],
            'Copago Especial':         round(f['copago_especial'], 2),
            'Codigo Nomenclador':      cod_nom,
            'Descripcion Nomenclador': desc_nom,
            'Estado':                  estado,
        })

df = pd.DataFrame(filas_resultado)
df.to_csv('resultado_sanatorio_nn_osmiss.csv', sep=';', index=False, encoding='latin-1')
print('CSV guardado: resultado_sanatorio_nn_osmiss.csv')

encontrados    = (df['Estado'] == 'Encontrado').sum()
no_encontrados = (df['Estado'] == 'No Encontrado').sum()
print('\n=== RESUMEN NN + N OSMISS ===')
print('  Total filas    : %d' % len(df))
print('  Encontrados    : %d' % encontrados)
print('  No encontrados : %d' % no_encontrados)
