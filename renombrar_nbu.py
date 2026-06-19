import pandas as pd

df = pd.read_csv('NBU.csv', sep=';', dtype=str, encoding='latin-1')
print('Columnas antes:', df.columns.tolist())

# Renombrar columna con espacios
df.columns = [c.strip() for c in df.columns]
df = df.rename(columns={
    'D E T E R M I N A C I O N E S': 'Determinaciones',
    'COD': 'CODIGO',
    'U. B.': 'UB',
    'Ref.': 'Ref',
})
print('Columnas despues:', df.columns.tolist())
df.to_csv('NBU.csv', sep=';', index=False, encoding='latin-1')
print('NBU.csv guardado OK')
