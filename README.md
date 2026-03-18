# Mapeo de Nomencladores Médicos

Script Python para cruzar prácticas médicas del archivo `rossi.csv` contra dos nomencladores:

- **NN.csv** — Nomenclador Nacional
- **NN_OSMISS.csv** — Nomenclador OSMISS

## Archivos

| Archivo | Descripción |
|---|---|
| `rossi.csv` | Listado de prácticas con código y nomenclador de referencia |
| `NN.csv` | Nomenclador Nacional |
| `NN_OSMISS.csv` | Nomenclador OSMISS |
| `mapeo.py` | Script principal de cruce |

## Uso

```bash
pip install pandas openpyxl
python mapeo.py
```

## Salida

- `resultado_mapeo.csv` — Resultado del cruce en formato CSV
- `resultado_mapeo.xlsx` — Resultado con columnas de código y descripción del nomenclador, columna **Estado** (`Encontrado` / `No Encontrado`), y filas no encontradas marcadas en **rojo con fondo rosa**

## Lógica

- Filas con `Nomenclador == "NN"` → se cruzan con `NN.csv`
- Filas con `Nomenclador == "N OSMISS"` → se cruzan con `NN_OSMISS.csv`
- Si el código no se encuentra en ninguno → Estado `No Encontrado`
