# Mapeo de Nomencladores Médicos

Sistema para cruzar prácticas médicas contra múltiples nomencladores, con interfaz gráfica de usuario (Tkinter).

## Archivos de datos

| Archivo | Descripción |
|---|---|
| `rossi.csv` | Listado de prácticas con código y nomenclador de referencia |
| `NN.csv` | Nomenclador Nacional |
| `NN_OSMISS.csv` | Nomenclador OSMISS |
| `NBU.csv` | Nomenclador Bioquímico Único |
| `sanatorio_del_oeste.csv` | Prácticas del Sanatorio del Oeste con rangos de códigos y copagos |
| `centro.csv` | Prácticas del Centro Médico para cruce con NBU |

## Scripts

| Script | Descripción |
|---|---|
| `mapeo.py` | Cruce de `rossi.csv` contra NN y NN_OSMISS |
| `sanatorio_nn_osmiss.py` | Cruce del sanatorio contra NN y N OSMISS |
| `sanatorio_nbu.py` | Cruce del sanatorio contra NBU con búsqueda fuzzy |
| `buscar_nbu.py` | Cruce de `centro.csv` contra NBU con TF-IDF |
| `app.py` | Interfaz gráfica Tkinter con todos los procesos |

## Interfaz gráfica (`app.py`)

```bash
python app.py
```

Ventana principal con 4 procesos:

### 1. Mapeo Centro Médico
- Archivos de entrada: `rossi.csv`, `NN.csv`, `NN_OSMISS.csv`
- Cruza por código según columna `Nomenclador` (`NN` → NN.csv / `N OSMISS` → NN_OSMISS.csv)
- Salida: `resultado_mapeo.csv` y `resultado_mapeo.xlsx`

### 2. Sanatorio NN / N OSMISS
- Archivos de entrada: `sanatorio_del_oeste.csv`, `NN.csv`, `NN_OSMISS.csv`
- Expande rangos de códigos (`010101 al 130304` → una fila por código)
- Suma gastos adicionales (Galeno, Otros Gastos) al `copago_especial` de la fila principal
- Salida: `resultado_sanatorio_nn_osmiss.csv` y `resultado_sanatorio_nn_osmiss.xlsx`

### 3. NBU
- Archivos de entrada: `sanatorio_del_oeste.csv`, `NBU.csv`
- Procesa filas con nomenclador `3` (NBU)
- Si el código existe en NBU lo usa directamente; si no, busca por similitud de nombre (difflib, umbral 0.4)
- Salida: `resultado_sanatorio_nbu.csv` y `resultado_sanatorio_nbu.xlsx`

### 4. Laboratorios NBU
- Archivos de entrada: `centro.csv`, `NBU.csv`
- Usa TF-IDF + cosine similarity (scikit-learn, umbral 0.5) para encontrar la determinación más parecida
- Puede retornar múltiples coincidencias por práctica ordenadas por similitud descendente
- Salida: `resultado_laboratorios_nbu.csv` y `resultado_laboratorios_nbu.xlsx`

## Funcionalidades de la interfaz

- Selector de archivos de entrada por proceso
- Selector de carpeta de salida (el usuario elige dónde guardar)
- Barra de progreso animada durante el procesamiento
- Log en tiempo real con mensajes de avance y resumen final
- Botón para abrir la carpeta de resultados al finalizar
- Acceso directo en el escritorio (`Mapeo Nomencladores.lnk`)

## Lógica general de cruce

- Filas `Nomenclador == "NN"` → se cruzan con `NN.csv`
- Filas `Nomenclador == "N OSMISS"` → se cruzan con `NN_OSMISS.csv`
- Filas `Nomenclador == "NBU"` → se cruzan con `NBU.csv`
- Si el código no se encuentra → Estado `No Encontrado`
- Filas **No Encontrado** → marcadas en **texto rojo y fondo rosa** en el Excel

## Columnas de salida (Excel / CSV)

| Columna | Descripción |
|---|---|
| CUIT | CUIT del prestador |
| Nombre | Nombre del prestador |
| Práctica | Descripción de la práctica original |
| Código | Código original |
| Nomenclador | Nomenclador de referencia |
| Copago Especial | Valor del copago (incluye suma de gastos adicionales) |
| Código Nomenclador | Código encontrado en el nomenclador |
| Descripción Nomenclador | Descripción del nomenclador |
| Estado | `Encontrado` / `No Encontrado` |

## Instalación de dependencias

```bash
pip install pandas openpyxl scikit-learn
```

## Acceso directo en el escritorio

Para regenerar el acceso directo ejecutar:

```bash
cscript crear_acceso_directo.vbs
```

## Estructura del proyecto

```
Mapeo/
├── app.py                        # Interfaz gráfica principal
├── mapeo.py                      # Script Mapeo Centro Médico
├── sanatorio_nn_osmiss.py        # Script Sanatorio NN / N OSMISS
├── sanatorio_nbu.py              # Script Sanatorio NBU
├── buscar_nbu.py                 # Script Laboratorios NBU
├── crear_acceso_directo.vbs      # Generador de acceso directo
├── NN.csv                        # Nomenclador Nacional
├── NN_OSMISS.csv                 # Nomenclador OSMISS
├── NBU.csv                       # Nomenclador Bioquímico Único
├── rossi.csv                     # Prácticas Centro Médico
├── sanatorio_del_oeste.csv       # Prácticas Sanatorio del Oeste
└── centro.csv                    # Prácticas laboratorio para NBU
```
