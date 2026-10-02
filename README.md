# Sistema de Planillas - Camboriú 2026

Genera planillas Excel de operaciones y programación a partir de los datos de
hasta 3 tipos de entrada diferentes (Receptivo, Operativo y Word). Produce el
mismo estilo de salida que los archivos modelo usados en la empresa.

## Estructura del proyecto

```
Marcaciones/
├── Planillas_GUI.exe          # Ejecutable: doble clic y a trabajar (no requiere Python)
├── RECORDATORIO.txt           # Nota: qué generador usar según el estilo que mande el cliente
├── requirements.txt           # Dependencias (solo para desarrolladores)
├── icono.ico                  # Ícono del ejecutable
├── codigo/                    # Código fuente (3 generadores + interfaz)
│   ├── gui_receptivo.py       #   Interfaz gráfica
│   ├── generador_receptivo.py #   Estilo PANAM  -> tipo "Receptivo"
│   ├── generador_operativo.py #   Estilo INKA   -> tipo "Operativo"
│   ├── generador_word.py      #   Estilo Word   -> tipo "Word"
│   ├── formato_modelo.py      #   Formatos compartidos
│   └── generar_icono.py       #   Regenera icono.ico
├── plantillas/                # Formato base de las planillas
│   ├── cristour_template.xlsx
│   └── inka_template.xlsx
├── entradas/                  # Datos de entrada (se ponen acá los archivos)
│   ├── COLEGIOS CANTIDAD DE PASAJEROS, DISCOS, MARCACIONES (3).xlsx
│   ├── Rooming 2026.xlsx
│   ├── Control Operativo.xlsx
│   └── OPERACIONES ... .docx
├── salidas/                   # Planillas generadas (no se sube a git)
└── manual/                    # Manual de usuario + imágenes
```

## Uso rápido

1. Abrir **`Planillas_GUI.exe`** (no requiere instalar nada).
2. Elegir el **tipo de planilla**: Receptivo, Operativo o Word.
3. Verificar los archivos de entrada/plantilla/salida (ya aparecen precargados).
4. Presionar **"Generar planillas"**, revisar la vista previa y **guardar**.

### Qué tipo usar según el estilo que llegue

| Estilo que manda el cliente | Tipo en el programa | Archivo de entrada |
|---|---|---|
| PANAM | Receptivo | `entradas/COLEGIOS CANTIDAD DE PASAJEROS...xlsx` |
| INKA | Operativo | `entradas/Control Operativo.xlsx` (hoja Reservas) |
| Operación en Word | Word | `entradas/OPERACIONES ... .docx` |

El **Operativo** detecta el Control Operativo automáticamente: usa el archivo
`Control Operativo.xlsx` si existe, o cualquier `.xlsx` de `entradas/` que tenga
una hoja llamada `Reservas`.

## Uso por línea de comandos (desarrolladores)

Requisitos: Python 3 + `pip install -r requirements.txt` (solo `openpyxl`).

```bash
cd codigo
python generador_receptivo.py
python generador_word.py
python generador_operativo.py
```

Cada script lee de `entradas/`, usa la plantilla de `plantillas/` y guarda en
`salidas/`.

## Regenerar el ejecutable

```bash
cd codigo
python -m PyInstaller --noconfirm --onefile --windowed \
    --name Planillas --icon ..\icono.ico gui_receptivo.py
copy dist\Planillas.exe ..
```

Más detalles en `manual/manual.html`.