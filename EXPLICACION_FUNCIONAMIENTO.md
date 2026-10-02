# Cómo funciona el sistema + guía de problemas

Este documento explica qué lee y qué busca cada generador, para poder diagnosticar
si algo se rompe o da un resultado raro.

## Flujo general

```
VOS                       EL PROGRAMA
------------------------  --------------------------------------------------
1. Abrís Planillas_GUI.exe
2. Elegís tipo (Receptivo, Operativo o Word)
3. "Generar planillas"  ->  lee el archivo de ENTRADA
                          ->  abre la PLANTILLA (formato base)
                          ->  parsea los datos
                          ->  completa la planilla en memoria
                          ->  muestra vista previa
4. "Guardar"            ->  escribe el .xlsx en salidas/
```

Las salidas SIEMPRE se guardan en `salidas/`. Las plantillas se leen de
`plantillas/`. Los datos de entrada se leen de `entradas/`.

---

## 1. RECEPTIVO (estilo PANAM) -> `generador_receptivo.py`

**Lee 2 archivos:**
- `entradas/COLEGIOS CANTIDAD DE PASAJEROS, DISCOS, MARCACIONES (3).xlsx`
- `entradas/Rooming 2026.xlsx` (edades/cantidades reales por colegio)

**Hojas que usa dentro del Excel de colegios:**
- Hoja **"Colegios totales"**: fila 2 en adelante, columnas:
  1 Nº, 2 File, 3 Colegio, 4 Curso, 5 Alumnos, 6 Guías, 7 Total, 8 Destino,
  9 Programa, 10 Salida (fecha), 11 Llegada, 12 Hotel, 13 Cant. discos, 14 Aérea, 15 Obs.
- Hoja **"Marcaciones"**: bloques semanales. Una fila con números = los días del
  bloque (el título de arriba indica el MES). Debajo de cada número van las
  atracciones, y debajo de éstas las sub-actividades/boliche (noche).

**Cómo asocia cada colegio:**
- Busca un bloque de "Marcaciones" cuya **fecha de inicio coincida con la fecha
  de Salida** del colegio (y si coincide la duración, mejor). Si no hay bloque
  para esa fecha, ese colegio queda sin planilla.

**Atracciones:** se ajustan a nombres estándar (Beto Carrero, Teleférico, Barco,
Cascata Carolina, Zacarías, Rueda, Multiparque, Cristo Luz...) con un diccionario
interno que aporta hora típica, sub-atracciones y restaurante del almuerzo cuando
la fuente no lo trae.

**Colegios que pueden ir a la misma atracción en varias fechas:** NO se repiten.
Se acumulan las fechas y la celda DATA muestra las dos juntas (`04/12 y 07/12`).
Sirve para:
- **Cristo Luz** cuando la semana tiene 2 noches con Cristo Luz (mitad va a una,
  mitad a la otra).
- **Barco / Teleferico** cuando la semana los tiene el mismo día repartidos
  (en el mapita: `Barco` + abajo `Teleferico`, y al revés al día siguiente):
  el colegio puede hacer cualquiera de los dos en cualquiera de los dos días.
- Si la atracción tiene una sola fecha posible, la celda DATA queda con esa fecha
  real (sigue siendo fecha, no texto).
El programa **no reparte** los colegios: no inventa qué mitad va a cada fecha, eso
se completa a mano cuando esté decidido.

**Salida:** `salidas/SALIDA_SISTEMA_RECEPTIVO.xlsx`

---

## 2. OPERATIVO (estilo INKA) -> `generador_operativo.py`

**Lee 1 archivo:** el "Control Operativo" dentro de `entradas/`.
Se detecta solo: usa `Control Operativo.xlsx`, o el primer `.xlsx` de `entradas/`
que tenga una hoja llamada **`Reservas`**.

**Hoja "Reservas" — qué busca:**
- **Grupos** (fila 4 en adelante, columna A): cada celda con texto es un grupo
  (ej. `1 van de 20`). Se leen: C=estudiantes, D=adultos, E=total, F=coordinadores,
  G=IN, H=OUT.
- **Itinerario Ideal** (debajo de los grupos): filas donde la columna A es un
  número (1,2,3...) y la columna B empieza con "IN". Las columnas B..I son los
  pasos de ese día (se toman como texto). Es la ÚNICA fuente del itinerario.
- **Notas** (fila 11+): busca texto `salida HH:MM` para la hora de salida.

**Mapeo del Itinerario a hojas:**
| En el itinerario          | Hoja          |
|---------------------------|---------------|
| Parque Unipraias          | UNIPRAIA      |
| Beto Carrero World        | BETO CARRERO  |
| Porto Belo                | PORTO BELO    |
| Campamento Americano      | ZACARIAS      |
| Cristo Luz / Cena Cristo Luz | CRISTO LUZ (día IN) |
| Barco Pirata              | BARCO PIRATA  |
| Día Libre                 | sin hoja      |

- Cada día se coloca en la columna de PROGRAMACION según su **fecha real**
  (aunque los grupos entren en días distintos).
- Si una atracción NO se reconoce, esa celda del día queda **en blanco** y no se
  crea hoja (el programa no inventa).
- El vuelo de OUT se deja `____` a propósito.

**Salida:** `salidas/SALIDA_OPERATIVO_2026.xlsx` (8 hojas: CRISTO LUZ, UNIPRAIA,
BETO CARRERO, PORTO BELO, ZACARIAS, BARCO PIRATA, PROGRAMACION, 5 COMIDA).

---

## 3. WORD -> `generador_word.py`

**Lee:** el primer archivo `.docx` que encuentre en `entradas/`.

**Qué extrae del texto del Word:**
- **Cabeceras de día**: líneas tipo `Dom. 11 de Oct. :` (día + mes en el título).
- **Nombre del colegio**: la primera línea en mayúsculas antes del primer día.
- **Cantidades**: patrón tipo `56M + 45H +10 PADRES + 2 AG. M = 113`.
- **Atracciones**: por palabras clave del texto (Unipraia→Teleférico, Beto
  Carrero, Barco, Campamento Americano, Multiparque, etc.).
- **Noches/discotecas**: por palabras clave (Eclipse, Green Valley, D'uff, etc.).
- **Vuelos**: aerolínea + número (ej. H2 0809 / LA 1242).
- **Hotel**: texto que empieza con "hotel ...".
- **Comidas**: desayuno / almuerzo / cena con su lugar, desde el texto.

**Salida:** `salidas/SALIDA_WORD_OPERACIONES {colegio}.xlsx`. Cómo el Word trae
horas reales, esas celdas salen completas (no `____H`).

---

## Guía de problemas (qué revisar si algo se rompe)

| SÍNTOMA | CAUSA MÁS COMÚN | QUÉ HACER |
|---|---|---|
| No encuentra el archivo de entrada | El archivo no está en `entradas/` | Copiarlo ahí; en el tipo se puede "Examinar..." |
| Error: Worksheet "Reservas" no existe | Renombraron o quitaron la hoja del Excel | Verificar el nombre de la hoja (Se debe llamar Reservas / Colegios totales / Marcaciones) |
| No se puede guardar (archivo en uso) | La salida está abierta en Excel | Cerrar Excel, o guardar una copia con otro nombre (lo ofrece solo) |
| A un colegio no le salen días | No hay bloque de "Marcaciones" con su fecha de salida | Chequear que exista semana para esa fecha en la hoja Marcaciones |
| Días en blanco en el Operativo | Atracción no reconocida en el Itinerario | Ver nombre en el itinerario y ver si está en el mapeo (sección 2) |
| Word: "No se pudo leer texto" | El .docx no tiene texto/líneas útiles | Revisar que el Word no sea imagen o esté vacío |
| `ModuleNotFoundError: openpyxl` | Python sin la librería | `pip install -r requirements.txt` |
| Aparece `____H` / `____` | La fuente no trae hora/vuelo | Se completa a mano (comportamiento normal) |
| Resultados con fechas corridas | Mes/año mal detectado en la hoja Marcaciones | Verificar el título del mes arriba del bloque de días |

## Cómo probar solo (consola)
Abrir CMD/PowerShell en `Marcaciones\codigo` y correr el generador que sea;
imprime un resumen de todo lo que leyó (grupos, atracciones, 5ª COMIDA, aviso de
guardado). Ahí se ve en qué paso falla o qué detectó mal.