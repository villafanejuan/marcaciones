# RESUMEN DEL SISTEMA

> **Este archivo es el índice del proyecto.** Si te piden "analizá el código" o
> cualquier pregunta general sobre el sistema, leé SOLO este archivo primero y
> abrí el código fuente únicamente lo que haga falta para responder en detalle.
>
> **Mantenimiento:** al hacer cambios, actualizá la sección correspondiente de
> este archivo en el mismo commit. No crecer más de ~250 líneas.

Última revisión: 2026-09-28 · commit inicial sin versionar en git (proyecto nuevo, 0 commits en ese repo)

---

## 1. Qué es

App de escritorio Windows (Python + CustomTkinter) para **asignar pasajeros a
habitaciones** de hoteles cuando llegan grupos (colegios, empresas, turismo
estudantil). Multi-hotel, con reglas por grupo, motor de empaquetado con
optimización, mapa visual de habitaciones, matriz de consumo por noche y
exportación a CSV / Excel / Rooming.

**Idioma de la UI y los comentarios: español rioplatense** (voseo: "usá",
"hacé", "agregá"). Los exports de Rooming mezclan español y portugués
("Limite Diário", "Pessoas", "Apartamentos", "Relatório gerado em").

---

## 2. Mapa de archivos (8.910 líneas Python)

| Archivo | Líneas | Responsabilidad | Calidad |
|---|---|---|---|
| `modelos.py` | 461 | Dataclasses puras: `Hotel`, `Habitacion`, `Grupo`, `Pax`, `Bloqueo`, `LoteCamas`, `Atomo`, `ConfigGrupo`, `EstadoHabitacion`, `ModoHospedaje`. Validación y errores. | Bueno |
| `reglas.py` | 164 | `atomizar(grupo)` → agrupa pax en `Atomo` (familia / pareja / single / suelto). Rechaza menores sin adulto, parejas con menores, duplicados. | Bueno |
| `empaquetador.py` | 470 | Algoritmos de bin-packing: BFD, BFD-AJUSTADA, FFD, MULTISTART, EXACTA (DFS con poda). `calcular_deficit()`. | Bueno |
| `motor.py` | 960 | Orquestación por fechas: `distribuir()`, reparto estudiantil/adulto, reservas entre grupos, confirmados congelados, planificador, rack, `verificar_reporte()`. | Bueno |
| `persistencia.py` | 341 | JSON versionado (`VERSION_FORMATO = 2`, acepta v1 y v2). `guardar_multi` / `cargar_area`. Compatible con JSON previo al modo estudiantil. | Bueno |
| `excel_exportador.py` | 688 | openpyxl: mapa, resumen, pasajeros, matriz de grupos, Rooming por hotel. | Medio |
| `exportador.py` | 132 | CSV de detalle y CSV de resumen. `utf-8-sig` (Excel-friendly). | Bueno |
| `informes.py` | 116 | `renderizar_reporte()` → reporte de texto monoespaciado. | Simple |
| `demo.py` | 123 | Datos de ejemplo (43 habitaciones, 4 grupos, 59 pax). | Referencia |
| `interfaz.py` | **3726** | **Toda la GUI.** 9 diálogos + clase `App` de ~1900 líneas / 90 métodos. | **El problema del proyecto** |
| `pruebas/` | 1.697 | 94 tests unitarios + `fabrica.py` (generadores) + `estres.py` (escenarios de estrés) + `test_modo_estudiantil.py`. | Bien |
| `docs/` | 489 | `manual.html` (~26 KB) + 3 scripts que generan capturas con PIL. | — |
| `compilar.bat` / `.spec` | — | PyInstaller onefile/windowed → `dist/DistribuidorHotelGUI.exe` (30 MB). | Funciona |

---

## 3. Conceptos del dominio (lo que hay que saber sí o sí)

- **`Atomo`**: unidad indivisible de asignación. Puede ser 1 pax o una
  familia/pareja entera. Tiene `tamano`, `sexos_no_x`, `tiene_menor`,
  `requiere_single`. El packer solo mueve átomos, nunca pax sueltos.
- **`EstadoHabitacion`**: estado mutable durante el packing, con
  `puede_recibir()` / `recibir()` / `quitar()` y `firma()` para deduplicar
  estados en la búsqueda exacta. Es el corazón de las reglas de cohabitación.
- **`separar_sexos`**: por defecto `True`. Mezclar sexos solo se permite con
  vínculo de familia (`exento_sexo`) o pareja.
- **`estudiantes_por_habitacion`**: cupo fijo k. Los estudiantes se salen del
  resto y se empaquetan aparte (`_empaquetar_estudiantes`). Con
  `estudiantes_estricto=True` exige que `total % k == 0` por sexo.
- **`ModoHospedaje`** (enum, `ConfigGrupo.modo_hospedaje`, default `ESTUDIANTIL`):
  se elige **por grupo**, así que conviven grupos estudiantiles y adultos en el
  mismo hotel.
  - `ESTUDIANTIL`: `_distribuir_estudiantil()` parte en dos bandos que **nunca
    se mezclan** — alumnos (se mezclan entre sí) y acompañantes (se agrupan por
    **rol + sexo** con `_cohortes_adultas`; si `adultos_por_rol=False` se junta
    todo del mismo sexo). Usa `_empaquetar_cohorte` para no llenar de más.
  - `ADULTOS`: `_distribuir_adultos()` conserva el reparto anterior, y sigue
    usando `_empaquetar_estudiantes` si el grupo trae `estudiantes_por_habitacion`.
  - **Fallback**: `_modo_efectivo()` devuelve `ADULTOS` si el grupo no tiene
    ningún alumno, así que `ESTUDIANTIL` sin alumnos no es un error.
  - **`max_plazas_estudiantes`** (default 4): techo de plazas para habitaciones de
    alumnos, más allá del cupo k. Evita sextuples aunque el hotel tenga de 6/8/10.
    Si no hay ninguna habitación con ≤ ese tope, se relaja a la más chica del
    hotel y se emite un aviso.
  - **`guias_choferes_single`** cambió de `True` a `False`: lo natural es agrupar
    choferes entre sí (2 en un duplo, 3 en un triplo), no aislar a cada uno.
- **Confirmados**: `grupo.confirmado = True` congela sus habitaciones. Con
  `modo_simulacion = CONSERVAR` se respeta lo anterior; `DESDE CERO` reoptimiza
  ignorando confirmaciones.
- **Flujo de la UI**: `simular` → revisar → `reservar` (convierte asignaciones
  en `Habitacion.bloqueos`). Editar una habitación invalida el reporte.
- **Múltiples representaciones del mismo dato**: reportes de texto
  (`informes.py`), CSV, Excel, y 3 renderizadores de canvas distintos
  (mapa de hotel, mapa de camas, matriz). Todos derivan del `ReporteGeneral`.

---

## 4. Comandos de verificación

```bash
python -m unittest discover -s pruebas -p "test_*.py" -t pruebas   # 94 tests, ~4s
python correr_pruebas.py                                           # tests + estres
python interfaz.py --selftest                                      # ~3.500 líneas, ~30s
python interfaz.py                                                 # app
python demo.py                                                     # reporte demo en consola
```

**Estado al 2026-09-28: 94/94 tests OK, selftest GUI OK, estrés OK (S1–S5).**
No hay linter, type checker ni CI configurado.

---

## 5. Problemas conocidos (priorizados)

| # | Severidad | Ubicación | Problema | Estado |
|---|---|---|---|---|
| 1 | 🔴 Alta | `interfaz.py:817`, `:900` | `except TclError:` **sin importar** (verificado con AST: no está en el namespace). Las otras 3 ocurrencias sí usan `tk.TclError`. Si `grab_set()` falla → `NameError`. | Pendiente |
| 2 | 🔴 Alta | `interfaz.py` entero | God-object: `App` con 90 métodos mezcla estado, layout, render, orquestación y autosave. Funciones de 113–313 líneas. | Pendiente |
| 3 | 🔴 Alta | `interfaz.py:2142` | `distribuir()` corre **sincrónico en el hilo de Tk** con `tiempo_max=4.0` por grupo. La UI se congela con muchos grupos. No hay `threading`/`queue` en el proyecto. | Pendiente |
| 4 | 🟠 Media | `interfaz.py:1435-1441` | 4 listas paralelas por hotel (`hoteles`, `grupos_por_hotel`, `reporte_por_hotel`, `seleccion_por_hotel`) + alias reasignados. El bloque de mutación ya está duplicado en `1577` y `2063`. | Pendiente |
| 5 | 🟠 Media | `interfaz.py:1521` | El autoguardado traga errores en silencio (`except ErrorDistribucion: pass`). El usuario puede perder trabajo sin aviso. Además reescribe los 212 KB de JSON cada 900 ms. | Pendiente |
| 6 | 🟠 Media | `empaquetador.py:314` | La semilla por defecto usa `hash(str)`, que **varía entre corridas** (PYTHONHASHSEED). `MULTISTART` no es reproducible. | Pendiente |
| 7 | 🟡 Baja | `empaquetador.py:170` | `sys.setrecursionlimit()` nunca se restaura. | Pendiente |
| 8 | 🟡 Baja | `interfaz.py:95` vs `excel_exportador.py:33` | `piso_de` y `_piso_de` duplicados **con comportamiento distinto**. Idem `formatear`/`_formatear` y `noches_entdesde_hasta`/`_dias_rango`. | Pendiente |
| 9 | 🟡 Baja | `interfaz.py:1776-1783` vs `2192,2197,2927` | Nombres de pestaña con emoji repetidos como strings mágicos. Romper un nombre rompe la navegación. | Pendiente |
| 10 | 🟡 Baja | `interfaz.py:2755` | Tope silencioso de 92 días en la matriz + geometría hardcodeada (`col_emp=16`, `col_dia0=420`, `ancho_dia=46`). | Pendiente |
| 11 | 🟡 Baja | `interfaz.py:10`, `87-88` | `sys.path.insert` en import time y constantes de ruta calculadas al importar. `_carpeta_base_datos` usa `dirname(sys.executable)` pero `ruta_log` usa `sys.path[0]`: bases inconsistentes en el build frozen. | Pendiente |
| 12 | 🟡 Baja | `interfaz.py:3265+` | El selftest usa `assert` bare (se desactiva con `python -O`) y números mágicos acoplados a `demo.py` (43 hab, 49 filas, 70/61, 36). | Pendiente |
| 13 | 🟡 Baja | `interfaz.py:_cambiar_hotel` | Sale temprano si `indice == self.indice_hotel`. Al agregar un hotel hay que pasar por `_crear_hotel`, que sí sincroniza las 4 listas paralelas. Usar las listas a mano desincroniza `self.hotel`. | Pendiente |
| 14 | 🟡 Baja | `motor.py:_distribuir_adultos` | `_empaquetar_estudiantes` tiene `cupo: int \| None = None` y con `None` **descarta el pool entero sin avisar** (caía en el invariante "pasajeros desaparecidos"). `_distribuir_adultos` debe siempre pasar `cupo=config.estudiantes_por_habitacion`. | Pendiente |
| 15 | ⚪ Cosmético | `informes.py` | Escribe sin acentos ("ocupacion", "anos") mientras el resto de la app sí los usa. | Pendiente |

---

## 6. Reglas de handcuffs (convenciones a respetar)

- **Imports entre capas solo hacia abajo:**
  `modelos` ← `reglas` ← `empaquetador` ← `motor` ← `informes`/`exportador`/`excel_exportador` ← `interfaz`.
  `modelos.py` no importa nada del proyecto.
- **Todo error de negocio es `ErrorDistribucion`** (o su subclase
  `ErrorValidacion`) de `modelos.py`, con mensaje en español para el usuario
  final. La GUI los muestra directo.
- **Fechas**: `datetime.date` en el dominio, ISO `AAAA-MM-DD` en JSON,
  `DD/MM/AAAA` en la UI y exports. `parsear_fecha()` acepta ambos formatos.
- **Enums como `str, Enum`**: `TipoPax`, `Sexo`, `Estrategia`, `ModoHospedaje`.
  Serializar con `.value`.
- **En la GUI, nunca llamar `.set()` directo sobre `CTkSegmentedButton`**: no dispara
  el callback `command`, así que la ayuda y los `state` quedan desfasados. Usar
  el setter dedicado (`DialogoGrupo._set_modo_hospedaje`).
- **Alias de texto**: `ALIAS_TIPO` y `ALIAS_SEXO` en `modelos.py` toleran
  variantes ("NIÑO"/"NINO", "M"/"HOMBRE"). Ampliar ahí, no en la GUI.
- **La red de seguridad es `verificar_reporte()`** (`motor.py:499`): valida
  capacidad, mezcla de sexos, menores sin adulto, doubles y solapamiento de
  fechas. `distribuir()` la llama y **aborta** si encuentra violaciones
  (`motor.py:490`). Si tocás el packer o las reglas, agregá un caso a
  `pruebas/test_motor.py` y verificá que el selftest siga pasando.
- **Selftest como red de seguridad del refactor**: antes de partir
  `interfaz.py`, confirmar que `python interfaz.py --selftest` pasa en verde.
  Cubre diálogo, multi-hotel, exports, canvas, matriz y modo solo-camas.
- **Al agregar un campo a un dataclass**: actualizar sí o sí
  `persistencia.py` (`*_a_dict` y `*_de_dict`) o el dato se pierde al guardar.

---

## 7. Cómo actualizar este resumen

1. Leé el diff de los archivos tocados.
2. Actualizá la tabla de la §2 si cambió una línea de responsabilidad.
3. Marcá como `Resuelto` en §5 el problema que se acaba de arreglar, y
   borralo en un refactor posterior.
4. Si cambió el flujo de la UI, actualizá §3.
5. No lo expands: si un módulo crece de 1000 líneas, partilo primero.
