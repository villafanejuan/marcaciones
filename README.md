# Distribuidor de Habitaciones para Grupos

Aplicación de escritorio para **asignar pasajeros a habitaciones** (apartamentos/camas) de hoteles
cuando llegan **grupos**: colegios, empresas, turismo estudiantil, etc.

Busca la mejor distribución respetando reglas por grupo (separar sexos, niños con adultos,
parejas, capacidad de cada habitación, disponibilidad por fecha, cupo de estudiantes por
habitación, etc.), muestra el resultado en **mapas de habitaciones**, una **matriz de consumo
por noche** y exporta a **Excel** (incluido un rooming estilo "File" con límites por noche).

![Pantalla principal](docs/imagenes/01_pantalla_principal.png)

## Funcionalidades

- **Multi-hotel**: trabajar con varios hoteles a la vez, cada uno con sus grupos y habitaciones.
- **Grupos con reglas**: nombre, empresa/escuela, fechas de ingreso/egreso (opcional: sin fecha =
  cola de espera), pasajeros con tipo (estudiante, adulto, niño, guía/chofer), sexo, edad, familia
  (apellido de grupo), parejas (cama matrimonial) y necesidad de habitación individual.
- **Modos por grupo**: **GRUPAL** (datos completos, reglas estrictas) o **INDIVIDUAL / provisional**
  (cantidades parciales, se marca como «ESTIMADO» hasta confirmarlo).
- **Modo de hospedaje por grupo**: **ADULTOS** (receptivo: parejas, familias, reparto libre) o
  **ESTUDIANTIL** (turismo de estudio: alumnos y acompañantes nunca comparten habitación, y cada
  bando se reparte aparte). Acompañantes por rol y sexo, y un **tope de plazas por habitación de
  alumnos** (4 por defecto) para evitar sextuples aunque el hotel tenga habitaciones grandes.
- **Confirmación de grupos**: marcar un grupo como **✔ Confirmado** congela sus habitaciones para
  que no se muevan en futuras simulaciones; los provisionales se pueden reacomodar.
- **Modo de simulación**: **CONSERVAR** (respeta confirmados y acomoda nuevos en lo libre) o
  **DESDE CERO** (reoptimiza todo ignorando confirmaciones).
- **Motor de optimización**: estrategias `AUTO`, `EXACTA`, `BFD`, `FFD`, `MULTISTART` y `TODAS`
  para encontrar la mejor ocupación, con tope configurable de plazas por habitación.
- **Habitaciones y camas**: alta de habitaciones por cantidad/capacidad/numeración automática,
  bloqueos por fechas, y un inventario de camas alternativo (modo "solo camas").
- **Rack / Cotizador**: estimar cuántos apartamentos/pasajeros entran en un rango de fechas.
- **Planificador / Estimador**: dado un total de pasajeros, estimar cuántas habitaciones y qué
  capacidad sugerir para cada grupo.
- **Mapas y matriz**: mapa del hotel con cada pax en su cama, mapa de camas, bloqueos, cotización
  marcada, y matriz de consumo grupo × noche (con pico y tope).
- **Exportaciones**: Excel con mapa+resumen+pasajeros, matriz a xlsx, y **Rooming** por hotel con
  hojas `Resumen por Grupo`, `Detalle por Categoría`, `Cambios` y `Límites por noche` (disponibles
  por noche con alerta de sobreventa).
- **Persistencia**: guardar/abrir JSON (todos los hoteles + grupos) y **espacio de trabajo**
  auto-guardado y auto-cargado al iniciar (hotel activo, grupos y rango de fechas de vista).

## Requisitos

- **Python 3.11+** (versión con Tkinter incluida).

### Dependencias (solo modo fuente)

```bash
pip install customtkinter openpyxl pillow
```

> `tkinter` suele venir instalado con Python en Windows. Si no, instalá el paquete
> `python-tk` correspondiente a tu sistema.

## Ejecutar desde el código

```bash
python interfaz.py                # abre la ventana principal
python interfaz.py --selftest     # prueba automática de toda la app (sin ventana)
```

En el primer inicio verás el hotel "Mi Hotel" vacío. Desde el botón **«🎬 Cargar demo completa»**
se cargan datos de ejemplo para explorar todas las funciones.

## Generar el ejecutable (.exe) para Windows

Con **PyInstaller** instalado:

```bash
pip install pyinstaller
pyinstaller DistribuidorHotelGUI.spec
```

O simplemente hacer doble clic en **`compilar.bat`**. El ejecutable queda en `dist\` como
`DistribuidorHotelGUI.exe`, listo para copiar a cualquier PC con Windows (no requiere Python).
La app guarda sus datos (exportaciones y espacio de trabajo) junto al ejecutable.

## Estructura del proyecto

| Archivo            | Descripción                                  |
|-------------------|----------------------------------------------|
| `interfaz.py`      | Ventana principal y diálogos (Tk/CustomTkinter) |
| `motor.py`         | Orquestación de la distribución por noches   |
| `empaquetador.py`  | Algoritmos de empaquetado (BFD/FFD/exacta/multistart) |
| `modelos.py`       | Entidades: Hotel, Habitacion, Grupo, Pax, Bloqueo |
| `reglas.py`        | Reglas por grupo (sexos, menores, cupos…)   |
| `persistencia.py`  | Guardado/carga JSON (hoteles, grupos, estado) |
| `exportador.py`    | Exportación CSV                               |
| `excel_exportador.py` | Exportaciones Excel (mapa, matriz, rooming) |
| `informes.py`      | Resúmenes de alto nivel                      |
| `demo.py`          | Datos de ejemplo para probar la app          |
| `pruebas/`         | Suite de pruebas unitarias (`correr_pruebas.py`) |
| `docs/`            | Manual de usuario en HTML y capturas        |
| `compilar.bat`     | Script de compilación del `.exe`             |

## Documentación

- [Manual de usuario (HTML)](docs/manual.html)

## Pruebas

```bash
python -m unittest discover -s pruebas -p "test_*.py" -t pruebas
```