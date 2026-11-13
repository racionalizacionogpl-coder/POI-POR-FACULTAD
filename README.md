# POI 2026 por facultad – desplegables AEI → AOI

Hay dos versiones del script, con los mismos datos del POI 2026 (sin OEI):

| Carpeta | Para qué libro | Dónde pone los desplegables |
|---|---|---|
| `apps-script-anexo1/` | **ANEXO 1 – Inventario de productos y procesos** (el libro original, con automatizaciones propias) | Solo en **D7 (AEI)** y **E7 (AOI)** de cada hoja de facultad; cada persona copia y pega después |
| `apps-script/` | Libro "POI POR FACULTADES" (hojas vacías) | Columnas A y B, 300 filas |

## Libro ANEXO 1 (`apps-script-anexo1/`)

1. Extensiones → Apps Script → crear dos archivos y pegar `POI_Datos.gs` y `POI_Codigo.gs`.
2. Ejecutar **una vez** la función `poiInstalar` (pide autorización). Crea:
   * dos disparadores instalables (`poiAlEditar` y `poiCrearMenu`),
   * el menú **POI AEI-AOI**,
   * la hoja oculta `POI_LISTAS` (origen de los desplegables),
   * los desplegables en D7/E7 de las 20 hojas de facultad.
3. Menú **POI AEI-AOI → Diagnóstico** para verificar hoja ↔ facultad.

Convive con las automatizaciones existentes: no define `onEdit` ni `onOpen`, todos sus nombres
llevan el prefijo `poi`/`POI_`, y **no modifica ningún valor**, solo la validación de D7 y E7.
Las pestañas "Copia de …" se omiten al aplicar en bloque. Si las hojas tienen rangos protegidos,
quien ejecute `poiInstalar` debe poder editarlos (los disparadores instalables corren con sus permisos).

**Copiar y pegar:** copiar D7:E7 (o solo D7) y pegar en cualquier celda de la hoja; el AEI pegado
—por su desplegable o por su texto— deja en la celda de **su derecha** la lista de actividades de ese AEI.
Si pega AEI + AOI y coinciden se conserva; si no, se limpia la AOI. Un texto propio del libro a la derecha
no se borra, pero esa celda recibe la lista de actividades. Pegar en otro archivo solo lleva el texto.

---

## Libro "POI POR FACULTADES" (`apps-script/`)

Apps Script para el Google Sheets del POI por facultad. En cada hoja de facultad
(`F01_FM`, `F02_FDCP`, …) crea dos desplegables encadenados:

1. **Acción estratégica (AEI)**, columna A: solo los AEI que la facultad tiene en el POI.
2. **Actividad operativa (AOI)**, columna B: al elegir un AEI solo aparecen sus actividades,
   como `AOI00008800777 - EJECUCIÓN DE LOS PROGRAMAS DE POSGRADO` (código + denominación).

Si se cambia el AEI, la AOI se limpia y su lista se recalcula.

Los datos salen del **Anexo B-5 del POI 2026** (PDF "POI - GENERAL", páginas 20-49,
centros de costo 01.15 a 01.34): 20 facultades, 649 actividades operativas.
Cada facultad solo muestra lo que tiene en el PDF.

## Copiar y pegar

* **Dentro del mismo Sheet (misma hoja o entre hojas):** se puede copiar el AEI, o AEI + AOI, y
  pegarlo **en cualquier lugar de la hoja**, incluso fuera de las columnas A:B o más abajo de la
  fila 301. El script reconoce el AEI pegado (por su desplegable o por su texto) y deja en la celda
  de **su derecha** la lista de actividades de ese AEI. Convención: la AOI va a la derecha del AEI.
* Si se pega AEI + AOI y la AOI pertenece a ese AEI, se conserva; si no, se limpia.
* Si se pega solo una AOI en una fila con otro AEI, se limpia (no coincide con su AEI).
* Un AEI que la facultad de destino no tiene queda marcado como inválido (triángulo rojo).
* No se pisa texto ajeno: si a la derecha de un AEI pegado hay una nota del usuario, se deja como está.
* **A otro archivo, Excel, Word o correo:** se pega el texto (`AEI.01.01 …` y
  `AOI0000… - …`). Los desplegables no viajan a otro archivo.
* **Pegar solo el valor:** `Ctrl+Shift+V` también funciona: el AEI se reconoce por su texto.
* **Pegar desde fuera hacia el Sheet:** el texto debe coincidir exactamente con el de las listas.
* No borrar la hoja oculta `POI_LISTAS`: es el origen de los desplegables.
* Si se hace *Archivo → Hacer una copia*, el script y las listas se copian con el archivo
  (hay que autorizarlo de nuevo la primera vez).

## Instalación (una sola vez)

1. Abrir el Sheet → **Extensiones → Apps Script**.
2. Crear dos archivos y pegar el contenido de `apps-script/Datos.gs` y `apps-script/Codigo.gs`.
3. En `Codigo.gs`, revisar `CFG` (sobre todo `FILA_INICIO`, `COL_AEI`, `COL_AOI`).
4. Guardar, recargar el Sheet y usar el menú **POI 2026 → Diagnóstico** para verificar que
   cada hoja se asocia a su facultad.
5. **POI 2026 → 1. Aplicar a TODAS las facultades** (la primera vez pide autorización).
6. Si las hojas ya se habían aplicado con la versión que tenía OEI: **POI 2026 → 3. Quitar la columna OEI**.

El `onEdit` es un disparador simple: no hay que instalar nada más.

## Cómo se asocia cada hoja a su facultad

Por el nombre de la hoja: debe contener la **sigla** (`F09_FCA`). La sigla manda; el número de
formulario solo se usa si ninguna sigla coincide, porque la numeración de las hojas puede diferir
del Anexo 1 (por ejemplo `F20_FISI`, aunque F20 sea el formulario de FII).

| Hoja | Facultad | Centro de costo (PDF) |
|---|---|---|
| F01_FM | Medicina | 01.30 |
| F02_FDCP | Derecho y Ciencia Política | 01.22 |
| F03_FLCH | Letras y Ciencias Humanas | 01.29 |
| F04_FFB | Farmacia y Bioquímica | 01.24 |
| F05_FO | Odontología | 01.32 |
| F06_FE | Educación | 01.23 |
| F07_FQIQ | Química e Ingeniería Química | 01.34 |
| F08_FMV | Medicina Veterinaria | 01.31 |
| F09_FCA | Ciencias Administrativas | 01.15 |
| F10_FCB | Ciencias Biológicas | 01.16 |
| F11_FCC | Ciencias Contables | 01.17 |
| F12_FCE | Ciencias Económicas | 01.18 |
| F13_FCF | Ciencias Físicas | 01.19 |
| F14_FCM | Ciencias Matemáticas | 01.20 |
| F15_FCCSS | Ciencias Sociales | 01.21 |
| F16_FIGMMG | Ing. Geológica, Minera, Metalúrgica y Geográfica | 01.27 |
| F17_FII | Ingeniería Industrial | 01.28 |
| F18_FPSIC | Psicología | 01.33 |
| F19_FIEE | Ing. Electrónica y Eléctrica | 01.26 |
| F20_FISI | Ing. de Sistemas e Informática | 01.25 |

## Hojas auxiliares que crea el script

* `POI_LISTAS` (oculta): origen de los desplegables. No editar.
* `POI_CATALOGO` (menú → *Generar hoja POI_CATALOGO*): tabla de consulta con todo el POI por facultad.

## Regenerar los datos

```bash
pip install pymupdf
python3 tools/extraer_poi.py "POI - GENERAL.pdf" apps-script/Datos.gs
python3 tools/extraer_poi.py "POI - GENERAL.pdf" apps-script-anexo1/POI_Datos.gs
node tools/test_codigo.js   # prueba apps-script/ con un Sheet simulado
node tools/test_anexo1.js   # prueba apps-script-anexo1/
```
