# POI 2026 por facultad – desplegables OEI → AEI → AOI

Apps Script para el Google Sheets del POI por facultad. En cada hoja de facultad
(`F01_FM`, `F02_FDCP`, …) crea tres desplegables encadenados:

1. **OEI** de la facultad.
2. **AEI**: al elegir un OEI solo aparecen los AEI de ese OEI.
3. **AOI**: al elegir un AEI solo aparecen sus actividades operativas, como
   `AOI00008800777 - EJECUCIÓN DE LOS PROGRAMAS DE POSGRADO` (código + denominación).

Si se cambia el OEI o el AEI, los desplegables que dependen de él se limpian y se recalculan.

Los datos salen del **Anexo B-5 del POI 2026** (PDF "POI - GENERAL", páginas 20-49,
centros de costo 01.15 a 01.34): 20 facultades, 649 actividades operativas.
Cada facultad solo muestra lo que tiene en el PDF.

## Instalación (una sola vez)

1. Abrir el Sheet → **Extensiones → Apps Script**.
2. Crear dos archivos y pegar el contenido de `apps-script/Datos.gs` y `apps-script/Codigo.gs`.
3. En `Codigo.gs`, revisar `CFG` (sobre todo `FILA_INICIO`, `COL_OEI`, `COL_AEI`, `COL_AOI`).
4. Guardar, recargar el Sheet y usar el menú **POI 2026 → Diagnóstico** para verificar que
   cada hoja se asocia a su facultad.
5. **POI 2026 → 1. Aplicar a TODAS las facultades** (la primera vez pide autorización).

El `onEdit` es un disparador simple: no hay que instalar nada más.

## Cómo se asocia cada hoja a su facultad

Por el nombre de la hoja: debe contener la **sigla** (`F09_FCA`) o el **formulario** (`F09`).
FPSIC, FIEE y FISI no tienen formulario en el Anexo 1, así que su hoja debe contener la sigla
(por ejemplo `F21_FPSIC`).

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
| F20_FII | Ingeniería Industrial | 01.28 |
| (FPSIC) | Psicología | 01.33 |
| (FIEE) | Ing. Electrónica y Eléctrica | 01.26 |
| (FISI) | Ing. de Sistemas e Informática | 01.25 |

## Hojas auxiliares que crea el script

* `POI_LISTAS` (oculta): origen de los desplegables. No editar.
* `POI_CATALOGO` (menú → *Generar hoja POI_CATALOGO*): tabla de consulta con todo el POI por facultad.

## Regenerar los datos

```bash
pip install pymupdf
python3 tools/extraer_poi.py "POI - GENERAL.pdf" apps-script/Datos.gs
node tools/test_codigo.js   # prueba la lógica con un Sheet simulado
```
