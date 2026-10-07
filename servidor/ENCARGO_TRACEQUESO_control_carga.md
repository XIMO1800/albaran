# Encargo para la conversación de TraceQueso: «Control carga autoventas» para calidad

Joaquín rellena el registro de calidad **FOR PR 08-10(1) «Control carga vehículos autoventas»** dentro de la app de pedidos/autoventa (ximo1800.github.io/albaran, pestaña «5 · Control carga» de cada cuadre). Calidad no debe entrar en esa app: tiene que **consultarlo e imprimirlo desde TraceQueso**, en su versión de escritorio de calidad, **solo lectura**.

## Dónde están los datos

- Hoja de Google **PEDIDOS TIENDA**, cuenta app.quesoselhidalgo@gmail.com
  ID: `1ectBsJ9eVZNDOOT7vDzSt7YoWey_og6zvESPQ7KBy9c`
- Pestaña **AUTOVENTA**. Columnas: `ID | VENDEDOR | FECHA | ESTADO | DATOS | ACTUALIZADO | POR`
  - La fila con `ID = CONFIG` es configuración: saltarla.
  - Cada otra fila es un cuadre. `DATOS` es un JSON con todo el cuadre; lo que interesa es `DATOS.control`.

## Formato de `DATOS.control`

| campo | qué es |
|---|---|
| `fecha` | día del control, `yyyy-mm-dd` |
| `trans` | nombre del transportista (= vendedor del cuadre; si falta, usar `DATOS.vendedor`) |
| `mat` | matrícula del camión |
| `bultos` | bultos escritos a mano; si es null/no existe, usar `bultosAuto` |
| `kilos` | kilos escritos a mano; si es null/no existe, usar `kilosAuto` |
| `temp` | temperatura ºC (número, p. ej. 3.8) |
| `conf` | conforme (true/false; si falta, true) |
| `obs` | observaciones (texto) |
| `por` | usuario que hizo el control (p. ej. JOAQUIN) |
| `firma` | firma a mano: trazos `"x,y x,y …|x,y …"` sobre un lienzo de 300×100 (opcional) |
| `firmado` | fecha-hora de la firma en milisegundos (opcional) |

Considerar registro válido cuando exista `DATOS.control` con `temp` o `mat` rellenos (o `ESTADO = cerrado`).

Firma como SVG (cada trazo separado por `|` es una polilínea):

```js
function firmaSvg(f){ if(!f) return "";
  return '<svg viewBox="0 0 300 100" width="60mm" height="20mm"><g fill="none" stroke="#123" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round">'
    + f.split("|").map(t => '<polyline points="'+t+'"/>').join("") + '</g></svg>'; }
```

## Qué hacer en TraceQueso

1. **Servidor (Apps Script de TraceQueso)**: nueva acción de solo lectura, p. ej. `controlesCarga(desde, hasta)`, solo para el rol de calidad y ADMIN:
   ```js
   var sh = SpreadsheetApp.openById('1ectBsJ9eVZNDOOT7vDzSt7YoWey_og6zvESPQ7KBy9c').getSheetByName('AUTOVENTA');
   var v = sh.getDataRange().getValues(); // fila 0 = cabecera
   // por cada fila con ID != 'CONFIG': JSON.parse(DATOS) → control → filtrar por fecha → devolver lista ordenada por fecha
   ```
   Misma cuenta de Google en ambas hojas, así que `openById` funciona (al desplegar pedirá autorizar acceso a hojas de cálculo si no lo tenía). **No escribir nunca** en PEDIDOS TIENDA.
2. **Escritorio de calidad**: pestaña «Control carga autoventas» con filtro Desde/Hasta y tabla
   `Fecha | Transportista | Matrícula | Bultos | Kilos | T (ºC) | Conforme | Controlado por | Firma`.
3. **Imprimir** con el formato del papel, **dos registros por hoja A4**: cabecera con logo y «CONTROL CARGA VEHÍCULOS AUTOVENTAS»; «FECHA: dd/mm/aaaa»; tabla `Nombre transportista | Matrícula coche | Bultos | Kilos`; cuadro centrado «LÁCTEOS CUQUERELLA, S.L. — CONTROL TRANSPORTE PREVIA CARGA — Higiene, Limpieza, Mantenimiento correcto, Ausencia de plagas, Suciedad, Humedad, Olores en interior del vehículo y Productos. T (ºC): … Conforme: SÍ/NO · Observaciones: … [firma] Controlado por: … · FOR PR 09-02/Enero 2015»; pie «FOR PR 08-10(1)».
4. Mantener la línea visual de TraceQueso (logo El Hidalgo y granates).
