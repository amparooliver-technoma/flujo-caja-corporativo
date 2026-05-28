# Flujo de Caja Corporativo

App web estatica para consolidar archivos Excel de presupuesto en una linea por proveedor por TECH.

## Uso

1. Abrir `index.html` en el navegador, o usar el servidor local si esta disponible.
2. Seleccionar o arrastrar uno o mas archivos `.xlsx`.
3. Presionar `Procesar`.
4. Revisar la vista previa, las advertencias y la tabla de trazabilidad.
5. Descargar el output o la trazabilidad.

## Reglas implementadas

- El TECH se detecta desde el nombre del archivo con formatos como `TECH11474` o `TECH-11880`.
- Se busca la mejor hoja `PClientes` por encabezados detectados.
- Se buscan hojas `BOM` o `Infra` para encontrar proveedor por item, codigo o descripcion.
- El output agrupa por `TECH + proveedor + moneda`.
- El monto de cada item se calcula exclusivamente con `Costo Asu sub-Total x 1,1`.
- Si no existe `Costo Asu sub-Total` numerico, el item no se suma y queda marcado como omitido en la trazabilidad.
- La columna `Moneda` se detecta desde el resumen/total de `PClientes`, normalmente en el bloque `Cantidad / Moneda / Total`.
- `Fecha inicio`, `EQ. USD` y `Probabilidad` quedan vacios.
- `SECTOR` siempre queda como `CORPORATIVO`.
- Si falta proveedor, se agrupa bajo `PROVEEDOR PENDIENTE` y se registra advertencia.
- La trazabilidad muestra cada item fuente usado en las sumas, incluyendo archivo, hoja, item, parte, descripcion, `Costo Asu Unitario`, calculo aplicado y criterio usado para detectar proveedor.

## Limitacion conocida

La app lee los valores guardados dentro del Excel. Si un archivo tiene formulas sin valores cacheados por Excel, el navegador no recalcula esas formulas. En ese caso conviene abrir y guardar el archivo en Excel antes de procesarlo, o evaluar una version local con Python/Excel.

## Dependencias

La libreria `xlsx.full.min.js` esta copiada localmente en `vendor/`, por lo que la app no necesita descargar dependencias externas al ejecutarse.
