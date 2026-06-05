# Flujo de Caja Corporativo

App web estatica para consolidar presupuestos Excel en una linea por proveedor por TECH.

## Uso

Abrir `index.html` o publicar la carpeta en GitHub Pages.

La app procesa los archivos localmente en el navegador. No sube los Excel ni los ZIP a ningun servidor.

## Calculo

- Detecta `TECH` desde el nombre del archivo.
- Acepta Excel individuales o un ZIP con multiples Excel.
- Detecta la hoja `PClientes`.
- Si no hay encabezado `PROVEEDOR`, toma proveedor desde la columna G de `PClientes`.
- Toma `Costo Asu sub-Total`.
- Multiplica cada item por `1,1`.
- Agrupa por `TECH + proveedor + moneda`.
- Detecta `Moneda` desde el resumen/total de la hoja `PClientes`: si el total tiene punto y coma se toma como `USD`; si usa separadores de miles sin decimal o indica `Gs.`, se toma como `PYG`.
- Exporta el formato `BASE COMPRAS`.
- Puede exportar todo el resultado o solo las filas visibles segun los filtros activos.

## Trazabilidad

La tabla de trazabilidad muestra cada item fuente incluido u omitido, el archivo de origen, hoja usada, proveedor detectado, monto tomado y criterio de calculo.
