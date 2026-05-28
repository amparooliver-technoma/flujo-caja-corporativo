# Se quiere hacer una proyeccion de flujo de caja para el sector corporativo

La idea es poder tomar distintos Excel que representan archivos de presupuesto para clientes. Dentro de cada archivo figuran distintas hojas de trabajo.

Cada nombre de archivo contiene el numero de TECH, por ejemplo "TECH11474" o "TECH-11880". Esto es una representacion interna para hacer seguimiento a que proyecto corresponde cada presupuesto. Hay veces que un proyecto tiene multiples TECH.

Las hojas dentro de cada Excel varian en nombres y cantidades. Usualmente, la primera hoja de trabajo es la hoja de cotizacion que se le comparte al cliente. Esa hoja siempre contiene la palabra PClientes, pero puede tener un agregado.

Ejemplos: PClientes, PClientes - DELL.

Esta hoja contiene una tabla con columnas como Item, Nro de Parte, Descripcion, Cant., Unit. y Sub Total.

Esta tabla va a estar enlazada a una hoja BOM, que tambien suele variar de nombre pero normalmente contiene la palabra BOM. La hoja BOM principal es la que esta enlazada a PClientes y donde esta especificado el Proveedor por item.

## Objetivo del output

Lo que se quiere lograr es un output unificado como el archivo "Base para proyectos a Futuros (1).xlsx", pero no a nivel item.

El resultado final debe tener UNA LINEA POR PROVEEDOR POR TECH.

Es decir:

- Primero se identifican los items de cada TECH.
- Luego se identifica el proveedor correspondiente a cada item.
- Despues se agrupan todos los items del mismo proveedor dentro del mismo TECH.
- Finalmente se suma el monto total por proveedor por TECH.

Ejemplo conceptual:

Si el TECH11474 tiene 10 items de proveedor A y 4 items de proveedor B, el output debe tener solo 2 lineas para ese TECH:

- TECH11474 - Proveedor A - total acumulado de sus items
- TECH11474 - Proveedor B - total acumulado de sus items

La granularidad correcta para el flujo de caja es proveedor por TECH, no item por TECH.

## Columnas del output

El archivo de salida debe respetar la estructura del archivo "Base para proyectos a Futuros (1).xlsx".

Columnas esperadas:

- Referencia del pedido - TECH: se completa con el TECH detectado.
- Proveedor: se completa con el proveedor agrupado.
- Cliente: se completa si se puede detectar desde el archivo o queda pendiente de definicion.
- Fecha inicio: se deja vacia.
- Total: suma total acumulada para ese proveedor dentro de ese TECH.
- Moneda: se completa con la moneda detectada del monto, si aplica.
- EQ. USD: se deja vacia.
- Probabilidad: se deja vacia.
- SECTOR: siempre se completa con CORPORATIVO.

## Reglas esperadas

- Si un archivo contiene mas de un TECH, el proceso debe separar los montos por TECH.
- Si un TECH aparece en mas de un archivo, el proceso debe poder acumular los montos del mismo proveedor bajo ese mismo TECH.
- Si no se encuentra proveedor para un item, no debe fallar todo el proceso. Debe marcarse como proveedor no encontrado o pendiente de revision.
- Si no se encuentra una hoja esperada, el proceso debe registrar una advertencia y continuar con los demas archivos.
- Los nombres de hojas y columnas deben detectarse de forma flexible, porque pueden variar entre archivos.

## Pregunta de implementacion

Como podemos lograr esto para que sea lo suficientemente robusto para aceptar nombres variables, que no falle si no encuentra, que no tenga dependencias mayores, e idealmente que este deployado gratuitamente, por ejemplo en GitHub Pages de ser posible?

Antes de hacer un ejecutable .exe, considerando que trabajamos en computadoras corporativas HP con Windows, lo ideal es que no sea una unica aplicacion instalada localmente.

## Archivos Excel de ejemplo

- C:\Users\amparooliver\Desktop\Renato\Flujo de Caja Corporativo\2026-TECH-11880_PCs Workstation-OP2 DELL.xlsx
- C:\Users\amparooliver\Desktop\Renato\Flujo de Caja Corporativo\Infra.CDA.SAX.CCEE-2026-TECH11474.xlsx
- C:\Users\amparooliver\Desktop\Renato\Flujo de Caja Corporativo\Starlink.Cerro.vs.Sporting.Cristal-2026-TECH11954.xlsx
- C:\Users\amparooliver\Desktop\Renato\Flujo de Caja Corporativo\V2-Infra.CCTV.SAX.CCE-2026-TECH11474.xlsx
- C:\Users\amparooliver\Desktop\Renato\Flujo de Caja Corporativo\V2-INFRA.CCTV.SAX.TE-2026-TECH11550.xlsx
- C:\Users\amparooliver\Desktop\Renato\Flujo de Caja Corporativo\V2-INFRA.CDA.SAX.TE-2026-TECH11550.xlsx
