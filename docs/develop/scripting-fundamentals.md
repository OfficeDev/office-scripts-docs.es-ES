---
title: Conceptos básicos sobre el scripting de los scripts de Office en Excel en la web
description: Información del modelo de objetos y otros conceptos básicos que se deben aprender antes de escribir scripts de Office.
ms.date: 01/27/2020
localization_priority: Priority
ms.openlocfilehash: 5a709c16e23c00ffc7ee7949a3cb11459dc2d530
ms.sourcegitcommit: b075eed5a6f275274fbbf6d62633219eac416f26
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 03/10/2020
ms.locfileid: "42700350"
---
# <a name="scripting-fundamentals-for-office-scripts-in-excel-on-the-web-preview"></a>Conceptos básicos sobre scripts para scripts de Office en Excel en la web (vista previa)

En este artículo se presentan los aspectos técnicos de las secuencias de comandos de Office. Aprenderá cómo funcionan conjuntamente los objetos de Excel y cómo se sincroniza el editor de código con un libro.

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="object-model"></a>Modelo de objetos

Para comprender las API de Excel, debe comprender cómo se relacionan entre sí los componentes de un libro.

- Un **libro** contiene una o varias **hojas de cálculo**.
- Una **hoja de cálculo** da acceso a celdas a través de objetos **Range** .
- Un **rango** representa un grupo de celdas contiguas.
- Los **rangos** se usan para crear y colocar **tablas**, **gráficos**, **formas**y otros objetos de visualización o de organización de datos.
- Una **hoja de cálculo** contiene colecciones de los objetos de datos que están presentes en la hoja individual.
- Los **libros** contienen colecciones de algunos de los objetos de datos (como **tablas**) para todo el **libro**.

### <a name="ranges"></a>Ranges

Un rango es un grupo de celdas contiguas del libro. Normalmente, los scripts usan la notación de estilo a1 (por ejemplo, **B3** para la única celda de la fila **B** y la columna **3** o **C2: F4** para las celdas de las filas **C** a **F** y de las columnas **2** a **4**) para definir rangos.

Los rangos tienen tres propiedades `values`principales `formulas`:, `format`y. Estas propiedades obtienen o establecen los valores de celda, las fórmulas que se van a evaluar y el formato visual de las celdas.

#### <a name="range-sample"></a>Ejemplo de intervalo

En el ejemplo siguiente se muestra cómo crear registros de ventas. Esta secuencia de `Range` comandos utiliza objetos para establecer los valores, las fórmulas y los formatos.

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the active worksheet.
  let sheet = context.workbook.worksheets.getActiveWorksheet();

  // Create the headers and format them to stand out.
  let headers = [
    ["Product", "Quantity", "Unit Price", "Totals"]
  ];
  let headerRange = sheet.getRange("B2:E2");
  headerRange.values = headers;
  headerRange.format.fill.color = "#4472C4";
  headerRange.format.font.color = "white";

  // Create the product data rows.
  let productData = [
    ["Almonds", 6, 7.5],
    ["Coffee", 20, 34.5],
    ["Chocolate", 10, 9.56],
  ];
  let dataRange = sheet.getRange("B3:D5");
  dataRange.values = productData;

  // Create the formulas to total the amounts sold.
  let totalFormulas = [
    ["=C3 * D3"],
    ["=C4 * D4"],
    ["=C5 * D5"],
    ["=SUM(E3:E5)"]
  ];
  let totalRange = sheet.getRange("E3:E6");
  totalRange.formulas = totalFormulas;
  totalRange.format.font.bold = true;

  // Display the totals as US dollar amounts.
  totalRange.numberFormat = [["$0.00"]];
}
```

Al ejecutar este script, se crean los siguientes datos en la hoja de cálculo actual:

![Un registro de ventas que muestra las filas de valores, una columna de fórmula y los encabezados con formato.](../images/range-sample.png)

### <a name="charts-tables-and-other-data-objects"></a>Gráficos, tablas y otros objetos de datos

Los scripts pueden crear y manipular las estructuras de datos y las visualizaciones en Excel. Las tablas y los gráficos son dos de los objetos que se usan con más frecuencia, pero las API admiten tablas dinámicas, formas, imágenes, etc.

#### <a name="creating-a-table"></a>Creación de una tabla

Cree tablas con rangos de datos rellenos. Los controles de formato y de tabla (como los filtros) se aplican automáticamente al rango.

La siguiente secuencia de comandos crea una tabla con los intervalos del ejemplo anterior.

```TypeScript
async function main(context: Excel.RequestContext) {
   let sheet = context.workbook.worksheets.getActiveWorksheet();
   sheet.tables.add("B2:E5", true);
}
```

Al ejecutar este script en la hoja de cálculo con los datos anteriores, se crea la siguiente tabla:

![Una tabla realizada a partir del registro de ventas anterior.](../images/table-sample.png)

#### <a name="creating-a-chart"></a>Crear un gráfico

Cree gráficos para visualizar los datos de un rango. Los scripts permiten docenas de variedades de gráficos, cada una de las cuales puede personalizarse según sus necesidades.

El script siguiente crea un gráfico de columnas simple para tres elementos y coloca 100 píxeles por debajo de la parte superior de la hoja de cálculo.

```TypeScript
async function main(context: Excel.RequestContext) {
  let sheet = context.workbook.worksheets.getActiveWorksheet();
  let chart = sheet.charts.add(Excel.ChartType.columnStacked, sheet.getRange("B3:C5"));
  chart.top = 100;
}
```

Al ejecutar este script en la hoja de cálculo con la tabla anterior, se crea el siguiente gráfico:

![Un gráfico de columnas que muestra las cantidades de tres elementos del registro de ventas anterior.](../images/chart-sample.png)

### <a name="further-reading-on-the-object-model"></a>Lecturas adicionales sobre el modelo de objetos

La [documentación de referencia de la API de scripts de Office](/javascript/api/office-scripts/overview) es una lista completa de los objetos que se usan en scripts de Office. Allí puede usar la tabla de contenido para navegar a cualquier clase de la que quiera obtener más información. A continuación se muestran varias páginas que se ven normalmente.

- [Chart](/javascript/api/office-scripts/excel/excel.chart)
- [Comment](/javascript/api/office-scripts/excel/excel.comment)
- [PivotTable](/javascript/api/office-scripts/excel/excel.pivottable)
- [Range](/javascript/api/office-scripts/excel/excel.range)
- [RangeFormat](/javascript/api/office-scripts/excel/excel.rangeformat)
- [Shape](/javascript/api/office-scripts/excel/excel.shape)
- [Table](/javascript/api/office-scripts/excel/excel.table)
- [Workbook](/javascript/api/office-scripts/excel/excel.workbook)
- [Worksheet](/javascript/api/office-scripts/excel/excel.worksheet)

## <a name="main-function"></a>`main`función

Cada script de Office debe contener `main` una función con la siguiente firma, incluida `Excel.RequestContext` la definición de tipo:

```TypeScript
async function main(context: Excel.RequestContext) {
    // Your Excel Script
}
```

El código dentro de `main` la función se ejecuta cuando se ejecuta el script. `main`puede llamar a otras funciones en el script, pero no se ejecutará el código que no está contenido en una función.

## <a name="context"></a>Contexto

La `main` función acepta un `Excel.RequestContext` parámetro, denominado `context`. Piense `context` en el puente entre el script y el libro. El script tiene acceso al libro con el `context` objeto y lo usa `context` para enviar datos de una a otra.

El `context` objeto es necesario porque el script y Excel se están ejecutando en diferentes procesos y ubicaciones. El script tendrá que realizar cambios o consultar datos del libro en la nube. El `context` objeto administra esas transacciones.

## <a name="sync-and-load"></a>Sincronización y carga

Dado que el script y el libro se ejecutan en diferentes ubicaciones, cualquier transferencia de datos entre los dos lleva tiempo. Para mejorar el rendimiento del script, los comandos se ponen en cola hasta que el `sync` script llame explícitamente a la operación para sincronizar el script y el libro. La secuencia de comandos puede funcionar de forma independiente hasta que tenga que realizar cualquiera de las siguientes acciones:

- Leer datos del libro (siguiendo una `load` operación).
- Escribir datos en el libro (normalmente, porque ha finalizado el script).

La imagen siguiente muestra un ejemplo de flujo de control entre el script y el libro:

![Un diagrama que muestra las operaciones de lectura y escritura que se dirigen al libro desde el script.](../images/load-sync.png)

### <a name="sync"></a>Sincronizar

Siempre que el script tenga que leer o escribir datos en el libro, llame al `RequestContext.sync` método como se muestra a continuación:

```TypeScript
await context.sync();
```

> [!NOTE]
> `context.sync()`se llama implícitamente cuando finaliza un script.

Una vez `sync` finalizada la operación, el libro se actualiza para reflejar las operaciones de escritura que haya especificado el script. Una operación de escritura está estableciendo cualquier propiedad en un objeto de Excel ( `range.format.fill.color = "red"`por ejemplo,) o llamando a un método que cambia una propiedad `range.format.autoFitColumns()`(por ejemplo,). La `sync` operación también lee los valores del libro que ha solicitado el script mediante una `load` operación (como se describe en la siguiente sección).

La sincronización del script con el libro puede tardar un tiempo, en función de la red. Debe minimizar el número de llamadas `sync` para que la secuencia de comandos pueda ejecutarse con rapidez.  

### <a name="load"></a>Volver

Un script debe cargar datos del libro antes de leerlo. Sin embargo, la carga de datos de todo el libro con frecuencia reducirá en gran medida la velocidad del script. En su lugar, `load` el método permite al script indicar específicamente qué datos deben recuperarse del libro.

El `load` método está disponible en todos los objetos de Excel. El script debe cargar las propiedades de un objeto antes de que pueda leerlas. Si no lo hace, se producirá un error.

En los ejemplos siguientes se `Range` usa un objeto para mostrar las tres `load` formas en que se puede usar el método para cargar datos.

|Intent |Comando de ejemplo | Efecto |
|:--|:--|:--|
|Cargar una propiedad |`myRange.load("values");` | Carga una única propiedad, en este caso la matriz bidimensional de valores de este intervalo. |
|Cargar varias propiedades |`myRange.load("values, rowCount, columnCount");`| Carga todas las propiedades de una lista delimitada por comas, en este ejemplo los valores, el recuento de filas y el número de columnas. |
|Carga todo | `myRange.load();`|Carga todas las propiedades del rango. Esta solución no es la recomendada, ya que se ralentizará el script al obtener datos innecesarios. Solo debe usar esto mientras prueba el script o si necesita todas las propiedades del objeto. |

El script debe llamar `context.sync()` antes de leer los valores cargados.

```TypeScript
let range = selectedSheet.getRange("A1:B3");
range.load ("rowCount"); // Load the property.
await context.sync(); // Synchronize with the workbook to get the property.
console.log(range.rowCount); // Read and log the property value (3).
```

También puede cargar propiedades en toda una colección. Cada objeto Collection tiene una `items` propiedad que es una matriz que contiene los objetos de esa colección. Usar `items` como inicio de una llamada jerárquica (`items\myProperty`) para `load` cargar las propiedades especificadas en cada uno de estos elementos. En el siguiente ejemplo, `resolved` se carga la `Comment` propiedad en cada `CommentCollection` objeto del objeto de una hoja de cálculo.

```TypeScript
let comments = selectedSheet.comments;
comments.load("items/resolved"); // Load the `resolved` property from every comment in this collection.
await context.sync(); // Synchronize with the workbook to get the properties.
```

> [!TIP]
> Para obtener más información sobre cómo trabajar con colecciones en scripts de Office, vea la [sección matriz del artículo usar objetos de JavaScript integrados en las secuencias de comandos de Office](javascript-objects.md#array) .

## <a name="see-also"></a>Vea también

- [Grabar, editar y crear scripts de Office en Excel en la web](../tutorials/excel-tutorial.md)
- [Leer datos de un libro con scripts de Office en Excel en la web](../tutorials/excel-read-tutorial.md)
- [Referencia de la API de scripts de Office](/javascript/api/office-scripts/overview)
- [Uso de objetos de JavaScript integrados en scripts de Office](javascript-objects.md)
