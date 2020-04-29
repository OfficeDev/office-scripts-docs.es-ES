---
title: Conceptos básicos de los scripts de Office en Excel en la Web
description: Información del modelo de objetos y otras nociones básicas necesarias antes de escribir scripts de Office.
ms.date: 04/24/2020
localization_priority: Priority
ms.openlocfilehash: 8449654e359f665677f3d416a8e28fa4d6930f26
ms.sourcegitcommit: 350bd2447f616fa87bb23ac826c7731fb813986b
ms.translationtype: HT
ms.contentlocale: es-ES
ms.lasthandoff: 04/28/2020
ms.locfileid: "43919801"
---
# <a name="scripting-fundamentals-for-office-scripts-in-excel-on-the-web-preview"></a>Conceptos básicos de los scripts de Office en Excel en la Web (vista previa)

En este artículo se presentan los aspectos técnicos de los scripts de Office. Obtendrá información sobre cómo funcionan conjuntamente los objetos de Excel y cómo se sincroniza el editor de código con un libro.

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="object-model"></a>Modelo de objetos

Para comprender las API de Excel, debe comprender cómo se relacionan entre sí los componentes de un libro.

- Un **Libro** contiene una o varias **Hojas de cálculo**.
- Una **Hoja de cálculo** proporciona acceso a las celdas mediante objetos de **Rango**.
- Un **Rango** representa un grupo de celdas adyacentes.
- Los **Rangos** se usan para crear y colocar **Tablas**, **Gráficos**, **Formas** y otros objetos de visualización u organización de datos.
- Una **Hoja de cálculo** contiene colecciones de aquellos objetos de datos presentes en la hoja individual.
- Los **Libros** contiene colecciones de algunos de esos objetos de datos (como **Tablas**) para todo el **Libro**.

### <a name="ranges"></a>Rangos

Un rango es un grupo de celdas adyacentes en el libro. Normalmente, los scripts usan la notación de estilo A1 (por ejemplo, **B3** para la única celda de la columna **B** y la fila **3** o **C2:F4** para las celdas de las columnas de **C** a **F** y las filas de **2** a **4**) para definir rangos. 

Los rangos tienen tres propiedades básicas: `values`, `formulas` y `format`. Estas propiedades obtienen o establecen los valores de celda, las fórmulas que se deben evaluar y el formato visual de las celdas.

#### <a name="range-sample"></a>Ejemplo de rango

En el siguiente ejemplo se muestra cómo crear registros de ventas. Este script usa objetos `Range` para establecer los valores, fórmulas y formatos.

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

Al ejecutar este script se crean los siguientes datos en la hoja de cálculo actual:

![Un registro de ventas que muestra filas de valores, una columna de fórmulas y los encabezados con formato.](../images/range-sample.png)

### <a name="charts-tables-and-other-data-objects"></a>Gráficos, tablas y otros objetos de datos

Los scripts pueden crear y manipular las estructuras y visualizaciones de datos en Excel. Las tablas y los gráficos son dos de los objetos más usados, pero las API son compatibles con tablas dinámicas, formas, imágenes, etc.

#### <a name="creating-a-table"></a>Crear una tabla

Cree tablas con rangos con datos. El formato y los controles de tabla (como filtros) se aplican automáticamente al rango.

El siguiente script crea una tabla con los rangos del ejemplo anterior.

```TypeScript
async function main(context: Excel.RequestContext) {
   let sheet = context.workbook.worksheets.getActiveWorksheet();
   sheet.tables.add("B2:E5", true);
}
```

Ejecutar este script en la hoja de cálculo con los datos anteriores crea la tabla siguiente:

![Una tabla creada con el registro de ventas anterior.](../images/table-sample.png)

#### <a name="creating-a-chart"></a>Crear un gráfico

Cree gráficos para visualizar los datos de un rango. Los scripts permiten decenas de tipos de gráficos, cada uno de los cuales se puede personalizar según sus necesidades.

El siguiente script crea un gráfico de columnas simple para tres elementos y los coloca 100 píxeles por debajo de la parte superior de la hoja de cálculo.

```TypeScript
async function main(context: Excel.RequestContext) {
  let sheet = context.workbook.worksheets.getActiveWorksheet();
  let chart = sheet.charts.add(Excel.ChartType.columnStacked, sheet.getRange("B3:C5"));
  chart.top = 100;
}
```

Ejecutar este script en la hoja de cálculo con la tabla anterior crea el gráfico siguiente:

![Un gráfico de columnas que muestra cantidades de tres elementos del registro de ventas anterior.](../images/chart-sample.png)

### <a name="further-reading-on-the-object-model"></a>Más información sobre el modelo de objetos

La [Documentación de referencia de las API de scripts de Office](/javascript/api/office-scripts/overview) es una lista completa de los objetos que se usan en los scripts de Office. Allí, puede usar la tabla de contenido para navegar hasta cualquier clase de la que quiera obtener más información. Las siguientes son algunas de las páginas habitualmente consultadas.

- [Chart](/javascript/api/office-scripts/excel/excel.chart)
- [Comment](/javascript/api/office-scripts/excel/excel.comment)
- [PivotTable](/javascript/api/office-scripts/excel/excel.pivottable)
- [Range](/javascript/api/office-scripts/excel/excel.range)
- [RangeFormat](/javascript/api/office-scripts/excel/excel.rangeformat)
- [Shape](/javascript/api/office-scripts/excel/excel.shape)
- [Table](/javascript/api/office-scripts/excel/excel.table)
- [Workbook](/javascript/api/office-scripts/excel/excel.workbook)
- [Worksheet](/javascript/api/office-scripts/excel/excel.worksheet)

## <a name="main-function"></a>Función `main`

Todos los scripts de Office deben contener una función `main` con la siguiente firma, incluyendo la definición de tipo de `Excel.RequestContext`:

```TypeScript
async function main(context: Excel.RequestContext) {
    // Your Excel Script
}
```

El código incluido en la función `main` se ejecuta cuando se ejecuta el script. `main` puede llamar a otras funciones en el script, pero no se ejecutará el código que no esté contenido en una función.

## <a name="context"></a>Context

La función `main` acepta un parámetro de `Excel.RequestContext`, denominado `context`. Considere `context` como el puente entre el script y el libro. El script obtiene acceso al libro con el objeto `context` y usa ese `context` para enviar datos hacia adelante y hacia atrás.

El objeto `context` es necesario porque el script y Excel se ejecutan en diferentes procesos y ubicaciones. El script tendrá que realizar cambios o consultar datos en el libro en la nube. El objeto `context` administra estas transacciones.

## <a name="sync-and-load"></a>Sync y Load

Como el script y el libro se ejecutan en distintas ubicaciones, cualquier transferencia de datos entre ambos necesita tiempo. Para mejorar el rendimiento del script, los comandos se ponen en cola hasta que el script llama explícitamente a la operación `sync` para sincronizar el script y el libro. El script puede funcionar de forma independiente hasta que necesite realizar cualquiera de las siguientes acciones:

- Lea los datos del libro (después de una operación `load` o método que devuelve un [ClientResult](/javascript/api/office-scripts/excel/excel.clientresult)).
- Escribir datos en el libro (por lo general, porque el script ha terminado).

En la imagen siguiente se muestra un ejemplo de flujo de control entre el script y el libro:

![Diagrama en el que se muestran las operaciones de lectura y escritura en el libro desde el script.](../images/load-sync.png)

### <a name="sync"></a>Sync

Siempre que el script tenga que leer o escribir datos en el libro, llama al método `RequestContext.sync` como se muestra a continuación:

```TypeScript
await context.sync();
```

> [!NOTE]
> Se llama de forma implícita a `context.sync()` cuando finaliza un script.

Una vez completada la operación `sync`, el libro se actualiza para reflejar las operaciones de escritura que haya especificado el script. Una operación de escritura consiste en establecer cualquier propiedad en un objeto de Excel (por ejemplo, `range.format.fill.color = "red"`) o llamar a un método para cambiar una propiedad (por ejemplo, `range.format.autoFitColumns()`). La operación `sync` también lee cualquier valor del libro solicitado por el script mediante una operación `load` o un método que devuelve un `ClientResult`(como se describe en la sección siguiente).

Sincronizar el script con el libro puede tardar un tiempo, según la red. Debe minimizar el número de llamadas `sync` para que el script se ejecute con rapidez.  

### <a name="load"></a>Load

Un script debe cargar los datos del libro antes de leerlo. Sin embargo, la carga frecuente de datos de todo el libro reducirá considerablemente la velocidad del script. En lugar de ello, el método `load` permite que el script indique específicamente qué datos se deben recuperar del libro.

El método `load` está disponible en cada objeto de Excel. El script debe cargar las propiedades de un objeto antes de poder leerlas. Si no, se producirá un error.

Los ejemplos siguientes usan un objeto `Range` para mostrar las tres formas en que se puede usar el método `load` para cargar datos.

|Objetivo |Comando de ejemplo | Efecto |
|:--|:--|:--|
|Cargar una propiedad |`myRange.load("values");` | Carga una única propiedad, en este caso la matriz bidimensional de valores en este rango. |
|Cargar varias propiedades |`myRange.load("values, rowCount, columnCount");`| Carga todas las propiedades de una lista delimitada por comas, en este ejemplo, los valores, el número de filas y el número de columnas. |
|Cargar todo | `myRange.load();`|Carga todas las propiedades en el rango. Esta no es una solución recomendable, ya que reducirá la velocidad del script al obtener datos innecesarios. Solo debería usarlo cuando pruebe el script o si necesita todas las propiedades del objeto. |

El script debe llamar a `context.sync()` antes de leer cualquier valor cargado.

```TypeScript
let range = selectedSheet.getRange("A1:B3");
range.load ("rowCount"); // Load the property.
await context.sync(); // Synchronize with the workbook to get the property.
console.log(range.rowCount); // Read and log the property value (3).
```

También puede cargar propiedades de toda la colección. Cada objeto de la colección tiene una propiedad `items` que es una matriz que contiene los objetos de esa colección. El uso de `items` como inicio de una llamada jerárquica (`items\myProperty`) a `load` carga las propiedades especificadas en cada uno de esos elementos. El ejemplo siguiente carga la propiedad `resolved` en cada objeto `Comment` del objeto `CommentCollection` de una hoja de cálculo.

```TypeScript
let comments = selectedSheet.comments;
comments.load("items/resolved"); // Load the `resolved` property from every comment in this collection.
await context.sync(); // Synchronize with the workbook to get the properties.
```

> [!TIP]
> Para obtener más información sobre cómo trabajar con colecciones en scripts de Office, consulte el artículo [Sección Array de Usar objetos integrados de JavaScript en los scripts de Office](javascript-objects.md#array).

### <a name="clientresult"></a>ClientResult

Los métodos que devuelven información del libro tienen un patrón similar al paradigma `load`/`sync`. Por ejemplo, `TableCollection.getCount` obtiene el número de tablas de la colección. `getCount` devuelve un `ClientResult<number>`, lo que significa que la propiedad `value` en el `ClientResult` de retorno es un número. El script no puede acceder a ese valor hasta que se llama a `context.sync()`. De forma muy similar a la carga de una propiedad, el `value` es un valor local "vacío" hasta esa llamada `sync`.

El siguiente script obtiene el número total de tablas en el libro y registra ese número en la consola.

```TypeScript
async function main(context: Excel.RequestContext) {
  let tableCount = context.workbook.tables.getCount();

  // This sync call implicitly loads tableCount.value.
  // Any other ClientResult values are loaded too.
  await context.sync();

  // Trying to log the value before calling sync would throw an error.
  console.log(tableCount.value);
}
```

## <a name="see-also"></a>Vea también

- [Grabar, editar y crear scripts de Office en Excel en la Web](../tutorials/excel-tutorial.md)
- [Leer datos de libros con scripts de Office en Excel en la Web](../tutorials/excel-read-tutorial.md)
- [Referencia de API de scripts de Office](/javascript/api/office-scripts/overview)
- [Usar objetos integrados de JavaScript en los scripts de Office](javascript-objects.md)
