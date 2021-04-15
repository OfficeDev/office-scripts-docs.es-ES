---
title: Conceptos básicos de los scripts de Office en Excel en la Web
description: Información del modelo de objetos y otras nociones básicas necesarias antes de escribir scripts de Office.
ms.date: 07/08/2020
localization_priority: Priority
ms.openlocfilehash: 2c2fd683e77a0dfbfd3e9df8c79db31e78ceee8b
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: HT
ms.contentlocale: es-ES
ms.lasthandoff: 04/14/2021
ms.locfileid: "51755066"
---
# <a name="scripting-fundamentals-for-office-scripts-in-excel-on-the-web-preview"></a>Conceptos básicos de los scripts de Office en Excel en la Web (vista previa)

En este artículo se presentan los aspectos técnicos de los scripts de Office. Obtendrá información sobre cómo funcionan conjuntamente los objetos de Excel y cómo se sincroniza el editor de código con un libro.

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="main-function"></a>`main` Función

Cada Script de Office debe contener una función `main` con el tipo de `ExcelScript.Workbook` como primer parámetro. Cuando se ejecuta la función, la aplicación Excel invoca a esta función `main` al proporcionar el libro como primer parámetro. Por lo tanto, es importante no modificar la firma básica de la función `main` una vez que se haya grabado el script o se haya creado un nuevo script desde el editor de código.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Your code goes here
}
```

El código incluido en la función `main` se ejecuta cuando se ejecuta el script. `main` puede llamar a otras funciones en el script, pero no se ejecutará el código que no esté contenido en una función.

> [!CAUTION]
> Si su función `main` es similar a `async function main(context: Excel.RequestContext)`, el script usa el modelo de API asincrónica antiguo. Para más información (incluido cómo convertir el script al modelo de API actual), consulte [Soporte de Scripts de Office antiguos que usan las API asincrónicas](excel-async-model.md).

## <a name="object-model"></a>Modelo de objetos

Para escribir un script, debe comprender cómo se encajan entre sí las API de Script de Office. Los componentes de un libro tienen relaciones específicas entre sí. En muchos aspectos, estas relaciones coinciden con las de la Interfaz de Usuario de Excel.

- Un **Libro** contiene una o varias **Hojas de cálculo**.
- Una **Hoja de cálculo** proporciona acceso a las celdas mediante objetos de **Rango**.
- Un **Rango** representa un grupo de celdas adyacentes.
- Los **Rangos** se usan para crear y colocar **Tablas**, **Gráficos**, **Formas** y otros objetos de visualización u organización de datos.
- Una **Hoja de cálculo** contiene colecciones de aquellos objetos de datos presentes en la hoja individual.
- Los **Libros** contiene colecciones de algunos de esos objetos de datos (como **Tablas**) para todo el **Libro**.

### <a name="workbook"></a>Libro de trabajo

Todas las secuencias de script proporcionan un objeto `workbook`de tipo`Workbook` por la función`main`. Esto representa el objeto de nivel superior con el cual su script interactúa con el libro de trabajo de Excel.

El siguiente script obtiene la hoja de cálculo activa del libro y registra su nombre.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the active worksheet.
    let sheet = workbook.getActiveWorksheet();

    // Display the current worksheet's name.
    console.log(sheet.getName());
}
```

### <a name="ranges"></a>Ranges

Un rango es un grupo de celdas adyacentes en el libro. Normalmente, los scripts usan la notación de estilo A1 (por ejemplo, **B3** para la única celda de la columna **B** y la fila **3** o **C2:F4** para las celdas de las columnas de **C** a **F** y las filas de **2** a **4**) para definir rangos.

Los rangos tienen tres propiedades fundamentales: valores, fórmulas y formato. Estas propiedades obtienen o establecen los valores de celda, las fórmulas que se deben evaluar y el formato visual de las celdas. Se obtiene acceso a ellos a través de `getValues`, `getFormulas`y `getFormat`. Se pueden cambiar los valores y las fórmulas con `setValues` y `setFormulas`, mientras que el formato es un objeto `RangeFormat` formado por varios objetos más pequeños que se configuran por separado.

Los rangos usan matrices bidimensionales para administrar la información. Para obtener más información sobre cómo administrar estas matrices en el marco de Scripts de Office, consulte la sección [que trabaja con rangos en el uso de objetos de JavaScript integrados en los scripts de Office](javascript-objects.md#working-with-ranges).

#### <a name="range-sample"></a>Ejemplo de rango

En el siguiente ejemplo se muestra cómo crear registros de ventas. Este script usa objetos `Range` para establecer los valores, fórmulas y formatos.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the active worksheet.
    let sheet = workbook.getActiveWorksheet();

    // Create the headers and format them to stand out.
    let headers = [["Product", "Quantity", "Unit Price", "Totals"]];
    let headerRange = sheet.getRange("B2:E2");
    headerRange.setValues(headers);
    headerRange.getFormat().getFill().setColor("#4472C4");
    headerRange.getFormat().getFont().setColor("white");

    // Create the product data rows.
    let productData = [
        ["Almonds", 6, 7.5],
        ["Coffee", 20, 34.5],
        ["Chocolate", 10, 9.56],
    ];
    let dataRange = sheet.getRange("B3:D5");
    dataRange.setValues(productData);

    // Create the formulas to total the amounts sold.
    let totalFormulas = [
        ["=C3 * D3"],
        ["=C4 * D4"],
        ["=C5 * D5"],
        ["=SUM(E3:E5)"],
    ];
    let totalRange = sheet.getRange("E3:E6");
    totalRange.setFormulas(totalFormulas);
    totalRange.getFormat().getFont().setBold(true);

    // Display the totals as US dollar amounts.
    totalRange.setNumberFormat("$0.00");
}
```

Al ejecutar este script se crean los siguientes datos en la hoja de cálculo actual:

:::image type="content" source="../images/range-sample.png" alt-text="Una hoja de cálculo que contiene un registro de ventas que se compone de filas de valor, una columna de fórmula y encabezados con formato.":::

### <a name="charts-tables-and-other-data-objects"></a>Gráficos, tablas y otros objetos de datos

Los scripts pueden crear y manipular las estructuras y visualizaciones de datos en Excel. Las tablas y los gráficos son dos de los objetos más usados, pero las API son compatibles con tablas dinámicas, formas, imágenes, etc. Se almacenan en colecciones, que se tratan más adelante en este artículo.

#### <a name="creating-a-table"></a>Crear una tabla

Cree tablas con rangos con datos. El formato y los controles de tabla (como filtros) se aplican automáticamente al rango.

El siguiente script crea una tabla con los rangos del ejemplo anterior.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the active worksheet.
    let sheet = workbook.getActiveWorksheet();

    // Add a table that has headers using the data from B2:E5.
    sheet.addTable("B2:E5", true);
}
```

Ejecutar este script en la hoja de cálculo con los datos anteriores crea la tabla siguiente:

:::image type="content" source="../images/table-sample.png" alt-text="Una hoja de cálculo que contiene una tabla hecha a partir del registro de ventas anterior.":::

#### <a name="creating-a-chart"></a>Crear un gráfico

Cree gráficos para visualizar los datos de un rango. Los scripts permiten decenas de tipos de gráficos, cada uno de los cuales se puede personalizar según sus necesidades.

El siguiente script crea un gráfico de columnas simple para tres elementos y los coloca 100 píxeles por debajo de la parte superior de la hoja de cálculo.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the active worksheet.
    let sheet = workbook.getActiveWorksheet();

    // Create a column chart using the data from B3:C5.
    let chart = sheet.addChart(
        ExcelScript.ChartType.columnStacked,
        sheet.getRange("B3:C5")
    );

    // Set the margin of the chart to be 100 pixels from the top of the screen.
    chart.setTop(100);
}
```

Ejecutar este script en la hoja de cálculo con la tabla anterior crea el gráfico siguiente:

:::image type="content" source="../images/chart-sample.png" alt-text="Un gráfico de columnas que muestra cantidades de tres elementos del registro de ventas anterior.":::

### <a name="collections-and-other-object-relations"></a>Colecciones y otras relaciones de objeto

Se puede acceder a cualquier objeto secundario a través de su objeto primario. Por ejemplo, puede leer `Worksheets` del objeto `Workbook`. Se producirá un método `get` relacionado en la clase principal (por ejemplo, `Workbook.getWorksheets()` o `Workbook.getWorksheet(name)`). Los métodos `get` que son singulares devuelven un único objeto y requieren un identificador o nombre para el objeto específico (como el nombre de una hoja de cálculo). Los métodos `get`que están en plural devuelven la colección de objetos completa como una matriz. Si la colección está vacía, obtendrá una matriz vacía (`[]`).

Una vez que se ha recuperado la colección, puede usar operaciones de matriz normales como obtener su `length` o usar `for`, `for..of`,`while`bucles para la iteración o usar métodos de matriz TypeScript como `map`, `forEach` en ellas. También puede obtener acceso a objetos individuales de la colección con el valor del índice de matriz. Por ejemplo, `workbook.getTables()[0]` devuelve la primera tabla de la colección. Para obtener más información sobre la funcionalidad de estas matrices integradas en el marco de Scripts de Office, consulte la sección [que trabaja con rangos en el uso de objetos de JavaScript integrados en los scripts de Office ](javascript-objects.md#working-with-collections)

El siguiente script obtiene todas las tablas del libro. Luego, asegura que se muestran los encabezados, que los botones de filtro están visibles y que el estilo de tabla está establecido en "TableStyleLight1".

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  /* Get table collection */
  const tables = workbook.getTables();
  /* Set table formatting properties */
  tables.forEach(table => {
    table.setShowHeaders(true);
    table.setShowFilterButton(true);
    table.setPredefinedTableStyle("TableStyleLight1");
  })
}
```

#### <a name="adding-excel-objects-with-a-script"></a>Agregar objetos de Excel con un script

Puede agregar mediante programación objetos de documento, como tablas o gráficos, llamando al método `add` correspondiente disponible en el objeto primario.

> [!NOTE]
> No agregue objetos manualmente a matrices de colección. Usar los métodos `add` en los objetos primarios por ejemplo, agregue un `Table` a un `Worksheet` con el método `Worksheet.addTable`.

El siguiente script crea una tabla en Excel en la primera hoja de cálculo del libro. Tenga en cuenta que el método `addTable` devuelve la tabla creada.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the first worksheet.
    let sheet = workbook.getWorksheets()[0];

    // Add a table that uses the data in C3:G10.
    let table = sheet.addTable(
      "C3:G10",
       true /* True because the table has headers. */
    );
}
```

## <a name="removing-excel-objects-with-a-script"></a>Quitar objetos de Excel con un script

Para eliminar un objeto, llame al método `delete`del objeto.

> [!NOTE]
> Al igual que con la adición de objetos, no elimine objetos de matrices de colecciones de forma manual. Use los métodos `delete` en los objetos de tipo de colección. Por ejemplo, quitar un `Table` de un `Worksheet` con `Table.delete`.

El siguiente script elimina la primera hoja de trabajo del libro de trabajo.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get first worksheet.
    let sheet = workbook.getWorksheets()[0];

    // Remove that worksheet from the workbook.
    sheet.delete();
}
```

### <a name="further-reading-on-the-object-model"></a>Más información sobre el modelo de objetos

La [Documentación de referencia de las API de scripts de Office](/javascript/api/office-scripts/overview) es una lista completa de los objetos que se usan en los scripts de Office. Allí, puede usar la tabla de contenido para navegar hasta cualquier clase de la que quiera obtener más información. Las siguientes son algunas de las páginas habitualmente consultadas.

- [Chart](/javascript/api/office-scripts/excelscript/excelscript.chart)
- [Comment](/javascript/api/office-scripts/excelscript/excelscript.comment)
- [PivotTable](/javascript/api/office-scripts/excelscript/excelscript.pivottable)
- [Range](/javascript/api/office-scripts/excelscript/excelscript.range)
- [RangeFormat](/javascript/api/office-scripts/excelscript/excelscript.rangeformat)
- [Shape](/javascript/api/office-scripts/excelscript/excelscript.shape)
- [Table](/javascript/api/office-scripts/excelscript/excelscript.table)
- [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook)
- [Worksheet](/javascript/api/office-scripts/excelscript/excelscript.worksheet)

## <a name="see-also"></a>Consulta también

- [Grabar, editar y crear scripts de Office en Excel en la Web](../tutorials/excel-tutorial.md)
- [Leer datos de libros con scripts de Office en Excel en la Web](../tutorials/excel-read-tutorial.md)
- [Referencia de API de scripts de Office](/javascript/api/office-scripts/overview)
- [Usar objetos integrados de JavaScript en los scripts de Office](javascript-objects.md)
