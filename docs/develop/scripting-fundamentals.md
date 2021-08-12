---
title: Conceptos básicos de los scripts de Office en Excel en la Web
description: Información del modelo de objetos y otras nociones básicas necesarias antes de escribir scripts de Office.
ms.date: 05/24/2021
localization_priority: Priority
ms.openlocfilehash: b5038dde38550e63bae872b39b9222d3defe9943ccefad85a469a5c0717fb2ef
ms.sourcegitcommit: 75f7ed8c2d23a104acc293f8ce29ea580b4fcdc5
ms.translationtype: HT
ms.contentlocale: es-ES
ms.lasthandoff: 08/11/2021
ms.locfileid: "57846704"
---
# <a name="scripting-fundamentals-for-office-scripts-in-excel-on-the-web"></a>Conceptos básicos de los Scripts de Office en Excel en la Web

En este artículo se presentan los aspectos técnicos de los scripts de Office. Obtendrá información sobre cómo funcionan conjuntamente los objetos de Excel y cómo se sincroniza el editor de código con un libro.

## <a name="typescript-the-language-of-office-scripts"></a>TypeScript: el lenguaje de Scripts de Office

Los Scripts de Office se escriben en [TypeScript](https://www.typescriptlang.org/docs/home.html), que es un superconjunto de [JavaScript](https://developer.mozilla.org/docs/Web/JavaScript). Si conoce JavaScript, parte con una gran ventaja porque la mayor parte del código es el mismo en los dos lenguajes. Recomendamos adquirir unos conocimientos de programación a nivel principiante antes de empezar con Scripts de Office. Los siguientes recursos pueden ayudarle a comprender la programación con Scripts de Office.

[!INCLUDE [Recommended coding resources](../includes/coding-basics-references.md)]

## <a name="main-function-the-scripts-starting-point"></a>Función `main`: el punto de origen del script

Cada Script de Office debe contener una función `main` con el tipo `ExcelScript.Workbook` como primer parámetro. Cuando se ejecuta la función, la aplicación Excel invoca a esta función `main` con el libro como primer parámetro. `ExcelScript.Workbook` debe ser siempre el primer parámetro.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Your code goes here
}
```

El código incluido en la función `main` se ejecuta cuando se ejecuta el script. `main` puede llamar a otras funciones en el script, pero no se ejecutará el código que no esté contenido en una función. Los scripts no pueden invocar ni llamar a otros Scripts de Office.

[Power Automate](https://flow.microsoft.com) permite conectar scripts en los flujos. Los datos se pasan entre los scripts y el flujo a través de los parámetros y se devuelve el método `main`. Encontrará información detallada sobre cómo integrar Scripts de Office con Power Automate en [Ejecutar Scripts de Office con Power Automate](power-automate-integration.md).

## <a name="object-model-overview"></a>Introducción al modelo de objetos

Para escribir un script, debe comprender cómo encajan entre sí las API de Scripts de Office. Los componentes de un libro tienen relaciones específicas entre sí. En muchos aspectos, estas relaciones coinciden con las de la Interfaz de Usuario de Excel.

- Un **Libro** contiene una o varias **Hojas de cálculo**.
- Una **Hoja de cálculo** proporciona acceso a las celdas mediante objetos de **Rango**.
- Un **Rango** representa un grupo de celdas adyacentes.
- Los **Rangos** se usan para crear y colocar **Tablas**, **Gráficos**, **Formas** y otros objetos de visualización u organización de datos.
- Una **Hoja de cálculo** contiene colecciones de aquellos objetos de datos presentes en la hoja individual.
- Los **Libros** contiene colecciones de algunos de esos objetos de datos (como **Tablas**) para todo el **Libro**.

## <a name="workbook"></a>Libro de trabajo

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

## <a name="ranges"></a>Ranges

Un rango es un grupo de celdas adyacentes en el libro. Normalmente, los scripts usan la notación de estilo A1 (por ejemplo, **B3** para la única celda de la columna **B** y la fila **3** o **C2:F4** para las celdas de las columnas de **C** a **F** y las filas de **2** a **4**) para definir rangos.

Los rangos tienen tres propiedades fundamentales: valores, fórmulas y formato. Estas propiedades obtienen o establecen los valores de celda, las fórmulas que se deben evaluar y el formato visual de las celdas. Se obtiene acceso a ellos a través de `getValues`, `getFormulas`y `getFormat`. Se pueden cambiar los valores y las fórmulas con `setValues` y `setFormulas`, mientras que el formato es un objeto `RangeFormat` formado por varios objetos más pequeños que se configuran por separado.

Los rangos usan matrices bidimensionales para administrar la información. Para obtener más información sobre el control de matrices en el marco de Scripts de Office, vea [Trabajar con rangos](javascript-objects.md#work-with-ranges).

### <a name="range-sample"></a>Ejemplo de rango

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
        ["Chocolate", 10, 9.54],
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

### <a name="the-types-of-range-values"></a>Los tipos de valores del intervalo

Cada celda tiene un valor. Este valor es el valor subyacente introducido en la celda, que puede ser diferente del texto que se muestra en Excel. Por ejemplo, puede que vea "02/05/2021" en la celda como una fecha, pero el valor real es 44318. Esta visualización puede cambiarse con el formato de número, pero el valor real y el tipo en la celda solo cambian cuando se establece un valor nuevo.

Cuando use el valor de la celda, es importante informar a TypeScript qué valor espera obtener de una celda o rango. Una celda contiene uno de los siguientes tipos: `string`. `number` o `boolean`. Para que el script trate los valores devueltos como uno de estos tipos, debe declarar el tipo.

El script siguiente obtiene el precio medio de la tabla en el ejemplo anterior. Anote el código `priceRange.getValues() as number[][]`. Esto [declara](https://www.typescriptlang.org/docs/handbook/2/everyday-types.html#type-assertions) que el tipo de los valores de rango sea un `number[][]`. Después, todos los valores de esa matriz se pueden tratar como números en el script.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the active worksheet.
  let sheet = workbook.getActiveWorksheet();

  // Get the "Unit Price" column. 
  // The result of calling getValues is declared to be a number[][] so that we can perform arithmetic operations.
  let priceRange = sheet.getRange("D3:D5");
  let prices = priceRange.getValues() as number[][];

  // Get the average price.
  let totalPrices = 0;
  prices.forEach((price) => totalPrices += price[0]);
  let averagePrice = totalPrices / prices.length;
  console.log(averagePrice);
}
```

## <a name="charts-tables-and-other-data-objects"></a>Gráficos, tablas y otros objetos de datos

Los scripts pueden crear y manipular las estructuras y visualizaciones de datos en Excel. Las tablas y los gráficos son dos de los objetos más usados, pero las API son compatibles con tablas dinámicas, formas, imágenes, etc. Se almacenan en colecciones, que se tratan más adelante en este artículo.

### <a name="create-a-table"></a>Crear una tabla

Cree tablas mediante rangos rellenos de datos. El formato y los controles de tabla (como los filtros) se aplican automáticamente al rango.

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

### <a name="create-a-chart"></a>Crear un gráfico

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

## <a name="collections"></a>Colecciones

Cuando un objeto de Excel tiene una colección de uno o varios objetos del mismo tipo, los almacena en una matriz. Por ejemplo, un objeto `Workbook` contiene una `Worksheet[]`. Se accede a esta matriz con el método `Workbook.getWorksheets()`. Los métodos `get` que están en plural, como `Worksheet.getCharts()`, devuelven la colección de objetos completa como una matriz. Verá este patrón en las API de Scripts de Office: el objeto `Worksheet` tiene un método `getTables()` que devuelve una `Table[]`, el objeto `Table` tiene un método `getColumns()` que devuelve una `TableColumn[]`, etc.

La matriz devuelta es una matriz normal, por lo que tiene todas las operaciones de matriz normales disponibles para su script. También puede obtener acceso a objetos individuales de la colección con el valor del índice de matriz. Por ejemplo, `workbook.getTables()[0]` devuelve la primera tabla de la colección. Para obtener más información sobre el uso de la funcionalidad de matriz integrada en el marco de Scripts de Office, vea [Trabajar con colecciones](javascript-objects.md#work-with-collections). 

También tiene acceso a los objetos individuales de la colección mediante un método `get`. Los métodos `get` que están en singular, como `Worksheet.getTable(name)`, devuelven un único objeto y requieren un identificador o nombre para el objeto específico. Este id. o nombre normalmente se establece en el script o en la interfaz de usuario de Excel.

El siguiente script obtiene todas las tablas del libro. Luego, garantiza que se muestren los encabezados, que los botones de filtro estén visibles y que el estilo de tabla esté establecido en "TableStyleLight1".

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the table collection.
  let tables = workbook.getTables();

  // Set the table formatting properties for every table.
  tables.forEach(table => {
    table.setShowHeaders(true);
    table.setShowFilterButton(true);
    table.setPredefinedTableStyle("TableStyleLight1");
  })
}
```

## <a name="add-excel-objects-with-a-script"></a>Agregar objetos de Excel con un script

Puede agregar mediante programación objetos de documento, como tablas o gráficos, llamando al método `add` correspondiente disponible en el objeto primario.

> [!IMPORTANT]
> No agregue objetos manualmente a matrices de colección. Usar los métodos `add` en los objetos primarios por ejemplo, agregue un `Table` a un `Worksheet` con el método `Worksheet.addTable`.

El siguiente script crea una tabla en Excel en la primera hoja de cálculo del libro. Tenga en cuenta que el método `addTable` devuelve la tabla creada.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the first worksheet.
    let sheet = workbook.getWorksheets()[0];

    // Add a table that uses the data in A1:G10.
    let table = sheet.addTable(
      "A1:G10",
       true /* True because the table has headers. */
    );
    
    // Give the table a name for easy reference in other scripts.
    table.setName("MyTable");
}
```

> [!TIP]
> La mayoría de los objetos de Excel tienen un método `setName`. Esto facilita acceder a los objetos de Excel más adelante en el script o en otros scripts para el mismo libro.

### <a name="verify-an-object-exists-in-the-collection"></a>Comprobar la existencia de un objeto en la colección

Los scripts a menudo necesitan comprobar si existe una tabla o un objeto similar antes de continuar. Use los nombres concedidos por los scripts o emplee la interfaz de usuario de Excel para identificar los objetos necesarios y actuar en consecuencia. Los métodos `get` devuelven `undefined` cuando el objeto solicitado no está en la colección.

El script siguiente solicita una tabla denominada "Mi tabla" y usa una instrucción `if...else` para comprobar si se ha encontrado la tabla.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the table named "MyTable".
  let myTable = workbook.getTable("MyTable");

  // If the table is in the workbook, myTable will have a value.
  // Otherwise, the variable will be undefined and go to the else clause.
  if (myTable) {
    let worksheetName = myTable.getWorksheet().getName();
    console.log(`MyTable is on the ${worksheetName} worksheet`);
  } else {
    console.log(`MyTable is not in the workbook.`);
  }
}
```

Un patrón común en Scripts de Office es volver a crear una tabla, gráfico u otro objeto cada vez que se ejecuta el script. Si no necesita los datos antiguos, es mejor eliminar el objeto antiguo antes de crear el nuevo. Esto evita conflictos de nombres u otros cambios que introdujeran potencialmente otros usuarios.

El script siguiente quita la tabla denominada "Mi tabla", si está presente, y después agrega una nueva tabla con el mismo nombre.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the table named "MyTable" from the first worksheet.
  let sheet = workbook.getWorksheets()[0];
  let tableName = "MyTable";
  let oldTable = sheet.getTable(tableName);

  // If the table exists, remove it.
  if (oldTable) {
    oldTable.delete();
  }

  // Add a new table with the same name.
  let newTable = sheet.addTable("A1:G10", true);
  newTable.setName(tableName);
}
```

## <a name="remove-excel-objects-with-a-script"></a>Quitar objetos de Excel con un script

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

## <a name="further-reading-on-the-object-model"></a>Más información sobre el modelo de objetos

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
- [Procedimientos recomendados para Scripts de Office](best-practices.md)
- [Centro para desarrolladores de Scripts de Office](https://developer.microsoft.com/office-scripts)
