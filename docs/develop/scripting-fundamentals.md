---
title: Conceptos básicos de los scripts de Office en Excel en la Web
description: Información del modelo de objetos y otras nociones básicas necesarias antes de escribir scripts de Office.
ms.date: 05/10/2021
localization_priority: Priority
ms.openlocfilehash: d930c9ee36933cb0458de8cce4f1d1adc7b6a001
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: HT
ms.contentlocale: es-ES
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545106"
---
# <a name="scripting-fundamentals-for-office-scripts-in-excel-on-the-web-preview"></a><span data-ttu-id="9728d-103">Conceptos básicos de los scripts de Office en Excel en la Web (vista previa)</span><span class="sxs-lookup"><span data-stu-id="9728d-103">Scripting fundamentals for Office Scripts in Excel on the web (preview)</span></span>

<span data-ttu-id="9728d-104">En este artículo se presentan los aspectos técnicos de los scripts de Office.</span><span class="sxs-lookup"><span data-stu-id="9728d-104">This article will introduce you to the technical aspects of Office Scripts.</span></span> <span data-ttu-id="9728d-105">Obtendrá información sobre cómo funcionan conjuntamente los objetos de Excel y cómo se sincroniza el editor de código con un libro.</span><span class="sxs-lookup"><span data-stu-id="9728d-105">You'll learn how the Excel objects work together and how the Code Editor synchronizes with a workbook.</span></span>

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="typescript-the-language-of-office-scripts"></a><span data-ttu-id="9728d-106">TypeScript: el lenguaje de Scripts de Office</span><span class="sxs-lookup"><span data-stu-id="9728d-106">TypeScript: The language of Office Scripts</span></span>

<span data-ttu-id="9728d-107">Los Scripts de Office se escriben en [TypeScript](https://www.typescriptlang.org/docs/home.html), que es un superconjunto de [JavaScript](https://developer.mozilla.org/docs/Web/JavaScript).</span><span class="sxs-lookup"><span data-stu-id="9728d-107">Office Scripts are written in [TypeScript](https://www.typescriptlang.org/docs/home.html), which is a superset of [JavaScript](https://developer.mozilla.org/docs/Web/JavaScript).</span></span> <span data-ttu-id="9728d-108">Si conoce JavaScript, parte con una gran ventaja porque la mayor parte del código es el mismo en los dos lenguajes.</span><span class="sxs-lookup"><span data-stu-id="9728d-108">If you're familiar with JavaScript, your knowledge will carry over because much of the code is the same in both languages.</span></span> <span data-ttu-id="9728d-109">Recomendamos adquirir unos conocimientos de programación a nivel principiante antes de empezar con Scripts de Office.</span><span class="sxs-lookup"><span data-stu-id="9728d-109">We recommend you have some beginner-level programming knowledge before starting your Office Scripts coding journey.</span></span> <span data-ttu-id="9728d-110">Los siguientes recursos pueden ayudarle a comprender la programación con Scripts de Office.</span><span class="sxs-lookup"><span data-stu-id="9728d-110">The following resources can help you understand the coding side of Office Scripts.</span></span>

[!INCLUDE [Preview note](../includes/coding-basics-references.md)]

## <a name="main-function-the-scripts-starting-point"></a><span data-ttu-id="9728d-111">Función `main`: el punto de origen del script</span><span class="sxs-lookup"><span data-stu-id="9728d-111">`main` function: The script's starting point</span></span>

<span data-ttu-id="9728d-112">Cada Script de Office debe contener una función `main` con el tipo `ExcelScript.Workbook` como primer parámetro.</span><span class="sxs-lookup"><span data-stu-id="9728d-112">Each script must contain a `main` function with the `ExcelScript.Workbook` type as its first parameter.</span></span> <span data-ttu-id="9728d-113">Cuando se ejecuta la función, la aplicación Excel invoca a esta función `main` con el libro como primer parámetro.</span><span class="sxs-lookup"><span data-stu-id="9728d-113">When the function runs, the Excel application invokes the `main` function by providing the workbook as its first parameter.</span></span> <span data-ttu-id="9728d-114">`ExcelScript.Workbook` debe ser siempre el primer parámetro.</span><span class="sxs-lookup"><span data-stu-id="9728d-114">An `ExcelScript.Workbook` should always be the first parameter.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Your code goes here
}
```

<span data-ttu-id="9728d-115">El código incluido en la función `main` se ejecuta cuando se ejecuta el script.</span><span class="sxs-lookup"><span data-stu-id="9728d-115">The code inside the `main` function runs when the script is run.</span></span> <span data-ttu-id="9728d-116">`main` puede llamar a otras funciones en el script, pero no se ejecutará el código que no esté contenido en una función.</span><span class="sxs-lookup"><span data-stu-id="9728d-116">`main` can call other functions in your script, but code that's not contained in a function will not run.</span></span> <span data-ttu-id="9728d-117">Los scripts no pueden invocar ni llamar a otros Scripts de Office.</span><span class="sxs-lookup"><span data-stu-id="9728d-117">Scripts cannot invoke or call other Office Scripts.</span></span>

<span data-ttu-id="9728d-118">[Power Automate](https://flow.microsoft.com) permite conectar scripts en los flujos.</span><span class="sxs-lookup"><span data-stu-id="9728d-118">[Power Automate](https://flow.microsoft.com) allows you to connect scripts in flows.</span></span> <span data-ttu-id="9728d-119">Los datos se pasan entre los scripts y el flujo a través de los parámetros y se devuelve el método `main`.</span><span class="sxs-lookup"><span data-stu-id="9728d-119">Data is passed between the scripts and the flow through the parameters and returns of the`main` method.</span></span> <span data-ttu-id="9728d-120">Encontrará información detallada sobre cómo integrar Scripts de Office con Power Automate en [Ejecutar Scripts de Office con Power Automate](power-automate-integration.md).</span><span class="sxs-lookup"><span data-stu-id="9728d-120">How to integrate Office Scripts with Power Automate is covered in detail in [Run Office Scripts with Power Automate](power-automate-integration.md).</span></span>

## <a name="object-model-overview"></a><span data-ttu-id="9728d-121">Introducción al modelo de objetos</span><span class="sxs-lookup"><span data-stu-id="9728d-121">Object model overview</span></span>

<span data-ttu-id="9728d-122">Para escribir un script, debe comprender cómo encajan entre sí las API de Scripts de Office.</span><span class="sxs-lookup"><span data-stu-id="9728d-122">To write a script, you need to understand how the Office Scripts APIs fit together.</span></span> <span data-ttu-id="9728d-123">Los componentes de un libro tienen relaciones específicas entre sí.</span><span class="sxs-lookup"><span data-stu-id="9728d-123">The components of a workbook have specific relations to one another.</span></span> <span data-ttu-id="9728d-124">En muchos aspectos, estas relaciones coinciden con las de la Interfaz de Usuario de Excel.</span><span class="sxs-lookup"><span data-stu-id="9728d-124">In many ways, these relations match those of the Excel UI.</span></span>

- <span data-ttu-id="9728d-125">Un **Libro** contiene una o varias **Hojas de cálculo**.</span><span class="sxs-lookup"><span data-stu-id="9728d-125">A **Workbook** contains one or more **Worksheets**.</span></span>
- <span data-ttu-id="9728d-126">Una **Hoja de cálculo** proporciona acceso a las celdas mediante objetos de **Rango**.</span><span class="sxs-lookup"><span data-stu-id="9728d-126">A **Worksheet** gives access to cells through **Range** objects.</span></span>
- <span data-ttu-id="9728d-127">Un **Rango** representa un grupo de celdas adyacentes.</span><span class="sxs-lookup"><span data-stu-id="9728d-127">A **Range** represents a group of contiguous cells.</span></span>
- <span data-ttu-id="9728d-128">Los **Rangos** se usan para crear y colocar **Tablas**, **Gráficos**, **Formas** y otros objetos de visualización u organización de datos.</span><span class="sxs-lookup"><span data-stu-id="9728d-128">**Ranges** are used to create and place **Tables**, **Charts**, **Shapes**, and other data visualization or organization objects.</span></span>
- <span data-ttu-id="9728d-129">Una **Hoja de cálculo** contiene colecciones de aquellos objetos de datos presentes en la hoja individual.</span><span class="sxs-lookup"><span data-stu-id="9728d-129">A **Worksheet** contains collections of those data objects that are present in the individual sheet.</span></span>
- <span data-ttu-id="9728d-130">Los **Libros** contiene colecciones de algunos de esos objetos de datos (como **Tablas**) para todo el **Libro**.</span><span class="sxs-lookup"><span data-stu-id="9728d-130">**Workbooks** contain collections of some of those data objects (such as **Tables**) for the entire **Workbook**.</span></span>

## <a name="workbook"></a><span data-ttu-id="9728d-131">Libro de trabajo</span><span class="sxs-lookup"><span data-stu-id="9728d-131">Workbook</span></span>

<span data-ttu-id="9728d-132">Todas las secuencias de script proporcionan un objeto `workbook`de tipo`Workbook` por la función`main`.</span><span class="sxs-lookup"><span data-stu-id="9728d-132">Every script is provided a `workbook` object of type `Workbook` by the `main` function.</span></span> <span data-ttu-id="9728d-133">Esto representa el objeto de nivel superior con el cual su script interactúa con el libro de trabajo de Excel.</span><span class="sxs-lookup"><span data-stu-id="9728d-133">This represents the top level object through which your script interacts with the Excel workbook.</span></span>

<span data-ttu-id="9728d-134">El siguiente script obtiene la hoja de cálculo activa del libro y registra su nombre.</span><span class="sxs-lookup"><span data-stu-id="9728d-134">The following script gets the active worksheet from the workbook and logs its name.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the active worksheet.
    let sheet = workbook.getActiveWorksheet();

    // Display the current worksheet's name.
    console.log(sheet.getName());
}
```

## <a name="ranges"></a><span data-ttu-id="9728d-135">Ranges</span><span class="sxs-lookup"><span data-stu-id="9728d-135">Ranges</span></span>

<span data-ttu-id="9728d-136">Un rango es un grupo de celdas adyacentes en el libro.</span><span class="sxs-lookup"><span data-stu-id="9728d-136">A range is a group of contiguous cells in the workbook.</span></span> <span data-ttu-id="9728d-137">Normalmente, los scripts usan la notación de estilo A1 (por ejemplo, **B3** para la única celda de la columna **B** y la fila **3** o **C2:F4** para las celdas de las columnas de **C** a **F** y las filas de **2** a **4**) para definir rangos.</span><span class="sxs-lookup"><span data-stu-id="9728d-137">Scripts typically use A1-style notation (e.g., **B3** for the single cell in column **B** and row **3** or **C2:F4** for the cells from columns **C** through **F** and rows **2** through **4**) to define ranges.</span></span>

<span data-ttu-id="9728d-138">Los rangos tienen tres propiedades fundamentales: valores, fórmulas y formato.</span><span class="sxs-lookup"><span data-stu-id="9728d-138">Ranges have three core properties: values, formulas, and format.</span></span> <span data-ttu-id="9728d-139">Estas propiedades obtienen o establecen los valores de celda, las fórmulas que se deben evaluar y el formato visual de las celdas.</span><span class="sxs-lookup"><span data-stu-id="9728d-139">These properties get or set the cell values, formulas to be evaluated, and the visual formatting of the cells.</span></span> <span data-ttu-id="9728d-140">Se obtiene acceso a ellos a través de `getValues`, `getFormulas`y `getFormat`.</span><span class="sxs-lookup"><span data-stu-id="9728d-140">They are accessed through `getValues`, `getFormulas`, and `getFormat`.</span></span> <span data-ttu-id="9728d-141">Se pueden cambiar los valores y las fórmulas con `setValues` y `setFormulas`, mientras que el formato es un objeto `RangeFormat` formado por varios objetos más pequeños que se configuran por separado.</span><span class="sxs-lookup"><span data-stu-id="9728d-141">Values and formulas can be changed with `setValues` and `setFormulas`, while the format is a `RangeFormat` object comprised of several smaller objects that are individually set.</span></span>

<span data-ttu-id="9728d-142">Los rangos usan matrices bidimensionales para administrar la información.</span><span class="sxs-lookup"><span data-stu-id="9728d-142">Ranges use two-dimensional arrays to manage information.</span></span> <span data-ttu-id="9728d-143">Para obtener más información sobre el control de matrices en el marco de Scripts de Office, vea [Trabajar con rangos](javascript-objects.md#work-with-ranges).</span><span class="sxs-lookup"><span data-stu-id="9728d-143">For more information on handling arrays in the Office Scripts framework, see [Work with ranges](javascript-objects.md#work-with-ranges).</span></span>

### <a name="range-sample"></a><span data-ttu-id="9728d-144">Ejemplo de rango</span><span class="sxs-lookup"><span data-stu-id="9728d-144">Range sample</span></span>

<span data-ttu-id="9728d-145">En el siguiente ejemplo se muestra cómo crear registros de ventas.</span><span class="sxs-lookup"><span data-stu-id="9728d-145">The following sample shows how to create sales records.</span></span> <span data-ttu-id="9728d-146">Este script usa objetos `Range` para establecer los valores, fórmulas y formatos.</span><span class="sxs-lookup"><span data-stu-id="9728d-146">This script uses `Range` objects to set the values, formulas, and parts of the format.</span></span>

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

<span data-ttu-id="9728d-147">Al ejecutar este script se crean los siguientes datos en la hoja de cálculo actual:</span><span class="sxs-lookup"><span data-stu-id="9728d-147">Running this script creates the following data in the current worksheet:</span></span>

:::image type="content" source="../images/range-sample.png" alt-text="Una hoja de cálculo que contiene un registro de ventas compuesto por filas de valor, una columna de fórmula y encabezados con formato.":::

## <a name="charts-tables-and-other-data-objects"></a><span data-ttu-id="9728d-149">Gráficos, tablas y otros objetos de datos</span><span class="sxs-lookup"><span data-stu-id="9728d-149">Charts, tables, and other data objects</span></span>

<span data-ttu-id="9728d-150">Los scripts pueden crear y manipular las estructuras y visualizaciones de datos en Excel.</span><span class="sxs-lookup"><span data-stu-id="9728d-150">Scripts can create and manipulate the data structures and visualizations within Excel.</span></span> <span data-ttu-id="9728d-151">Las tablas y los gráficos son dos de los objetos más usados, pero las API son compatibles con tablas dinámicas, formas, imágenes, etc.</span><span class="sxs-lookup"><span data-stu-id="9728d-151">Tables and charts are two of the more commonly used objects, but the APIs support PivotTables, shapes, images, and more.</span></span> <span data-ttu-id="9728d-152">Se almacenan en colecciones, que se tratan más adelante en este artículo.</span><span class="sxs-lookup"><span data-stu-id="9728d-152">These are stored in collections, which will be discussed later in this article.</span></span>

### <a name="create-a-table"></a><span data-ttu-id="9728d-153">Crear una tabla</span><span class="sxs-lookup"><span data-stu-id="9728d-153">Create a table</span></span>

<span data-ttu-id="9728d-p113">Cree tablas mediante rangos rellenos de datos. El formato y los controles de tabla (como los filtros) se aplican automáticamente al rango.</span><span class="sxs-lookup"><span data-stu-id="9728d-p113">Create tables by using data-filled ranges. Formatting and table controls (such as filters) are automatically applied to the range.</span></span>

<span data-ttu-id="9728d-156">El siguiente script crea una tabla con los rangos del ejemplo anterior.</span><span class="sxs-lookup"><span data-stu-id="9728d-156">The following script creates a table using the ranges from the previous sample.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the active worksheet.
    let sheet = workbook.getActiveWorksheet();

    // Add a table that has headers using the data from B2:E5.
    sheet.addTable("B2:E5", true);
}
```

<span data-ttu-id="9728d-157">Ejecutar este script en la hoja de cálculo con los datos anteriores crea la tabla siguiente:</span><span class="sxs-lookup"><span data-stu-id="9728d-157">Running this script on the worksheet with the previous data creates the following table:</span></span>

:::image type="content" source="../images/table-sample.png" alt-text="Una hoja de cálculo que contiene una tabla hecha a partir del registro de ventas anterior.":::

### <a name="create-a-chart"></a><span data-ttu-id="9728d-159">Crear un gráfico</span><span class="sxs-lookup"><span data-stu-id="9728d-159">Create a chart</span></span>

<span data-ttu-id="9728d-160">Cree gráficos para visualizar los datos de un rango.</span><span class="sxs-lookup"><span data-stu-id="9728d-160">Create charts to visualize the data in a range.</span></span> <span data-ttu-id="9728d-161">Los scripts permiten decenas de tipos de gráficos, cada uno de los cuales se puede personalizar según sus necesidades.</span><span class="sxs-lookup"><span data-stu-id="9728d-161">Scripts allow for dozens of chart varieties, each of which can be customized to suit your needs.</span></span>

<span data-ttu-id="9728d-162">El siguiente script crea un gráfico de columnas simple para tres elementos y los coloca 100 píxeles por debajo de la parte superior de la hoja de cálculo.</span><span class="sxs-lookup"><span data-stu-id="9728d-162">The following script creates a simple column chart for three items and places it 100 pixels below the top of the worksheet.</span></span>

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

<span data-ttu-id="9728d-163">Ejecutar este script en la hoja de cálculo con la tabla anterior crea el gráfico siguiente:</span><span class="sxs-lookup"><span data-stu-id="9728d-163">Running this script on the worksheet with the previous table creates the following chart:</span></span>

:::image type="content" source="../images/chart-sample.png" alt-text="Un gráfico de columnas que muestra cantidades de tres elementos del registro de ventas anterior.":::

## <a name="collections"></a><span data-ttu-id="9728d-165">Colecciones</span><span class="sxs-lookup"><span data-stu-id="9728d-165">Collections</span></span>

<span data-ttu-id="9728d-166">Cuando un objeto de Excel tiene una colección de uno o varios objetos del mismo tipo, los almacena en una matriz.</span><span class="sxs-lookup"><span data-stu-id="9728d-166">When an Excel object has a collection of one or more objects of the same type, it stores them in an array.</span></span> <span data-ttu-id="9728d-167">Por ejemplo, un objeto `Workbook` contiene una `Worksheet[]`.</span><span class="sxs-lookup"><span data-stu-id="9728d-167">For example, a `Workbook` object contains a `Worksheet[]`.</span></span> <span data-ttu-id="9728d-168">Se accede a esta matriz con el método `Workbook.getWorksheets()`.</span><span class="sxs-lookup"><span data-stu-id="9728d-168">This array is accessed by the `Workbook.getWorksheets()` method.</span></span> <span data-ttu-id="9728d-169">Los métodos `get` que están en plural, como `Worksheet.getCharts()`, devuelven la colección de objetos completa como una matriz.</span><span class="sxs-lookup"><span data-stu-id="9728d-169">`get` methods that are plural, such as `Worksheet.getCharts()`, return the entire object collection as an array.</span></span> <span data-ttu-id="9728d-170">Verá este patrón en las API de Scripts de Office: el objeto `Worksheet` tiene un método `getTables()` que devuelve una `Table[]`, el objeto `Table` tiene un método `getColumns()` que devuelve una `TableColumn[]`, etc.</span><span class="sxs-lookup"><span data-stu-id="9728d-170">You'll see this pattern throughout the Office Scripts APIs: the `Worksheet` object has a `getTables()` method that returns a `Table[]`, the `Table` object has a `getColumns()` method that returns a `TableColumn[]`, as so on.</span></span>

<span data-ttu-id="9728d-171">La matriz devuelta es una matriz normal, por lo que tiene todas las operaciones de matriz normales disponibles para su script.</span><span class="sxs-lookup"><span data-stu-id="9728d-171">The returned array is a normal array, so all the regular array operations are available for your script.</span></span> <span data-ttu-id="9728d-172">También puede obtener acceso a objetos individuales de la colección con el valor del índice de matriz.</span><span class="sxs-lookup"><span data-stu-id="9728d-172">You can also access individual objects within the collection using the array index value.</span></span> <span data-ttu-id="9728d-173">Por ejemplo, `workbook.getTables()[0]` devuelve la primera tabla de la colección.</span><span class="sxs-lookup"><span data-stu-id="9728d-173">For example, `workbook.getTables()[0]` returns the first table in the collection.</span></span> <span data-ttu-id="9728d-174">Para obtener más información sobre el uso de la funcionalidad de matriz integrada en el marco de Scripts de Office, vea [Trabajar con colecciones](javascript-objects.md#work-with-collections).</span><span class="sxs-lookup"><span data-stu-id="9728d-174">For more information on using the built-in array functionality with the Office Scripts framework, see [Work with collections](javascript-objects.md#work-with-collections).</span></span> 

<span data-ttu-id="9728d-175">También tiene acceso a los objetos individuales de la colección mediante un método `get`.</span><span class="sxs-lookup"><span data-stu-id="9728d-175">Individual objects are also accessed from the collection through a `get` method.</span></span> <span data-ttu-id="9728d-176">Los métodos `get` que están en singular, como `Worksheet.getTable(name)`, devuelven un único objeto y requieren un identificador o nombre para el objeto específico.</span><span class="sxs-lookup"><span data-stu-id="9728d-176">`get` methods that are singular, such as `Worksheet.getTable(name)`, return a single object and require an ID or name for the specific object.</span></span> <span data-ttu-id="9728d-177">Este id. o nombre normalmente se establece en el script o en la interfaz de usuario de Excel.</span><span class="sxs-lookup"><span data-stu-id="9728d-177">This ID or name is usually set by the script or through the Excel UI.</span></span>

<span data-ttu-id="9728d-p118">El siguiente script obtiene todas las tablas del libro. Luego, garantiza que se muestren los encabezados, que los botones de filtro estén visibles y que el estilo de tabla esté establecido en "TableStyleLight1".</span><span class="sxs-lookup"><span data-stu-id="9728d-p118">The following script gets all tables in the workbook. It then ensures the headers are displays, the filter buttons are visible, and the table style is set to "TableStyleLight1".</span></span>

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

## <a name="add-excel-objects-with-a-script"></a><span data-ttu-id="9728d-180">Agregar objetos de Excel con un script</span><span class="sxs-lookup"><span data-stu-id="9728d-180">Add Excel objects with a script</span></span>

<span data-ttu-id="9728d-181">Puede agregar mediante programación objetos de documento, como tablas o gráficos, llamando al método `add` correspondiente disponible en el objeto primario.</span><span class="sxs-lookup"><span data-stu-id="9728d-181">You can programmatically add document objects, such as tables or charts, by calling the corresponding `add` method available on the parent object.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="9728d-182">No agregue objetos manualmente a matrices de colección.</span><span class="sxs-lookup"><span data-stu-id="9728d-182">Do not manually add objects to collection arrays.</span></span> <span data-ttu-id="9728d-183">Usar los métodos `add` en los objetos primarios por ejemplo, agregue un `Table` a un `Worksheet` con el método `Worksheet.addTable`.</span><span class="sxs-lookup"><span data-stu-id="9728d-183">Use the `add` methods on the parent objects For example, add a `Table` to a `Worksheet` with the `Worksheet.addTable` method.</span></span>

<span data-ttu-id="9728d-184">El siguiente script crea una tabla en Excel en la primera hoja de cálculo del libro.</span><span class="sxs-lookup"><span data-stu-id="9728d-184">The following script creates a table in Excel on the first worksheet in the workbook.</span></span> <span data-ttu-id="9728d-185">Tenga en cuenta que el método `addTable` devuelve la tabla creada.</span><span class="sxs-lookup"><span data-stu-id="9728d-185">Note that the created table is returned by the `addTable` method.</span></span>

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
> <span data-ttu-id="9728d-186">La mayoría de los objetos de Excel tienen un método `setName`.</span><span class="sxs-lookup"><span data-stu-id="9728d-186">Most Excel objects have a `setName` method.</span></span> <span data-ttu-id="9728d-187">Esto facilita acceder a los objetos de Excel más adelante en el script o en otros scripts para el mismo libro.</span><span class="sxs-lookup"><span data-stu-id="9728d-187">This gives you an easy way to access Excel objects later in the script or in other scripts for the same workbook.</span></span>

### <a name="verify-an-object-exists-in-the-collection"></a><span data-ttu-id="9728d-188">Comprobar la existencia de un objeto en la colección</span><span class="sxs-lookup"><span data-stu-id="9728d-188">Verify an object exists in the collection</span></span>

<span data-ttu-id="9728d-189">Los scripts a menudo necesitan comprobar si existe una tabla o un objeto similar antes de continuar.</span><span class="sxs-lookup"><span data-stu-id="9728d-189">Scripts often need to check if a table or similar object exists before continuing.</span></span> <span data-ttu-id="9728d-190">Use los nombres concedidos por los scripts o emplee la interfaz de usuario de Excel para identificar los objetos necesarios y actuar en consecuencia.</span><span class="sxs-lookup"><span data-stu-id="9728d-190">Use the names given by scripts or through the Excel UI to identify necessary objects and act accordingly.</span></span> <span data-ttu-id="9728d-191">Los métodos `get` devuelven `undefined` cuando el objeto solicitado no está en la colección.</span><span class="sxs-lookup"><span data-stu-id="9728d-191">`get` methods return `undefined` when the requested object is not in the collection.</span></span>

<span data-ttu-id="9728d-192">El script siguiente solicita una tabla denominada "Mi tabla" y usa una instrucción `if...else` para comprobar si se ha encontrado la tabla.</span><span class="sxs-lookup"><span data-stu-id="9728d-192">The following script requests a table named "MyTable" and uses an `if...else` statement to check if the table was found.</span></span>

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

<span data-ttu-id="9728d-193">Un patrón común en Scripts de Office es volver a crear una tabla, gráfico u otro objeto cada vez que se ejecuta el script.</span><span class="sxs-lookup"><span data-stu-id="9728d-193">A common pattern in Office Scripts is to recreate a table, chart, or other object every time the script is run.</span></span> <span data-ttu-id="9728d-194">Si no necesita los datos antiguos, es mejor eliminar el objeto antiguo antes de crear el nuevo.</span><span class="sxs-lookup"><span data-stu-id="9728d-194">If you don't need the old data, it's best to delete the old object before creating the new one.</span></span> <span data-ttu-id="9728d-195">Esto evita conflictos de nombres u otros cambios que introdujeran potencialmente otros usuarios.</span><span class="sxs-lookup"><span data-stu-id="9728d-195">This avoids name conflicts or other differences that may have been introduced by other users.</span></span>

<span data-ttu-id="9728d-196">El script siguiente quita la tabla denominada "Mi tabla", si está presente, y después agrega una nueva tabla con el mismo nombre.</span><span class="sxs-lookup"><span data-stu-id="9728d-196">The following script removes the table named "MyTable", if it is present, then adds a new table with the same name.</span></span>

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

## <a name="remove-excel-objects-with-a-script"></a><span data-ttu-id="9728d-197">Quitar objetos de Excel con un script</span><span class="sxs-lookup"><span data-stu-id="9728d-197">Remove Excel objects with a script</span></span>

<span data-ttu-id="9728d-198">Para eliminar un objeto, llame al método `delete`del objeto.</span><span class="sxs-lookup"><span data-stu-id="9728d-198">To delete an object, call the object's `delete` method.</span></span>

> [!NOTE]
> <span data-ttu-id="9728d-199">Al igual que con la adición de objetos, no elimine objetos de matrices de colecciones de forma manual.</span><span class="sxs-lookup"><span data-stu-id="9728d-199">As with adding objects, do not manually remove objects from collection arrays.</span></span> <span data-ttu-id="9728d-200">Use los métodos `delete` en los objetos de tipo de colección.</span><span class="sxs-lookup"><span data-stu-id="9728d-200">Use the `delete` methods on the collection-type objects.</span></span> <span data-ttu-id="9728d-201">Por ejemplo, quitar un `Table` de un `Worksheet` con `Table.delete`.</span><span class="sxs-lookup"><span data-stu-id="9728d-201">For example, remove a `Table` from a `Worksheet` using `Table.delete`.</span></span>

<span data-ttu-id="9728d-202">El siguiente script elimina la primera hoja de trabajo del libro de trabajo.</span><span class="sxs-lookup"><span data-stu-id="9728d-202">The following script removes the first worksheet in the workbook.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get first worksheet.
    let sheet = workbook.getWorksheets()[0];

    // Remove that worksheet from the workbook.
    sheet.delete();
}
```

## <a name="further-reading-on-the-object-model"></a><span data-ttu-id="9728d-203">Más información sobre el modelo de objetos</span><span class="sxs-lookup"><span data-stu-id="9728d-203">Further reading on the object model</span></span>

<span data-ttu-id="9728d-204">La [Documentación de referencia de las API de scripts de Office](/javascript/api/office-scripts/overview) es una lista completa de los objetos que se usan en los scripts de Office.</span><span class="sxs-lookup"><span data-stu-id="9728d-204">The [Office Scripts API reference documentation](/javascript/api/office-scripts/overview) is a comprehensive listing of the objects used in Office Scripts.</span></span> <span data-ttu-id="9728d-205">Allí, puede usar la tabla de contenido para navegar hasta cualquier clase de la que quiera obtener más información.</span><span class="sxs-lookup"><span data-stu-id="9728d-205">There, you can use the table of contents to navigate to any class you'd like to learn more about.</span></span> <span data-ttu-id="9728d-206">Las siguientes son algunas de las páginas habitualmente consultadas.</span><span class="sxs-lookup"><span data-stu-id="9728d-206">The following are several commonly viewed pages.</span></span>

- [<span data-ttu-id="9728d-207">Chart</span><span class="sxs-lookup"><span data-stu-id="9728d-207">Chart</span></span>](/javascript/api/office-scripts/excelscript/excelscript.chart)
- [<span data-ttu-id="9728d-208">Comment</span><span class="sxs-lookup"><span data-stu-id="9728d-208">Comment</span></span>](/javascript/api/office-scripts/excelscript/excelscript.comment)
- [<span data-ttu-id="9728d-209">PivotTable</span><span class="sxs-lookup"><span data-stu-id="9728d-209">PivotTable</span></span>](/javascript/api/office-scripts/excelscript/excelscript.pivottable)
- [<span data-ttu-id="9728d-210">Range</span><span class="sxs-lookup"><span data-stu-id="9728d-210">Range</span></span>](/javascript/api/office-scripts/excelscript/excelscript.range)
- [<span data-ttu-id="9728d-211">RangeFormat</span><span class="sxs-lookup"><span data-stu-id="9728d-211">RangeFormat</span></span>](/javascript/api/office-scripts/excelscript/excelscript.rangeformat)
- [<span data-ttu-id="9728d-212">Shape</span><span class="sxs-lookup"><span data-stu-id="9728d-212">Shape</span></span>](/javascript/api/office-scripts/excelscript/excelscript.shape)
- [<span data-ttu-id="9728d-213">Table</span><span class="sxs-lookup"><span data-stu-id="9728d-213">Table</span></span>](/javascript/api/office-scripts/excelscript/excelscript.table)
- [<span data-ttu-id="9728d-214">Workbook</span><span class="sxs-lookup"><span data-stu-id="9728d-214">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook)
- [<span data-ttu-id="9728d-215">Worksheet</span><span class="sxs-lookup"><span data-stu-id="9728d-215">Worksheet</span></span>](/javascript/api/office-scripts/excelscript/excelscript.worksheet)

## <a name="see-also"></a><span data-ttu-id="9728d-216">Consulta también</span><span class="sxs-lookup"><span data-stu-id="9728d-216">See also</span></span>

- [<span data-ttu-id="9728d-217">Grabar, editar y crear scripts de Office en Excel en la Web</span><span class="sxs-lookup"><span data-stu-id="9728d-217">Record, edit, and create Office Scripts in Excel on the web</span></span>](../tutorials/excel-tutorial.md)
- [<span data-ttu-id="9728d-218">Leer datos de libros con scripts de Office en Excel en la Web</span><span class="sxs-lookup"><span data-stu-id="9728d-218">Read workbook data with Office Scripts in Excel on the web</span></span>](../tutorials/excel-read-tutorial.md)
- [<span data-ttu-id="9728d-219">Referencia de API de scripts de Office</span><span class="sxs-lookup"><span data-stu-id="9728d-219">Office Scripts API reference</span></span>](/javascript/api/office-scripts/overview)
- [<span data-ttu-id="9728d-220">Usar objetos integrados de JavaScript en los scripts de Office</span><span class="sxs-lookup"><span data-stu-id="9728d-220">Using built-in JavaScript objects in Office Scripts</span></span>](javascript-objects.md)
- [<span data-ttu-id="9728d-221">Procedimientos recomendados para Scripts de Office</span><span class="sxs-lookup"><span data-stu-id="9728d-221">Best practices in Office Scripts</span></span>](best-practices.md)
