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
# <a name="scripting-fundamentals-for-office-scripts-in-excel-on-the-web-preview"></a><span data-ttu-id="37b7d-103">Conceptos básicos sobre scripts para scripts de Office en Excel en la web (vista previa)</span><span class="sxs-lookup"><span data-stu-id="37b7d-103">Scripting fundamentals for Office Scripts in Excel on the web (preview)</span></span>

<span data-ttu-id="37b7d-104">En este artículo se presentan los aspectos técnicos de las secuencias de comandos de Office.</span><span class="sxs-lookup"><span data-stu-id="37b7d-104">This article will introduce you to the technical aspects of Office Scripts.</span></span> <span data-ttu-id="37b7d-105">Aprenderá cómo funcionan conjuntamente los objetos de Excel y cómo se sincroniza el editor de código con un libro.</span><span class="sxs-lookup"><span data-stu-id="37b7d-105">You'll learn how the Excel objects work together and how the Code Editor synchronizes with a workbook.</span></span>

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="object-model"></a><span data-ttu-id="37b7d-106">Modelo de objetos</span><span class="sxs-lookup"><span data-stu-id="37b7d-106">Object model</span></span>

<span data-ttu-id="37b7d-107">Para comprender las API de Excel, debe comprender cómo se relacionan entre sí los componentes de un libro.</span><span class="sxs-lookup"><span data-stu-id="37b7d-107">To understand the Excel APIs, you must understand how the components of a workbook are related to one another.</span></span>

- <span data-ttu-id="37b7d-108">Un **libro** contiene una o varias **hojas de cálculo**.</span><span class="sxs-lookup"><span data-stu-id="37b7d-108">A **Workbook** contains one or more **Worksheets**.</span></span>
- <span data-ttu-id="37b7d-109">Una **hoja de cálculo** da acceso a celdas a través de objetos **Range** .</span><span class="sxs-lookup"><span data-stu-id="37b7d-109">A **Worksheet** gives access to cells through **Range** objects.</span></span>
- <span data-ttu-id="37b7d-110">Un **rango** representa un grupo de celdas contiguas.</span><span class="sxs-lookup"><span data-stu-id="37b7d-110">A **Range** represents a group of contiguous cells.</span></span>
- <span data-ttu-id="37b7d-111">Los **rangos** se usan para crear y colocar **tablas**, **gráficos**, **formas**y otros objetos de visualización o de organización de datos.</span><span class="sxs-lookup"><span data-stu-id="37b7d-111">**Ranges** are used to create and place **Tables**, **Charts**, **Shapes**, and other data visualization or organization objects.</span></span>
- <span data-ttu-id="37b7d-112">Una **hoja de cálculo** contiene colecciones de los objetos de datos que están presentes en la hoja individual.</span><span class="sxs-lookup"><span data-stu-id="37b7d-112">A **Worksheet** contains collections of those data objects that are present in the individual sheet.</span></span>
- <span data-ttu-id="37b7d-113">Los **libros** contienen colecciones de algunos de los objetos de datos (como **tablas**) para todo el **libro**.</span><span class="sxs-lookup"><span data-stu-id="37b7d-113">**Workbooks** contain collections of some of those data objects (such as **Tables**) for the entire **Workbook**.</span></span>

### <a name="ranges"></a><span data-ttu-id="37b7d-114">Ranges</span><span class="sxs-lookup"><span data-stu-id="37b7d-114">Ranges</span></span>

<span data-ttu-id="37b7d-115">Un rango es un grupo de celdas contiguas del libro.</span><span class="sxs-lookup"><span data-stu-id="37b7d-115">A range is a group of contiguous cells in the workbook.</span></span> <span data-ttu-id="37b7d-116">Normalmente, los scripts usan la notación de estilo a1 (por ejemplo, **B3** para la única celda de la fila **B** y la columna **3** o **C2: F4** para las celdas de las filas **C** a **F** y de las columnas **2** a **4**) para definir rangos.</span><span class="sxs-lookup"><span data-stu-id="37b7d-116">Scripts typically use A1-style notation (e.g. **B3** for the single cell in row **B** and column **3** or **C2:F4** for the cells from rows **C** through **F** and columns **2** through **4**) to define ranges.</span></span>

<span data-ttu-id="37b7d-117">Los rangos tienen tres propiedades `values`principales `formulas`:, `format`y.</span><span class="sxs-lookup"><span data-stu-id="37b7d-117">Ranges have three core properties: `values`, `formulas`, and `format`.</span></span> <span data-ttu-id="37b7d-118">Estas propiedades obtienen o establecen los valores de celda, las fórmulas que se van a evaluar y el formato visual de las celdas.</span><span class="sxs-lookup"><span data-stu-id="37b7d-118">These properties get or set the cell values, formulas to be evaluated, and the visual formatting of the cells.</span></span>

#### <a name="range-sample"></a><span data-ttu-id="37b7d-119">Ejemplo de intervalo</span><span class="sxs-lookup"><span data-stu-id="37b7d-119">Range sample</span></span>

<span data-ttu-id="37b7d-120">En el ejemplo siguiente se muestra cómo crear registros de ventas.</span><span class="sxs-lookup"><span data-stu-id="37b7d-120">The following sample shows how to create sales records.</span></span> <span data-ttu-id="37b7d-121">Esta secuencia de `Range` comandos utiliza objetos para establecer los valores, las fórmulas y los formatos.</span><span class="sxs-lookup"><span data-stu-id="37b7d-121">This script uses `Range` objects to set the values, formulas, and formats.</span></span>

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

<span data-ttu-id="37b7d-122">Al ejecutar este script, se crean los siguientes datos en la hoja de cálculo actual:</span><span class="sxs-lookup"><span data-stu-id="37b7d-122">Running this script creates the following data in the current worksheet:</span></span>

![Un registro de ventas que muestra las filas de valores, una columna de fórmula y los encabezados con formato.](../images/range-sample.png)

### <a name="charts-tables-and-other-data-objects"></a><span data-ttu-id="37b7d-124">Gráficos, tablas y otros objetos de datos</span><span class="sxs-lookup"><span data-stu-id="37b7d-124">Charts, tables, and other data objects</span></span>

<span data-ttu-id="37b7d-125">Los scripts pueden crear y manipular las estructuras de datos y las visualizaciones en Excel.</span><span class="sxs-lookup"><span data-stu-id="37b7d-125">Scripts can create and manipulate the data structures and visualizations within Excel.</span></span> <span data-ttu-id="37b7d-126">Las tablas y los gráficos son dos de los objetos que se usan con más frecuencia, pero las API admiten tablas dinámicas, formas, imágenes, etc.</span><span class="sxs-lookup"><span data-stu-id="37b7d-126">Tables and charts are two of the more commonly used objects, but the APIs support PivotTables, shapes, images, and more.</span></span>

#### <a name="creating-a-table"></a><span data-ttu-id="37b7d-127">Creación de una tabla</span><span class="sxs-lookup"><span data-stu-id="37b7d-127">Creating a table</span></span>

<span data-ttu-id="37b7d-128">Cree tablas con rangos de datos rellenos.</span><span class="sxs-lookup"><span data-stu-id="37b7d-128">Create tables by using data-filled ranges.</span></span> <span data-ttu-id="37b7d-129">Los controles de formato y de tabla (como los filtros) se aplican automáticamente al rango.</span><span class="sxs-lookup"><span data-stu-id="37b7d-129">Formatting and table controls (such as filters) are automatically applied to the range.</span></span>

<span data-ttu-id="37b7d-130">La siguiente secuencia de comandos crea una tabla con los intervalos del ejemplo anterior.</span><span class="sxs-lookup"><span data-stu-id="37b7d-130">The following script creates a table using the ranges from the previous sample.</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
   let sheet = context.workbook.worksheets.getActiveWorksheet();
   sheet.tables.add("B2:E5", true);
}
```

<span data-ttu-id="37b7d-131">Al ejecutar este script en la hoja de cálculo con los datos anteriores, se crea la siguiente tabla:</span><span class="sxs-lookup"><span data-stu-id="37b7d-131">Running this script on the worksheet with the previous data creates the following table:</span></span>

![Una tabla realizada a partir del registro de ventas anterior.](../images/table-sample.png)

#### <a name="creating-a-chart"></a><span data-ttu-id="37b7d-133">Crear un gráfico</span><span class="sxs-lookup"><span data-stu-id="37b7d-133">Creating a chart</span></span>

<span data-ttu-id="37b7d-134">Cree gráficos para visualizar los datos de un rango.</span><span class="sxs-lookup"><span data-stu-id="37b7d-134">Create charts to visualize the data in a range.</span></span> <span data-ttu-id="37b7d-135">Los scripts permiten docenas de variedades de gráficos, cada una de las cuales puede personalizarse según sus necesidades.</span><span class="sxs-lookup"><span data-stu-id="37b7d-135">Scripts allow for dozens of chart varieties, each of which can be customized to suit your needs.</span></span>

<span data-ttu-id="37b7d-136">El script siguiente crea un gráfico de columnas simple para tres elementos y coloca 100 píxeles por debajo de la parte superior de la hoja de cálculo.</span><span class="sxs-lookup"><span data-stu-id="37b7d-136">The following script creates a simple column chart for three items and places it 100 pixels below the top of the worksheet.</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
  let sheet = context.workbook.worksheets.getActiveWorksheet();
  let chart = sheet.charts.add(Excel.ChartType.columnStacked, sheet.getRange("B3:C5"));
  chart.top = 100;
}
```

<span data-ttu-id="37b7d-137">Al ejecutar este script en la hoja de cálculo con la tabla anterior, se crea el siguiente gráfico:</span><span class="sxs-lookup"><span data-stu-id="37b7d-137">Running this script on the worksheet with the previous table creates the following chart:</span></span>

![Un gráfico de columnas que muestra las cantidades de tres elementos del registro de ventas anterior.](../images/chart-sample.png)

### <a name="further-reading-on-the-object-model"></a><span data-ttu-id="37b7d-139">Lecturas adicionales sobre el modelo de objetos</span><span class="sxs-lookup"><span data-stu-id="37b7d-139">Further reading on the object model</span></span>

<span data-ttu-id="37b7d-140">La [documentación de referencia de la API de scripts de Office](/javascript/api/office-scripts/overview) es una lista completa de los objetos que se usan en scripts de Office.</span><span class="sxs-lookup"><span data-stu-id="37b7d-140">The [Office Scripts API reference documentation](/javascript/api/office-scripts/overview) is a comprehensive listing of the objects used in Office Scripts.</span></span> <span data-ttu-id="37b7d-141">Allí puede usar la tabla de contenido para navegar a cualquier clase de la que quiera obtener más información.</span><span class="sxs-lookup"><span data-stu-id="37b7d-141">There, you can use the table of contents to navigate to any class you'd like to learn more about.</span></span> <span data-ttu-id="37b7d-142">A continuación se muestran varias páginas que se ven normalmente.</span><span class="sxs-lookup"><span data-stu-id="37b7d-142">The following are several commonly viewed pages.</span></span>

- [<span data-ttu-id="37b7d-143">Chart</span><span class="sxs-lookup"><span data-stu-id="37b7d-143">Chart</span></span>](/javascript/api/office-scripts/excel/excel.chart)
- [<span data-ttu-id="37b7d-144">Comment</span><span class="sxs-lookup"><span data-stu-id="37b7d-144">Comment</span></span>](/javascript/api/office-scripts/excel/excel.comment)
- [<span data-ttu-id="37b7d-145">PivotTable</span><span class="sxs-lookup"><span data-stu-id="37b7d-145">PivotTable</span></span>](/javascript/api/office-scripts/excel/excel.pivottable)
- [<span data-ttu-id="37b7d-146">Range</span><span class="sxs-lookup"><span data-stu-id="37b7d-146">Range</span></span>](/javascript/api/office-scripts/excel/excel.range)
- [<span data-ttu-id="37b7d-147">RangeFormat</span><span class="sxs-lookup"><span data-stu-id="37b7d-147">RangeFormat</span></span>](/javascript/api/office-scripts/excel/excel.rangeformat)
- [<span data-ttu-id="37b7d-148">Shape</span><span class="sxs-lookup"><span data-stu-id="37b7d-148">Shape</span></span>](/javascript/api/office-scripts/excel/excel.shape)
- [<span data-ttu-id="37b7d-149">Table</span><span class="sxs-lookup"><span data-stu-id="37b7d-149">Table</span></span>](/javascript/api/office-scripts/excel/excel.table)
- [<span data-ttu-id="37b7d-150">Workbook</span><span class="sxs-lookup"><span data-stu-id="37b7d-150">Workbook</span></span>](/javascript/api/office-scripts/excel/excel.workbook)
- [<span data-ttu-id="37b7d-151">Worksheet</span><span class="sxs-lookup"><span data-stu-id="37b7d-151">Worksheet</span></span>](/javascript/api/office-scripts/excel/excel.worksheet)

## <a name="main-function"></a><span data-ttu-id="37b7d-152">`main`función</span><span class="sxs-lookup"><span data-stu-id="37b7d-152">`main` function</span></span>

<span data-ttu-id="37b7d-153">Cada script de Office debe contener `main` una función con la siguiente firma, incluida `Excel.RequestContext` la definición de tipo:</span><span class="sxs-lookup"><span data-stu-id="37b7d-153">Every Office Script must contain a `main` function with the following signature, including the `Excel.RequestContext` type definition:</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
    // Your Excel Script
}
```

<span data-ttu-id="37b7d-154">El código dentro de `main` la función se ejecuta cuando se ejecuta el script.</span><span class="sxs-lookup"><span data-stu-id="37b7d-154">The code inside the `main` function runs when the script is run.</span></span> <span data-ttu-id="37b7d-155">`main`puede llamar a otras funciones en el script, pero no se ejecutará el código que no está contenido en una función.</span><span class="sxs-lookup"><span data-stu-id="37b7d-155">`main` can call other functions in your script, but code that's not contained in a function will not run.</span></span>

## <a name="context"></a><span data-ttu-id="37b7d-156">Contexto</span><span class="sxs-lookup"><span data-stu-id="37b7d-156">Context</span></span>

<span data-ttu-id="37b7d-157">La `main` función acepta un `Excel.RequestContext` parámetro, denominado `context`.</span><span class="sxs-lookup"><span data-stu-id="37b7d-157">The `main` function accepts an `Excel.RequestContext` parameter, named `context`.</span></span> <span data-ttu-id="37b7d-158">Piense `context` en el puente entre el script y el libro.</span><span class="sxs-lookup"><span data-stu-id="37b7d-158">Think of `context` as the bridge between your script and the workbook.</span></span> <span data-ttu-id="37b7d-159">El script tiene acceso al libro con el `context` objeto y lo usa `context` para enviar datos de una a otra.</span><span class="sxs-lookup"><span data-stu-id="37b7d-159">Your script accesses the workbook with the `context` object and uses that `context` to send data back and forth.</span></span>

<span data-ttu-id="37b7d-160">El `context` objeto es necesario porque el script y Excel se están ejecutando en diferentes procesos y ubicaciones.</span><span class="sxs-lookup"><span data-stu-id="37b7d-160">The `context` object is necessary because the script and Excel are running in different processes and locations.</span></span> <span data-ttu-id="37b7d-161">El script tendrá que realizar cambios o consultar datos del libro en la nube.</span><span class="sxs-lookup"><span data-stu-id="37b7d-161">The script will need to make changes to or query data from the workbook in the cloud.</span></span> <span data-ttu-id="37b7d-162">El `context` objeto administra esas transacciones.</span><span class="sxs-lookup"><span data-stu-id="37b7d-162">The `context` object manages those transactions.</span></span>

## <a name="sync-and-load"></a><span data-ttu-id="37b7d-163">Sincronización y carga</span><span class="sxs-lookup"><span data-stu-id="37b7d-163">Sync and Load</span></span>

<span data-ttu-id="37b7d-164">Dado que el script y el libro se ejecutan en diferentes ubicaciones, cualquier transferencia de datos entre los dos lleva tiempo.</span><span class="sxs-lookup"><span data-stu-id="37b7d-164">Because your script and workbook run in different locations, any data transfer between the two takes time.</span></span> <span data-ttu-id="37b7d-165">Para mejorar el rendimiento del script, los comandos se ponen en cola hasta que el `sync` script llame explícitamente a la operación para sincronizar el script y el libro.</span><span class="sxs-lookup"><span data-stu-id="37b7d-165">To improve script performance, commands are queued up until the script explicitly calls the `sync` operation to synchronize the script and workbook.</span></span> <span data-ttu-id="37b7d-166">La secuencia de comandos puede funcionar de forma independiente hasta que tenga que realizar cualquiera de las siguientes acciones:</span><span class="sxs-lookup"><span data-stu-id="37b7d-166">Your script can work independently until it needs to do either of the following:</span></span>

- <span data-ttu-id="37b7d-167">Leer datos del libro (siguiendo una `load` operación).</span><span class="sxs-lookup"><span data-stu-id="37b7d-167">Read data from the workbook (following a `load` operation).</span></span>
- <span data-ttu-id="37b7d-168">Escribir datos en el libro (normalmente, porque ha finalizado el script).</span><span class="sxs-lookup"><span data-stu-id="37b7d-168">Write data to the workbook (usually because the script has finished).</span></span>

<span data-ttu-id="37b7d-169">La imagen siguiente muestra un ejemplo de flujo de control entre el script y el libro:</span><span class="sxs-lookup"><span data-stu-id="37b7d-169">The following image shows an example control flow between the script and workbook:</span></span>

![Un diagrama que muestra las operaciones de lectura y escritura que se dirigen al libro desde el script.](../images/load-sync.png)

### <a name="sync"></a><span data-ttu-id="37b7d-171">Sincronizar</span><span class="sxs-lookup"><span data-stu-id="37b7d-171">Sync</span></span>

<span data-ttu-id="37b7d-172">Siempre que el script tenga que leer o escribir datos en el libro, llame al `RequestContext.sync` método como se muestra a continuación:</span><span class="sxs-lookup"><span data-stu-id="37b7d-172">Whenever your script needs to read data from or write data to the workbook, call the `RequestContext.sync` method as shown here:</span></span>

```TypeScript
await context.sync();
```

> [!NOTE]
> <span data-ttu-id="37b7d-173">`context.sync()`se llama implícitamente cuando finaliza un script.</span><span class="sxs-lookup"><span data-stu-id="37b7d-173">`context.sync()` is implicitly called when a script ends.</span></span>

<span data-ttu-id="37b7d-174">Una vez `sync` finalizada la operación, el libro se actualiza para reflejar las operaciones de escritura que haya especificado el script.</span><span class="sxs-lookup"><span data-stu-id="37b7d-174">After the `sync` operation completes, the workbook updates to reflect any write operations that script has specified.</span></span> <span data-ttu-id="37b7d-175">Una operación de escritura está estableciendo cualquier propiedad en un objeto de Excel ( `range.format.fill.color = "red"`por ejemplo,) o llamando a un método que cambia una propiedad `range.format.autoFitColumns()`(por ejemplo,).</span><span class="sxs-lookup"><span data-stu-id="37b7d-175">A write operation is setting any property on a Excel object (e.g. `range.format.fill.color = "red"`) or calling a method that changes a property (e.g., `range.format.autoFitColumns()`).</span></span> <span data-ttu-id="37b7d-176">La `sync` operación también lee los valores del libro que ha solicitado el script mediante una `load` operación (como se describe en la siguiente sección).</span><span class="sxs-lookup"><span data-stu-id="37b7d-176">The `sync` operation also reads any values from the workbook that the script requested by using a `load` operation (as discussed in the next section).</span></span>

<span data-ttu-id="37b7d-177">La sincronización del script con el libro puede tardar un tiempo, en función de la red.</span><span class="sxs-lookup"><span data-stu-id="37b7d-177">Synchronizing your script with the workbook can take time, depending on your network.</span></span> <span data-ttu-id="37b7d-178">Debe minimizar el número de llamadas `sync` para que la secuencia de comandos pueda ejecutarse con rapidez.</span><span class="sxs-lookup"><span data-stu-id="37b7d-178">You should minimize the number of `sync` calls to help your script run fast.</span></span>  

### <a name="load"></a><span data-ttu-id="37b7d-179">Volver</span><span class="sxs-lookup"><span data-stu-id="37b7d-179">Load</span></span>

<span data-ttu-id="37b7d-180">Un script debe cargar datos del libro antes de leerlo.</span><span class="sxs-lookup"><span data-stu-id="37b7d-180">A script must load data from the workbook before reading it.</span></span> <span data-ttu-id="37b7d-181">Sin embargo, la carga de datos de todo el libro con frecuencia reducirá en gran medida la velocidad del script.</span><span class="sxs-lookup"><span data-stu-id="37b7d-181">However, frequently loading data from the entire workbook would greatly reduce the script's speed.</span></span> <span data-ttu-id="37b7d-182">En su lugar, `load` el método permite al script indicar específicamente qué datos deben recuperarse del libro.</span><span class="sxs-lookup"><span data-stu-id="37b7d-182">Instead, the `load` method lets your script state specifically which data should be retrieved from the workbook.</span></span>

<span data-ttu-id="37b7d-183">El `load` método está disponible en todos los objetos de Excel.</span><span class="sxs-lookup"><span data-stu-id="37b7d-183">The `load` method is available on every Excel object.</span></span> <span data-ttu-id="37b7d-184">El script debe cargar las propiedades de un objeto antes de que pueda leerlas.</span><span class="sxs-lookup"><span data-stu-id="37b7d-184">Your script must load an object's properties before it can read them.</span></span> <span data-ttu-id="37b7d-185">Si no lo hace, se producirá un error.</span><span class="sxs-lookup"><span data-stu-id="37b7d-185">Not doing so will result in an error.</span></span>

<span data-ttu-id="37b7d-186">En los ejemplos siguientes se `Range` usa un objeto para mostrar las tres `load` formas en que se puede usar el método para cargar datos.</span><span class="sxs-lookup"><span data-stu-id="37b7d-186">The following examples use a `Range` object to show the three ways the `load` method can be used to load data.</span></span>

|<span data-ttu-id="37b7d-187">Intent</span><span class="sxs-lookup"><span data-stu-id="37b7d-187">Intent</span></span> |<span data-ttu-id="37b7d-188">Comando de ejemplo</span><span class="sxs-lookup"><span data-stu-id="37b7d-188">Example Command</span></span> | <span data-ttu-id="37b7d-189">Efecto</span><span class="sxs-lookup"><span data-stu-id="37b7d-189">Effect</span></span> |
|:--|:--|:--|
|<span data-ttu-id="37b7d-190">Cargar una propiedad</span><span class="sxs-lookup"><span data-stu-id="37b7d-190">Load one property</span></span> |`myRange.load("values");` | <span data-ttu-id="37b7d-191">Carga una única propiedad, en este caso la matriz bidimensional de valores de este intervalo.</span><span class="sxs-lookup"><span data-stu-id="37b7d-191">Loads a single property, in this case the two-dimensional array of values in this range.</span></span> |
|<span data-ttu-id="37b7d-192">Cargar varias propiedades</span><span class="sxs-lookup"><span data-stu-id="37b7d-192">Load multiple properties</span></span> |`myRange.load("values, rowCount, columnCount");`| <span data-ttu-id="37b7d-193">Carga todas las propiedades de una lista delimitada por comas, en este ejemplo los valores, el recuento de filas y el número de columnas.</span><span class="sxs-lookup"><span data-stu-id="37b7d-193">Loads all the properties from a comma-delimited list, in this example the values, row count, and column count.</span></span> |
|<span data-ttu-id="37b7d-194">Carga todo</span><span class="sxs-lookup"><span data-stu-id="37b7d-194">Load everything</span></span> | `myRange.load();`|<span data-ttu-id="37b7d-195">Carga todas las propiedades del rango.</span><span class="sxs-lookup"><span data-stu-id="37b7d-195">Loads all the properties on the range.</span></span> <span data-ttu-id="37b7d-196">Esta solución no es la recomendada, ya que se ralentizará el script al obtener datos innecesarios.</span><span class="sxs-lookup"><span data-stu-id="37b7d-196">This is not a recommended solution, since it will slow down your script by getting unnecessary data.</span></span> <span data-ttu-id="37b7d-197">Solo debe usar esto mientras prueba el script o si necesita todas las propiedades del objeto.</span><span class="sxs-lookup"><span data-stu-id="37b7d-197">You should only use this while testing your script or if you need every property from the object.</span></span> |

<span data-ttu-id="37b7d-198">El script debe llamar `context.sync()` antes de leer los valores cargados.</span><span class="sxs-lookup"><span data-stu-id="37b7d-198">Your script must call `context.sync()` before reading any loaded values.</span></span>

```TypeScript
let range = selectedSheet.getRange("A1:B3");
range.load ("rowCount"); // Load the property.
await context.sync(); // Synchronize with the workbook to get the property.
console.log(range.rowCount); // Read and log the property value (3).
```

<span data-ttu-id="37b7d-199">También puede cargar propiedades en toda una colección.</span><span class="sxs-lookup"><span data-stu-id="37b7d-199">You can also load properties across an entire collection.</span></span> <span data-ttu-id="37b7d-200">Cada objeto Collection tiene una `items` propiedad que es una matriz que contiene los objetos de esa colección.</span><span class="sxs-lookup"><span data-stu-id="37b7d-200">Every collection object has an `items` property that is an array containing the objects in that collection.</span></span> <span data-ttu-id="37b7d-201">Usar `items` como inicio de una llamada jerárquica (`items\myProperty`) para `load` cargar las propiedades especificadas en cada uno de estos elementos.</span><span class="sxs-lookup"><span data-stu-id="37b7d-201">Using `items` as the start of a hierarchical call (`items\myProperty`) to `load` loads the specified properties on each of those items.</span></span> <span data-ttu-id="37b7d-202">En el siguiente ejemplo, `resolved` se carga la `Comment` propiedad en cada `CommentCollection` objeto del objeto de una hoja de cálculo.</span><span class="sxs-lookup"><span data-stu-id="37b7d-202">The following example loads the `resolved` property on every `Comment` object in the `CommentCollection` object of a worksheet.</span></span>

```TypeScript
let comments = selectedSheet.comments;
comments.load("items/resolved"); // Load the `resolved` property from every comment in this collection.
await context.sync(); // Synchronize with the workbook to get the properties.
```

> [!TIP]
> <span data-ttu-id="37b7d-203">Para obtener más información sobre cómo trabajar con colecciones en scripts de Office, vea la [sección matriz del artículo usar objetos de JavaScript integrados en las secuencias de comandos de Office](javascript-objects.md#array) .</span><span class="sxs-lookup"><span data-stu-id="37b7d-203">To learn more about working with collections in Office Scripts, see the [Array section of the Using built-in JavaScript objects in Office Scripts](javascript-objects.md#array) article.</span></span>

## <a name="see-also"></a><span data-ttu-id="37b7d-204">Vea también</span><span class="sxs-lookup"><span data-stu-id="37b7d-204">See also</span></span>

- [<span data-ttu-id="37b7d-205">Grabar, editar y crear scripts de Office en Excel en la web</span><span class="sxs-lookup"><span data-stu-id="37b7d-205">Record, edit, and create Office Scripts in Excel on the web</span></span>](../tutorials/excel-tutorial.md)
- [<span data-ttu-id="37b7d-206">Leer datos de un libro con scripts de Office en Excel en la web</span><span class="sxs-lookup"><span data-stu-id="37b7d-206">Read workbook data with Office Scripts in Excel on the web</span></span>](../tutorials/excel-read-tutorial.md)
- [<span data-ttu-id="37b7d-207">Referencia de la API de scripts de Office</span><span class="sxs-lookup"><span data-stu-id="37b7d-207">Office Scripts API reference</span></span>](/javascript/api/office-scripts/overview)
- [<span data-ttu-id="37b7d-208">Uso de objetos de JavaScript integrados en scripts de Office</span><span class="sxs-lookup"><span data-stu-id="37b7d-208">Using built-in JavaScript objects in Office Scripts</span></span>](javascript-objects.md)
