---
title: Scripts de ejemplo para scripts de Office en Excel en la web
description: Una colección de ejemplos de código para usar con scripts de Office en Excel en la Web.
ms.date: 08/04/2020
localization_priority: Normal
ms.openlocfilehash: 4f8d6f2395a841a8dcba2ea0e712e645a84a6d91
ms.sourcegitcommit: 1c88abcf5df16a05913f12df89490ce843cfebe2
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 08/13/2020
ms.locfileid: "46665232"
---
# <a name="sample-scripts-for-office-scripts-in-excel-on-the-web-preview"></a><span data-ttu-id="65daf-103">Scripts de ejemplo para scripts de Office en Excel en la web (vista previa)</span><span class="sxs-lookup"><span data-stu-id="65daf-103">Sample scripts for Office Scripts in Excel on the web (preview)</span></span>

<span data-ttu-id="65daf-104">Los siguientes ejemplos son scripts sencillos que puede probar en sus propios libros.</span><span class="sxs-lookup"><span data-stu-id="65daf-104">The following samples are simple scripts for you to try on your own workbooks.</span></span> <span data-ttu-id="65daf-105">Para usarlas en Excel en la web:</span><span class="sxs-lookup"><span data-stu-id="65daf-105">To use them in Excel on the web:</span></span>

1. <span data-ttu-id="65daf-106">Abra la pestaña **Automatizar**.</span><span class="sxs-lookup"><span data-stu-id="65daf-106">Open the **Automate** tab.</span></span>
2. <span data-ttu-id="65daf-107">Presione el **Editor de código**.</span><span class="sxs-lookup"><span data-stu-id="65daf-107">Press **Code Editor**.</span></span>
3. <span data-ttu-id="65daf-108">Presione **nueva secuencia de comandos** en el panel de tareas del editor de código.</span><span class="sxs-lookup"><span data-stu-id="65daf-108">Press **New Script** in the Code Editor's task pane.</span></span>
4. <span data-ttu-id="65daf-109">Reemplace todo el script por el ejemplo de su elección.</span><span class="sxs-lookup"><span data-stu-id="65daf-109">Replace the entire script with the sample of your choice.</span></span>
5. <span data-ttu-id="65daf-110">Presione **Ejecutar** en el panel de tareas del editor de código.</span><span class="sxs-lookup"><span data-stu-id="65daf-110">Press **Run** in the Code Editor's task pane.</span></span>

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="scripting-basics"></a><span data-ttu-id="65daf-111">Conceptos básicos de scripting</span><span class="sxs-lookup"><span data-stu-id="65daf-111">Scripting basics</span></span>

<span data-ttu-id="65daf-112">Estos ejemplos muestran bloques de creación fundamentales para los scripts de Office.</span><span class="sxs-lookup"><span data-stu-id="65daf-112">These samples demonstrate fundamental building blocks for Office Scripts.</span></span> <span data-ttu-id="65daf-113">Agréguelos a los scripts para ampliar la solución y resolver problemas comunes.</span><span class="sxs-lookup"><span data-stu-id="65daf-113">Add these to your scripts to extend your solution and solve common problems.</span></span>

### <a name="read-and-log-one-cell"></a><span data-ttu-id="65daf-114">Leer e iniciar sesión en una celda</span><span class="sxs-lookup"><span data-stu-id="65daf-114">Read and log one cell</span></span>

<span data-ttu-id="65daf-115">En este ejemplo se lee el valor de **a1** y se imprime en la consola.</span><span class="sxs-lookup"><span data-stu-id="65daf-115">This sample reads the value of **A1** and prints it to the console.</span></span>

```typescript
function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  let selectedSheet = workbook.getActiveWorksheet();

  // Get the value of cell A1.
  let range = selectedSheet.getRange("A1");
  
  // Print the value of A1.
  console.log(range.getValue());
}
```

### <a name="read-the-active-cell"></a><span data-ttu-id="65daf-116">Leer la celda activa</span><span class="sxs-lookup"><span data-stu-id="65daf-116">Read the active cell</span></span>

<span data-ttu-id="65daf-117">Esta secuencia de comandos registra el valor de la celda activa actual.</span><span class="sxs-lookup"><span data-stu-id="65daf-117">This script logs the value of the current active cell.</span></span> <span data-ttu-id="65daf-118">Si se seleccionan varias celdas, se registrará la celda superior izquierda.</span><span class="sxs-lookup"><span data-stu-id="65daf-118">If multiple cells are selected, the top-leftmost cell will be logged.</span></span>

```typescript
function main(workbook: ExcelScript.Workbook) {
  // Get the current active cell in the workbook.
  let cell = workbook.getActiveCell();

  // Log that cell's value.
  console.log(`The current cell's value is ${cell.getValue()}`);
}
```

### <a name="change-an-adjacent-cell"></a><span data-ttu-id="65daf-119">Cambiar una celda adyacente</span><span class="sxs-lookup"><span data-stu-id="65daf-119">Change an adjacent cell</span></span>

<span data-ttu-id="65daf-120">Este script obtiene las celdas adyacentes mediante referencias relativas.</span><span class="sxs-lookup"><span data-stu-id="65daf-120">This script gets adjacent cells using relative references.</span></span> <span data-ttu-id="65daf-121">Tenga en cuenta que si la celda activa está en la fila superior, se producirá un error en parte del script porque hace referencia a la celda por encima de la selección actualmente seleccionada.</span><span class="sxs-lookup"><span data-stu-id="65daf-121">Note that if the active cell is on the top row, part of the script fails, because it references the cell above the currently selected one.</span></span>

```typescript
function main(workbook: ExcelScript.Workbook) {
  // Get the currently active cell in the workbook.
  let activeCell = workbook.getActiveCell();
  console.log(`The active cell's address is: ${activeCell.getAddress()}`);

  // Get the cell to the right of the active cell and set its value and color.
  let rightCell = activeCell.getOffsetRange(0,1);
  rightCell.setValue("Right cell");
  console.log(`The right cell's address is: ${rightCell.getAddress()}`);
  rightCell.getFormat().getFont().setColor("Magenta");
  rightCell.getFormat().getFill().setColor("Cyan");

  // Get the cell to the above of the active cell and set its value and color.
  // Note that this operation will fail if the active cell is in the top row.
  let aboveCell = activeCell.getOffsetRange(-1, 0);
  aboveCell.setValue("Above cell");
  console.log(`The above cell's address is: ${aboveCell.getAddress()}`);
  aboveCell.getFormat().getFont().setColor("White");
  aboveCell.getFormat().getFill().setColor("Black");
}
```

### <a name="change-all-adjacent-cells"></a><span data-ttu-id="65daf-122">Cambiar todas las celdas adyacentes</span><span class="sxs-lookup"><span data-stu-id="65daf-122">Change all adjacent cells</span></span>

<span data-ttu-id="65daf-123">Este script copia el formato de la celda activa en las celdas vecinas.</span><span class="sxs-lookup"><span data-stu-id="65daf-123">This script copies the formatting in the active cell to the neighboring cells.</span></span> <span data-ttu-id="65daf-124">Tenga en cuenta que este script sólo funciona cuando la celda activa no está en un borde de la hoja de cálculo.</span><span class="sxs-lookup"><span data-stu-id="65daf-124">Note that this script only works when the active cell isn't on an edge of the worksheet.</span></span>

```typescript
function main(workbook: ExcelScript.Workbook) {
  // Get the active cell.
  let activeCell = workbook.getActiveCell();

  // Get the cell that's one row above and one column to the left of the active cell.
  let cornerCell = activeCell.getOffsetRange(-1,-1);

  // Get a range that includes all the cells surrounding the active cell.
  let surroundingRange = cornerCell.getResizedRange(2, 2)

  // Copy the formatting from the active cell to the new range.
  surroundingRange.copyFrom(
    activeCell, /* The source range. */
    ExcelScript.RangeCopyType.formats /* What to copy. */
    );
}
```

### <a name="change-each-individual-cell-in-a-range"></a><span data-ttu-id="65daf-125">Cambiar cada celda individual de un rango</span><span class="sxs-lookup"><span data-stu-id="65daf-125">Change each individual cell in a range</span></span>

<span data-ttu-id="65daf-126">Este script recorre el rango seleccionado actualmente.</span><span class="sxs-lookup"><span data-stu-id="65daf-126">This script loops over the currently select range.</span></span> <span data-ttu-id="65daf-127">Borra el formato actual y establece el color de relleno de cada celda en un color aleatorio.</span><span class="sxs-lookup"><span data-stu-id="65daf-127">It clears the current formatting and sets the fill color in each cell to a random color.</span></span>

```typescript
function main(workbook: ExcelScript.Workbook) {
  // Get the currently selected range.
  let range = workbook.getSelectedRange();

  // Get the size boundaries of the range.
  let rows = range.getRowCount();
  let cols = range.getColumnCount();

  // Clear any existing formatting
  range.clear(ExcelScript.ClearApplyTo.formats);

  // Iterate over the range.
  for (let row = 0; row < rows; row++) {
    for (let col = 0; col < cols; col++) {
      // Generate a random color hex-code.
      let colorString = `#${Math.random().toString(16).substr(-6)}`;

      // Set the color of the current cell to that random hex-code.
      range.getCell(row, col).getFormat().getFill().setColor(colorString);
    }
  }
}
```

## <a name="collections"></a><span data-ttu-id="65daf-128">Colecciones</span><span class="sxs-lookup"><span data-stu-id="65daf-128">Collections</span></span>

<span data-ttu-id="65daf-129">Estos ejemplos funcionan con colecciones de objetos en el libro.</span><span class="sxs-lookup"><span data-stu-id="65daf-129">These samples work with collections of objects in the workbook.</span></span>

### <a name="iterating-over-collections"></a><span data-ttu-id="65daf-130">Recorrer en iteración las colecciones</span><span class="sxs-lookup"><span data-stu-id="65daf-130">Iterating over collections</span></span>

<span data-ttu-id="65daf-131">Este script obtiene y registra los nombres de todas las hojas de cálculo del libro.</span><span class="sxs-lookup"><span data-stu-id="65daf-131">This script gets and logs the names of all the worksheets in the workbook.</span></span> <span data-ttu-id="65daf-132">También establece los colores de las pestañas en un color aleatorio.</span><span class="sxs-lookup"><span data-stu-id="65daf-132">It also sets the their tab colors to a random color.</span></span>

```typescript
function main(workbook: ExcelScript.Workbook) {
  // Get all the worksheets in the workbook.
  let sheets = workbook.getWorksheets();

  // Get a list of all the worksheet names.
  let names = sheets.map ((sheet) => sheet.getName());

  // Write in the console all the worksheet names and the total count.
  console.log(names);
  console.log(`Total worksheets inside of this workbook: ${sheets.length}`);
  
  // Set the tab color each worksheet to a random color
  for (let sheet of sheets) {
    // Generate a random color hex-code.
    let colorString = `#${Math.random().toString(16).substr(-6)}`;

    // Set the color of the current worksheet's tab to that random hex-code.
    sheet.setTabColor(colorString);
  }
}
```

### <a name="querying-and-deleting-from-a-collection"></a><span data-ttu-id="65daf-133">Consulta y eliminación de una colección</span><span class="sxs-lookup"><span data-stu-id="65daf-133">Querying and deleting from a collection</span></span>

<span data-ttu-id="65daf-134">Este script crea una hoja de cálculo nueva.</span><span class="sxs-lookup"><span data-stu-id="65daf-134">This script creates a new worksheet.</span></span> <span data-ttu-id="65daf-135">Busca una copia existente de la hoja de cálculo y la elimina antes de crear una hoja nueva.</span><span class="sxs-lookup"><span data-stu-id="65daf-135">It checks for an existing copy of the worksheet and deletes it before making a new sheet.</span></span>

```typescript
function main(workbook: ExcelScript.Workbook) {
  // Name of the worksheet to be added.
  let name = "Index";

  // Get any worksheet with that name.
  let sheet = workbook.getWorksheet("Index");
  
  // If `null` wasn't returned, then there's already a worksheet with that name.
  if (sheet) {
    console.log(`Worksheet by the name ${name} already exists. Deleting it.`);
    // Delete the sheet.
    sheet.delete();
  }
  
  // Add a blank worksheet with the name "Index".
  // Note that this code runs regardless of whether an existing sheet was deleted.
  console.log(`Adding the worksheet named ${name}.`);
  let newSheet = workbook.addWorksheet("Index");

  // Switch to the new worksheet.
  newSheet.activate();
}
```

## <a name="dates"></a><span data-ttu-id="65daf-136">Fechas</span><span class="sxs-lookup"><span data-stu-id="65daf-136">Dates</span></span>

<span data-ttu-id="65daf-137">Los ejemplos de esta sección muestran cómo usar el objeto [Date](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date) de JavaScript.</span><span class="sxs-lookup"><span data-stu-id="65daf-137">The samples in this section show how to use the JavaScript [Date](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date) object.</span></span>

<span data-ttu-id="65daf-138">En el ejemplo siguiente se obtiene la fecha y hora actuales y, a continuación, se escriben los valores en dos celdas de la hoja de cálculo activa.</span><span class="sxs-lookup"><span data-stu-id="65daf-138">The following sample gets the current date and time and then writes those values to two cells in the active worksheet.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the cells at A1 and B1.
  let dateRange = workbook.getActiveWorksheet().getRange("A1");
  let timeRange = workbook.getActiveWorksheet().getRange("B1");

  // Get the current date and time with the JavaScript Date object.
  let date = new Date(Date.now());

  // Add the date string to A1.
  dateRange.setValue(date.toLocaleDateString());

  // Add the time string to B1.
  timeRange.setValue(date.toLocaleTimeString());
}
```

<span data-ttu-id="65daf-139">El siguiente ejemplo lee una fecha que está almacenada en Excel y la convierte en un objeto Date de JavaScript.</span><span class="sxs-lookup"><span data-stu-id="65daf-139">The next sample reads a date that's stored in Excel and translates it to a JavaScript Date object.</span></span> <span data-ttu-id="65daf-140">Usa el [número de serie numérico de la fecha](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) como entrada para la fecha de JavaScript.</span><span class="sxs-lookup"><span data-stu-id="65daf-140">It uses the [date's numeric serial number](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) as input for the JavaScript Date.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Read a date at cell A1 from Excel.
  let dateRange = workbook.getActiveWorksheet().getRange("A1");

  // Convert the Excel date to a JavaScript Date object.
  let excelDateValue = dateRange.getValue();
  let javaScriptDate = new Date(Math.round((excelDateValue - 25569) * 86400 * 1000));
  console.log(javaScriptDate);
}
```

## <a name="display-data"></a><span data-ttu-id="65daf-141">Mostrar datos</span><span class="sxs-lookup"><span data-stu-id="65daf-141">Display data</span></span>

<span data-ttu-id="65daf-142">En estos ejemplos se muestra cómo trabajar con los datos de la hoja de cálculo y proporcionar a los usuarios una vista o organización mejor.</span><span class="sxs-lookup"><span data-stu-id="65daf-142">These samples demonstrate how to work with worksheet data and provide users with a better view or organization.</span></span>

### <a name="apply-conditional-formatting"></a><span data-ttu-id="65daf-143">Aplicar formato condicional</span><span class="sxs-lookup"><span data-stu-id="65daf-143">Apply conditional formatting</span></span>

<span data-ttu-id="65daf-144">En este ejemplo se aplica formato condicional al intervalo que se usa actualmente en la hoja de cálculo.</span><span class="sxs-lookup"><span data-stu-id="65daf-144">This sample applies conditional formatting to the currently used range in the worksheet.</span></span> <span data-ttu-id="65daf-145">El formato condicional es un relleno verde para el 10% de los valores principales.</span><span class="sxs-lookup"><span data-stu-id="65daf-145">The conditional formatting is a green fill for the top 10% of values.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  let selectedSheet = workbook.getActiveWorksheet();

  // Get the used range in the worksheet.
  let range = selectedSheet.getUsedRange();

  // Set the fill color to green for the top 10% of values in the range.
  let conditionalFormat = range.addConditionalFormat(ExcelScript.ConditionalFormatType.topBottom)
  conditionalFormat.getTopBottom().getFormat().getFill().setColor("green");
  conditionalFormat.getTopBottom().setRule({
    rank: 10, // The percentage threshold.
    type: ExcelScript.ConditionalTopBottomCriterionType.topPercent // The type of the top/bottom condition.
  });
}
```

### <a name="create-a-sorted-table"></a><span data-ttu-id="65daf-146">Crear una tabla ordenada</span><span class="sxs-lookup"><span data-stu-id="65daf-146">Create a sorted table</span></span>

<span data-ttu-id="65daf-147">En este ejemplo se crea una tabla a partir del rango usado de la hoja de cálculo actual y, a continuación, se ordena basándose en la primera columna.</span><span class="sxs-lookup"><span data-stu-id="65daf-147">This sample creates a table from the current worksheet's used range, then sorts it based on the first column.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  let selectedSheet = workbook.getActiveWorksheet();

  // Create a table with the used cells.
  let usedRange = selectedSheet.getUsedRange();
  let newTable = selectedSheet.addTable(usedRange, true);

  // Sort the table using the first column.
  newTable.getSort().apply([{ key: 0, ascending: true }]);
}
```

### <a name="log-the-grand-total-values-from-a-pivottable"></a><span data-ttu-id="65daf-148">Registrar los valores de "total general" de una tabla dinámica</span><span class="sxs-lookup"><span data-stu-id="65daf-148">Log the "Grand Total" values from a PivotTable</span></span>

<span data-ttu-id="65daf-149">En este ejemplo se busca la primera tabla dinámica del libro y se registran los valores de las celdas de "Grand total" (resaltado en verde en la imagen siguiente).</span><span class="sxs-lookup"><span data-stu-id="65daf-149">This sample finds the first PivotTable in the workbook and logs the values in the "Grand Total" cells (as highlighted in green in the image below).</span></span>

![Una tabla dinámica ventas de frutas con la fila total general resaltada en verde.](../images/sample-pivottable-grand-total-row.png)

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the first PivotTable in the workbook.
  let pivotTable = workbook.getPivotTables()[0];

  // Get the names of each data column in the PivotTable.
  let pivotColumnLabelRange = pivotTable.getLayout().getColumnLabelRange();

  // Get the range displaying the pivoted data.
  let pivotDataRange = pivotTable.getLayout().getRangeBetweenHeaderAndTotal();

  // Get the range with the "grand totals" for the PivotTable columns.
  let grandTotalRange = pivotDataRange.getLastRow();

  // Print each of the "Grand Totals" to the console.
  grandTotalRange.getValues()[0].forEach((column, columnIndex) => {
    console.log(`Grand total of ${pivotColumnLabelRange.getValues()[0][columnIndex]}: ${grandTotalRange.getValues()[0][columnIndex]}`);
    // Example log: "Grand total of Sum of Crates Sold Wholesale: 11000"
  });
}
```

## <a name="formulas"></a><span data-ttu-id="65daf-151">Éstas</span><span class="sxs-lookup"><span data-stu-id="65daf-151">Formulas</span></span>

<span data-ttu-id="65daf-152">Estos ejemplos utilizan fórmulas de Excel y muestran cómo trabajar con ellos en scripts.</span><span class="sxs-lookup"><span data-stu-id="65daf-152">These samples use Excel formulas and show how to work with them in scripts.</span></span>

## <a name="single-formula"></a><span data-ttu-id="65daf-153">Fórmula única</span><span class="sxs-lookup"><span data-stu-id="65daf-153">Single formula</span></span>

<span data-ttu-id="65daf-154">Esta secuencia de comandos establece la fórmula de una celda y, a continuación, muestra cómo Excel almacena por separado la fórmula y el valor de la celda.</span><span class="sxs-lookup"><span data-stu-id="65daf-154">This script sets a cell's formula, then displays how Excel stores the cell's formula and value separately.</span></span>

```typescript
function main(workbook: ExcelScript.Workbook) {
  let selectedSheet = workbook.getActiveWorksheet();

  // Set A1 to 2.
  let a1 = selectedSheet.getRange("A1");
  a1.setValue(2);

  // Set B1 to the formula =(2*A1), which should equal 4.
  let b1 = selectedSheet.getRange("B1")
  b1.setFormula("=(2*A1)");

  // Log the current results for `getFormula` and `getValue` at B1.
  console.log(`B1 - Formula: ${b1.getFormula()} | Value: ${b1.getValue()}`);
}
```

### <a name="spilling-results-from-a-formula"></a><span data-ttu-id="65daf-155">Derrame de resultados a partir de una fórmula</span><span class="sxs-lookup"><span data-stu-id="65daf-155">Spilling results from a formula</span></span>

<span data-ttu-id="65daf-156">Este script transpone el intervalo "a1: D2" en "A4: B7" mediante la función transponer.</span><span class="sxs-lookup"><span data-stu-id="65daf-156">This script transposes the range "A1:D2" to "A4:B7" by using the TRANSPOSE function.</span></span> <span data-ttu-id="65daf-157">Si la TRANSPOSE da como resultado un error de #SPILL, borra el rango de destino y aplica la fórmula de nuevo.</span><span class="sxs-lookup"><span data-stu-id="65daf-157">If the transpose results in a #SPILL error, it clears the target range and applies the formula again.</span></span>

```typescript
function main(workbook: ExcelScript.Workbook) {
  let sheet = workbook.getActiveWorksheet();
  // Use the data in A1:D2 for the sample.
  let dataAddress = "A1:D2"
  let inputRange = sheet.getRange(dataAddress);

  // Place the transposed data starting at A4.
  let targetStartCell = sheet.getRange("A4");

  // Compute the target range.
  let targetRange = targetStartCell.getResizedRange(inputRange.getColumnCount() - 1, inputRange.getRowCount() - 1);

  // Call the transpose helper function.
  targetStartCell.setFormula(`=TRANSPOSE(${dataAddress})`);

  // Check if the range update resulted in a spill error.
  let checkValue = targetStartCell.getValue() as string;
  if (checkValue === '#SPILL!') {
    // Clear the target range and call the transpose function again.
    console.log("Target range has data that is preventing update. Clearing target range.");
    targetRange.clear();
    targetStartCell.setFormula(`=TRANSPOSE(${dataAddress})`);
  }

  // Select the transposed range to highlight it.
  targetRange.select();
}
```

## <a name="scenario-samples"></a><span data-ttu-id="65daf-158">Ejemplos de escenario</span><span class="sxs-lookup"><span data-stu-id="65daf-158">Scenario samples</span></span>

<span data-ttu-id="65daf-159">Para obtener ejemplos que muestren soluciones de gran tamaño para el mundo real, visite ejemplos [de escenarios de Office scripts](scenarios/sample-scenario-overview.md).</span><span class="sxs-lookup"><span data-stu-id="65daf-159">For samples showcasing larger, real-world solutions, visit [Sample scenarios for Office Scripts](scenarios/sample-scenario-overview.md).</span></span>

## <a name="suggest-new-samples"></a><span data-ttu-id="65daf-160">Sugerir nuevos ejemplos</span><span class="sxs-lookup"><span data-stu-id="65daf-160">Suggest new samples</span></span>

<span data-ttu-id="65daf-161">Agradecemos las sugerencias para los nuevos ejemplos.</span><span class="sxs-lookup"><span data-stu-id="65daf-161">We welcome suggestions for new samples.</span></span> <span data-ttu-id="65daf-162">Si hay un escenario común que ayudaría a otros programadores de scripts, indíquenos en la sección Comentarios a continuación.</span><span class="sxs-lookup"><span data-stu-id="65daf-162">If there is a common scenario that would help other script developers, please tell us in the feedback section below.</span></span>
