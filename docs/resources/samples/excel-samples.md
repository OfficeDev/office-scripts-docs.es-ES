---
title: Scripts básicos para Office scripts en Excel en la Web
description: Una colección de ejemplos de código para usar con Office scripts en Excel en la Web.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 3bf3bd5acd10bc5999db4746a2ed62af85237e48
ms.sourcegitcommit: 4693c8f79428ec74695328275703af0ba1bfea8f
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 06/23/2021
ms.locfileid: "53074560"
---
# <a name="basic-scripts-for-office-scripts-in-excel-on-the-web"></a><span data-ttu-id="90067-103">Scripts básicos para Office scripts en Excel en la Web</span><span class="sxs-lookup"><span data-stu-id="90067-103">Basic scripts for Office Scripts in Excel on the web</span></span>

<span data-ttu-id="90067-104">Los ejemplos siguientes son scripts sencillos para que pruebe en sus propios libros.</span><span class="sxs-lookup"><span data-stu-id="90067-104">The following samples are simple scripts for you to try on your own workbooks.</span></span> <span data-ttu-id="90067-105">Para usarlos en Excel en la Web:</span><span class="sxs-lookup"><span data-stu-id="90067-105">To use them in Excel on the web:</span></span>

1. <span data-ttu-id="90067-106">Abra la pestaña **Automatizar**.</span><span class="sxs-lookup"><span data-stu-id="90067-106">Open the **Automate** tab.</span></span>
2. <span data-ttu-id="90067-107">Presione **editor de código**.</span><span class="sxs-lookup"><span data-stu-id="90067-107">Press **Code Editor**.</span></span>
3. <span data-ttu-id="90067-108">Presione **Nuevo script** en el panel de tareas del Editor de código.</span><span class="sxs-lookup"><span data-stu-id="90067-108">Press **New Script** in the Code Editor's task pane.</span></span>
4. <span data-ttu-id="90067-109">Reemplace todo el script por el ejemplo que prefiera.</span><span class="sxs-lookup"><span data-stu-id="90067-109">Replace the entire script with the sample of your choice.</span></span>
5. <span data-ttu-id="90067-110">Presione **Ejecutar** en el panel de tareas del Editor de código.</span><span class="sxs-lookup"><span data-stu-id="90067-110">Press **Run** in the Code Editor's task pane.</span></span>

## <a name="script-basics"></a><span data-ttu-id="90067-111">Conceptos básicos del script</span><span class="sxs-lookup"><span data-stu-id="90067-111">Script basics</span></span>

<span data-ttu-id="90067-112">Estos ejemplos muestran bloques de creación fundamentales para Office scripts.</span><span class="sxs-lookup"><span data-stu-id="90067-112">These samples demonstrate fundamental building blocks for Office Scripts.</span></span> <span data-ttu-id="90067-113">Agrégalos a los scripts para ampliar la solución y resolver problemas comunes.</span><span class="sxs-lookup"><span data-stu-id="90067-113">Add these to your scripts to extend your solution and solve common problems.</span></span>

### <a name="read-and-log-one-cell"></a><span data-ttu-id="90067-114">Leer y registrar una celda</span><span class="sxs-lookup"><span data-stu-id="90067-114">Read and log one cell</span></span>

<span data-ttu-id="90067-115">En este ejemplo se lee el valor **de A1** y se imprime en la consola.</span><span class="sxs-lookup"><span data-stu-id="90067-115">This sample reads the value of **A1** and prints it to the console.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  let selectedSheet = workbook.getActiveWorksheet();

  // Get the value of cell A1.
  let range = selectedSheet.getRange("A1");
  
  // Print the value of A1.
  console.log(range.getValue());
}
```

### <a name="read-the-active-cell"></a><span data-ttu-id="90067-116">Leer la celda activa</span><span class="sxs-lookup"><span data-stu-id="90067-116">Read the active cell</span></span>

<span data-ttu-id="90067-117">Este script registra el valor de la celda activa actual.</span><span class="sxs-lookup"><span data-stu-id="90067-117">This script logs the value of the current active cell.</span></span> <span data-ttu-id="90067-118">Si se seleccionan varias celdas, se registrará la celda superior izquierda.</span><span class="sxs-lookup"><span data-stu-id="90067-118">If multiple cells are selected, the top-leftmost cell will be logged.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the current active cell in the workbook.
  let cell = workbook.getActiveCell();

  // Log that cell's value.
  console.log(`The current cell's value is ${cell.getValue()}`);
}
```

### <a name="change-an-adjacent-cell"></a><span data-ttu-id="90067-119">Cambiar una celda adyacente</span><span class="sxs-lookup"><span data-stu-id="90067-119">Change an adjacent cell</span></span>

<span data-ttu-id="90067-120">Este script obtiene celdas adyacentes mediante referencias relativas.</span><span class="sxs-lookup"><span data-stu-id="90067-120">This script gets adjacent cells using relative references.</span></span> <span data-ttu-id="90067-121">Tenga en cuenta que si la celda activa está en la fila superior, se produce un error en parte del script, ya que hace referencia a la celda por encima de la seleccionada actualmente.</span><span class="sxs-lookup"><span data-stu-id="90067-121">Note that if the active cell is on the top row, part of the script fails, because it references the cell above the currently selected one.</span></span>

```TypeScript
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

### <a name="change-all-adjacent-cells"></a><span data-ttu-id="90067-122">Cambiar todas las celdas adyacentes</span><span class="sxs-lookup"><span data-stu-id="90067-122">Change all adjacent cells</span></span>

<span data-ttu-id="90067-123">Este script copia el formato de la celda activa en las celdas adyacentes.</span><span class="sxs-lookup"><span data-stu-id="90067-123">This script copies the formatting in the active cell to the neighboring cells.</span></span> <span data-ttu-id="90067-124">Tenga en cuenta que este script solo funciona cuando la celda activa no está en un borde de la hoja de cálculo.</span><span class="sxs-lookup"><span data-stu-id="90067-124">Note that this script only works when the active cell isn't on an edge of the worksheet.</span></span>

```TypeScript
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

### <a name="change-each-individual-cell-in-a-range"></a><span data-ttu-id="90067-125">Cambiar cada celda individual de un intervalo</span><span class="sxs-lookup"><span data-stu-id="90067-125">Change each individual cell in a range</span></span>

<span data-ttu-id="90067-126">Este script recorre el intervalo seleccionado actualmente.</span><span class="sxs-lookup"><span data-stu-id="90067-126">This script loops over the currently select range.</span></span> <span data-ttu-id="90067-127">Borra el formato actual y establece el color de relleno de cada celda en un color aleatorio.</span><span class="sxs-lookup"><span data-stu-id="90067-127">It clears the current formatting and sets the fill color in each cell to a random color.</span></span>

```TypeScript
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

### <a name="get-groups-of-cells-based-on-special-criteria"></a><span data-ttu-id="90067-128">Obtener grupos de celdas según criterios especiales</span><span class="sxs-lookup"><span data-stu-id="90067-128">Get groups of cells based on special criteria</span></span>

<span data-ttu-id="90067-129">Este script obtiene todas las celdas en blanco en el rango usado de la hoja de cálculo actual.</span><span class="sxs-lookup"><span data-stu-id="90067-129">This script gets all the blank cells in the current worksheet's used range.</span></span> <span data-ttu-id="90067-130">A continuación, resalta todas las celdas con un fondo amarillo.</span><span class="sxs-lookup"><span data-stu-id="90067-130">It then highlights all those cells with a yellow background.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the current used range.
    let range = workbook.getActiveWorksheet().getUsedRange();
    
    // Get all the blank cells.
    let blankCells = range.getSpecialCells(ExcelScript.SpecialCellType.blanks);

    // Highlight the blank cells with a yellow background.
    blankCells.getFormat().getFill().setColor("yellow");
}
```

## <a name="collections"></a><span data-ttu-id="90067-131">Colecciones</span><span class="sxs-lookup"><span data-stu-id="90067-131">Collections</span></span>

<span data-ttu-id="90067-132">Estos ejemplos funcionan con colecciones de objetos en el libro.</span><span class="sxs-lookup"><span data-stu-id="90067-132">These samples work with collections of objects in the workbook.</span></span>

### <a name="iterate-over-collections"></a><span data-ttu-id="90067-133">Iteración sobre colecciones</span><span class="sxs-lookup"><span data-stu-id="90067-133">Iterate over collections</span></span>

<span data-ttu-id="90067-134">Este script obtiene y registra los nombres de todas las hojas de cálculo del libro.</span><span class="sxs-lookup"><span data-stu-id="90067-134">This script gets and logs the names of all the worksheets in the workbook.</span></span> <span data-ttu-id="90067-135">También establece los colores de sus pestañas en un color aleatorio.</span><span class="sxs-lookup"><span data-stu-id="90067-135">It also sets the their tab colors to a random color.</span></span>

```TypeScript
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

### <a name="query-and-delete-from-a-collection"></a><span data-ttu-id="90067-136">Consulta y eliminación de una colección</span><span class="sxs-lookup"><span data-stu-id="90067-136">Query and delete from a collection</span></span>

<span data-ttu-id="90067-137">Este script crea una nueva hoja de cálculo.</span><span class="sxs-lookup"><span data-stu-id="90067-137">This script creates a new worksheet.</span></span> <span data-ttu-id="90067-138">Comprueba una copia existente de la hoja de cálculo y la elimina antes de crear una hoja nueva.</span><span class="sxs-lookup"><span data-stu-id="90067-138">It checks for an existing copy of the worksheet and deletes it before making a new sheet.</span></span>

```TypeScript
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

## <a name="dates"></a><span data-ttu-id="90067-139">Fechas</span><span class="sxs-lookup"><span data-stu-id="90067-139">Dates</span></span>

<span data-ttu-id="90067-140">Los ejemplos de esta sección muestran cómo usar el objeto [Date](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date) de JavaScript.</span><span class="sxs-lookup"><span data-stu-id="90067-140">The samples in this section show how to use the JavaScript [Date](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date) object.</span></span>

<span data-ttu-id="90067-141">En el ejemplo siguiente se obtiene la fecha y hora actuales y, a continuación, se escriben esos valores en dos celdas de la hoja de cálculo activa.</span><span class="sxs-lookup"><span data-stu-id="90067-141">The following sample gets the current date and time and then writes those values to two cells in the active worksheet.</span></span>

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

<span data-ttu-id="90067-142">En el ejemplo siguiente se lee una fecha que se almacena en Excel y se traduce a un objeto Date de JavaScript.</span><span class="sxs-lookup"><span data-stu-id="90067-142">The next sample reads a date that's stored in Excel and translates it to a JavaScript Date object.</span></span> <span data-ttu-id="90067-143">Usa el número [de serie numérico de la](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) fecha como entrada para la fecha de JavaScript.</span><span class="sxs-lookup"><span data-stu-id="90067-143">It uses the [date's numeric serial number](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) as input for the JavaScript Date.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Read a date at cell A1 from Excel.
  let dateRange = workbook.getActiveWorksheet().getRange("A1");

  // Convert the Excel date to a JavaScript Date object.
  let excelDateValue = dateRange.getValue() as number;
  let javaScriptDate = new Date(Math.round((excelDateValue - 25569) * 86400 * 1000));
  console.log(javaScriptDate);
}
```

## <a name="display-data"></a><span data-ttu-id="90067-144">Mostrar datos</span><span class="sxs-lookup"><span data-stu-id="90067-144">Display data</span></span>

<span data-ttu-id="90067-145">Estos ejemplos muestran cómo trabajar con datos de hoja de cálculo y proporcionar a los usuarios una mejor vista u organización.</span><span class="sxs-lookup"><span data-stu-id="90067-145">These samples demonstrate how to work with worksheet data and provide users with a better view or organization.</span></span>

### <a name="apply-conditional-formatting"></a><span data-ttu-id="90067-146">Aplicar formato condicional</span><span class="sxs-lookup"><span data-stu-id="90067-146">Apply conditional formatting</span></span>

<span data-ttu-id="90067-147">En este ejemplo se aplica formato condicional al intervalo usado actualmente en la hoja de cálculo.</span><span class="sxs-lookup"><span data-stu-id="90067-147">This sample applies conditional formatting to the currently used range in the worksheet.</span></span> <span data-ttu-id="90067-148">El formato condicional es un relleno verde para el 10 % superior de los valores.</span><span class="sxs-lookup"><span data-stu-id="90067-148">The conditional formatting is a green fill for the top 10% of values.</span></span>

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

### <a name="create-a-sorted-table"></a><span data-ttu-id="90067-149">Crear una tabla ordenada</span><span class="sxs-lookup"><span data-stu-id="90067-149">Create a sorted table</span></span>

<span data-ttu-id="90067-150">En este ejemplo se crea una tabla a partir del intervalo usado de la hoja de cálculo actual y, a continuación, se ordena en función de la primera columna.</span><span class="sxs-lookup"><span data-stu-id="90067-150">This sample creates a table from the current worksheet's used range, then sorts it based on the first column.</span></span>

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

### <a name="log-the-grand-total-values-from-a-pivottable"></a><span data-ttu-id="90067-151">Registrar los valores "Total general" de una tabla dinámica</span><span class="sxs-lookup"><span data-stu-id="90067-151">Log the "Grand Total" values from a PivotTable</span></span>

<span data-ttu-id="90067-152">En este ejemplo se busca la primera tabla dinámica del libro y se registra los valores en las celdas "Total general" (como se resalta en verde en la imagen siguiente).</span><span class="sxs-lookup"><span data-stu-id="90067-152">This sample finds the first PivotTable in the workbook and logs the values in the "Grand Total" cells (as highlighted in green in the image below).</span></span>

:::image type="content" source="../../images/sample-pivottable-grand-total-row.png" alt-text="Tabla dinámica que muestra las ventas de frutas con la fila Grand Total resaltada en verde.":::

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the first PivotTable in the workbook.
  let pivotTable = workbook.getPivotTables()[0];

  // Get the names of each data column in the PivotTable.
  let pivotColumnLabelRange = pivotTable.getLayout().getColumnLabelRange();

  // Get the range displaying the pivoted data.
  let pivotDataRange = pivotTable.getLayout().getBodyAndTotalRange();

  // Get the range with the "grand totals" for the PivotTable columns.
  let grandTotalRange = pivotDataRange.getLastRow();

  // Print each of the "Grand Totals" to the console.
  grandTotalRange.getValues()[0].forEach((column, columnIndex) => {
    console.log(`Grand total of ${pivotColumnLabelRange.getValues()[0][columnIndex]}: ${grandTotalRange.getValues()[0][columnIndex]}`);
    // Example log: "Grand total of Sum of Crates Sold Wholesale: 11000"
  });
}
```

### <a name="create-a-drop-down-list-using-data-validation"></a><span data-ttu-id="90067-154">Crear una lista desplegable con validación de datos</span><span class="sxs-lookup"><span data-stu-id="90067-154">Create a drop-down list using data validation</span></span>

<span data-ttu-id="90067-155">Este script crea una lista de selección desplegable para una celda.</span><span class="sxs-lookup"><span data-stu-id="90067-155">This script creates a drop-down selection list for a cell.</span></span> <span data-ttu-id="90067-156">Usa los valores existentes del intervalo seleccionado como opciones para la lista.</span><span class="sxs-lookup"><span data-stu-id="90067-156">It uses the existing values of the selected range as the choices for the list.</span></span>

:::image type="content" source="../../images/sample-data-validation.png" alt-text="Una hoja de cálculo que muestra un rango de tres celdas que contienen opciones de color &quot;rojo, azul, verde&quot; y junto a ella, las mismas opciones que se muestran en una lista desplegable.":::

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the values for data validation.
  let selectedRange = workbook.getSelectedRange();
  let rangeValues = selectedRange.getValues();

  // Convert the values into a comma-delimited string.
  let dataValidationListString = "";
  rangeValues.forEach((rangeValueRow) => {
    rangeValueRow.forEach((value) => {
      dataValidationListString += value + ",";
    });
  });

  // Clear the old range.
  selectedRange.clear(ExcelScript.ClearApplyTo.contents);

  // Apply the data validation to the first cell in the selected range.
  let targetCell = selectedRange.getCell(0,0);
  let dataValidation = targetCell.getDataValidation();

  // Set the content of the drop-down list.
  dataValidation.setRule({
      list: {
        inCellDropDown: true,
        source: dataValidationListString
      }
    });
}
```

## <a name="formulas"></a><span data-ttu-id="90067-158">Fórmulas</span><span class="sxs-lookup"><span data-stu-id="90067-158">Formulas</span></span>

<span data-ttu-id="90067-159">Estos ejemplos usan Excel fórmulas y muestran cómo trabajar con ellas en scripts.</span><span class="sxs-lookup"><span data-stu-id="90067-159">These samples use Excel formulas and show how to work with them in scripts.</span></span>

### <a name="single-formula"></a><span data-ttu-id="90067-160">Fórmula única</span><span class="sxs-lookup"><span data-stu-id="90067-160">Single formula</span></span>

<span data-ttu-id="90067-161">Este script establece la fórmula de una celda y, a continuación, muestra cómo Excel la fórmula y el valor de la celda por separado.</span><span class="sxs-lookup"><span data-stu-id="90067-161">This script sets a cell's formula, then displays how Excel stores the cell's formula and value separately.</span></span>

```TypeScript
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

### <a name="handle-a-spill-error-returned-from-a-formula"></a><span data-ttu-id="90067-162">Controlar un `#SPILL!` error devuelto desde una fórmula</span><span class="sxs-lookup"><span data-stu-id="90067-162">Handle a `#SPILL!` error returned from a formula</span></span>

<span data-ttu-id="90067-163">Este script transpone el intervalo "A1:D2" a "A4:B7" mediante la función TRANSPOSE.</span><span class="sxs-lookup"><span data-stu-id="90067-163">This script transposes the range "A1:D2" to "A4:B7" by using the TRANSPOSE function.</span></span> <span data-ttu-id="90067-164">Si la transponer da como resultado un error, borra el intervalo de destino y vuelve a aplicar `#SPILL` la fórmula.</span><span class="sxs-lookup"><span data-stu-id="90067-164">If the transpose results in a `#SPILL` error, it clears the target range and applies the formula again.</span></span>

```TypeScript
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

## <a name="suggest-new-samples"></a><span data-ttu-id="90067-165">Sugerir nuevas muestras</span><span class="sxs-lookup"><span data-stu-id="90067-165">Suggest new samples</span></span>

<span data-ttu-id="90067-166">Le damos la bienvenida a las sugerencias de nuevos ejemplos.</span><span class="sxs-lookup"><span data-stu-id="90067-166">We welcome suggestions for new samples.</span></span> <span data-ttu-id="90067-167">Si hay un escenario común que podría ayudar a otros desarrolladores de scripts, díganoslo en la sección de comentarios de la parte inferior de la página.</span><span class="sxs-lookup"><span data-stu-id="90067-167">If there is a common scenario that would help other script developers, please tell us in the feedback section at the bottom of the page.</span></span>

## <a name="see-also"></a><span data-ttu-id="90067-168">Consulte también</span><span class="sxs-lookup"><span data-stu-id="90067-168">See also</span></span>

* [<span data-ttu-id="90067-169">"Conceptos básicos del rango" de Sudhi Ramamurthy en YouTube</span><span class="sxs-lookup"><span data-stu-id="90067-169">Sudhi Ramamurthy's "Range basics" on YouTube</span></span>](https://youtu.be/4emjkOFdLBA)
* [<span data-ttu-id="90067-170">Office Ejemplos y escenarios de scripts</span><span class="sxs-lookup"><span data-stu-id="90067-170">Office Scripts samples and scenarios</span></span>](samples-overview.md)
* [<span data-ttu-id="90067-171">Grabar, editar y crear scripts de Office en Excel en la Web</span><span class="sxs-lookup"><span data-stu-id="90067-171">Record, edit, and create Office Scripts in Excel on the web</span></span>](../../tutorials/excel-tutorial.md)
