---
title: Scripts de ejemplo para scripts de Office en Excel en la web
description: Una colección de ejemplos de código para usar con scripts de Office en Excel en la Web.
ms.date: 06/18/2020
localization_priority: Normal
ms.openlocfilehash: bfa6679595e6e28cc5d2ae3e3e487fd3e77738aa
ms.sourcegitcommit: aec3c971c6640429f89b6bb99d2c95ea06725599
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 06/25/2020
ms.locfileid: "44878678"
---
# <a name="sample-scripts-for-office-scripts-in-excel-on-the-web-preview"></a><span data-ttu-id="c4290-103">Scripts de ejemplo para scripts de Office en Excel en la web (vista previa)</span><span class="sxs-lookup"><span data-stu-id="c4290-103">Sample scripts for Office Scripts in Excel on the web (preview)</span></span>

<span data-ttu-id="c4290-104">Los siguientes ejemplos son scripts sencillos que puede probar en sus propios libros.</span><span class="sxs-lookup"><span data-stu-id="c4290-104">The following samples are simple scripts for you to try on your own workbooks.</span></span> <span data-ttu-id="c4290-105">Para usarlas en Excel en la web:</span><span class="sxs-lookup"><span data-stu-id="c4290-105">To use them in Excel on the web:</span></span>

1. <span data-ttu-id="c4290-106">Abra la pestaña **Automatizar**.</span><span class="sxs-lookup"><span data-stu-id="c4290-106">Open the **Automate** tab.</span></span>
2. <span data-ttu-id="c4290-107">Presione el **Editor de código**.</span><span class="sxs-lookup"><span data-stu-id="c4290-107">Press **Code Editor**.</span></span>
3. <span data-ttu-id="c4290-108">Presione **nueva secuencia de comandos** en el panel de tareas del editor de código.</span><span class="sxs-lookup"><span data-stu-id="c4290-108">Press **New Script** in the Code Editor's task pane.</span></span>
4. <span data-ttu-id="c4290-109">Reemplace todo el script por el ejemplo de su elección.</span><span class="sxs-lookup"><span data-stu-id="c4290-109">Replace the entire script with the sample of your choice.</span></span>
5. <span data-ttu-id="c4290-110">Presione **Ejecutar** en el panel de tareas del editor de código.</span><span class="sxs-lookup"><span data-stu-id="c4290-110">Press **Run** in the Code Editor's task pane.</span></span>

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="scripting-basics"></a><span data-ttu-id="c4290-111">Conceptos básicos de scripting</span><span class="sxs-lookup"><span data-stu-id="c4290-111">Scripting basics</span></span>

<span data-ttu-id="c4290-112">Estos ejemplos muestran bloques de creación fundamentales para los scripts de Office.</span><span class="sxs-lookup"><span data-stu-id="c4290-112">These samples demonstrate fundamental building blocks for Office Scripts.</span></span> <span data-ttu-id="c4290-113">Agréguelos a los scripts para ampliar la solución y resolver problemas comunes.</span><span class="sxs-lookup"><span data-stu-id="c4290-113">Add these to your scripts to extend your solution and solve common problems.</span></span>

### <a name="read-and-log-one-cell"></a><span data-ttu-id="c4290-114">Leer e iniciar sesión en una celda</span><span class="sxs-lookup"><span data-stu-id="c4290-114">Read and log one cell</span></span>

<span data-ttu-id="c4290-115">En este ejemplo se lee el valor de **a1** y se imprime en la consola.</span><span class="sxs-lookup"><span data-stu-id="c4290-115">This sample reads the value of **A1** and prints it to the console.</span></span>

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

### <a name="read-the-active-cell"></a><span data-ttu-id="c4290-116">Leer la celda activa</span><span class="sxs-lookup"><span data-stu-id="c4290-116">Read the active cell</span></span>

<span data-ttu-id="c4290-117">Esta secuencia de comandos registra el valor de la celda activa actual.</span><span class="sxs-lookup"><span data-stu-id="c4290-117">This script logs the value of the current active cell.</span></span> <span data-ttu-id="c4290-118">Si se seleccionan varias celdas, se registrará la celda superior izquierda.</span><span class="sxs-lookup"><span data-stu-id="c4290-118">If multiple cells are selected, the top-leftmost cell will be logged.</span></span>

```typescript
function main(workbook: ExcelScript.Workbook) {
  // Get the current active cell in the workbook.
  let cell = workbook.getActiveCell();

  // Log that cell's value.
  console.log(`The current cell's value is ${cell.getValue()}`);
}
```

### <a name="change-an-adjacent-cell"></a><span data-ttu-id="c4290-119">Cambiar una celda adyacente</span><span class="sxs-lookup"><span data-stu-id="c4290-119">Change an adjacent cell</span></span>

<span data-ttu-id="c4290-120">Este script obtiene las celdas adyacentes mediante referencias relativas.</span><span class="sxs-lookup"><span data-stu-id="c4290-120">This script gets adjacent cells using relative references.</span></span> <span data-ttu-id="c4290-121">Tenga en cuenta que si la celda activa está en la fila superior, se producirá un error en parte del script porque hace referencia a la celda por encima de la selección actualmente seleccionada.</span><span class="sxs-lookup"><span data-stu-id="c4290-121">Note that if the active cell is on the top row, part of the script fails, because it references the cell above the currently selected one.</span></span>

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

### <a name="change-all-adjacent-cells"></a><span data-ttu-id="c4290-122">Cambiar todas las celdas adyacentes</span><span class="sxs-lookup"><span data-stu-id="c4290-122">Change all adjacent cells</span></span>

<span data-ttu-id="c4290-123">Este script copia el formato de la celda activa en las celdas vecinas.</span><span class="sxs-lookup"><span data-stu-id="c4290-123">This script copies the formatting in the active cell to the neighboring cells.</span></span> <span data-ttu-id="c4290-124">Tenga en cuenta que este script sólo funciona cuando la celda activa no está en un borde de la hoja de cálculo.</span><span class="sxs-lookup"><span data-stu-id="c4290-124">Note that this script only works when the active cell isn't on an edge of the worksheet.</span></span>

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

### <a name="work-with-dates"></a><span data-ttu-id="c4290-125">Trabajar con fechas</span><span class="sxs-lookup"><span data-stu-id="c4290-125">Work with dates</span></span>

<span data-ttu-id="c4290-126">Los ejemplos de esta sección muestran cómo usar el objeto [Date](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date) de JavaScript.</span><span class="sxs-lookup"><span data-stu-id="c4290-126">The samples in this section show how to use the JavaScript [Date](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date) object.</span></span>

<span data-ttu-id="c4290-127">En el ejemplo siguiente se obtiene la fecha y hora actuales y, a continuación, se escriben los valores en dos celdas de la hoja de cálculo activa.</span><span class="sxs-lookup"><span data-stu-id="c4290-127">The following sample gets the current date and time and then writes those values to two cells in the active worksheet.</span></span>

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

<span data-ttu-id="c4290-128">El siguiente ejemplo lee una fecha que está almacenada en Excel y la convierte en un objeto Date de JavaScript.</span><span class="sxs-lookup"><span data-stu-id="c4290-128">The next sample reads a date that's stored in Excel and translates it to a JavaScript Date object.</span></span> <span data-ttu-id="c4290-129">Usa el [número de serie numérico de la fecha](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) como entrada para la fecha de JavaScript.</span><span class="sxs-lookup"><span data-stu-id="c4290-129">It uses the [date's numeric serial number](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) as input for the JavaScript Date.</span></span>

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

## <a name="display-data"></a><span data-ttu-id="c4290-130">Mostrar datos</span><span class="sxs-lookup"><span data-stu-id="c4290-130">Display data</span></span>

<span data-ttu-id="c4290-131">En estos ejemplos se muestra cómo trabajar con los datos de la hoja de cálculo y proporcionar a los usuarios una vista o organización mejor.</span><span class="sxs-lookup"><span data-stu-id="c4290-131">These samples demonstrate how to work with worksheet data and provide users with a better view or organization.</span></span>

### <a name="apply-conditional-formatting"></a><span data-ttu-id="c4290-132">Aplicar formato condicional</span><span class="sxs-lookup"><span data-stu-id="c4290-132">Apply conditional formatting</span></span>

<span data-ttu-id="c4290-133">En este ejemplo se aplica formato condicional al intervalo que se usa actualmente en la hoja de cálculo.</span><span class="sxs-lookup"><span data-stu-id="c4290-133">This sample applies conditional formatting to the currently used range in the worksheet.</span></span> <span data-ttu-id="c4290-134">El formato condicional es un relleno verde para el 10% de los valores principales.</span><span class="sxs-lookup"><span data-stu-id="c4290-134">The conditional formatting is a green fill for the top 10% of values.</span></span>

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

### <a name="create-a-sorted-table"></a><span data-ttu-id="c4290-135">Crear una tabla ordenada</span><span class="sxs-lookup"><span data-stu-id="c4290-135">Create a sorted table</span></span>

<span data-ttu-id="c4290-136">En este ejemplo se crea una tabla a partir del rango usado de la hoja de cálculo actual y, a continuación, se ordena basándose en la primera columna.</span><span class="sxs-lookup"><span data-stu-id="c4290-136">This sample creates a table from the current worksheet's used range, then sorts it based on the first column.</span></span>

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

### <a name="log-the-grand-total-values-from-a-pivottable"></a><span data-ttu-id="c4290-137">Registrar los valores de "total general" de una tabla dinámica</span><span class="sxs-lookup"><span data-stu-id="c4290-137">Log the "Grand Total" values from a PivotTable</span></span>

<span data-ttu-id="c4290-138">En este ejemplo se busca la primera tabla dinámica del libro y se registran los valores de las celdas de "Grand total" (resaltado en verde en la imagen siguiente).</span><span class="sxs-lookup"><span data-stu-id="c4290-138">This sample finds the first PivotTable in the workbook and logs the values in the "Grand Total" cells (as highlighted in green in the image below).</span></span>

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

## <a name="scenario-samples"></a><span data-ttu-id="c4290-140">Ejemplos de escenario</span><span class="sxs-lookup"><span data-stu-id="c4290-140">Scenario samples</span></span>

<span data-ttu-id="c4290-141">Para obtener ejemplos que muestren soluciones de gran tamaño para el mundo real, visite ejemplos [de escenarios de Office scripts](scenarios/sample-scenario-overview.md).</span><span class="sxs-lookup"><span data-stu-id="c4290-141">For samples showcasing larger, real-world solutions, visit [Sample scenarios for Office Scripts](scenarios/sample-scenario-overview.md).</span></span>

## <a name="suggest-new-samples"></a><span data-ttu-id="c4290-142">Sugerir nuevos ejemplos</span><span class="sxs-lookup"><span data-stu-id="c4290-142">Suggest new samples</span></span>

<span data-ttu-id="c4290-143">Agradecemos las sugerencias para los nuevos ejemplos.</span><span class="sxs-lookup"><span data-stu-id="c4290-143">We welcome suggestions for new samples.</span></span> <span data-ttu-id="c4290-144">Si hay un escenario común que ayudaría a otros programadores de scripts, indíquenos en la sección Comentarios a continuación.</span><span class="sxs-lookup"><span data-stu-id="c4290-144">If there is a common scenario that would help other script developers, please tell us in the feedback section below.</span></span>
