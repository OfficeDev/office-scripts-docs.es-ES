---
title: Scripts básicos para scripts de Office en Excel
description: Colección de ejemplos de código que se usarán con scripts de Office en Excel.
ms.date: 03/24/2022
ms.localizationpriority: medium
ms.openlocfilehash: 8e28026b7a3498d477cce8b6dc5940da33a30f53
ms.sourcegitcommit: 34c7740c9bff0e4c7426e01029f967724bfee566
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 05/13/2022
ms.locfileid: "65393659"
---
# <a name="basic-scripts-for-office-scripts-in-excel"></a>Scripts básicos para scripts de Office en Excel

Los ejemplos siguientes son scripts sencillos para probar en sus propios libros. Para usarlos en Excel:

1. Abra un libro en Excel en la Web.
1. Abra la pestaña **Automatizar**.
1. Seleccione **Nuevo script**.
1. Reemplace todo el script por el ejemplo que prefiera.
1. Seleccione **Ejecutar** en el panel de tareas del Editor de código.

## <a name="script-basics"></a>Conceptos básicos del script

En estos ejemplos se muestran los bloques de creación fundamentales para scripts de Office. Expanda estos scripts para ampliar la solución y resolver problemas comunes.

### <a name="read-and-log-one-cell"></a>Leer y registrar una celda

En este ejemplo se lee el valor de **A1** y se imprime en la consola.

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

### <a name="read-the-active-cell"></a>Leer la celda activa

Este script registra el valor de la celda activa actual. Si se seleccionan varias celdas, se registrará la celda superior izquierda.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the current active cell in the workbook.
  let cell = workbook.getActiveCell();

  // Log that cell's value.
  console.log(`The current cell's value is ${cell.getValue()}`);
}
```

### <a name="change-an-adjacent-cell"></a>Cambio de una celda adyacente

Este script obtiene celdas adyacentes mediante referencias relativas. Tenga en cuenta que si la celda activa está en la fila superior, se produce un error en parte del script, ya que hace referencia a la celda situada encima de la seleccionada actualmente.

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

### <a name="change-all-adjacent-cells"></a>Cambiar todas las celdas adyacentes

Este script copia el formato de la celda activa en las celdas vecinas. Tenga en cuenta que este script solo funciona cuando la celda activa no está en un borde de la hoja de cálculo.

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

### <a name="change-each-individual-cell-in-a-range"></a>Cambiar cada celda individual de un rango

Este script recorre en bucle el intervalo seleccionado actualmente. Borra el formato actual y establece el color de relleno de cada celda en un color aleatorio.

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

### <a name="get-groups-of-cells-based-on-special-criteria"></a>Obtención de grupos de celdas según criterios especiales

Este script obtiene todas las celdas en blanco del rango usado de la hoja de cálculo actual. A continuación, resalta todas esas celdas con un fondo amarillo.

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

## <a name="collections"></a>Colecciones

Estos ejemplos funcionan con colecciones de objetos del libro.

### <a name="iterate-over-collections"></a>Recorrer en iteración colecciones

Este script obtiene y registra los nombres de todas las hojas de cálculo del libro. También establece los colores de sus pestañas en un color aleatorio.

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

### <a name="query-and-delete-from-a-collection"></a>Consulta y eliminación de una colección

Este script crea una nueva hoja de cálculo. Comprueba si hay una copia existente de la hoja de cálculo y la elimina antes de crear una nueva hoja.

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

## <a name="dates"></a>Fechas

Los ejemplos de esta sección muestran cómo usar el objeto [Date](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date) de JavaScript.

En el ejemplo siguiente se obtiene la fecha y hora actuales y, a continuación, se escriben esos valores en dos celdas de la hoja de cálculo activa.

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

En el ejemplo siguiente se lee una fecha que se almacena en Excel y se traduce en un objeto Date de JavaScript. Usa el número de serie numérico de la fecha como entrada para la fecha de JavaScript. Este número de serie se describe en el artículo de la [función NOW().](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46)

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

## <a name="display-data"></a>Mostrar datos

En estos ejemplos se muestra cómo trabajar con datos de hoja de cálculo y proporcionar a los usuarios una mejor vista u organización.

### <a name="apply-conditional-formatting"></a>Aplicar formato condicional

Este ejemplo aplica formato condicional al rango usado actualmente en la hoja de cálculo. El formato condicional es un relleno verde para el 10 % superior de los valores.

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

### <a name="create-a-sorted-table"></a>Creación de una tabla ordenada

En este ejemplo se crea una tabla a partir del rango utilizado de la hoja de cálculo actual y, a continuación, se ordena en función de la primera columna.

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

### <a name="log-the-grand-total-values-from-a-pivottable"></a>Registrar los valores "Total general" de una tabla dinámica

En este ejemplo se encuentra la primera tabla dinámica del libro y se registran los valores de las celdas "Grand Total" (como se resalta en verde en la imagen siguiente).

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

### <a name="create-a-drop-down-list-using-data-validation"></a>Creación de una lista desplegable mediante la validación de datos

Este script crea una lista desplegable de selección para una celda. Usa los valores existentes del intervalo seleccionado como las opciones de la lista.

:::image type="content" source="../../images/sample-data-validation.png" alt-text="Hoja de cálculo que muestra un rango de tres celdas que contienen opciones de color &quot;rojo, azul, verde&quot; y junto a ella, las mismas opciones que se muestran en una lista desplegable.":::

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

## <a name="formulas"></a>Fórmulas

Estos ejemplos usan fórmulas Excel y muestran cómo trabajar con ellas en scripts.

### <a name="single-formula"></a>Fórmula única

Este script establece la fórmula de una celda y, a continuación, muestra cómo Excel almacena la fórmula y el valor de la celda por separado.

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

### <a name="handle-a-spill-error-returned-from-a-formula"></a>Controlar un `#SPILL!` error devuelto desde una fórmula

Este script transpone el intervalo "A1:D2" a "A4:B7" mediante la función TRANSPOSE. Si la transponer produce un `#SPILL` error, borra el intervalo de destino y vuelve a aplicar la fórmula.

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

## <a name="suggest-new-samples"></a>Sugerir nuevos ejemplos

Agradecemos las sugerencias para nuevos ejemplos. Si hay un escenario común que ayude a otros desarrolladores de scripts, díganoslo en la sección de comentarios de la parte inferior de la página.

## <a name="see-also"></a>Consulte también

* ["Conceptos básicos del rango" de Sudhi Ramamurthy en YouTube](https://youtu.be/4emjkOFdLBA)
* [escenarios y ejemplos de scripts de Office](samples-overview.md)
* [Grabar, editar y crear Scripts de Office en Excel para la Web](../../tutorials/excel-tutorial.md)
