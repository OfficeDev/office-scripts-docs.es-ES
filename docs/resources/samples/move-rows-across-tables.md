---
title: Mover filas entre tablas mediante Office scripts
description: Aprenda a mover filas entre tablas guardando filtros y procesando y reaprobando los filtros.
ms.date: 05/06/2021
localization_priority: Normal
ms.openlocfilehash: 02fa99ff0444924bd2d44ad4fa421fe66fbd7272
ms.sourcegitcommit: 763d341857bcb209b2f2c278a82fdb63d0e18f0a
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 05/08/2021
ms.locfileid: "52285944"
---
# <a name="move-rows-across-tables-by-saving-filters-then-processing-and-reapplying-the-filters"></a>Mueva filas entre tablas guardando filtros y procesando y reaprobando los filtros.

Este script hace lo siguiente:

* Selecciona filas de la tabla de origen donde el valor de una columna es igual a _algún valor_.
* Mueve todas las filas seleccionadas a otra tabla (destino) de otra hoja de cálculo.
* Vuelve a aplicar los filtros relevantes en la tabla de origen.

:::image type="content" source="../../images/table-filter-before-after.png" alt-text="Capturas de pantalla del libro antes y después":::

## <a name="sample-excel-file"></a>Archivo Excel ejemplo

Descargue el archivo <a href="input-table-filters.xlsx">input-table-filters.xlsx</a> se usa en esta solución para probarlo usted mismo.

## <a name="sample-code-move-rows-using-range-values"></a>Código de ejemplo: mover filas con valores de intervalo

```TypeScript
function main(workbook: ExcelScript.Workbook) {

  // You can change these names to match the data in your workbook.
  const TargetTableName = 'Table1';
  const SourceTableName = 'Table2';
  const IndexOfColumnToFilterOn = 1;
  const NameOfColumnToFilterOn = 'Category';
  const ValueToFilterOn = 'Clothing';

  // Get the Table objects.
  let targetTable = workbook.getTable(TargetTableName);
  let sourceTable = workbook.getTable(SourceTableName);

  // If either table is missing, report that information and stop the script.
  if (!targetTable || !sourceTable) {
    console.log(`Tables missing - Check to make sure both source (${TargetTableName}) and target table (${SourceTableName}) are present before running the script. `);
    return;
  }

  // Save the filter criteria.
  const tableFilters = {};
  // For each table column, collect the filter criteria on that column.
  sourceTable.getColumns().forEach((column) => {
    let colFilterCriteria = column.getFilter().getCriteria();
    if (colFilterCriteria) {
      tableFilters[column.getName()] = colFilterCriteria;
    }
  });

  // Get all the data from the table.
  const sourceRange = sourceTable.getRangeBetweenHeaderAndTotal();
  const dataRows: (number | string | boolean)[][] = sourceTable.getRangeBetweenHeaderAndTotal().getValues();

  // Create variables to hold the rows to be moved and their addresses.
  let rowsToMoveValues: (number | string | boolean)[][] = [];
  let rowAddressToRemove: string[] = [];

  // Get the data values from the source table.
  for (let i = 0; i < dataRows.length; i++) { 
    if (dataRows[i][IndexOfColumnToFilterOn] === ValueToFilterOn) {
      rowsToMoveValues.push(dataRows[i]);

      // Get the intersection between table address and the entire row where we found the match. This provides the address of the range to remove.
      let address = sourceRange.getIntersection(sourceRange.getCell(i,0).getEntireRow()).getAddress();
      rowAddressToRemove.push(address);
    }
  }

  // If there are no data rows to process, end the script.
  if (rowsToMoveValues.length < 1) {
    console.log('No rows selected from the source table match the filter criteria.');
    return;
  }

  console.log(`Adding ${rowsToMoveValues.length} rows to target table.`);

  // Insert rows at the end of target table.
  targetTable.addRows(-1, rowsToMoveValues)

  // Remove the rows from the source table.
  const sheet = sourceTable.getWorksheet();

  // Remove all filters before removing rows.
  sourceTable.getAutoFilter().clearCriteria();

  // Important: Remove the rows starting at the bottom of the table.
  // Otherwise, the lower rows change position before they are deleted.
  console.log(`Removing ${rowAddressToRemove.length} rows from the source table.`);
  rowAddressToRemove.reverse().forEach((address) => {
    sheet.getRange(address).delete(ExcelScript.DeleteShiftDirection.up);
  });

  // Reapply the original filters. 
  Object.keys(tableFilters).forEach((columnName) => {
      sourceTable.getColumnByName(columnName).getFilter().apply(tableFilters[columnName]);
    });
}
```

## <a name="training-video-move-rows-across-tables"></a>Vídeo de aprendizaje: Mover filas entre tablas

[Vea el recorrido de Sudhi Ramamurthy por este ejemplo en YouTube](https://youtu.be/_3t3Pk4i2L0). Hay dos scripts que se muestran en la solución del vídeo. La diferencia principal es cómo se seleccionan las filas.

* En la primera variante, las filas se seleccionan aplicando el filtro de tabla y leyendo el intervalo visible.
* En el segundo, las filas se seleccionan leyendo los valores y extrayendo los valores de fila (que es lo que usa el ejemplo de esta página).
