---
title: Mover filas entre tablas mediante scripts de Office
description: Obtenga información sobre cómo mover filas entre tablas guardando filtros y procesando y volviendo a aplicar los filtros.
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: a7c28c4fef91402b8889d749a03f3aab5e615521
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572622"
---
# <a name="move-rows-across-tables"></a>Mover filas entre tablas

Este script hace lo siguiente:

* Selecciona filas de la tabla de origen donde el valor de una columna es igual a algún valor (`FILTER_VALUE` en el script).
* Mueve todas las filas seleccionadas a la tabla de destino de otra hoja de cálculo.
* Vuelve a aplicar los filtros pertinentes a la tabla de origen.

## <a name="sample-excel-file"></a>Archivo de Excel de ejemplo

Descargue el [ archivoinput-table-filters.xlsx](input-table-filters.xlsx) de un libro listo para usar. Agregue el siguiente script para probar el ejemplo usted mismo.

## <a name="sample-code-move-rows-using-range-values"></a>Código de ejemplo: Mover filas con valores de intervalo

```TypeScript
function main(workbook: ExcelScript.Workbook) {

  // You can change these names to match the data in your workbook.
  const TARGET_TABLE_NAME = 'Table1';
  const SOURCE_TABLE_NAME = 'Table2';

  // Select what will be moved between tables.
  const FILTER_COLUMN_INDEX = 1;
  const FILTER_VALUE = 'Clothing';

  // Get the Table objects.
  let targetTable = workbook.getTable(TARGET_TABLE_NAME);
  let sourceTable = workbook.getTable(SOURCE_TABLE_NAME);

  // If either table is missing, report that information and stop the script.
  if (!targetTable || !sourceTable) {
    console.log(`Tables missing - Check to make sure both source (${TARGET_TABLE_NAME}) and target table (${SOURCE_TABLE_NAME}) are present before running the script. `);
    return;
  }

  // Save the filter criteria currently on the source table.
  const originalTableFilters = {};
  // For each table column, collect the filter criteria on that column.
  sourceTable.getColumns().forEach((column) => {
    let originalColumnFilter = column.getFilter().getCriteria();
    if (originalColumnFilter) {
      originalTableFilters[column.getName()] = originalColumnFilter;
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
    if (dataRows[i][FILTER_COLUMN_INDEX] === FILTER_VALUE) {
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
  Object.keys(originalTableFilters).forEach((columnName) => {
      sourceTable.getColumnByName(columnName).getFilter().apply(originalTableFilters[columnName]);
    });
}
```

## <a name="training-video-move-rows-across-tables"></a>Vídeo de entrenamiento: Traslado de filas entre tablas

[Vea cómo Sudhi Ramamurthy recorre este ejemplo en YouTube](https://youtu.be/_3t3Pk4i2L0). En la solución del vídeo se muestran dos scripts. La principal diferencia es cómo se seleccionan las filas.

* En la primera variante, las filas se seleccionan aplicando el filtro de tabla y leyendo el intervalo visible.
* En el segundo, las filas se seleccionan leyendo los valores y extrayendo los valores de fila (que es lo que usa el ejemplo de esta página).
