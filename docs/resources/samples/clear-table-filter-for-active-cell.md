---
title: Quitar filtros de columna de tabla
description: Obtenga información sobre cómo borrar el filtro de columna de tabla en función de la ubicación de celda activa.
ms.date: 07/15/2022
ms.localizationpriority: medium
ms.openlocfilehash: e016f7f2af9e7553229f3b3b19007e011879de8e
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572524"
---
# <a name="remove-table-column-filters"></a>Quitar filtros de columna de tabla

En este ejemplo se quitan los filtros de una columna de tabla, en función de la ubicación de celda activa. El script detecta si la celda forma parte de una tabla, determina la columna de tabla y borra cualquier filtro que se aplique a ella.

Si desea obtener más información sobre cómo guardar el filtro antes de borrarlo (y volver a aplicarlo más adelante), consulte [Mover filas entre tablas guardando filtros](move-rows-across-tables.md), un ejemplo más avanzado.

## <a name="sample-excel-file"></a>Archivo de Excel de ejemplo

Descargue [table-with-filter.xlsx](table-with-filter.xlsx) de un libro listo para usar. Agregue el siguiente script para probar el ejemplo usted mismo.

## <a name="sample-code-clear-table-column-filter-based-on-active-cell"></a>Código de ejemplo: Borrar filtro de columna de tabla en función de la celda activa

El siguiente script borra el filtro de columna de tabla en función de la ubicación de celda activa y se puede aplicar a cualquier archivo de Excel con una tabla.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the active cell.
  const cell = workbook.getActiveCell();

  // Get the tables associated with that cell.
  // Since tables can't overlap, this will be one table at most.
  const currentTable = cell.getTables()[0];

  // If there is no table on the selection, end the script.
  if (!currentTable) {
    console.log("The selection is not in a table.");
    return;
  }

  // Get the table header above the current cell by referencing its column.
  const entireColumn = cell.getEntireColumn();
  const intersect = entireColumn.getIntersection(currentTable.getRange());
  const headerCellValue = intersect.getCell(0, 0).getValue() as string;

  // Get the TableColumn object matching that header.
  const tableColumn = currentTable.getColumnByName(headerCellValue);

  // Clear the filters on that table column.
  tableColumn.getFilter().clear();
}
```

## <a name="before-clearing-column-filter-notice-the-active-cell"></a>Antes de borrar el filtro de columna (observe la celda activa)

:::image type="content" source="../../images/before-filter-applied.png" alt-text="Celda activa antes de borrar el filtro de columna.":::

## <a name="after-clearing-column-filter"></a>Después de borrar el filtro de columna

:::image type="content" source="../../images/after-filter-cleared.png" alt-text="Celda activa después de borrar el filtro de columna.":::
