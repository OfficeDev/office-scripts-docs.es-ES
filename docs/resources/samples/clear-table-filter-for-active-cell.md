---
title: Borrar filtro de columna de tabla en función de la ubicación de celda activa
description: Obtenga información sobre cómo borrar el filtro de columna de tabla en función de la ubicación de celda activa.
ms.date: 03/04/2021
localization_priority: Normal
ms.openlocfilehash: bbca4adce1de2cfade2c4f84273bf0bc06b5cc4b
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232504"
---
# <a name="clear-table-column-filter-based-on-active-cell-location"></a>Borrar filtro de columna de tabla en función de la ubicación de celda activa

En este ejemplo se borra el filtro de columna de tabla en función de la ubicación de la celda activa. El script detecta si la celda forma parte de una tabla, determina la columna de tabla y borra cualquier filtro que se aplique en ella.

Si desea obtener más información sobre cómo guardar el filtro antes de borrarlo (y volver a aplicarlo más adelante), vea [Mover](move-rows-across-tables.md)filas entre tablas guardando filtros, un ejemplo más avanzado.

_Antes de borrar el filtro de columna (observe la celda activa)_

:::image type="content" source="../../images/before-filter-applied.png" alt-text="Una celda activa antes de borrar el filtro de columna":::

_Después de borrar el filtro de columna_

:::image type="content" source="../../images/after-filter-cleared.png" alt-text="Una celda activa después de borrar el filtro de columna":::

## <a name="sample-code-clear-table-column-filter-based-on-active-cell"></a>Código de ejemplo: borrar el filtro de columna de tabla en función de la celda activa

El siguiente script borra el filtro de columna de tabla en función de la ubicación de celda activa y se puede aplicar a cualquier archivo Excel con una tabla. Para mayor comodidad, puede descargar y usar <a href="table-with-filter.xlsx">table-with-filter.xlsx</a>.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get active cell.
    const cell = workbook.getActiveCell();

    // Get all tables associated with that cell.
    const tables = cell.getTables();
    
    // If there is no table on the selection, return/exit.
    if (tables.length !== 1) {
      console.log("The selection is not in a table.");
      return;
    }

    // Get table (since it is already determined that there is only
    // a single table part of the selection).
    const currentTable = tables[0];

    console.log(currentTable.getName());
    console.log(currentTable.getRange().getAddress());

    const entireCol = cell.getEntireColumn();
    const intersect = entireCol.getIntersection(currentTable.getRange());
    console.log(intersect.getAddress());

    const headerCellValue = intersect.getCell(0,0).getValue() as string;
    console.log(headerCellValue);

    // Get column.
    const col = currentTable.getColumnByName(headerCellValue);

    // Clear filter.
    col.getFilter().clear();
}
```
