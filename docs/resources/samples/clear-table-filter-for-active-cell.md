---
title: Borrar filtro de columna de tabla en función de la ubicación de celda activa
description: Obtenga información sobre cómo borrar el filtro de columna de tabla en función de la ubicación de celda activa.
ms.date: 05/03/2021
localization_priority: Normal
ms.openlocfilehash: 06fba191a79f4641d4d1017bda332c7559b50e6d
ms.sourcegitcommit: 763d341857bcb209b2f2c278a82fdb63d0e18f0a
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 05/08/2021
ms.locfileid: "52285930"
---
# <a name="clear-table-column-filter-based-on-active-cell-location"></a><span data-ttu-id="6a871-103">Borrar filtro de columna de tabla en función de la ubicación de celda activa</span><span class="sxs-lookup"><span data-stu-id="6a871-103">Clear table column filter based on active cell location</span></span>

<span data-ttu-id="6a871-104">En este ejemplo se borra el filtro de columna de tabla en función de la ubicación de la celda activa.</span><span class="sxs-lookup"><span data-stu-id="6a871-104">This sample clears the table column filter based on the active cell location.</span></span> <span data-ttu-id="6a871-105">El script detecta si la celda forma parte de una tabla, determina la columna de tabla y borra cualquier filtro que se aplique en ella.</span><span class="sxs-lookup"><span data-stu-id="6a871-105">The script detects if the cell is part of a table, determines the table column, and clears any filter that are applied on it.</span></span>

<span data-ttu-id="6a871-106">Si desea obtener más información sobre cómo guardar el filtro antes de borrarlo (y volver a aplicarlo más adelante), vea [Mover](move-rows-across-tables.md)filas entre tablas guardando filtros, un ejemplo más avanzado.</span><span class="sxs-lookup"><span data-stu-id="6a871-106">If you wish to learn more about how to save the filter prior to clearing it (and re-apply later), see [Move rows across tables by saving filters](move-rows-across-tables.md), a more advanced sample.</span></span>

<span data-ttu-id="6a871-107">_Antes de borrar el filtro de columna (observe la celda activa)_</span><span class="sxs-lookup"><span data-stu-id="6a871-107">_Before clearing column filter (notice the active cell)_</span></span>

:::image type="content" source="../../images/before-filter-applied.png" alt-text="Una celda activa antes de borrar el filtro de columna":::

<span data-ttu-id="6a871-109">_Después de borrar el filtro de columna_</span><span class="sxs-lookup"><span data-stu-id="6a871-109">_After clearing column filter_</span></span>

:::image type="content" source="../../images/after-filter-cleared.png" alt-text="Una celda activa después de borrar el filtro de columna":::

## <a name="sample-code-clear-table-column-filter-based-on-active-cell"></a><span data-ttu-id="6a871-111">Código de ejemplo: borrar el filtro de columna de tabla en función de la celda activa</span><span class="sxs-lookup"><span data-stu-id="6a871-111">Sample code: Clear table column filter based on active cell</span></span>

<span data-ttu-id="6a871-112">El siguiente script borra el filtro de columna de tabla en función de la ubicación de celda activa y se puede aplicar a cualquier archivo Excel con una tabla.</span><span class="sxs-lookup"><span data-stu-id="6a871-112">The following script clears the table column filter based on active cell location and can be applied to any Excel file with a table.</span></span> <span data-ttu-id="6a871-113">Para mayor comodidad, puede descargar y usar <a href="table-with-filter.xlsx">table-with-filter.xlsx</a>.</span><span class="sxs-lookup"><span data-stu-id="6a871-113">For convenience, you can download and use <a href="table-with-filter.xlsx">table-with-filter.xlsx</a>.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the active cell.
    const cell = workbook.getActiveCell();

    // Get all tables associated with that cell.
    const tables = cell.getTables();
    
    // If there is no table on the selection, end the script.
    if (tables.length !== 1) {
      console.log("The selection is not in a table.");
      return;
    }

    // Get the first table associated with the active cell.
    const currentTable = tables[0];

    // Log key information about the table.
    console.log(currentTable.getName());
    console.log(currentTable.getRange().getAddress());

    // Get the table header above the current cell by referencing its column.
    const entireColumn = cell.getEntireColumn();
    const intersect = entireColumn.getIntersection(currentTable.getRange());
    console.log(intersect.getAddress());

    const headerCellValue = intersect.getCell(0,0).getValue() as string;
    console.log(headerCellValue);

    // Get the TableColumn object matching that header.
    const tableColumn = currentTable.getColumnByName(headerCellValue);

    // Clear the filter on that table column.
    tableColumn.getFilter().clear();
}
```
