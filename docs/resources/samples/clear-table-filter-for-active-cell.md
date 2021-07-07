---
title: Borrar filtro de columna de tabla en función de la ubicación de celda activa
description: Obtenga información sobre cómo borrar el filtro de columna de tabla en función de la ubicación de celda activa.
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: f10e23b4ad948a28c5b749533ddedefe164d7142
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 07/07/2021
ms.locfileid: "53313893"
---
# <a name="clear-table-column-filter-based-on-active-cell-location"></a><span data-ttu-id="95e8d-103">Borrar filtro de columna de tabla en función de la ubicación de celda activa</span><span class="sxs-lookup"><span data-stu-id="95e8d-103">Clear table column filter based on active cell location</span></span>

<span data-ttu-id="95e8d-104">En este ejemplo se borra el filtro de columna de tabla en función de la ubicación de la celda activa.</span><span class="sxs-lookup"><span data-stu-id="95e8d-104">This sample clears the table column filter based on the active cell location.</span></span> <span data-ttu-id="95e8d-105">El script detecta si la celda forma parte de una tabla, determina la columna de tabla y borra cualquier filtro que se aplique en ella.</span><span class="sxs-lookup"><span data-stu-id="95e8d-105">The script detects if the cell is part of a table, determines the table column, and clears any filter that are applied on it.</span></span>

<span data-ttu-id="95e8d-106">Si desea obtener más información sobre cómo guardar el filtro antes de borrarlo (y volver a aplicarlo más adelante), vea [Mover](move-rows-across-tables.md)filas entre tablas guardando filtros, un ejemplo más avanzado.</span><span class="sxs-lookup"><span data-stu-id="95e8d-106">If you wish to learn more about how to save the filter prior to clearing it (and re-apply later), see [Move rows across tables by saving filters](move-rows-across-tables.md), a more advanced sample.</span></span>

<span data-ttu-id="95e8d-107">_Antes de borrar el filtro de columna (observe la celda activa)_</span><span class="sxs-lookup"><span data-stu-id="95e8d-107">_Before clearing column filter (notice the active cell)_</span></span>

:::image type="content" source="../../images/before-filter-applied.png" alt-text="Una celda activa antes de borrar el filtro de columna.":::

<span data-ttu-id="95e8d-109">_Después de borrar el filtro de columna_</span><span class="sxs-lookup"><span data-stu-id="95e8d-109">_After clearing column filter_</span></span>

:::image type="content" source="../../images/after-filter-cleared.png" alt-text="Una celda activa después de borrar el filtro de columna.":::

## <a name="sample-excel-file"></a><span data-ttu-id="95e8d-111">Archivo Excel ejemplo</span><span class="sxs-lookup"><span data-stu-id="95e8d-111">Sample Excel file</span></span>

<span data-ttu-id="95e8d-112">Descargue <a href="table-with-filter.xlsx">table-with-filter.xlsx</a> para un libro listo para usar.</span><span class="sxs-lookup"><span data-stu-id="95e8d-112">Download <a href="table-with-filter.xlsx">table-with-filter.xlsx</a> for a ready-to-use workbook.</span></span> <span data-ttu-id="95e8d-113">Agregue el siguiente script para probar el ejemplo usted mismo.</span><span class="sxs-lookup"><span data-stu-id="95e8d-113">Add the following script to try the sample yourself!</span></span>

## <a name="sample-code-clear-table-column-filter-based-on-active-cell"></a><span data-ttu-id="95e8d-114">Código de ejemplo: borrar el filtro de columna de tabla en función de la celda activa</span><span class="sxs-lookup"><span data-stu-id="95e8d-114">Sample code: Clear table column filter based on active cell</span></span>

<span data-ttu-id="95e8d-115">El siguiente script borra el filtro de columna de tabla en función de la ubicación de celda activa y se puede aplicar a cualquier archivo Excel con una tabla.</span><span class="sxs-lookup"><span data-stu-id="95e8d-115">The following script clears the table column filter based on active cell location and can be applied to any Excel file with a table.</span></span>

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
