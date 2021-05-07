---
title: Combinar datos de varias Excel tablas en una sola tabla
description: Obtenga información sobre cómo usar Office scripts para combinar datos de varias Excel en una sola tabla.
ms.date: 04/28/2021
localization_priority: Normal
ms.openlocfilehash: ac8c7d0a3f0f4f3d7d3217ffac31aff1a5595d17
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232448"
---
# <a name="combine-data-from-multiple-excel-tables-into-a-single-table"></a><span data-ttu-id="e2747-103">Combinar datos de varias Excel tablas en una sola tabla</span><span class="sxs-lookup"><span data-stu-id="e2747-103">Combine data from multiple Excel tables into a single table</span></span>

<span data-ttu-id="e2747-104">En este ejemplo se combinan los datos de varias Excel tablas en una sola tabla que incluye todas las filas.</span><span class="sxs-lookup"><span data-stu-id="e2747-104">This sample combines data from multiple Excel tables into a single table that includes all the rows.</span></span> <span data-ttu-id="e2747-105">Se supone que todas las tablas que se usan tienen la misma estructura.</span><span class="sxs-lookup"><span data-stu-id="e2747-105">It assumes that all tables being used have the same structure.</span></span>

<span data-ttu-id="e2747-106">Hay dos variaciones de este script:</span><span class="sxs-lookup"><span data-stu-id="e2747-106">There are two variations of this script:</span></span>

1. <span data-ttu-id="e2747-107">El [primer script](#sample-code-combine-data-from-multiple-excel-tables-into-a-single-table) combina todas las tablas del Excel archivo.</span><span class="sxs-lookup"><span data-stu-id="e2747-107">The [first script](#sample-code-combine-data-from-multiple-excel-tables-into-a-single-table) combines all tables in the Excel file.</span></span>
1. <span data-ttu-id="e2747-108">El [segundo script obtiene](#sample-code-combine-data-from-multiple-excel-tables-in-select-worksheets-into-a-single-table) selectivamente tablas dentro de un conjunto de hojas de cálculo.</span><span class="sxs-lookup"><span data-stu-id="e2747-108">The [second script](#sample-code-combine-data-from-multiple-excel-tables-in-select-worksheets-into-a-single-table) selectively gets tables within a set of worksheets.</span></span>

## <a name="sample-code-combine-data-from-multiple-excel-tables-into-a-single-table"></a><span data-ttu-id="e2747-109">Código de ejemplo: combinar datos de varias Excel tablas en una sola tabla</span><span class="sxs-lookup"><span data-stu-id="e2747-109">Sample code: Combine data from multiple Excel tables into a single table</span></span>

<span data-ttu-id="e2747-110">Descargue el archivo de <a href="tables-copy.xlsx">tables-copy.xlsx</a> y ústelo con el siguiente script para probarlo usted mismo.</span><span class="sxs-lookup"><span data-stu-id="e2747-110">Download the sample file <a href="tables-copy.xlsx">tables-copy.xlsx</a> and use it with the following script to try it out yourself!</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    workbook.getWorksheet('Combined')?.delete();
    const newSheet = workbook.addWorksheet('Combined');
    
    const tables = workbook.getTables();    
    const headerValues = tables[0].getHeaderRowRange().getTexts();
    console.log(headerValues);
    const targetRange = updateRange(newSheet, headerValues);
    const combinedTable = newSheet.addTable(targetRange.getAddress(), true);
    for (let table of tables) {      
      let dataValues = table.getRangeBetweenHeaderAndTotal().getTexts();
      let rowCount = table.getRowCount();
      if (rowCount > 0) {
        combinedTable.addRows(-1, dataValues);
      }
    }
}

function updateRange(sheet: ExcelScript.Worksheet, data: string[][]): ExcelScript.Range {
  const targetRange = sheet.getRange('A1').getResizedRange(data.length-1, data[0].length-1);
  targetRange.setValues(data);
  return targetRange;
}
```

## <a name="sample-code-combine-data-from-multiple-excel-tables-in-select-worksheets-into-a-single-table"></a><span data-ttu-id="e2747-111">Código de ejemplo: combinar datos de varias tablas Excel en hojas de cálculo selectas en una sola tabla</span><span class="sxs-lookup"><span data-stu-id="e2747-111">Sample code: Combine data from multiple Excel tables in select worksheets into a single table</span></span>

<span data-ttu-id="e2747-112">Descargue el archivo de <a href="tables-select-copy.xlsx">tables-select-copy.xlsx</a> y ústelo con el siguiente script para probarlo usted mismo.</span><span class="sxs-lookup"><span data-stu-id="e2747-112">Download the sample file <a href="tables-select-copy.xlsx">tables-select-copy.xlsx</a> and use it with the following script to try it out yourself!</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    const sheetNames = ['Sheet1', 'Sheet2', 'Sheet3'];
    
    workbook.getWorksheet('Combined')?.delete();
    const newSheet = workbook.addWorksheet('Combined');
    let targetTableCreated = false;
    let combinedTable;
    sheetNames.forEach((sheet) => {
      const tables = workbook.getWorksheet(sheet).getTables();
      if (!targetTableCreated) {
        const headerValues = tables[0].getHeaderRowRange().getTexts();
        const targetRange = updateRange(newSheet, headerValues);
        combinedTable = newSheet.addTable(targetRange.getAddress(), true);
        targetTableCreated = true;
      }      
      for (let table of tables) {
        let dataValues = table.getRangeBetweenHeaderAndTotal().getTexts();
        let rowCount = table.getRowCount();
        if (rowCount > 0) {
        combinedTable.addRows(-1, dataValues);
        }
      }
    })
}

function updateRange(sheet: ExcelScript.Worksheet, data: string[][]): ExcelScript.Range {
  const targetRange = sheet.getRange('A1').getResizedRange(data.length-1, data[0].length-1);
  targetRange.setValues(data);
  return targetRange;
}
```

## <a name="training-video-combine-data-from-multiple-excel-tables-into-a-single-table"></a><span data-ttu-id="e2747-113">Vídeo de aprendizaje: combinar datos de varias Excel tablas en una sola tabla</span><span class="sxs-lookup"><span data-stu-id="e2747-113">Training video: Combine data from multiple Excel tables into a single table</span></span>

<span data-ttu-id="e2747-114">[Vea el recorrido de Sudhi Ramamurthy por este ejemplo en YouTube](https://youtu.be/di-8JukK3Lc).</span><span class="sxs-lookup"><span data-stu-id="e2747-114">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/di-8JukK3Lc).</span></span>
