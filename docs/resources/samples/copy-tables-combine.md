---
title: Combinar datos de varias Excel tablas en una sola tabla
description: Obtenga información sobre cómo usar Office scripts para combinar datos de varias Excel en una sola tabla.
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: d178da87820efb368968492b1af66a2afd80393f
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 07/07/2021
ms.locfileid: "53313802"
---
# <a name="combine-data-from-multiple-excel-tables-into-a-single-table"></a><span data-ttu-id="f5859-103">Combinar datos de varias Excel tablas en una sola tabla</span><span class="sxs-lookup"><span data-stu-id="f5859-103">Combine data from multiple Excel tables into a single table</span></span>

<span data-ttu-id="f5859-104">En este ejemplo se combinan los datos de varias Excel tablas en una sola tabla que incluye todas las filas.</span><span class="sxs-lookup"><span data-stu-id="f5859-104">This sample combines data from multiple Excel tables into a single table that includes all the rows.</span></span> <span data-ttu-id="f5859-105">Se supone que todas las tablas que se usan tienen la misma estructura.</span><span class="sxs-lookup"><span data-stu-id="f5859-105">It assumes that all tables being used have the same structure.</span></span>

<span data-ttu-id="f5859-106">Hay dos variaciones de este script:</span><span class="sxs-lookup"><span data-stu-id="f5859-106">There are two variations of this script:</span></span>

1. <span data-ttu-id="f5859-107">El [primer script](#sample-code-combine-data-from-multiple-excel-tables-into-a-single-table) combina todas las tablas del Excel archivo.</span><span class="sxs-lookup"><span data-stu-id="f5859-107">The [first script](#sample-code-combine-data-from-multiple-excel-tables-into-a-single-table) combines all tables in the Excel file.</span></span>
1. <span data-ttu-id="f5859-108">El [segundo script obtiene](#sample-code-combine-data-from-multiple-excel-tables-in-select-worksheets-into-a-single-table) selectivamente tablas dentro de un conjunto de hojas de cálculo.</span><span class="sxs-lookup"><span data-stu-id="f5859-108">The [second script](#sample-code-combine-data-from-multiple-excel-tables-in-select-worksheets-into-a-single-table) selectively gets tables within a set of worksheets.</span></span>

## <a name="sample-excel-file"></a><span data-ttu-id="f5859-109">Archivo Excel ejemplo</span><span class="sxs-lookup"><span data-stu-id="f5859-109">Sample Excel file</span></span>

<span data-ttu-id="f5859-110">Descargue <a href="tables-copy.xlsx">tables-copy.xlsx</a> para un libro listo para usar.</span><span class="sxs-lookup"><span data-stu-id="f5859-110">Download <a href="tables-copy.xlsx">tables-copy.xlsx</a> for a ready-to-use workbook.</span></span> <span data-ttu-id="f5859-111">Agregue los siguientes scripts para probar el ejemplo usted mismo.</span><span class="sxs-lookup"><span data-stu-id="f5859-111">Add the following scripts to try the sample yourself!</span></span>

## <a name="sample-code-combine-data-from-multiple-excel-tables-into-a-single-table"></a><span data-ttu-id="f5859-112">Código de ejemplo: combinar datos de varias Excel tablas en una sola tabla</span><span class="sxs-lookup"><span data-stu-id="f5859-112">Sample code: Combine data from multiple Excel tables into a single table</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Delete the "Combined" worksheet, if it's present.
  workbook.getWorksheet('Combined')?.delete();

  // Create a new worksheet named "Combined" for the combined table.
  const newSheet = workbook.addWorksheet('Combined');
  
  // Get the header values for the first table in the workbook.
  // This also saves the table list before we add the new, combined table.
  const tables = workbook.getTables();    
  const headerValues = tables[0].getHeaderRowRange().getTexts();
  console.log(headerValues);

  // Copy the headers on a new worksheet to an equal-sized range.
  const targetRange = newSheet.getRange('A1').getResizedRange(headerValues.length-1, headerValues[0].length-1);
  targetRange.setValues(headerValues);

  // Add the data from each table in the workbook to the new table.
  const combinedTable = newSheet.addTable(targetRange.getAddress(), true);
  for (let table of tables) {      
    let dataValues = table.getRangeBetweenHeaderAndTotal().getTexts();
    let rowCount = table.getRowCount();

    // If the table is not empty, add its rows to the combined table.
    if (rowCount > 0) {
      combinedTable.addRows(-1, dataValues);
    }
  }
}
```

## <a name="sample-code-combine-data-from-multiple-excel-tables-in-select-worksheets-into-a-single-table"></a><span data-ttu-id="f5859-113">Código de ejemplo: combinar datos de varias tablas Excel en hojas de cálculo selectas en una sola tabla</span><span class="sxs-lookup"><span data-stu-id="f5859-113">Sample code: Combine data from multiple Excel tables in select worksheets into a single table</span></span>

<span data-ttu-id="f5859-114">Descargue el archivo de <a href="tables-select-copy.xlsx">tables-select-copy.xlsx</a> y ústelo con el siguiente script para probarlo usted mismo.</span><span class="sxs-lookup"><span data-stu-id="f5859-114">Download the sample file <a href="tables-select-copy.xlsx">tables-select-copy.xlsx</a> and use it with the following script to try it out yourself!</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Set the worksheet names to get tables from.
  const sheetNames = ['Sheet1', 'Sheet2', 'Sheet3'];
    
  // Delete the "Combined" worksheet, if it's present.
  workbook.getWorksheet('Combined')?.delete();

  // Create a new worksheet named "Combined" for the combined table.
  const newSheet = workbook.addWorksheet('Combined');

  // Create a new table with the same headers as the other tables.
  const headerValues = workbook.getWorksheet(sheetNames[0]).getTables()[0].getHeaderRowRange().getTexts();
  const targetRange = newSheet.getRange('A1').getResizedRange(headerValues.length-1, headerValues[0].length-1);
  targetRange.setValues(headerValues);
  const combinedTable = newSheet.addTable(targetRange.getAddress(), true);

  // Go through each listed worksheet and get their tables.
  sheetNames.forEach((sheet) => {
    const tables = workbook.getWorksheet(sheet).getTables();     
    for (let table of tables) {
      // Get the rows from the tables.
      let dataValues = table.getRangeBetweenHeaderAndTotal().getTexts();
      let rowCount = table.getRowCount();

      // If there's data in the table, add it to the combined table.
      if (rowCount > 0) {
          combinedTable.addRows(-1, dataValues);
      }
    }
  });
}
```

## <a name="training-video-combine-data-from-multiple-excel-tables-into-a-single-table"></a><span data-ttu-id="f5859-115">Vídeo de aprendizaje: combinar datos de varias Excel tablas en una sola tabla</span><span class="sxs-lookup"><span data-stu-id="f5859-115">Training video: Combine data from multiple Excel tables into a single table</span></span>

<span data-ttu-id="f5859-116">[Vea el recorrido de Sudhi Ramamurthy por este ejemplo en YouTube](https://youtu.be/di-8JukK3Lc).</span><span class="sxs-lookup"><span data-stu-id="f5859-116">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/di-8JukK3Lc).</span></span>
