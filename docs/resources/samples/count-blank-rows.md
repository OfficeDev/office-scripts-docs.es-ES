---
title: Contar filas en blanco en hojas
description: Obtenga información sobre cómo usar Office Scripts para detectar si hay filas en blanco en lugar de datos en las hojas de cálculo y, a continuación, informe del recuento de filas en blanco que se usará en un flujo Power Automate.
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: e5b60779d2ca2de5f4cf4e03ddd6ff7372515ad6
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 07/07/2021
ms.locfileid: "53313809"
---
# <a name="count-blank-rows-on-sheets"></a><span data-ttu-id="7f9b7-103">Contar filas en blanco en hojas</span><span class="sxs-lookup"><span data-stu-id="7f9b7-103">Count blank rows on sheets</span></span>

<span data-ttu-id="7f9b7-104">Este proyecto incluye dos scripts:</span><span class="sxs-lookup"><span data-stu-id="7f9b7-104">This project includes two scripts:</span></span>

* <span data-ttu-id="7f9b7-105">[Contar filas en blanco en una hoja determinada:](#sample-code-count-blank-rows-on-a-given-sheet)recorre el intervalo usado en una hoja de cálculo determinada y devuelve un recuento de filas en blanco.</span><span class="sxs-lookup"><span data-stu-id="7f9b7-105">[Count blank rows on a given sheet](#sample-code-count-blank-rows-on-a-given-sheet): Traverses the used range on a given worksheet and returns a blank row count.</span></span>
* <span data-ttu-id="7f9b7-106">[Contar filas en blanco en todas las hojas:](#sample-code-count-blank-rows-on-all-sheets)recorre el intervalo usado en todas las hojas de cálculo _y_ devuelve un recuento de filas en blanco.</span><span class="sxs-lookup"><span data-stu-id="7f9b7-106">[Count blank rows on all sheets](#sample-code-count-blank-rows-on-all-sheets): Traverses the used range on _all of the worksheets_ and returns a blank row count.</span></span>

> [!NOTE]
> <span data-ttu-id="7f9b7-107">Para nuestro script, una fila en blanco es cualquier fila donde no hay datos.</span><span class="sxs-lookup"><span data-stu-id="7f9b7-107">For our script, a blank row is any row where there's no data.</span></span> <span data-ttu-id="7f9b7-108">La fila puede tener formato.</span><span class="sxs-lookup"><span data-stu-id="7f9b7-108">The row can have formatting.</span></span>

<span data-ttu-id="7f9b7-109">_Esta hoja devuelve el recuento de 4 filas en blanco_</span><span class="sxs-lookup"><span data-stu-id="7f9b7-109">_This sheet returns count of 4 blank rows_</span></span>

:::image type="content" source="../../images/blank-rows.png" alt-text="Hoja de cálculo que muestra datos con filas en blanco.":::

<span data-ttu-id="7f9b7-111">_Esta hoja devuelve el recuento de 0 filas en blanco (todas las filas tienen algunos datos)_</span><span class="sxs-lookup"><span data-stu-id="7f9b7-111">_This sheet returns count of 0 blank rows (all rows have some data)_</span></span>

:::image type="content" source="../../images/no-blank-rows.png" alt-text="Hoja de cálculo que muestra datos sin filas en blanco.":::

## <a name="sample-code-count-blank-rows-on-a-given-sheet"></a><span data-ttu-id="7f9b7-113">Código de ejemplo: contar filas en blanco en una hoja determinada</span><span class="sxs-lookup"><span data-stu-id="7f9b7-113">Sample code: Count blank rows on a given sheet</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): number
{
  // Get the worksheet named "Sheet1".
  const sheet = workbook.getWorksheet('Sheet1'); 
  
  // Get the entire data range.
  const range = sheet.getUsedRange(true);

  // If the used range is empty, end the script.
  if (!range) {
    console.log(`No data on this sheet.`);
    return;
  }
  
  // Log the address of the used range.
  console.log(`Used range for the worksheet: ${range.getAddress()}`);
    
  // Look through the values in the range for blank rows.
  const values = range.getValues();
  let emptyRows = 0;
  for (let row of values) {
    let emptyRow = true;
    
    // Look at every cell in the row for one with a value.
    for (let cell of row) {
      if (cell.toString().length > 0) {
        emptyRow = false
      }
    }

    // If no cell had a value, the row is empty.
    if (emptyRow) {
      emptyRows++;
    }
  }

  // Log the number of empty rows.
  console.log(`Total empty rows: ${emptyRows}`);

  // Return the number of empty rows for use in a Power Automate flow.
  return emptyRows;
}
```

## <a name="sample-code-count-blank-rows-on-all-sheets"></a><span data-ttu-id="7f9b7-114">Código de ejemplo: contar filas en blanco en todas las hojas</span><span class="sxs-lookup"><span data-stu-id="7f9b7-114">Sample code: Count blank rows on all sheets</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): number
{
  // Loop through every worksheet in the workbook.
  const sheets = workbook.getWorksheets();
  let emptyRows = 0;
  for (let sheet of sheets) {     
    // Get the entire data range.
    const range = sheet.getUsedRange(true);
  
    // If the used range is empty, skip to the next worksheet.
    if (!range) {
      console.log(`No data on this sheet.`);
      continue;
    }
    
    // Log the address of the used range.
    console.log(`Used range for the worksheet: ${range.getAddress()}`);
      
    // Look through the values in the range for blank rows.
    const values = range.getValues();
    for (let row of values) {
      let emptyRow = true;
      
      // Look at every cell in the row for one with a value.
      for (let cell of row) {
        if (cell.toString().length > 0) {
          emptyRow = false
        }
      }
  
      // If no cell had a value, the row is empty.
      if (emptyRow) {
        emptyRows++;
      }
    }
  }

  // Log the number of empty rows.
  console.log(`Total empty rows: ${emptyRows}`);

  // Return the number of empty rows for use in a Power Automate flow.
  return emptyRows;
}
```
