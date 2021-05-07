---
title: Contar filas en blanco en hojas
description: Obtenga información sobre cómo usar Office Scripts para detectar si hay filas en blanco en lugar de datos en las hojas de cálculo y, a continuación, informe del recuento de filas en blanco que se usará en un flujo Power Automate.
ms.date: 03/31/2021
localization_priority: Normal
ms.openlocfilehash: db84f2446c168f867c325a05129fe982c9645731
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232588"
---
# <a name="count-blank-rows-on-sheets"></a><span data-ttu-id="1ff8d-103">Contar filas en blanco en hojas</span><span class="sxs-lookup"><span data-stu-id="1ff8d-103">Count blank rows on sheets</span></span>

<span data-ttu-id="1ff8d-104">Este proyecto incluye dos scripts:</span><span class="sxs-lookup"><span data-stu-id="1ff8d-104">This project includes two scripts:</span></span>

* <span data-ttu-id="1ff8d-105">[Contar filas en blanco en una hoja determinada:](#sample-code-count-blank-rows-on-a-given-sheet)recorre el intervalo usado en una hoja de cálculo determinada y devuelve un recuento de filas en blanco.</span><span class="sxs-lookup"><span data-stu-id="1ff8d-105">[Count blank rows on a given sheet](#sample-code-count-blank-rows-on-a-given-sheet): Traverses the used range on a given worksheet and returns a blank row count.</span></span>
* <span data-ttu-id="1ff8d-106">[Contar filas en blanco en todas las hojas:](#sample-code-count-blank-rows-on-all-sheets)recorre el intervalo usado en todas las hojas de cálculo _y_ devuelve un recuento de filas en blanco.</span><span class="sxs-lookup"><span data-stu-id="1ff8d-106">[Count blank rows on all sheets](#sample-code-count-blank-rows-on-all-sheets): Traverses the used range on _all of the worksheets_ and returns a blank row count.</span></span>

> [!NOTE]
> <span data-ttu-id="1ff8d-107">Para nuestro script, una fila en blanco es cualquier fila donde no hay datos.</span><span class="sxs-lookup"><span data-stu-id="1ff8d-107">For our script, a blank row is any row where there's no data.</span></span> <span data-ttu-id="1ff8d-108">La fila puede tener formato.</span><span class="sxs-lookup"><span data-stu-id="1ff8d-108">The row can have formatting.</span></span>

<span data-ttu-id="1ff8d-109">_Esta hoja devuelve el recuento de 4 filas en blanco_</span><span class="sxs-lookup"><span data-stu-id="1ff8d-109">_This sheet returns count of 4 blank rows_</span></span>

:::image type="content" source="../../images/blank-rows.png" alt-text="Hoja de cálculo que muestra datos con filas en blanco":::

<span data-ttu-id="1ff8d-111">_Esta hoja devuelve el recuento de 0 filas en blanco (todas las filas tienen algunos datos)_</span><span class="sxs-lookup"><span data-stu-id="1ff8d-111">_This sheet returns count of 0 blank rows (all rows have some data)_</span></span>

:::image type="content" source="../../images/no-blank-rows.png" alt-text="Una hoja de cálculo que muestra datos sin filas en blanco":::

## <a name="sample-code-count-blank-rows-on-a-given-sheet"></a><span data-ttu-id="1ff8d-113">Código de ejemplo: contar filas en blanco en una hoja determinada</span><span class="sxs-lookup"><span data-stu-id="1ff8d-113">Sample code: Count blank rows on a given sheet</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): number
{
  const sheet = workbook.getWorksheet('Sheet1'); 
  // Getting the active worksheet is not suitable for a script used by Power Automate.
  // const sheet = workbook.getActiveWorksheet();
  
  const range = sheet.getUsedRange(true); // Get value only.
  if (!range) {
    console.log(`No data on this sheet. `);
    return;
  }
  console.log(`Used range for the worksheet: ${range.getAddress()}`);
  const values = range.getValues();
  let emptyRows = 0;
  for (let row of values) {
    let len = 0; 
    for (let cell of row) {
      len = len + cell.toString().length;
    }
    if (len === 0) { 
      emptyRows++;
    }
  }
  console.log(`Total empty row: ` + emptyRows);
  return emptyRows;
}
```

## <a name="sample-code-count-blank-rows-on-all-sheets"></a><span data-ttu-id="1ff8d-114">Código de ejemplo: contar filas en blanco en todas las hojas</span><span class="sxs-lookup"><span data-stu-id="1ff8d-114">Sample code: Count blank rows on all sheets</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): number
{
  const sheets = workbook.getWorksheets();
  let emptyRows = 0;
  for (let sheet of sheets) { 
    const range = sheet.getUsedRange(true); // Get value only.
    if (!range) {
      console.log(`No data on this sheet. `);
      continue;
    }
    console.log(`Used range for the worksheet ${sheet.getName()}: ${range.getAddress()}`);
    const values = range.getValues();

    for (let row of values) {
      let len = 0;
      for (let cell of row) {
        len = len + cell.toString().length;
      }
      if (len === 0) {
        emptyRows++;
      }
    }
  }
  console.log(`Total empty row: ` + emptyRows);
  return emptyRows;
}
```

## <a name="use-with-power-automate"></a><span data-ttu-id="1ff8d-115">Usar con Power Automate</span><span class="sxs-lookup"><span data-stu-id="1ff8d-115">Use with Power Automate</span></span>

:::image type="content" source="../../images/use-in-power-automate.png" alt-text="Flujo de Power Automate que muestra cómo configurar para ejecutar un script Office script":::
