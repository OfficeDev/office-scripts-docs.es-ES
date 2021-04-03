---
title: Contar filas en blanco en hojas
description: Obtenga información sobre cómo usar scripts de Office para detectar si hay filas en blanco en lugar de datos en las hojas de cálculo y, a continuación, informe del recuento de filas en blanco que se usará en un flujo de Power Automate.
ms.date: 03/31/2021
localization_priority: Normal
ms.openlocfilehash: 1f52b9c4d538d5d3e64dc61dae3e27d046b56862
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571584"
---
# <a name="count-blank-rows-on-sheets"></a><span data-ttu-id="a0263-103">Contar filas en blanco en hojas</span><span class="sxs-lookup"><span data-stu-id="a0263-103">Count blank rows on sheets</span></span>

<span data-ttu-id="a0263-104">Este proyecto incluye dos scripts:</span><span class="sxs-lookup"><span data-stu-id="a0263-104">This project includes two scripts:</span></span>

* <span data-ttu-id="a0263-105">[Contar filas en blanco en una hoja determinada:](#sample-code-count-blank-rows-on-a-given-sheet)recorre el intervalo usado en una hoja de cálculo determinada y devuelve un recuento de filas en blanco.</span><span class="sxs-lookup"><span data-stu-id="a0263-105">[Count blank rows on a given sheet](#sample-code-count-blank-rows-on-a-given-sheet): Traverses the used range on a given worksheet and returns a blank row count.</span></span>
* <span data-ttu-id="a0263-106">[Contar filas en blanco en todas las hojas:](#sample-code-count-blank-rows-on-all-sheets)recorre el intervalo usado en todas las hojas de cálculo _y_ devuelve un recuento de filas en blanco.</span><span class="sxs-lookup"><span data-stu-id="a0263-106">[Count blank rows on all sheets](#sample-code-count-blank-rows-on-all-sheets): Traverses the used range on _all of the worksheets_ and returns a blank row count.</span></span>

> [!NOTE]
> <span data-ttu-id="a0263-107">Para nuestro script, una fila en blanco es cualquier fila donde no hay datos.</span><span class="sxs-lookup"><span data-stu-id="a0263-107">For our script, a blank row is any row where there's no data.</span></span> <span data-ttu-id="a0263-108">La fila puede tener formato.</span><span class="sxs-lookup"><span data-stu-id="a0263-108">The row can have formatting.</span></span>

<span data-ttu-id="a0263-109">_Esta hoja devuelve el recuento de 4 filas en blanco_</span><span class="sxs-lookup"><span data-stu-id="a0263-109">_This sheet returns count of 4 blank rows_</span></span>

![Datos con filas en blanco](../../images/blank-rows.png)

<span data-ttu-id="a0263-111">_Esta hoja devuelve el recuento de 0 filas en blanco (todas las filas tienen algunos datos)_</span><span class="sxs-lookup"><span data-stu-id="a0263-111">_This sheet returns count of 0 blank rows (all rows have some data)_</span></span>

![Datos sin filas en blanco](../../images/no-blank-rows.png)

## <a name="sample-code-count-blank-rows-on-a-given-sheet"></a><span data-ttu-id="a0263-113">Código de ejemplo: contar filas en blanco en una hoja determinada</span><span class="sxs-lookup"><span data-stu-id="a0263-113">Sample code: Count blank rows on a given sheet</span></span>

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

## <a name="sample-code-count-blank-rows-on-all-sheets"></a><span data-ttu-id="a0263-114">Código de ejemplo: contar filas en blanco en todas las hojas</span><span class="sxs-lookup"><span data-stu-id="a0263-114">Sample code: Count blank rows on all sheets</span></span>

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

## <a name="use-with-power-automate"></a><span data-ttu-id="a0263-115">Usar con Power Automate</span><span class="sxs-lookup"><span data-stu-id="a0263-115">Use with Power Automate</span></span>

![Captura de pantalla que muestra cómo configurar en Power Automate](../../images/use-in-power-automate.png)