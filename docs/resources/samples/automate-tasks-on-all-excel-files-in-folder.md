---
title: Ejecutar un script en todos los archivos de Excel de una carpeta
description: Obtenga información sobre cómo ejecutar un script en todos los archivos de Excel de una carpeta de OneDrive para la Empresa.
ms.date: 03/31/2021
localization_priority: Normal
ms.openlocfilehash: a11876e8241a069a7c640bbcf2c36b4842d3bd90
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571614"
---
# <a name="run-a-script-on-all-excel-files-in-a-folder"></a><span data-ttu-id="c7754-103">Ejecutar un script en todos los archivos de Excel de una carpeta</span><span class="sxs-lookup"><span data-stu-id="c7754-103">Run a script on all Excel files in a folder</span></span>

<span data-ttu-id="c7754-104">Este proyecto realiza un conjunto de tareas de automatización en todos los archivos situados en una carpeta de OneDrive para la Empresa.</span><span class="sxs-lookup"><span data-stu-id="c7754-104">This project performs a set of automation tasks on all files situated in a folder on OneDrive for Business.</span></span> <span data-ttu-id="c7754-105">También se puede usar en una carpeta de SharePoint.</span><span class="sxs-lookup"><span data-stu-id="c7754-105">It could also be used on a SharePoint folder.</span></span>
<span data-ttu-id="c7754-106">Realiza cálculos en los archivos de Excel, agrega formato e inserta un comentario que @mentions [compañero.](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7)</span><span class="sxs-lookup"><span data-stu-id="c7754-106">It performs calculations on the Excel files, adds formatting, and inserts a comment that [@mentions](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) a colleague.</span></span>

## <a name="sample-code-add-formatting-and-insert-comment"></a><span data-ttu-id="c7754-107">Código de ejemplo: Agregar formato e insertar comentario</span><span class="sxs-lookup"><span data-stu-id="c7754-107">Sample code: Add formatting and insert comment</span></span>

<span data-ttu-id="c7754-108">Descargue el archivo <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/highlight-alert-excel-files.zip?raw=true">highlight-alert-excel-files.zip</a>, extraiga los archivos en una carpeta titulada **Ventas** usada en este ejemplo y pruébalo usted mismo.</span><span class="sxs-lookup"><span data-stu-id="c7754-108">Download the file <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/highlight-alert-excel-files.zip?raw=true">highlight-alert-excel-files.zip</a>, extract the files to a folder titled **Sales** used in this sample, and try it out yourself!</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  let table1 = workbook.getTable("Table1");
  const rowCount = table1.getRowCount();
  if (rowCount === 0) {
    return;
  }
  workbook.getApplication().calculate(ExcelScript.CalculationType.full);

  const amountDueCol = table1.getColumnByName('Amount Due');
  const amountDueValues = amountDueCol.getRangeBetweenHeaderAndTotal().getValues();

  let highestValue = amountDueValues[0][0];
  let row = 0;
  for (let i = 1; i < amountDueValues.length; i++) {
    if (amountDueValues[i][0] > highestValue) {
      highestValue = amountDueValues[i][0];
      row = i;
    }
  }
  // Set fill color to FFFF00 for range in table Table1 cell in row 0 on column "Amount due".
  table1.getColumn("Amount due")
    .getRangeBetweenHeaderAndTotal()
    .getRow(row)
    .getFormat()
    .getFill()
    .setColor("FFFF00");
  let selectedSheet = workbook.getActiveWorksheet();
  // Insert comment at cell InvoiceAmounts!F2.
  workbook.addComment(table1.getColumn("Amount due")
    .getRangeBetweenHeaderAndTotal()
    .getRow(row), {
    mentions: [{
      email: "AdeleV@M365x904181.OnMicrosoft.com",
      id: 0,
      name: "Adele Vance"
    }],
    richContent: "<at id=\"0\">Adele Vance</at> Please review this amount"
  }, ExcelScript.ContentType.mention);
}
```

## <a name="training-video-run-a-script-on-all-excel-files-in-a-folder"></a><span data-ttu-id="c7754-109">Vídeo de aprendizaje: ejecutar un script en todos los archivos de Excel de una carpeta</span><span class="sxs-lookup"><span data-stu-id="c7754-109">Training video: Run a script on all Excel files in a folder</span></span>

<span data-ttu-id="c7754-110">[Vea el vídeo paso](https://youtu.be/xMg711o7k6w) a paso sobre cómo ejecutar un script en todos los archivos de Excel en una carpeta de OneDrive para la Empresa o SharePoint.</span><span class="sxs-lookup"><span data-stu-id="c7754-110">[Watch step-by-step video](https://youtu.be/xMg711o7k6w) on how to run a script on all Excel files in a OneDrive for Business or SharePoint folder.</span></span>
