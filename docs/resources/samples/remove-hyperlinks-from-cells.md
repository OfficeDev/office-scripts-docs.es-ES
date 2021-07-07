---
title: Quitar hipervínculos de cada celda de una hoja Excel hoja de cálculo
description: Obtenga información sobre cómo usar Office scripts para quitar hipervínculos de cada celda de una hoja Excel trabajo.
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: dc33eb639edac8ada29824a53440031942e59179
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 07/07/2021
ms.locfileid: "53313753"
---
# <a name="remove-hyperlinks-from-each-cell-in-an-excel-worksheet"></a>Quitar hipervínculos de cada celda de una hoja Excel hoja de cálculo

 En este ejemplo se borran todos los hipervínculos de la hoja de cálculo actual. Recorre la hoja de cálculo y, si hay algún hipervínculo asociado a la celda, borra el hipervínculo pero conserva el valor de la celda tal como está. También registra el tiempo necesario para completar el recorrido.

> [!NOTE]
> Esto solo funciona si el recuento de celdas < 10k.

## <a name="sample-excel-file"></a>Archivo Excel ejemplo

Descargue el archivo <a href="remove-hyperlinks.xlsx">remove-hyperlinks.xlsx</a> para un libro listo para usar. Agregue el siguiente script para probar el ejemplo usted mismo.

## <a name="sample-code-remove-hyperlinks"></a>Código de ejemplo: Quitar hipervínculos

```TypeScript
function main(workbook: ExcelScript.Workbook, sheetName: string = 'Sheet1') {
  // Get the active worksheet. 
  let sheet = workbook.getWorksheet(sheetName);

  // Get the used range to operate on.
  // For large ranges (over 10000 entries), consider splitting the operation into batches for performance.
  const targetRange = sheet.getUsedRange(true);
  console.log(`Target Range to clear hyperlinks from: ${targetRange.getAddress()}`);

  const rowCount = targetRange.getRowCount();
  const colCount = targetRange.getColumnCount();
  console.log(`Searching for hyperlinks in ${targetRange.getAddress()} which contains ${(rowCount * colCount)} cells`);

  // Go through each individual cell looking for a hyperlink. 
  // This allows us to limit the formatting changes to only the cells with hyperlink formatting.
  let clearedCount = 0;
  for (let i = 0; i < rowCount; i++) {
    for (let j = 0; j < colCount; j++) {
      const cell = targetRange.getCell(i, j);
      const hyperlink = cell.getHyperlink();
      if (hyperlink) {
        cell.clear(ExcelScript.ClearApplyTo.hyperlinks);
        cell.getFormat().getFont().setUnderline(ExcelScript.RangeUnderlineStyle.none);
        cell.getFormat().getFont().setColor('Black');
        clearedCount++;
      }
    }
  }

  console.log(`Done. Cleared hyperlinks from ${clearedCount} cells`);
}
```

## <a name="training-video-remove-hyperlinks-from-each-cell-in-an-excel-worksheet"></a>Vídeo de aprendizaje: quitar hipervínculos de cada celda de una hoja Excel trabajo

[Vea el recorrido de Sudhi Ramamurthy por este ejemplo en YouTube](https://youtu.be/v20fdinxpHU).
