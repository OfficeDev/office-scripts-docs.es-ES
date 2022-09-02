---
title: Quitar hipervínculos de cada celda de una hoja de cálculo de Excel
description: Obtenga información sobre cómo usar scripts de Office para quitar hipervínculos de cada celda de una hoja de cálculo de Excel.
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: 1445988b1e6a85fcab8914ffeaaef80a07a52f5e
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572629"
---
# <a name="remove-hyperlinks-from-each-cell-in-an-excel-worksheet"></a>Quitar hipervínculos de cada celda de una hoja de cálculo de Excel

 En este ejemplo se borran todos los hipervínculos de la hoja de cálculo actual. Recorre la hoja de cálculo y, si hay algún hipervínculo asociado a la celda, borra el hipervínculo, pero conserva el valor de celda tal cual. También registra el tiempo necesario para completar el recorrido.

> [!NOTE]
> Esto solo funciona si el recuento de celdas es < 10 000.

## <a name="sample-excel-file"></a>Archivo de Excel de ejemplo

Descargue el [ archivoremove-hyperlinks.xlsx](remove-hyperlinks.xlsx) de un libro listo para usar. Agregue el siguiente script para probar el ejemplo usted mismo.

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

## <a name="training-video-remove-hyperlinks-from-each-cell-in-an-excel-worksheet"></a>Vídeo de entrenamiento: Eliminación de hipervínculos de cada celda de una hoja de cálculo de Excel

[Vea cómo Sudhi Ramamurthy recorre este ejemplo en YouTube](https://youtu.be/v20fdinxpHU).
