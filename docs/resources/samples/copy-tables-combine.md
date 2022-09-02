---
title: Combinación de datos de varias tablas de Excel en una sola tabla
description: Aprenda a usar scripts de Office para combinar datos de varias tablas de Excel en una sola tabla.
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: 3db510514c676b9012fd47abc2a7e92492a9cf87
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572454"
---
# <a name="combine-data-from-multiple-excel-tables-into-a-single-table"></a>Combinación de datos de varias tablas de Excel en una sola tabla

Este ejemplo combina datos de varias tablas de Excel en una sola tabla que incluye todas las filas. Se supone que todas las tablas que se usan tienen la misma estructura.

Hay dos variaciones de este script:

1. El [primer script](#sample-code-combine-data-from-multiple-excel-tables-into-a-single-table) combina todas las tablas del archivo de Excel.
1. El [segundo script](#sample-code-combine-data-from-multiple-excel-tables-in-select-worksheets-into-a-single-table) obtiene de forma selectiva las tablas dentro de un conjunto de hojas de cálculo.

## <a name="sample-excel-file"></a>Archivo de Excel de ejemplo

Descargue [tables-copy.xlsx](tables-copy.xlsx) de un libro listo para usar. Agregue los siguientes scripts para probar el ejemplo usted mismo.

## <a name="sample-code-combine-data-from-multiple-excel-tables-into-a-single-table"></a>Código de ejemplo: Combinar datos de varias tablas de Excel en una sola tabla

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

## <a name="sample-code-combine-data-from-multiple-excel-tables-in-select-worksheets-into-a-single-table"></a>Código de ejemplo: Combinar datos de varias tablas de Excel en hojas de cálculo seleccionadas en una sola tabla

Descargue el archivo de ejemplo [tables-select-copy.xlsx](tables-select-copy.xlsx) y úselo con el siguiente script para probarlo usted mismo.

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

## <a name="training-video-combine-data-from-multiple-excel-tables-into-a-single-table"></a>Vídeo de entrenamiento: Combinación de datos de varias tablas de Excel en una sola tabla

[Vea cómo Sudhi Ramamurthy recorre este ejemplo en YouTube](https://youtu.be/di-8JukK3Lc).
