---
title: Datos Excel salida como JSON
description: Obtenga información sobre cómo generar Excel datos de tabla como JSON para usarlos en Power Automate.
ms.date: 03/18/2021
localization_priority: Normal
ms.openlocfilehash: c6b033a68fdbde2b053f65d1a54db58da6c93b2e
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232539"
---
# <a name="output-excel-table-data-as-json-for-usage-in-power-automate"></a>Salida Excel datos de tabla como JSON para su uso en Power Automate

Excel datos de tabla se pueden representar como una matriz de objetos en forma de JSON. Cada objeto representa una fila de la tabla. Esto ayuda a extraer los datos Excel en un formato coherente que sea visible para el usuario. A continuación, los datos se pueden dar a otros sistemas a través Power Automate flujos.

_Datos de tabla de entrada_

:::image type="content" source="../../images/table-input.png" alt-text="Una hoja de cálculo que muestra los datos de la tabla de entrada":::

Una variación de este ejemplo también incluye los hipervínculos en una de las columnas de la tabla. Esto permite que se presenten niveles adicionales de datos de celda en el JSON.

_Datos de tabla de entrada que incluyen hipervínculos_

:::image type="content" source="../../images/table-hyperlink-view.png" alt-text="Hoja de cálculo que muestra una columna de datos de tabla con formato de hipervínculos":::

_Cuadro de diálogo para editar hipervínculo_

:::image type="content" source="../../images/table-hyperlink-edit.png" alt-text="Cuadro de diálogo Editar hipervínculo que muestra opciones para cambiar hipervínculos":::

## <a name="sample-excel-file"></a>Archivo Excel ejemplo

Descarga el archivo <a href="table-data-with-hyperlinks.xlsx">table-data-with-hyperlinks.xlsx</a> usado en estos ejemplos y pruébalo tú mismo.

## <a name="sample-code-return-table-data-as-json"></a>Código de ejemplo: devolver datos de tabla como JSON

> [!NOTE]
> Puede cambiar la estructura `interface TableData` para que coincida con las columnas de la tabla. Tenga en cuenta que para los nombres de columna con espacios, asegúrese de colocar la clave entre comillas, como con `"Event ID"` en el ejemplo.

```TypeScript
function main(workbook: ExcelScript.Workbook): TableData[] {
  const table = workbook.getWorksheet('PlainTable').getTables()[0];
  // If you know the table name, you can also do the following:
  // const table = workbook.getTable('Table13436');
  const texts = table.getRange().getTexts();
  let returnObjects: TableData[] = [];
  if (table.getRowCount() > 0)  {
    returnObjects = returnObjectFromValues(texts);
  } 
  console.log(JSON.stringify(returnObjects));  
  return returnObjects
}

function returnObjectFromValues(values: string[][]): TableData[] {
  let objArray = [];
  let objKeys: string[] = [];
  for (let i = 0; i < values.length; i++) {
    if (i === 0) {
      objKeys = values[i]
      continue;
    }
    let obj = {}
    for (let j = 0; j < values[i].length; j++) {
      obj[objKeys[j]] = values[i][j]
    }
    objArray.push(obj);
  }
  return objArray as TableData[];
}

interface BasicObj {
  [key: string]: string
}

interface TableData {
  "Event ID": string
  Date: string
  Location: string
  Capacity: string
  Speakers: string
}
```

### <a name="sample-output"></a>Salida de ejemplo

```json
[{
    "Event ID": "E107",
    "Date": "2020-12-10",
    "Location": "Montgomery",
    "Capacity": "10",
    "Speakers&quot;: &quot;Debra Berger"
}, {
    "Event ID": "E108",
    "Date": "2020-12-11",
    "Location": "Montgomery",
    "Capacity": "10",
    "Speakers&quot;: &quot;Delia Dennis"
}, {
    "Event ID": "E109",
    "Date": "2020-12-12",
    "Location": "Montgomery",
    "Capacity": "10",
    "Speakers&quot;: &quot;Diego Siciliani"
}, {
    "Event ID": "E110",
    "Date": "2020-12-13",
    "Location": "Boise",
    "Capacity": "25",
    "Speakers&quot;: &quot;Gerhart Moller"
}, {
    "Event ID": "E111",
    "Date": "2020-12-14",
    "Location": "Salt Lake City",
    "Capacity": "20",
    "Speakers&quot;: &quot;Grady Archie"
}, {
    "Event ID": "E112",
    "Date": "2020-12-15",
    "Location": "Fremont",
    "Capacity": "25",
    "Speakers&quot;: &quot;Irvin Sayers"
}, {
    "Event ID": "E113",
    "Date": "2020-12-16",
    "Location": "Salt Lake City",
    "Capacity": "20",
    "Speakers&quot;: &quot;Isaiah Langer"
}, {
    "Event ID": "E114",
    "Date": "2020-12-17",
    "Location": "Salt Lake City",
    "Capacity": "20",
    "Speakers&quot;: &quot;Johanna Lorenz"
}]
```

## <a name="sample-code-return-table-data-as-json-with-hyperlink-text"></a>Código de ejemplo: devolver datos de tabla como JSON con texto de hipervínculo

> [!NOTE]
> El script siempre extrae hipervínculos de la 4ª columna (índice 0) de la tabla. Puede cambiar ese orden o incluir varias columnas como datos de hipervínculo modificando el código en el comentario `// For the 4th column (0 index), extract the hyperlink and use that instead of text.`

```TypeScript
function main(workbook: ExcelScript.Workbook): TableData[] {
  const table = workbook.getWorksheet('WithHyperLink').getTables()[0];
  const range = table.getRange();
  // If you know the table name, you can also do the following:
  // const table = workbook.getTable('Table13436');
  const texts = table.getRange().getTexts();
  let returnObjects: TableData[] = [];
  if (table.getRowCount() > 0)  {
    returnObjects = returnObjectFromValues(texts, range);
  } 
  console.log(JSON.stringify(returnObjects));  
  return returnObjects
}

function returnObjectFromValues(values: string[][], range: ExcelScript.Range): TableData[] {
  let objArray = [];
  let objKeys: string[] = [];
  for (let i = 0; i < values.length; i++) {
    if (i === 0) {
      objKeys = values[i]
      continue;
    }
    let obj = {}
    for (let j = 0; j < values[i].length; j++) {
      // For the 4th column (0 index), extract the hyperlink and use that instead of text. 
      if (j === 4) {
        obj[objKeys[j]] = range.getCell(i, j).getHyperlink().address;
      } else {
        obj[objKeys[j]] = values[i][j];
      }
    }
    objArray.push(obj);
  }
  return objArray as TableData[];
}

interface BasicObj {
  [key: string]: string
}

interface TableData {
  "Event ID": string
  Date: string
  Location: string
  Capacity: string
  "Search link": string
  Speakers: string
}
```

### <a name="sample-output"></a>Salida de ejemplo

```json
[{
    "Event ID": "E107",
    "Date": "2020-12-10",
    "Location": "Montgomery",
    "Capacity": "10",
    "Search link": "https://www.google.com/search?q=Montgomery",
    "Speakers&quot;: &quot;Debra Berger"
}, {
    "Event ID": "E108",
    "Date": "2020-12-11",
    "Location": "Montgomery",
    "Capacity": "10",
    "Search link": "https://www.google.com/search?q=Montgomery",
    "Speakers&quot;: &quot;Delia Dennis"
}, {
    "Event ID": "E109",
    "Date": "2020-12-12",
    "Location": "Montgomery",
    "Capacity": "10",
    "Search link": "https://www.google.com/search?q=Montgomery",
    "Speakers&quot;: &quot;Diego Siciliani"
}, {
    "Event ID": "E110",
    "Date": "2020-12-13",
    "Location": "Boise",
    "Capacity": "25",
    "Search link": "https://www.google.com/search?q=Boise",
    "Speakers&quot;: &quot;Gerhart Moller"
}, {
    "Event ID": "E111",
    "Date": "2020-12-14",
    "Location": "Salt Lake City",
    "Capacity": "20",
    "Search link": "https://www.google.com/search?q=salt+lake+city",
    "Speakers&quot;: &quot;Grady Archie"
}, {
    "Event ID": "E112",
    "Date": "2020-12-15",
    "Location": "Fremont",
    "Capacity": "25",
    "Search link": "https://www.google.com/search?q=Fremont",
    "Speakers&quot;: &quot;Irvin Sayers"
}, {
    "Event ID": "E113",
    "Date": "2020-12-16",
    "Location": "Salt Lake City",
    "Capacity": "20",
    "Search link": "https://www.google.com/search?q=salt+lake+city",
    "Speakers&quot;: &quot;Isaiah Langer"
}, {
    "Event ID": "E114",
    "Date": "2020-12-17",
    "Location": "Salt Lake City",
    "Capacity": "20",
    "Search link": "https://www.google.com/search?q=salt+lake+city",
    "Speakers&quot;: &quot;Johanna Lorenz"
}]
```

## <a name="use-in-power-automate"></a>Usar en Power Automate

Para obtener información sobre cómo usar este script en Power Automate, vea [Create an automated workflow with Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).
