---
title: Salida de datos de Excel como JSON
description: Obtenga información sobre cómo generar datos de tabla de Excel como JSON para usarlos en Power Automate.
ms.date: 03/18/2021
localization_priority: Normal
ms.openlocfilehash: 678506fee0b6a41ede8245fb360d485d635e2d64
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571594"
---
# <a name="output-excel-table-data-as-json-for-usage-in-power-automate"></a><span data-ttu-id="b0736-103">Salida de datos de tabla de Excel como JSON para su uso en Power Automate</span><span class="sxs-lookup"><span data-stu-id="b0736-103">Output Excel table data as JSON for usage in Power Automate</span></span>

<span data-ttu-id="b0736-104">Los datos de tabla de Excel se pueden representar como una matriz de objetos en forma de JSON.</span><span class="sxs-lookup"><span data-stu-id="b0736-104">Excel table data can be represented as an array of objects in the form of JSON.</span></span> <span data-ttu-id="b0736-105">Cada objeto representa una fila de la tabla.</span><span class="sxs-lookup"><span data-stu-id="b0736-105">Each object represents a row in the table.</span></span> <span data-ttu-id="b0736-106">Esto ayuda a extraer los datos de Excel en un formato coherente que sea visible para el usuario.</span><span class="sxs-lookup"><span data-stu-id="b0736-106">This helps extract the data from Excel in a consistent format that is visible to the user.</span></span> <span data-ttu-id="b0736-107">A continuación, los datos se pueden dar a otros sistemas a través de flujos de Power Automate.</span><span class="sxs-lookup"><span data-stu-id="b0736-107">The data can then be given to other systems through Power Automate flows.</span></span>

<span data-ttu-id="b0736-108">_Datos de tabla de entrada_</span><span class="sxs-lookup"><span data-stu-id="b0736-108">_Input table data_</span></span>

![Captura de pantalla que muestra los datos de la tabla de entrada](../../images/table-input.png)

<span data-ttu-id="b0736-110">Una variación de este ejemplo también incluye los hipervínculos en una de las columnas de la tabla.</span><span class="sxs-lookup"><span data-stu-id="b0736-110">A variation of this sample also includes the hyperlinks in one of the table columns.</span></span> <span data-ttu-id="b0736-111">Esto permite que se presenten niveles adicionales de datos de celda en el JSON.</span><span class="sxs-lookup"><span data-stu-id="b0736-111">This allows additional levels of cell data to be surfaced in the JSON.</span></span>

<span data-ttu-id="b0736-112">_Datos de tabla de entrada que incluyen hipervínculos_</span><span class="sxs-lookup"><span data-stu-id="b0736-112">_Input table data that includes hyperlinks_</span></span>

![Captura de pantalla que muestra datos de tabla que incluye hipervínculos](../../images/table-hyperlink-view.png)

<span data-ttu-id="b0736-114">_Cuadro de diálogo para editar hipervínculo_</span><span class="sxs-lookup"><span data-stu-id="b0736-114">_Dialog to edit hyperlink_</span></span>

![Captura de pantalla que muestra el cuadro de diálogo para editar hipervínculo](../../images/table-hyperlink-edit.png)

## <a name="sample-excel-file"></a><span data-ttu-id="b0736-116">Archivo de Excel de ejemplo</span><span class="sxs-lookup"><span data-stu-id="b0736-116">Sample Excel file</span></span>

<span data-ttu-id="b0736-117">Descarga el archivo <a href="table-data-with-hyperlinks.xlsx">table-data-with-hyperlinks.xlsx</a> usado en estos ejemplos y pruébalo tú mismo.</span><span class="sxs-lookup"><span data-stu-id="b0736-117">Download the file <a href="table-data-with-hyperlinks.xlsx">table-data-with-hyperlinks.xlsx</a> used in these samples and try it out yourself!</span></span>

## <a name="sample-code-return-table-data-as-json"></a><span data-ttu-id="b0736-118">Código de ejemplo: devolver datos de tabla como JSON</span><span class="sxs-lookup"><span data-stu-id="b0736-118">Sample code: Return table data as JSON</span></span>

> [!NOTE]
> <span data-ttu-id="b0736-119">Puede cambiar la estructura `interface TableData` para que coincida con las columnas de la tabla.</span><span class="sxs-lookup"><span data-stu-id="b0736-119">You can change the `interface TableData` structure to match your table columns.</span></span> <span data-ttu-id="b0736-120">Tenga en cuenta que para los nombres de columna con espacios, asegúrese de colocar la clave entre comillas, como con `"Event ID"` en el ejemplo.</span><span class="sxs-lookup"><span data-stu-id="b0736-120">Note that for column names with spaces, be sure to place your key in quotation marks, such as with `"Event ID"` in the sample.</span></span>

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

### <a name="sample-output"></a><span data-ttu-id="b0736-121">Salida de ejemplo</span><span class="sxs-lookup"><span data-stu-id="b0736-121">Sample output</span></span>

```json
[{
    "Event ID": "E107",
    "Date": "2020-12-10",
    "Location": "Montgomery",
    "Capacity": "10",
    "Speakers": "Debra Berger"
}, {
    "Event ID": "E108",
    "Date": "2020-12-11",
    "Location": "Montgomery",
    "Capacity": "10",
    "Speakers": "Delia Dennis"
}, {
    "Event ID": "E109",
    "Date": "2020-12-12",
    "Location": "Montgomery",
    "Capacity": "10",
    "Speakers": "Diego Siciliani"
}, {
    "Event ID": "E110",
    "Date": "2020-12-13",
    "Location": "Boise",
    "Capacity": "25",
    "Speakers": "Gerhart Moller"
}, {
    "Event ID": "E111",
    "Date": "2020-12-14",
    "Location": "Salt Lake City",
    "Capacity": "20",
    "Speakers": "Grady Archie"
}, {
    "Event ID": "E112",
    "Date": "2020-12-15",
    "Location": "Fremont",
    "Capacity": "25",
    "Speakers": "Irvin Sayers"
}, {
    "Event ID": "E113",
    "Date": "2020-12-16",
    "Location": "Salt Lake City",
    "Capacity": "20",
    "Speakers": "Isaiah Langer"
}, {
    "Event ID": "E114",
    "Date": "2020-12-17",
    "Location": "Salt Lake City",
    "Capacity": "20",
    "Speakers": "Johanna Lorenz"
}]
```

## <a name="sample-code-return-table-data-as-json-with-hyperlink-text"></a><span data-ttu-id="b0736-122">Código de ejemplo: devolver datos de tabla como JSON con texto de hipervínculo</span><span class="sxs-lookup"><span data-stu-id="b0736-122">Sample code: Return table data as JSON with hyperlink text</span></span>

> [!NOTE]
> <span data-ttu-id="b0736-123">El script siempre extrae hipervínculos de la 4ª columna (índice 0) de la tabla.</span><span class="sxs-lookup"><span data-stu-id="b0736-123">The script always extracts hyperlinks from the 4th column (0 index) of the table.</span></span> <span data-ttu-id="b0736-124">Puede cambiar ese orden o incluir varias columnas como datos de hipervínculo modificando el código en el comentario `// For the 4th column (0 index), extract the hyperlink and use that instead of text.`</span><span class="sxs-lookup"><span data-stu-id="b0736-124">You can change that order or include multiple columns as hyperlink data by modifying the code under the comment `// For the 4th column (0 index), extract the hyperlink and use that instead of text.`</span></span>

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

### <a name="sample-output"></a><span data-ttu-id="b0736-125">Salida de ejemplo</span><span class="sxs-lookup"><span data-stu-id="b0736-125">Sample output</span></span>

```json
[{
    "Event ID": "E107",
    "Date": "2020-12-10",
    "Location": "Montgomery",
    "Capacity": "10",
    "Search link": "https://www.google.com/search?q=Montgomery",
    "Speakers": "Debra Berger"
}, {
    "Event ID": "E108",
    "Date": "2020-12-11",
    "Location": "Montgomery",
    "Capacity": "10",
    "Search link": "https://www.google.com/search?q=Montgomery",
    "Speakers": "Delia Dennis"
}, {
    "Event ID": "E109",
    "Date": "2020-12-12",
    "Location": "Montgomery",
    "Capacity": "10",
    "Search link": "https://www.google.com/search?q=Montgomery",
    "Speakers": "Diego Siciliani"
}, {
    "Event ID": "E110",
    "Date": "2020-12-13",
    "Location": "Boise",
    "Capacity": "25",
    "Search link": "https://www.google.com/search?q=Boise",
    "Speakers": "Gerhart Moller"
}, {
    "Event ID": "E111",
    "Date": "2020-12-14",
    "Location": "Salt Lake City",
    "Capacity": "20",
    "Search link": "https://www.google.com/search?q=salt+lake+city",
    "Speakers": "Grady Archie"
}, {
    "Event ID": "E112",
    "Date": "2020-12-15",
    "Location": "Fremont",
    "Capacity": "25",
    "Search link": "https://www.google.com/search?q=Fremont",
    "Speakers": "Irvin Sayers"
}, {
    "Event ID": "E113",
    "Date": "2020-12-16",
    "Location": "Salt Lake City",
    "Capacity": "20",
    "Search link": "https://www.google.com/search?q=salt+lake+city",
    "Speakers": "Isaiah Langer"
}, {
    "Event ID": "E114",
    "Date": "2020-12-17",
    "Location": "Salt Lake City",
    "Capacity": "20",
    "Search link": "https://www.google.com/search?q=salt+lake+city",
    "Speakers": "Johanna Lorenz"
}]
```

## <a name="use-in-power-automate"></a><span data-ttu-id="b0736-126">Uso en Power Automate</span><span class="sxs-lookup"><span data-stu-id="b0736-126">Use in Power Automate</span></span>

<span data-ttu-id="b0736-127">Para obtener información sobre cómo usar este script en Power Automate, consulte [Create an automated workflow with Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).</span><span class="sxs-lookup"><span data-stu-id="b0736-127">For how to use such a script in Power Automate, see [Create an automated workflow with Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).</span></span>
