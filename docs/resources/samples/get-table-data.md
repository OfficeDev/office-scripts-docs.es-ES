---
title: Datos Excel salida como JSON
description: Obtenga información sobre cómo generar Excel datos de tabla como JSON para usarlos en Power Automate.
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: 63379d1323f5e2084f4aa39af3f4b6e5e6d7e7bb
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 07/07/2021
ms.locfileid: "53313949"
---
# <a name="output-excel-table-data-as-json-for-usage-in-power-automate"></a><span data-ttu-id="c1bea-103">Salida Excel datos de tabla como JSON para su uso en Power Automate</span><span class="sxs-lookup"><span data-stu-id="c1bea-103">Output Excel table data as JSON for usage in Power Automate</span></span>

<span data-ttu-id="c1bea-104">Excel datos de tabla se pueden representar como una matriz de objetos en forma de JSON.</span><span class="sxs-lookup"><span data-stu-id="c1bea-104">Excel table data can be represented as an array of objects in the form of JSON.</span></span> <span data-ttu-id="c1bea-105">Cada objeto representa una fila de la tabla.</span><span class="sxs-lookup"><span data-stu-id="c1bea-105">Each object represents a row in the table.</span></span> <span data-ttu-id="c1bea-106">Esto ayuda a extraer los datos Excel en un formato coherente que sea visible para el usuario.</span><span class="sxs-lookup"><span data-stu-id="c1bea-106">This helps extract the data from Excel in a consistent format that is visible to the user.</span></span> <span data-ttu-id="c1bea-107">A continuación, los datos se pueden dar a otros sistemas a través Power Automate flujos.</span><span class="sxs-lookup"><span data-stu-id="c1bea-107">The data can then be given to other systems through Power Automate flows.</span></span>

<span data-ttu-id="c1bea-108">_Datos de tabla de entrada_</span><span class="sxs-lookup"><span data-stu-id="c1bea-108">_Input table data_</span></span>

:::image type="content" source="../../images/table-input.png" alt-text="Hoja de cálculo que muestra los datos de la tabla de entrada.":::

<span data-ttu-id="c1bea-110">Una variación de este ejemplo también incluye los hipervínculos en una de las columnas de la tabla.</span><span class="sxs-lookup"><span data-stu-id="c1bea-110">A variation of this sample also includes the hyperlinks in one of the table columns.</span></span> <span data-ttu-id="c1bea-111">Esto permite que se presenten niveles adicionales de datos de celda en el JSON.</span><span class="sxs-lookup"><span data-stu-id="c1bea-111">This allows additional levels of cell data to be surfaced in the JSON.</span></span>

<span data-ttu-id="c1bea-112">_Datos de tabla de entrada que incluyen hipervínculos_</span><span class="sxs-lookup"><span data-stu-id="c1bea-112">_Input table data that includes hyperlinks_</span></span>

:::image type="content" source="../../images/table-hyperlink-view.png" alt-text="Hoja de cálculo que muestra una columna de datos de tabla con formato de hipervínculos.":::

<span data-ttu-id="c1bea-114">_Cuadro de diálogo para editar hipervínculo_</span><span class="sxs-lookup"><span data-stu-id="c1bea-114">_Dialog to edit hyperlink_</span></span>

:::image type="content" source="../../images/table-hyperlink-edit.png" alt-text="Cuadro de diálogo Editar hipervínculo que muestra opciones para cambiar hipervínculos.":::

## <a name="sample-excel-file"></a><span data-ttu-id="c1bea-116">Archivo Excel ejemplo</span><span class="sxs-lookup"><span data-stu-id="c1bea-116">Sample Excel file</span></span>

<span data-ttu-id="c1bea-117">Descargue el archivo <a href="table-data-with-hyperlinks.xlsx">table-data-with-hyperlinks.xlsx</a> para un libro listo para usar.</span><span class="sxs-lookup"><span data-stu-id="c1bea-117">Download the file <a href="table-data-with-hyperlinks.xlsx">table-data-with-hyperlinks.xlsx</a> for a ready-to-use workbook.</span></span> <span data-ttu-id="c1bea-118">Agregue el siguiente script para probar el ejemplo usted mismo.</span><span class="sxs-lookup"><span data-stu-id="c1bea-118">Add the following script to try the sample yourself!</span></span>

## <a name="sample-code-return-table-data-as-json"></a><span data-ttu-id="c1bea-119">Código de ejemplo: devolver datos de tabla como JSON</span><span class="sxs-lookup"><span data-stu-id="c1bea-119">Sample code: Return table data as JSON</span></span>

> [!NOTE]
> <span data-ttu-id="c1bea-120">Puede cambiar la estructura `interface TableData` para que coincida con las columnas de la tabla.</span><span class="sxs-lookup"><span data-stu-id="c1bea-120">You can change the `interface TableData` structure to match your table columns.</span></span> <span data-ttu-id="c1bea-121">Tenga en cuenta que para los nombres de columna con espacios, asegúrese de colocar la clave entre comillas, como con `"Event ID"` en el ejemplo.</span><span class="sxs-lookup"><span data-stu-id="c1bea-121">Note that for column names with spaces, be sure to place your key in quotation marks, such as with `"Event ID"` in the sample.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): TableData[] {
  // Get the first table in the "PlainTable" worksheet.
  // If you know the table name, use `workbook.getTable('TableName')` instead.
  const table = workbook.getWorksheet('PlainTable').getTables()[0];

  // Get all the values from the table as text.
  const texts = table.getRange().getTexts();

  // Create an array of JSON objects that match the row structure.
  let returnObjects: TableData[] = [];
  if (table.getRowCount() > 0)  {
    returnObjects = returnObjectFromValues(texts);
  }

  // Log the information and return it for a Power Automate flow.
  console.log(JSON.stringify(returnObjects));  
  return returnObjects
}

// This function converts a 2D-array of values into a generic JSON object.
// In this case, we have defined the TableData object, but any similar interface would work.
function returnObjectFromValues(values: string[][]): TableData[] {
  let objectArray = [];
  let objectKeys: string[] = [];
  for (let i = 0; i < values.length; i++) {
    if (i === 0) {
      objectKeys = values[i]
      continue;
    }

    let object = {}
    for (let j = 0; j < values[i].length; j++) {
      object[objectKeys[j]] = values[i][j]
    }

    objectArray.push(object);
  }

  return objectArray as TableData[];
}

interface TableData {
  "Event ID": string
  Date: string
  Location: string
  Capacity: string
  Speakers: string
}
```

### <a name="sample-output-from-the-plaintable-worksheet"></a><span data-ttu-id="c1bea-122">Salida de ejemplo de la hoja de cálculo "PlainTable"</span><span class="sxs-lookup"><span data-stu-id="c1bea-122">Sample output from the "PlainTable" worksheet</span></span>

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

## <a name="sample-code-return-table-data-as-json-with-hyperlink-text"></a><span data-ttu-id="c1bea-123">Código de ejemplo: devolver datos de tabla como JSON con texto de hipervínculo</span><span class="sxs-lookup"><span data-stu-id="c1bea-123">Sample code: Return table data as JSON with hyperlink text</span></span>

> [!NOTE]
> <span data-ttu-id="c1bea-124">El script siempre extrae hipervínculos de la 4ª columna (índice 0) de la tabla.</span><span class="sxs-lookup"><span data-stu-id="c1bea-124">The script always extracts hyperlinks from the 4th column (0 index) of the table.</span></span> <span data-ttu-id="c1bea-125">Puede cambiar ese orden o incluir varias columnas como datos de hipervínculo modificando el código en el comentario `// For the 4th column (0 index), extract the hyperlink and use that instead of text.`</span><span class="sxs-lookup"><span data-stu-id="c1bea-125">You can change that order or include multiple columns as hyperlink data by modifying the code under the comment `// For the 4th column (0 index), extract the hyperlink and use that instead of text.`</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): TableData[] {
  // Get the first table in the "WithHyperLink" worksheet.
  // If you know the table name, use `workbook.getTable('TableName')` instead.
  const table = workbook.getWorksheet('WithHyperLink').getTables()[0];

  // Get all the values from the table as text.
  const range = table.getRange();

  // Create an array of JSON objects that match the row structure.
  let returnObjects: TableData[] = [];
  if (table.getRowCount() > 0)  {
    returnObjects = returnObjectFromValues(range);
  }

  // Log the information and return it for a Power Automate flow.
  console.log(JSON.stringify(returnObjects));  
  return returnObjects
}

function returnObjectFromValues(range: ExcelScript.Range): TableData[] {
  let values = range.getTexts();
  let objectArray = [];
  let objectKeys: string[] = [];
  for (let i = 0; i < values.length; i++) {
    if (i === 0) {
      objectKeys = values[i]
      continue;
    }

    let object = {}
    for (let j = 0; j < values[i].length; j++) {
      // For the 4th column (0 index), extract the hyperlink and use that instead of text. 
      if (j === 4) {
        object[objectKeys[j]] = range.getCell(i, j).getHyperlink().address;
      } else {
        object[objectKeys[j]] = values[i][j];
      }
    }

    objectArray.push(object);
  }
  return objectArray as TableData[];
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

### <a name="sample-output-from-the-withhyperlink-worksheet"></a><span data-ttu-id="c1bea-126">Salida de ejemplo de la hoja de cálculo "WithHyperLink"</span><span class="sxs-lookup"><span data-stu-id="c1bea-126">Sample output from the "WithHyperLink" worksheet</span></span>

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

## <a name="use-in-power-automate"></a><span data-ttu-id="c1bea-127">Usar en Power Automate</span><span class="sxs-lookup"><span data-stu-id="c1bea-127">Use in Power Automate</span></span>

<span data-ttu-id="c1bea-128">Para obtener información sobre cómo usar este script en Power Automate, vea [Create an automated workflow with Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).</span><span class="sxs-lookup"><span data-stu-id="c1bea-128">For how to use such a script in Power Automate, see [Create an automated workflow with Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).</span></span>
