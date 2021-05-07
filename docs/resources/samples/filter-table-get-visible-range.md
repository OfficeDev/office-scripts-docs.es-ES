---
title: Filtrar Excel tabla y obtener rango visible
description: Obtenga información sobre cómo usar Office scripts para filtrar una tabla Excel y obtener el rango visible como una matriz de objetos.
ms.date: 04/28/2021
localization_priority: Normal
ms.openlocfilehash: a310857e6055b3da57c353dc7ad78a6fbdd86d4e
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232378"
---
# <a name="filter-excel-table-and-get-visible-range-as-a-json-object"></a><span data-ttu-id="d7553-103">Filtrar Excel tabla y obtener rango visible como un objeto JSON</span><span class="sxs-lookup"><span data-stu-id="d7553-103">Filter Excel table and get visible range as a JSON object</span></span>

<span data-ttu-id="d7553-104">En este ejemplo se filtra Excel tabla y se devuelve el intervalo visible como un objeto JSON.</span><span class="sxs-lookup"><span data-stu-id="d7553-104">This sample filters an Excel table and returns the visible range as a JSON object.</span></span> <span data-ttu-id="d7553-105">Este JSON podría proporcionarse a un flujo Power Automate como parte de una solución más grande.</span><span class="sxs-lookup"><span data-stu-id="d7553-105">This JSON could be provided to a Power Automate flow as part of a larger solution.</span></span>

## <a name="example-scenario"></a><span data-ttu-id="d7553-106">Ejemplo ficticio</span><span class="sxs-lookup"><span data-stu-id="d7553-106">Example scenario</span></span>

* <span data-ttu-id="d7553-107">Aplicar un filtro a una columna de tabla.</span><span class="sxs-lookup"><span data-stu-id="d7553-107">Apply a filter to a table column.</span></span>
* <span data-ttu-id="d7553-108">Extraer el intervalo visible después de filtrar.</span><span class="sxs-lookup"><span data-stu-id="d7553-108">Extract the visible range after filtering.</span></span>
* <span data-ttu-id="d7553-109">Ensamblar y devolver un objeto con una [estructura JSON específica.](#sample-json)</span><span class="sxs-lookup"><span data-stu-id="d7553-109">Assemble and return an object with a [specific JSON structure](#sample-json).</span></span>

## <a name="sample-code-filter-a-table-and-get-visible-range"></a><span data-ttu-id="d7553-110">Código de ejemplo: filtrar una tabla y obtener rango visible</span><span class="sxs-lookup"><span data-stu-id="d7553-110">Sample code: Filter a table and get visible range</span></span>

<span data-ttu-id="d7553-111">El siguiente script filtra una tabla y obtiene el intervalo visible.</span><span class="sxs-lookup"><span data-stu-id="d7553-111">The following script filters a table and gets the visible range.</span></span>

<span data-ttu-id="d7553-112">Descargue el archivo de <a href="table-filter.xlsx">table-filter.xlsx</a> y ústelo con este script para probarlo usted mismo.</span><span class="sxs-lookup"><span data-stu-id="d7553-112">Download the sample file <a href="table-filter.xlsx">table-filter.xlsx</a> and use it with this script to try it out yourself!</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): ReturnTemplate {
  const table1 = workbook.getTable("Table1");
  const keyColumnValues: string [] = table1.getColumnByName('Station').getRangeBetweenHeaderAndTotal().getValues().map(v => v[0] as string);
  const uniqueKeys = keyColumnValues.filter((v, i, a) => a.indexOf(v) === i);

  console.log(uniqueKeys);
  const returnObj: ReturnTemplate = {}

  uniqueKeys.forEach((key: string) => {
    table1.getColumnByName('Station').getFilter()
      .applyValuesFilter([key]);
    const rangeView = table1.getRange().getVisibleView();
    returnObj[key] = returnObjectFromValues(rangeView.getValues() as string[][]);
  })
  table1.getColumnByName('Station').getFilter().clear();
  console.log(JSON.stringify(returnObj));
  return returnObj
}

function returnObjectFromValues(values: string[][]): BasicObj[] {
  let objArray = [];
  let objKeys: string[] = [];
  for (let i=0; i < values.length; i++) {
    if (i===0) {
      objKeys = values[i]
      continue;
    }
    let obj = {}
    for (let j=0; j < values[i].length; j++) {
      obj[objKeys[j]] = values[i][j]
    }
    objArray.push(obj);
  }
  return objArray;
}

interface BasicObj {
  [key: string] : string
}

interface ReturnTemplate {
  [key: string]: BasicObj[]
}
```

### <a name="sample-json"></a><span data-ttu-id="d7553-113">JSON de ejemplo</span><span class="sxs-lookup"><span data-stu-id="d7553-113">Sample JSON</span></span>

<span data-ttu-id="d7553-114">Cada clave representa un valor único de una tabla.</span><span class="sxs-lookup"><span data-stu-id="d7553-114">Each key represents a unique value of a table.</span></span> <span data-ttu-id="d7553-115">Cada instancia de matriz representa la fila que está visible cuando se aplica el filtro correspondiente.</span><span class="sxs-lookup"><span data-stu-id="d7553-115">Each array instance represents the row that is visible when the corresponding filter is applied.</span></span>

```json
{
  "Station-1": [{
    "Station": "Station-1",
    "Shift": "Morning",
    "Date": "27-Oct-20",
    "Responsible": "Debra Berger",
    "Reason&quot;: &quot;"
  }, {
    "Station": "Station-1",
    "Shift": "Afternoon",
    "Date": "27-Oct-20",
    "Responsible": "Delia Dennis",
    "Reason&quot;: &quot;"
  }, {
    "Station": "Station-1",
    "Shift": "Afternoon",
    "Date": "28-Oct-20",
    "Responsible": "Lidia Holloway",
    "Reason&quot;: &quot;"
  }],
  "Station-2": [{
    "Station": "Station-2",
    "Shift": "Morning",
    "Date": "27-Oct-20",
    "Responsible": "Gerhart Moller",
    "Reason&quot;: &quot;"
  }, {
    "Station": "Station-2",
    "Shift": "Afternoon",
    "Date": "28-Oct-20",
    "Responsible": "Grady Archie",
    "Reason&quot;: &quot;"
  }],
  "Station-3": [{
    "Station": "Station-3",
    "Shift": "Morning",
    "Date": "27-Oct-20",
    "Responsible": "Isaiah Langer",
    "Reason&quot;: &quot;"
  }]
}
```

## <a name="training-video-filter-an-excel-table-and-get-the-visible-range"></a><span data-ttu-id="d7553-116">Vídeo de aprendizaje: filtrar una Excel tabla y obtener el intervalo visible</span><span class="sxs-lookup"><span data-stu-id="d7553-116">Training video: Filter an Excel table and get the visible range</span></span>

<span data-ttu-id="d7553-117">[Vea el recorrido de Sudhi Ramamurthy por este ejemplo en YouTube](https://youtu.be/Mv7BrvPq84A).</span><span class="sxs-lookup"><span data-stu-id="d7553-117">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/Mv7BrvPq84A).</span></span>
