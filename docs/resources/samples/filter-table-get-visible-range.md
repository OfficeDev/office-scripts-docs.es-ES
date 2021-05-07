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
# <a name="filter-excel-table-and-get-visible-range-as-a-json-object"></a>Filtrar Excel tabla y obtener rango visible como un objeto JSON

En este ejemplo se filtra Excel tabla y se devuelve el intervalo visible como un objeto JSON. Este JSON podría proporcionarse a un flujo Power Automate como parte de una solución más grande.

## <a name="example-scenario"></a>Ejemplo ficticio

* Aplicar un filtro a una columna de tabla.
* Extraer el intervalo visible después de filtrar.
* Ensamblar y devolver un objeto con una [estructura JSON específica.](#sample-json)

## <a name="sample-code-filter-a-table-and-get-visible-range"></a>Código de ejemplo: filtrar una tabla y obtener rango visible

El siguiente script filtra una tabla y obtiene el intervalo visible.

Descargue el archivo de <a href="table-filter.xlsx">table-filter.xlsx</a> y ústelo con este script para probarlo usted mismo.

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

### <a name="sample-json"></a>JSON de ejemplo

Cada clave representa un valor único de una tabla. Cada instancia de matriz representa la fila que está visible cuando se aplica el filtro correspondiente.

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

## <a name="training-video-filter-an-excel-table-and-get-the-visible-range"></a>Vídeo de aprendizaje: filtrar una Excel tabla y obtener el intervalo visible

[Vea el recorrido de Sudhi Ramamurthy por este ejemplo en YouTube](https://youtu.be/Mv7BrvPq84A).
