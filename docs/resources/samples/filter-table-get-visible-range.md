---
title: Filtrar tabla de Excel y obtener intervalo visible
description: Obtenga información sobre cómo usar scripts de Office para filtrar una tabla de Excel y obtener el intervalo visible como una matriz de objetos.
ms.date: 03/16/2021
localization_priority: Normal
ms.openlocfilehash: c0a5842af4a62162225e3fc10203c261b91e010a
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571607"
---
# <a name="filter-excel-table-and-get-visible-range-as-a-json-object"></a>Filtrar tabla de Excel y obtener rango visible como un objeto JSON

En este ejemplo se filtra una tabla de Excel y se devuelve el intervalo visible como un objeto JSON. Este JSON podría proporcionarse a un flujo de Power Automate como parte de una solución más grande.

## <a name="example-scenario"></a>Escenario de ejemplo

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
    "Reason": ""
  }, {
    "Station": "Station-1",
    "Shift": "Afternoon",
    "Date": "27-Oct-20",
    "Responsible": "Delia Dennis",
    "Reason": ""
  }, {
    "Station": "Station-1",
    "Shift": "Afternoon",
    "Date": "28-Oct-20",
    "Responsible": "Lidia Holloway",
    "Reason": ""
  }],
  "Station-2": [{
    "Station": "Station-2",
    "Shift": "Morning",
    "Date": "27-Oct-20",
    "Responsible": "Gerhart Moller",
    "Reason": ""
  }, {
    "Station": "Station-2",
    "Shift": "Afternoon",
    "Date": "28-Oct-20",
    "Responsible": "Grady Archie",
    "Reason": ""
  }],
  "Station-3": [{
    "Station": "Station-3",
    "Shift": "Morning",
    "Date": "27-Oct-20",
    "Responsible": "Isaiah Langer",
    "Reason": ""
  }]
}
```

## <a name="training-video-filter-an-excel-table-and-get-the-visible-range"></a>Vídeo de aprendizaje: filtrar una tabla de Excel y obtener el intervalo visible

[![Ver vídeo paso a paso sobre cómo filtrar una tabla de Excel y obtener el rango visible](../../images/visible-range-as-objects-vid.jpg)](https://youtu.be/Mv7BrvPq84A "Vídeo paso a paso sobre cómo filtrar una tabla de Excel y obtener el intervalo visible")
