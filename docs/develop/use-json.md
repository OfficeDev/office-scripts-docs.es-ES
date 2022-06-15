---
title: Uso de JSON para pasar datos hacia y desde scripts de Office
description: Aprenda a estructurar datos en objetos JSON para su uso con llamadas externas y Power Automate
ms.date: 06/10/2022
ms.localizationpriority: medium
ms.openlocfilehash: 753097183a18f5d20ca2c78a3748c7a1d968ad42
ms.sourcegitcommit: dd01979d34b3499360d2f79a56f8a8f24f480eed
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 06/15/2022
ms.locfileid: "66088161"
---
# <a name="use-json-to-pass-data-to-and-from-office-scripts"></a>Uso de JSON para pasar datos hacia y desde scripts de Office

[JSON (notación de objetos JavaScript)](https://www.w3schools.com/whatis/whatis_json.asp) es un formato para almacenar y transferir datos. Cada objeto JSON es una colección de pares nombre-valor que se pueden definir cuando se crean. JSON es útil con scripts de Office porque puede controlar la complejidad arbitraria de intervalos, tablas y otros patrones de datos en Excel. JSON permite analizar datos entrantes de [servicios web](external-calls.md) y pasar objetos complejos a través [de flujos de Power Automate](power-automate-integration.md).

Este artículo se centra en el uso de JSON con scripts de Office. Le recomendamos que primero obtenga más información sobre el formato en artículos como [Introducción a JSON](https://www.w3schools.com/js/js_json_intro.asp) de W3 Schools.

## <a name="parse-json-data-into-a-range-or-table"></a>Análisis de datos JSON en un intervalo o tabla

Las matrices de objetos JSON proporcionan una manera coherente de pasar filas de datos de tabla entre aplicaciones y servicios web. En estos casos, cada objeto JSON representa una fila, mientras que las propiedades representan las columnas. Un script Office puede recorrer en bucle una matriz JSON y volver a ensamblarla como una matriz 2D. A continuación, esta matriz se establece como los valores de un intervalo y se almacena en un libro. Los nombres de propiedad también se pueden agregar como encabezados para crear una tabla.

El siguiente script muestra los datos JSON que se convierten en una tabla. Tenga en cuenta que los datos no se toman de un origen externo. Esto se trata más adelante en este artículo.

```typescript
/**
 * Sample JSON data. This would be replaced by external calls or
 * parameters getting data from Power Automate in a production script.
 */
const jsonData = [
  { "Action": "Edit", /* Action property with value of "Edit". */
    "N": 3370, /* N property with value of 3370. */
    "Percent": 17.85 /* Percent property with value of 17.85. */
  },
  // The rest of the object entries follow the same pattern.
  { "Action": "Paste", "N": 1171, "Percent": 6.2 },
  { "Action": "Clear", "N": 599, "Percent": 3.17 },
  { "Action": "Insert", "N": 352, "Percent": 1.86 },
  { "Action": "Delete", "N": 350, "Percent": 1.85 },
  { "Action": "Refresh", "N": 314, "Percent": 1.66 },
  { "Action": "Fill", "N": 286, "Percent": 1.51 },
];

/**
 * This script converts JSON data to an Excel table.
 */
function main(workbook: ExcelScript.Workbook) {
  // Create a new worksheet to store the imported data.
  const newSheet = workbook.addWorksheet();
  newSheet.activate();

  // Determine the data's shape by getting the properties in one object.
  // This assumes all the JSON objects have the same properties.
  const columnNames = getPropertiesFromJson(jsonData[0]);

  // Create the table headers using the property names.
  const headerRange = newSheet.getRangeByIndexes(0, 0, 1, columnNames.length);
  headerRange.setValues([columnNames]);

  // Create a new table with the headers.
  const newTable = newSheet.addTable(headerRange, true);

  // Add each object in the array of JSON objects to the table.
  const tableValues = jsonData.map(row => convertJsonToRow(row));
  newTable.addRows(-1, tableValues);
}

/**
 * This function turns a JSON object into an array to be used as a table row.
 */
function convertJsonToRow(obj: object) {
  const array: (string | number)[] = [];

  // Loop over each property and get the value. Their order will be the same as the column headers.
  for (let value in obj) {
    array.push(obj[value]);
  }
  return array;
}

/**
 * This function gets the property names from a single JSON object.
 */
function getPropertiesFromJson(obj: object) {
  const propertyArray: string[] = [];
  
  // Loop over each property in the object and store the property name in an array.
  for (let property in obj) {
    propertyArray.push(property);
  }

  return propertyArray;
}
```

> [!TIP]
> Si conoce la estructura del JSON, puede crear su propia interfaz para facilitar la obtención de propiedades específicas. Puede reemplazar los pasos de conversión json a matriz por referencias seguras para tipos. El siguiente fragmento de código muestra esos pasos (ahora comentados) reemplazados por llamadas que usan una nueva `ActionRow` interfaz. Tenga en cuenta que esto hace que la `convertJsonToRow` función ya no sea necesaria.
>
> ```typescript
>   // const tableValues = jsonData.map(row => convertJsonToRow(row));
>   // newTable.addRows(-1, tableValues);
>   // }
>
>      const actionRows: ActionRow[] = jsonData as ActionRow[];
>      // Add each object in the array of JSON objects to the table.
>      const tableValues = actionRows.map(row => [row.Action, row.N, row.Percent]);
>      newTable.addRows(-1, tableValues);
>    }
>    
>    interface ActionRow {
>      Action: string;
>      N: number;
>      Percent: number;
>    }
> ```

### <a name="get-json-data-from-external-sources"></a>Obtención de datos JSON de orígenes externos

Hay dos maneras de importar datos JSON en el libro a través de un script de Office.

- Como [parámetro](power-automate-integration.md#main-parameters-pass-data-to-a-script) con un flujo de Power Automate.
- Con una `fetch` llamada a un [servicio web externo](external-calls.md).

#### <a name="modify-the-sample-to-work-with-power-automate"></a>Modifique el ejemplo para que funcione con Power Automate

Los datos JSON de Power Automate se pueden pasar como una matriz de objetos genéricos. Agregue una `object[]` propiedad al script para aceptar esos datos.

```typescript
// For Power Automate, replace the main signature in the previous sample with this one
// and remove the sample data.
function main(workbook: ExcelScript.Workbook, jsonData: object[]) {
```

A continuación, verá una opción en el conector de Power Automate para agregarla `jsonData` a la acción **Ejecutar script**.

:::image type="content" source="../images/json-parameter-power-automate.png" alt-text="Un conector de Excel Online (Empresa) que muestra una acción Ejecutar script con el parámetro jsonData.":::

#### <a name="modify-the-sample-to-use-a-fetch-call"></a>Modificación del ejemplo para usar una `fetch` llamada

Los servicios web pueden responder a `fetch` llamadas con datos JSON. Esto proporciona al script los datos que necesita y, al mismo tiempo, le mantiene en Excel. Para obtener más información sobre `fetch` las llamadas externas y , consulte [Compatibilidad con llamadas API externas en scripts de Office](external-calls.md).

```typescript
// For external services, replace the main signature in the previous sample with this one,
// add the fetch call, and remove the sample data.
async function main(workbook: ExcelScript.Workbook) {
  // Replace WEB_SERVICE_URL with the URL of whatever service you need to call.
  const response = await fetch('WEB_SERVICE_URL');
  const jsonData: object[] = await response.json();
```

## <a name="create-json-from-a-range"></a>Creación de JSON a partir de un intervalo

Las filas y columnas de una hoja de cálculo suelen implicar relaciones entre sus valores de datos. Una fila de una tabla se asigna conceptualmente a un objeto de programación, con cada columna como propiedad de ese objeto. Tenga en cuenta la siguiente tabla de datos. Cada fila representa una transacción registrada en la hoja de cálculo.

|ID |Fecha     |Amount |Proveedor                        |
|:--|:--------|:------|:-----------------------------|
|1  |6/1/2022 |$43.54 |Mejor para usted Organics Company |
|2  |6/3/2022 |$67.23 |Liberty Bakery and Cafe       |
|3  |6/3/2022 |$37.12 |Mejor para usted Organics Company |
|4   |6/6/2022 |$86.95 |Coho Vineyard                 |
|5   |6/7/2022 |$13.64 |Liberty Bakery and Cafe       |

Cada transacción (cada fila) tiene un conjunto de propiedades asociadas: "ID", "Date", "Amount" y "Vendor". Esto se puede modelar en un script de Office como un objeto.

```typescript
// An interface that wraps transaction details as JSON.
interface Transaction {
  "ID": string;
  "Date": number;
  "Amount": number;
  "Vendor": string;
}
```

Las filas de la tabla de ejemplo corresponden a las propiedades de la interfaz, por lo que un script puede convertir fácilmente cada fila en un `Transaction` objeto. Esto resulta útil al generar los datos de Power Automate. El siguiente script recorre en iteración cada fila de la tabla y lo agrega a .`Transaction[]`

```typescript
function main(workbook: ExcelScript.Workbook) {
  // Get the table on the current worksheet.
  const table = workbook.getActiveWorksheet().getTables()[0];

  // Create an array of Transactions and add each row to it.
  let transactions: Transaction[] = [];
  const dataValues = table.getRangeBetweenHeaderAndTotal().getValues();
  for (let i = 0; i < dataValues.length; i++) {
    let row = dataValues[i];
    let currentTransaction: Transaction = {
      ID: row[table.getColumnByName("ID").getIndex()] as string,
      Date: row[table.getColumnByName("Date").getIndex()] as number,
      Amount: row[table.getColumnByName("Amount").getIndex()] as number,
      Vendor: row[table.getColumnByName("Vendor").getIndex()] as string
    };
    transactions.push(currentTransaction);
  }

  // Do something with the Transaction objects, such as return them to a Power Automate flow.
  console.log(transactions);
}

// An interface that wraps transaction details as JSON.
interface Transaction {
  "ID": string;
  "Date": number;
  "Amount": number;
  "Vendor": string;
}
```

:::image type="content" source="../images/create-json-console-output.png" alt-text="Salida de la consola del script anterior que muestra los valores de propiedad del objeto.":::

### <a name="use-a-generic-object"></a>Uso de un objeto genérico

En el ejemplo anterior se supone que los valores de encabezado de tabla son coherentes. Si la tabla tiene columnas variables, deberá crear un objeto JSON genérico. El script siguiente muestra un script que registra cualquier tabla como JSON.

```typescript
function main(workbook: ExcelScript.Workbook) {
  // Get the table on the current worksheet.
  const table = workbook.getActiveWorksheet().getTables()[0];

  // Use the table header names as JSON properties.
  const tableHeaders = table.getHeaderRowRange().getValues()[0] as string[];
  
  // Get each data row in the table.
  const dataValues = table.getRangeBetweenHeaderAndTotal().getValues();
  let jsonArray: object[] = [];

  // For each row, create a JSON object and assign each property to it based on the table headers.
  for (let i = 0; i < dataValues.length; i++) {
    // Create a blank generic JSON object.
    let jsonObject: { [key: string]: string } = {};
    for (let j = 0; j < dataValues[i].length; j++) {
      jsonObject[tableHeaders[j]] = dataValues[i][j] as string;
    }

    jsonArray.push(jsonObject);
  }

  // Do something with the objects, such as return them to a Power Automate flow.
  console.log(jsonArray);
}

```

## <a name="see-also"></a>Consulte también

- [Soporte de llamadas de API externas en Scripts de Office](external-calls.md)
- [Ejemplo: Uso de llamadas de captura externas en scripts de Office](../resources/samples/external-fetch-calls.md)
- [Ejecución de scripts de Office con Power Automate](power-automate-integration.md)