---
title: Usar objetos integrados de JavaScript en los scripts de Office
description: Cómo llamar a las API de JavaScript integradas desde un script de Office en Excel en la Web.
ms.date: 06/29/2020
localization_priority: Normal
ms.openlocfilehash: 1c8ac757574e8c4be64b373f8d4bf421ddfa0c79
ms.sourcegitcommit: 30750c4392db3ef057075a5702abb92863c93eda
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 07/01/2020
ms.locfileid: "44999263"
---
# <a name="using-built-in-javascript-objects-in-office-scripts"></a>Usar objetos integrados de JavaScript en los scripts de Office

JavaScript proporciona varios objetos integrados que puede usar en los scripts de Office, independientemente de si está creando scripts en JavaScript o [TypeScript](../overview/code-editor-environment.md) (un superconjunto de JavaScript). En este artículo se describe cómo se pueden usar algunos de los objetos de JavaScript integrados en los scripts de Office para Excel en la Web.

> [!NOTE]
> Para obtener una lista completa de todos los objetos de JavaScript integrados, consulte el artículo de [objetos integrados estándar](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects) de Mozilla.

## <a name="array"></a>Matriz

El objeto [array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) proporciona una forma estandarizada de trabajar con matrices en el script. Aunque las matrices son construcciones estándar de JavaScript, están relacionadas con los scripts de Office de dos maneras principales: Ranges y Collections.

### <a name="working-with-ranges"></a>Trabajar con rangos

Los rangos contienen varias matrices bidimensionales que se asignan directamente a las celdas de ese intervalo. Estas matrices contienen información específica sobre cada celda de ese intervalo. Por ejemplo, `Range.getValues` devuelve todos los valores de esas celdas (con las filas y columnas de la matriz bidimensional asignada a las filas y columnas de esa subsección de la hoja de cálculo). `Range.getFormulas`y `Range.getNumberFormats` son otros métodos usados con frecuencia que devuelven matrices como `Range.getValues` .

La siguiente secuencia de comandos busca el intervalo **a1: D4** para cualquier formato de número que contenga un "$". La secuencia de comandos establece el color de relleno de esas celdas en "amarillo".

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the range From A1 to D4.
  let range = workbook.getActiveWorksheet().getRange("A1:D4");

  // Get the number formats for each cell in the range.
  let rangeNumberFormats = range.getNumberFormats();
  // Iterate through the arrays of rows and columns corresponding to those in the range.
  rangeNumberFormats.forEach((rowItem, rowIndex) => {
    rangeNumberFormats[rowIndex].forEach((columnItem, columnIndex) => {
      // Treat the numberFormat as a string so we can do text comparisons.
      let columnItemText = columnItem as string;
      if (columnItemText.indexOf("$") >= 0) {
        // Set the cell's fill to yellow.
        range.getCell(rowIndex, columnIndex).getFormat().getFill().setColor("yellow");
      }
    });
  });
}
```

### <a name="working-with-collections"></a>Trabajar con colecciones

Muchos objetos de Excel están incluidos en una colección. La colección se administra mediante la API de scripts de Office y se expone como una matriz. Por ejemplo, todas las [formas](/javascript/api/office-scripts/excelscript/excelscript.shape) de una hoja de cálculo están contenidas en un `Shape[]` devuelto por el `Worksheet.getShapes` método. Puede usar esta matriz para leer valores de la colección o puede obtener acceso a objetos específicos desde los métodos del objeto primario `get*` .

> [!NOTE]
> No agregue ni quite objetos manualmente de estas matrices de colecciones. Use los `add` métodos de los objetos primarios y los `delete` métodos de los objetos de tipo de colección. Por ejemplo, agregue una [tabla](/javascript/api/office-scripts/excelscript/excelscript.table) a una [hoja de cálculo](/javascript/api/office-scripts/excelscript/excelscript.worksheet) con el `Worksheet.addTable` método y quite el `Table` using `Table.delete` .

La siguiente secuencia de comandos registra el tipo de cada forma de la hoja de cálculo actual.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  let selectedSheet = workbook.getActiveWorksheet();

  // Get the shapes in this worksheet.
  let shapes = selectedSheet.getShapes();

  // Log the type of every shape in the collection.
  shapes.forEach((shape) => {
    console.log(shape.getType());
  });
}
```

La siguiente secuencia de comandos elimina la forma más antigua de la hoja de cálculo actual.

```Typescript
function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  let selectedSheet = workbook.getActiveWorksheet();

  // Get the first (oldest) shape in the worksheet.
  // Note that this script will thrown an error if there are no shapes.
  let shape = selectedSheet.getShapes()[0];

  // Remove the shape from the worksheet.
  shape.delete();
}
```

## <a name="date"></a>Fecha

El objeto [Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) proporciona una forma estandarizada de trabajar con fechas en el script. `Date.now()`genera un objeto con la fecha y hora actuales, lo que resulta útil cuando se agregan marcas de tiempo a la entrada de datos del script.

La siguiente secuencia de comandos agrega la fecha actual a la hoja de cálculo. Tenga en cuenta que, al usar el `toLocaleDateString` método, Excel reconoce el valor como una fecha y cambia automáticamente el formato de número de la celda.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the range for cell A1.
  let range = workbook.getActiveWorksheet().getRange("A1");

  // Get the current date and time.
  let date = new Date(Date.now());

  // Set the value at A1 to the current date, using a localized string.
  range.setValue(date.toLocaleDateString());
}
```

La sección [work with](../resources/excel-samples.md#work-with-dates) Dates de los ejemplos tiene más scripts relacionados con la fecha.

## <a name="math"></a>Matemáticas

El objeto [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math) proporciona métodos y constantes para las operaciones matemáticas comunes. Estos proporcionan muchas funciones que también están disponibles en Excel, sin necesidad de usar el motor de cálculo del libro. Esto evita que el script tenga que consultar el libro, lo que mejora el rendimiento.

El siguiente script usa `Math.min` para buscar y registrar el número menor del intervalo de **a1: D4** . Tenga en cuenta que en este ejemplo se supone que el rango completo contiene sólo números, no cadenas.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the range from A1 to D4.
  let comparisonRange = workbook.getActiveWorksheet().getRange("A1:D4");

  // Load the range's values.
  let comparisonRangeValues = comparisonRange.getValues();

  // Set the minimum values as the first value.
  let minimum = comparisonRangeValues[0][0];

  // Iterate over each row looking for the smallest value.
  comparisonRangeValues.forEach((rowItem, rowIndex) => {
    // Iterate over each column looking for the smallest value.
    comparisonRangeValues[rowIndex].forEach((columnItem) => {
      // Use `Math.min` to set the smallest value as either the current cell's value or the previous minimum.
      minimum = Math.min(minimum, columnItem);
    });
  });

  console.log(minimum);
}

```

## <a name="use-of-external-javascript-libraries-is-not-supported"></a>No se admite el uso de bibliotecas de JavaScript externas

Los scripts de Office no admiten el uso de bibliotecas externas de terceros. El script solo puede usar los objetos de JavaScript integrados y las API de scripts de Office.

## <a name="see-also"></a>Ver también

- [Objetos integrados estándar](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects)
- [Entorno de editor de código de scripts de Office](../overview/code-editor-environment.md)
