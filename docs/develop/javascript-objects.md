---
title: Usar objetos integrados de JavaScript en los scripts de Office
description: Cómo llamar a las API de JavaScript integradas desde un script de Office en Excel en la Web.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 680dd326e357bd06e2fc66cba5bd6745bbd33c24
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545050"
---
# <a name="use-built-in-javascript-objects-in-office-scripts"></a>Utilice objetos JavaScript integrados en scripts de Office

JavaScript proporciona varios objetos integrados que puede usar en los scripts de Office, independientemente de si está guionando en JavaScript o [TypeScript](../overview/code-editor-environment.md) (un superconjunto de JavaScript). En este artículo se describe cómo puede utilizar algunos de los objetos JavaScript integrados en scripts Office para Excel en la Web.

> [!NOTE]
> Para obtener una lista completa de todos los objetos JavaScript integrados, consulte el artículo [Objetos integrados estándar de](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects) Mozilla.

## <a name="array"></a>Matriz

El objeto [Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) proporciona una forma estandarizada de trabajar con matrices en el script. Aunque las matrices son construcciones JavaScript estándar, se relacionan con scripts Office de dos maneras principales: rangos y colecciones.

### <a name="work-with-ranges"></a>Trabajar con gamas

Los rangos contienen varias matrices bidimensionales que se asignan directamente a las celdas de ese rango. Estas matrices contienen información específica sobre cada celda de ese intervalo. Por ejemplo, `Range.getValues` devuelve todos los valores de esas celdas (con las filas y columnas de la asignación de matriz bidimensional a las filas y columnas de esa subsección de hoja de cálculo). `Range.getFormulas` y `Range.getNumberFormats` son otros métodos utilizados con frecuencia que devuelven matrices como `Range.getValues` .

El siguiente script busca en el intervalo **A1:D4** cualquier formato numérico que contenga un "$". El script establece el color de relleno en esas celdas en "amarillo".

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

### <a name="work-with-collections"></a>Trabajar con colecciones

Muchos objetos Excel están contenidos en una colección. La colección es administrada por la API de scripts de Office y expuesta como una matriz. Por ejemplo, todas las [formas](/javascript/api/office-scripts/excelscript/excelscript.shape) de una hoja de cálculo están contenidas en una `Shape[]` que devuelve el `Worksheet.getShapes` método. Puede usar esta matriz para leer valores de la colección o puede tener acceso a objetos específicos desde los métodos del objeto `get*` primario.

> [!NOTE]
> No agregue ni quite manualmente objetos de estas matrices de colección. Utilice los `add` métodos de los objetos primarios y los métodos de los objetos de tipo `delete` colección. Por ejemplo, agregue una [tabla](/javascript/api/office-scripts/excelscript/excelscript.table) a una [hoja de cálculo](/javascript/api/office-scripts/excelscript/excelscript.worksheet) con el método y quite el uso `Worksheet.addTable` `Table` `Table.delete` .

El siguiente script registra el tipo de cada forma de la hoja de cálculo actual.

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

El siguiente script elimina la forma más antigua de la hoja de cálculo actual.

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

El objeto [Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) proporciona una forma estandarizada de trabajar con fechas en el script. `Date.now()` genera un objeto con la fecha y hora actuales, lo que resulta útil al agregar marcas de tiempo a la entrada de datos del script.

El siguiente script agrega la fecha actual a la hoja de cálculo. Tenga en cuenta que mediante el uso del `toLocaleDateString` método, Excel reconoce el valor como una fecha y cambia el formato de número de la celda automáticamente.

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

La sección [Trabajar con fechas](../resources/samples/excel-samples.md#dates) de los ejemplos tiene más scripts relacionados con la fecha.

## <a name="math"></a>Matemáticas

El objeto [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math) proporciona métodos y constantes para operaciones matemáticas comunes. Estos proporcionan muchas funciones también disponibles en Excel, sin necesidad de utilizar el motor de cálculo del libro de trabajo. Esto evita que el script tenga que consultar el libro, lo que mejora el rendimiento.

El siguiente script utiliza `Math.min` para buscar y registrar el número más pequeño en el intervalo **A1:D4.** Tenga en cuenta que en este ejemplo se supone que todo el intervalo solo contiene números, no cadenas.

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

## <a name="use-of-external-javascript-libraries-is-not-supported"></a>No se admite el uso de bibliotecas JavaScript externas

Office Los scripts no admiten el uso de bibliotecas externas de terceros. El script solo puede utilizar los objetos JavaScript integrados y las API de scripts de Office.

## <a name="see-also"></a>Vea también

- [Objetos incorporados estándar](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects)
- [Office Entorno del Editor de código de scripts](../overview/code-editor-environment.md)
