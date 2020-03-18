---
title: Uso de objetos de JavaScript integrados en scripts de Office
description: Cómo llamar a las API de JavaScript integradas desde un script de Office en Excel en la Web.
ms.date: 01/21/2020
localization_priority: Normal
ms.openlocfilehash: e0fcd98117125ead18e55675e195415ff59c0c5d
ms.sourcegitcommit: b075eed5a6f275274fbbf6d62633219eac416f26
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 03/10/2020
ms.locfileid: "42700355"
---
# <a name="using-built-in-javascript-objects-in-office-scripts"></a>Uso de objetos de JavaScript integrados en scripts de Office

JavaScript proporciona varios objetos integrados que puede usar en los scripts de Office, independientemente de si está creando scripts en JavaScript o [TypeScript](../overview/code-editor-environment.md) (un superconjunto de JavaScript). En este artículo se describe cómo se pueden usar algunos de los objetos de JavaScript integrados en los scripts de Office para Excel en la Web.

> [!NOTE]
> Para obtener una lista completa de todos los objetos de JavaScript integrados, consulte el artículo de [objetos integrados estándar](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects) de Mozilla.

## <a name="array"></a>Matriz

El objeto [array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) proporciona una forma estandarizada de trabajar con matrices en el script. Aunque las matrices son construcciones estándar de JavaScript, están relacionadas con los scripts de Office de dos maneras principales: Ranges y Collections.

### <a name="working-with-ranges"></a>Trabajar con rangos

Los rangos contienen varias matrices bidimensionales que se asignan directamente a las celdas de ese intervalo. Entre ellas se incluyen propiedades `values`como `formulas`, y `numberFormat`. Las propiedades de tipo de matriz deben [cargarse](scripting-fundamentals.md#sync-and-load) como cualquier otra propiedad.

La siguiente secuencia de comandos busca el intervalo **a1: D4** para cualquier formato de número que contenga un "$". La secuencia de comandos establece el color de relleno de esas celdas en "amarillo".

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the range From A1 to D4.
  let range = context.workbook.worksheets.getActiveWorksheet().getRange("A1:D4");

  // Load the numberFormat property on the range.
  range.load("numberFormat");
  await context.sync();

  // Iterate through the arrays of rows and columns corresponding to those in the range.
  range.numberFormat.forEach((rowItem, rowIndex) => {
    range.numberFormat[rowIndex].forEach((columnItem, columnIndex) => {
      // Treat the numberFormat as a string so we can do text comparisons.
      let columnItemText = columnItem as string;
      if (columnItemText.indexOf("$") >= 0) {
        // Set the cell's fill to yellow.
        range.getCell(rowIndex, columnIndex).format.fill.color = "yellow";
      }
    });
  });
}
```

### <a name="working-with-collections"></a>Trabajar con colecciones

Muchos objetos de Excel están incluidos en una colección. Por ejemplo, todas las [formas](/javascript/api/office-scripts/excel/excel.shape) de una hoja de cálculo están contenidas en `Worksheet.shapes` [ShapeCollection](/javascript/api/office-scripts/excel/excel.shapecollection) (como la propiedad). Cada `*Collection` objeto contiene una `items` propiedad, que es una matriz que almacena los objetos dentro de dicha colección. Esto puede tratarse como una matriz de JavaScript normal, pero los elementos de la colección deben cargarse primero. Si necesita trabajar con una propiedad en cada objeto de la colección, use una instrucción de carga jerárquica (`items/propertyName`).

La siguiente secuencia de comandos registra el tipo de cada forma de la hoja de cálculo actual.

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the current worksheet.
  let selectedSheet = context.workbook.worksheets.getActiveWorksheet();

  // Get the shapes in this worksheet.
  let shapes = selectedSheet.shapes;
  shapes.load("items/type");
  await context.sync();

  // Log the type of every shape in the collection.
  shapes.items.forEach((shape) => {
    console.log(shape.type);
  });
}
```

Puede cargar objetos individuales de una colección mediante los `getItem` métodos o `getItemAt` . `getItem`Obtiene un objeto mediante un identificador único como un nombre (a menudo, los nombres se especifican en el script). `getItemAt`Obtiene un objeto mediante su índice en la colección. Cada llamada debe ir seguida de un `await context.sync();` comando para que se pueda usar el objeto.

La siguiente secuencia de comandos elimina la forma más antigua de la hoja de cálculo actual.

```Typescript
async function main(context: Excel.RequestContext) {
  // Get the current worksheet.
  let selectedSheet = context.workbook.worksheets.getActiveWorksheet();

  // Get the first (oldest) shape in the worksheet.
  // Note that this script will thrown an error if there are no shapes.
  let shape = selectedSheet.shapes.getItemAt(0);

  // Sync to load `shape` from the collection.
  await context.sync();

  // Remove the shape from the worksheet.
  shape.delete();
}
```

## <a name="date"></a>Fecha

El objeto [Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) proporciona una forma estandarizada de trabajar con fechas en el script. `Date.now()`genera un objeto con la fecha y hora actuales, lo que resulta útil cuando se agregan marcas de tiempo a la entrada de datos del script.

La siguiente secuencia de comandos agrega la fecha actual a la hoja de cálculo. Tenga en cuenta que, `toLocaleDateString` al usar el método, Excel reconoce el valor como una fecha y cambia automáticamente el formato de número de la celda.

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the range for cell A1.
  let range = context.workbook.worksheets.getActiveWorksheet().getRange("A1");

  // Get the current date and time.
  let date = new Date(Date.now());

  // Set the value at A1 to the current date, using a localized string.
  range.values = [[date.toLocaleDateString()]];
}
```

## <a name="math"></a>Matemáticas

El objeto [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math) proporciona métodos y constantes para las operaciones matemáticas comunes. Estos proporcionan muchas funciones que también están disponibles en Excel, sin necesidad de usar el motor de cálculo del libro. Esto evita que el script tenga que consultar el libro, lo que mejora el rendimiento.

El siguiente script usa `Math.min` para buscar y registrar el número menor del intervalo de **a1: D4** . Tenga en cuenta que en este ejemplo se supone que el rango completo contiene sólo números, no cadenas.

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the range from A1 to D4.
  let comparisonRange = context.workbook.worksheets.getActiveWorksheet().getRange("A1:D4");
  
  // Load the range's values.
  comparisonRange.load("values");
  await context.sync();

  // Set the minimum values as the first value.
  let minimum = comparisonRange.values[0][0];

  // Iterate over each row looking for the smallest value.
  comparisonRange.values.forEach((rowItem, rowIndex) => {
    // Iterate over each column looking for the smallest value.
    comparisonRange.values[rowIndex].forEach((columnItem) => {
      // Use `Math.min` to set the smallest value as either the current cell's value or the previous minimum.
      minimum = Math.min(minimum, columnItem);
    });
  });
  
  console.log(minimum);
}

```

## <a name="see-also"></a>Vea también

- [Objetos integrados estándar](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects)
- [Entorno de editor de código de scripts de Office](../overview/code-editor-environment.md)
