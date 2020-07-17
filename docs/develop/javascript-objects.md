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
# <a name="using-built-in-javascript-objects-in-office-scripts"></a><span data-ttu-id="f6013-103">Usar objetos integrados de JavaScript en los scripts de Office</span><span class="sxs-lookup"><span data-stu-id="f6013-103">Using built-in JavaScript objects in Office Scripts</span></span>

<span data-ttu-id="f6013-104">JavaScript proporciona varios objetos integrados que puede usar en los scripts de Office, independientemente de si está creando scripts en JavaScript o [TypeScript](../overview/code-editor-environment.md) (un superconjunto de JavaScript).</span><span class="sxs-lookup"><span data-stu-id="f6013-104">JavaScript provides several built-in objects that you can use in your Office Scripts, regardless of whether you're scripting in JavaScript or [TypeScript](../overview/code-editor-environment.md) (a superset of JavaScript).</span></span> <span data-ttu-id="f6013-105">En este artículo se describe cómo se pueden usar algunos de los objetos de JavaScript integrados en los scripts de Office para Excel en la Web.</span><span class="sxs-lookup"><span data-stu-id="f6013-105">This article describes how you can use some of the built-in JavaScript objects in Office Scripts for Excel on the web.</span></span>

> [!NOTE]
> <span data-ttu-id="f6013-106">Para obtener una lista completa de todos los objetos de JavaScript integrados, consulte el artículo de [objetos integrados estándar](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects) de Mozilla.</span><span class="sxs-lookup"><span data-stu-id="f6013-106">For a complete list of all built-in JavaScript objects, see Mozilla's [Standard built-in objects](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects) article.</span></span>

## <a name="array"></a><span data-ttu-id="f6013-107">Matriz</span><span class="sxs-lookup"><span data-stu-id="f6013-107">Array</span></span>

<span data-ttu-id="f6013-108">El objeto [array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) proporciona una forma estandarizada de trabajar con matrices en el script.</span><span class="sxs-lookup"><span data-stu-id="f6013-108">The [Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) object provides a standardized way to work with arrays in your script.</span></span> <span data-ttu-id="f6013-109">Aunque las matrices son construcciones estándar de JavaScript, están relacionadas con los scripts de Office de dos maneras principales: Ranges y Collections.</span><span class="sxs-lookup"><span data-stu-id="f6013-109">While arrays are standard JavaScript constructs, they relate to Office Scripts in two major ways: ranges and collections.</span></span>

### <a name="working-with-ranges"></a><span data-ttu-id="f6013-110">Trabajar con rangos</span><span class="sxs-lookup"><span data-stu-id="f6013-110">Working with ranges</span></span>

<span data-ttu-id="f6013-111">Los rangos contienen varias matrices bidimensionales que se asignan directamente a las celdas de ese intervalo.</span><span class="sxs-lookup"><span data-stu-id="f6013-111">Ranges contain several two-dimensional arrays that directly map to the cells in that range.</span></span> <span data-ttu-id="f6013-112">Estas matrices contienen información específica sobre cada celda de ese intervalo.</span><span class="sxs-lookup"><span data-stu-id="f6013-112">These arrays contain specific information about each cell in that range.</span></span> <span data-ttu-id="f6013-113">Por ejemplo, `Range.getValues` devuelve todos los valores de esas celdas (con las filas y columnas de la matriz bidimensional asignada a las filas y columnas de esa subsección de la hoja de cálculo).</span><span class="sxs-lookup"><span data-stu-id="f6013-113">For example, `Range.getValues` returns all the values in those cells (with the rows and columns of the two-dimensional array mapping to the rows and columns of that worksheet subsection).</span></span> <span data-ttu-id="f6013-114">`Range.getFormulas`y `Range.getNumberFormats` son otros métodos usados con frecuencia que devuelven matrices como `Range.getValues` .</span><span class="sxs-lookup"><span data-stu-id="f6013-114">`Range.getFormulas` and `Range.getNumberFormats` are other frequently used methods that return arrays like `Range.getValues`.</span></span>

<span data-ttu-id="f6013-115">La siguiente secuencia de comandos busca el intervalo **a1: D4** para cualquier formato de número que contenga un "$".</span><span class="sxs-lookup"><span data-stu-id="f6013-115">The following script searches the **A1:D4** range for any number format containing a "$".</span></span> <span data-ttu-id="f6013-116">La secuencia de comandos establece el color de relleno de esas celdas en "amarillo".</span><span class="sxs-lookup"><span data-stu-id="f6013-116">The script sets the fill color in those cells to "yellow".</span></span>

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

### <a name="working-with-collections"></a><span data-ttu-id="f6013-117">Trabajar con colecciones</span><span class="sxs-lookup"><span data-stu-id="f6013-117">Working with collections</span></span>

<span data-ttu-id="f6013-118">Muchos objetos de Excel están incluidos en una colección.</span><span class="sxs-lookup"><span data-stu-id="f6013-118">Many Excel objects are contained in a collection.</span></span> <span data-ttu-id="f6013-119">La colección se administra mediante la API de scripts de Office y se expone como una matriz.</span><span class="sxs-lookup"><span data-stu-id="f6013-119">The collection is managed by the Office Scripts API and exposed as an array.</span></span> <span data-ttu-id="f6013-120">Por ejemplo, todas las [formas](/javascript/api/office-scripts/excelscript/excelscript.shape) de una hoja de cálculo están contenidas en un `Shape[]` devuelto por el `Worksheet.getShapes` método.</span><span class="sxs-lookup"><span data-stu-id="f6013-120">For example, all [Shapes](/javascript/api/office-scripts/excelscript/excelscript.shape) in a worksheet are contained in a `Shape[]` that is returned by the `Worksheet.getShapes` method.</span></span> <span data-ttu-id="f6013-121">Puede usar esta matriz para leer valores de la colección o puede obtener acceso a objetos específicos desde los métodos del objeto primario `get*` .</span><span class="sxs-lookup"><span data-stu-id="f6013-121">You can use this array to read values from the collection, or you can access specific objects from the parent object's `get*` methods.</span></span>

> [!NOTE]
> <span data-ttu-id="f6013-122">No agregue ni quite objetos manualmente de estas matrices de colecciones.</span><span class="sxs-lookup"><span data-stu-id="f6013-122">Do not manually add or remove objects from these collection arrays.</span></span> <span data-ttu-id="f6013-123">Use los `add` métodos de los objetos primarios y los `delete` métodos de los objetos de tipo de colección.</span><span class="sxs-lookup"><span data-stu-id="f6013-123">Use the `add` methods on the parent objects and the `delete` methods on the collection-type objects.</span></span> <span data-ttu-id="f6013-124">Por ejemplo, agregue una [tabla](/javascript/api/office-scripts/excelscript/excelscript.table) a una [hoja de cálculo](/javascript/api/office-scripts/excelscript/excelscript.worksheet) con el `Worksheet.addTable` método y quite el `Table` using `Table.delete` .</span><span class="sxs-lookup"><span data-stu-id="f6013-124">For example, add a [Table](/javascript/api/office-scripts/excelscript/excelscript.table) to a [Worksheet](/javascript/api/office-scripts/excelscript/excelscript.worksheet) with the `Worksheet.addTable` method and remove the `Table` using `Table.delete`.</span></span>

<span data-ttu-id="f6013-125">La siguiente secuencia de comandos registra el tipo de cada forma de la hoja de cálculo actual.</span><span class="sxs-lookup"><span data-stu-id="f6013-125">The following script logs the type of every shape in the current worksheet.</span></span>

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

<span data-ttu-id="f6013-126">La siguiente secuencia de comandos elimina la forma más antigua de la hoja de cálculo actual.</span><span class="sxs-lookup"><span data-stu-id="f6013-126">The following script deletes the oldest shape in the current worksheet.</span></span>

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

## <a name="date"></a><span data-ttu-id="f6013-127">Fecha</span><span class="sxs-lookup"><span data-stu-id="f6013-127">Date</span></span>

<span data-ttu-id="f6013-128">El objeto [Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) proporciona una forma estandarizada de trabajar con fechas en el script.</span><span class="sxs-lookup"><span data-stu-id="f6013-128">The [Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) object provides a standardized way to work with dates in your script.</span></span> <span data-ttu-id="f6013-129">`Date.now()`genera un objeto con la fecha y hora actuales, lo que resulta útil cuando se agregan marcas de tiempo a la entrada de datos del script.</span><span class="sxs-lookup"><span data-stu-id="f6013-129">`Date.now()` generates an object with the current date and time, which is useful when adding timestamps to your script's data entry.</span></span>

<span data-ttu-id="f6013-130">La siguiente secuencia de comandos agrega la fecha actual a la hoja de cálculo.</span><span class="sxs-lookup"><span data-stu-id="f6013-130">The following script adds the current date to the worksheet.</span></span> <span data-ttu-id="f6013-131">Tenga en cuenta que, al usar el `toLocaleDateString` método, Excel reconoce el valor como una fecha y cambia automáticamente el formato de número de la celda.</span><span class="sxs-lookup"><span data-stu-id="f6013-131">Note that by using the `toLocaleDateString` method, Excel recognizes the value as a date and changes the number format of the cell automatically.</span></span>

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

<span data-ttu-id="f6013-132">La sección [work with](../resources/excel-samples.md#work-with-dates) Dates de los ejemplos tiene más scripts relacionados con la fecha.</span><span class="sxs-lookup"><span data-stu-id="f6013-132">The [Work with dates](../resources/excel-samples.md#work-with-dates) section of the samples has more Date-related scripts.</span></span>

## <a name="math"></a><span data-ttu-id="f6013-133">Matemáticas</span><span class="sxs-lookup"><span data-stu-id="f6013-133">Math</span></span>

<span data-ttu-id="f6013-134">El objeto [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math) proporciona métodos y constantes para las operaciones matemáticas comunes.</span><span class="sxs-lookup"><span data-stu-id="f6013-134">The [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math) object provides methods and constants for common mathematical operations.</span></span> <span data-ttu-id="f6013-135">Estos proporcionan muchas funciones que también están disponibles en Excel, sin necesidad de usar el motor de cálculo del libro.</span><span class="sxs-lookup"><span data-stu-id="f6013-135">These provide many functions also available in Excel, without the need to use the workbook's calculation engine.</span></span> <span data-ttu-id="f6013-136">Esto evita que el script tenga que consultar el libro, lo que mejora el rendimiento.</span><span class="sxs-lookup"><span data-stu-id="f6013-136">This saves your script from having to query the workbook, which improves performance.</span></span>

<span data-ttu-id="f6013-137">El siguiente script usa `Math.min` para buscar y registrar el número menor del intervalo de **a1: D4** .</span><span class="sxs-lookup"><span data-stu-id="f6013-137">The following script uses `Math.min` to find and log the smallest number in the **A1:D4** range.</span></span> <span data-ttu-id="f6013-138">Tenga en cuenta que en este ejemplo se supone que el rango completo contiene sólo números, no cadenas.</span><span class="sxs-lookup"><span data-stu-id="f6013-138">Note that this sample assumes the entire range contains only numbers, not strings.</span></span>

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

## <a name="use-of-external-javascript-libraries-is-not-supported"></a><span data-ttu-id="f6013-139">No se admite el uso de bibliotecas de JavaScript externas</span><span class="sxs-lookup"><span data-stu-id="f6013-139">Use of external JavaScript libraries is not supported</span></span>

<span data-ttu-id="f6013-140">Los scripts de Office no admiten el uso de bibliotecas externas de terceros.</span><span class="sxs-lookup"><span data-stu-id="f6013-140">Office Scripts don't support the use of external, third-party libraries.</span></span> <span data-ttu-id="f6013-141">El script solo puede usar los objetos de JavaScript integrados y las API de scripts de Office.</span><span class="sxs-lookup"><span data-stu-id="f6013-141">Your script can only use the built-in JavaScript objects and the Office Scripts APIs.</span></span>

## <a name="see-also"></a><span data-ttu-id="f6013-142">Ver también</span><span class="sxs-lookup"><span data-stu-id="f6013-142">See also</span></span>

- [<span data-ttu-id="f6013-143">Objetos integrados estándar</span><span class="sxs-lookup"><span data-stu-id="f6013-143">Standard built-in objects</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects)
- [<span data-ttu-id="f6013-144">Entorno de editor de código de scripts de Office</span><span class="sxs-lookup"><span data-stu-id="f6013-144">Office Scripts Code Editor environment</span></span>](../overview/code-editor-environment.md)
