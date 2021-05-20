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
# <a name="use-built-in-javascript-objects-in-office-scripts"></a><span data-ttu-id="a885e-103">Utilice objetos JavaScript integrados en scripts de Office</span><span class="sxs-lookup"><span data-stu-id="a885e-103">Use built-in JavaScript objects in Office Scripts</span></span>

<span data-ttu-id="a885e-104">JavaScript proporciona varios objetos integrados que puede usar en los scripts de Office, independientemente de si está guionando en JavaScript o [TypeScript](../overview/code-editor-environment.md) (un superconjunto de JavaScript).</span><span class="sxs-lookup"><span data-stu-id="a885e-104">JavaScript provides several built-in objects that you can use in your Office Scripts, regardless of whether you're scripting in JavaScript or [TypeScript](../overview/code-editor-environment.md) (a superset of JavaScript).</span></span> <span data-ttu-id="a885e-105">En este artículo se describe cómo puede utilizar algunos de los objetos JavaScript integrados en scripts Office para Excel en la Web.</span><span class="sxs-lookup"><span data-stu-id="a885e-105">This article describes how you can use some of the built-in JavaScript objects in Office Scripts for Excel on the web.</span></span>

> [!NOTE]
> <span data-ttu-id="a885e-106">Para obtener una lista completa de todos los objetos JavaScript integrados, consulte el artículo [Objetos integrados estándar de](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects) Mozilla.</span><span class="sxs-lookup"><span data-stu-id="a885e-106">For a complete list of all built-in JavaScript objects, see Mozilla's [Standard built-in objects](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects) article.</span></span>

## <a name="array"></a><span data-ttu-id="a885e-107">Matriz</span><span class="sxs-lookup"><span data-stu-id="a885e-107">Array</span></span>

<span data-ttu-id="a885e-108">El objeto [Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) proporciona una forma estandarizada de trabajar con matrices en el script.</span><span class="sxs-lookup"><span data-stu-id="a885e-108">The [Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) object provides a standardized way to work with arrays in your script.</span></span> <span data-ttu-id="a885e-109">Aunque las matrices son construcciones JavaScript estándar, se relacionan con scripts Office de dos maneras principales: rangos y colecciones.</span><span class="sxs-lookup"><span data-stu-id="a885e-109">While arrays are standard JavaScript constructs, they relate to Office Scripts in two major ways: ranges and collections.</span></span>

### <a name="work-with-ranges"></a><span data-ttu-id="a885e-110">Trabajar con gamas</span><span class="sxs-lookup"><span data-stu-id="a885e-110">Work with ranges</span></span>

<span data-ttu-id="a885e-111">Los rangos contienen varias matrices bidimensionales que se asignan directamente a las celdas de ese rango.</span><span class="sxs-lookup"><span data-stu-id="a885e-111">Ranges contain several two-dimensional arrays that directly map to the cells in that range.</span></span> <span data-ttu-id="a885e-112">Estas matrices contienen información específica sobre cada celda de ese intervalo.</span><span class="sxs-lookup"><span data-stu-id="a885e-112">These arrays contain specific information about each cell in that range.</span></span> <span data-ttu-id="a885e-113">Por ejemplo, `Range.getValues` devuelve todos los valores de esas celdas (con las filas y columnas de la asignación de matriz bidimensional a las filas y columnas de esa subsección de hoja de cálculo).</span><span class="sxs-lookup"><span data-stu-id="a885e-113">For example, `Range.getValues` returns all the values in those cells (with the rows and columns of the two-dimensional array mapping to the rows and columns of that worksheet subsection).</span></span> <span data-ttu-id="a885e-114">`Range.getFormulas` y `Range.getNumberFormats` son otros métodos utilizados con frecuencia que devuelven matrices como `Range.getValues` .</span><span class="sxs-lookup"><span data-stu-id="a885e-114">`Range.getFormulas` and `Range.getNumberFormats` are other frequently used methods that return arrays like `Range.getValues`.</span></span>

<span data-ttu-id="a885e-115">El siguiente script busca en el intervalo **A1:D4** cualquier formato numérico que contenga un "$".</span><span class="sxs-lookup"><span data-stu-id="a885e-115">The following script searches the **A1:D4** range for any number format containing a "$".</span></span> <span data-ttu-id="a885e-116">El script establece el color de relleno en esas celdas en "amarillo".</span><span class="sxs-lookup"><span data-stu-id="a885e-116">The script sets the fill color in those cells to "yellow".</span></span>

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

### <a name="work-with-collections"></a><span data-ttu-id="a885e-117">Trabajar con colecciones</span><span class="sxs-lookup"><span data-stu-id="a885e-117">Work with collections</span></span>

<span data-ttu-id="a885e-118">Muchos objetos Excel están contenidos en una colección.</span><span class="sxs-lookup"><span data-stu-id="a885e-118">Many Excel objects are contained in a collection.</span></span> <span data-ttu-id="a885e-119">La colección es administrada por la API de scripts de Office y expuesta como una matriz.</span><span class="sxs-lookup"><span data-stu-id="a885e-119">The collection is managed by the Office Scripts API and exposed as an array.</span></span> <span data-ttu-id="a885e-120">Por ejemplo, todas las [formas](/javascript/api/office-scripts/excelscript/excelscript.shape) de una hoja de cálculo están contenidas en una `Shape[]` que devuelve el `Worksheet.getShapes` método.</span><span class="sxs-lookup"><span data-stu-id="a885e-120">For example, all [Shapes](/javascript/api/office-scripts/excelscript/excelscript.shape) in a worksheet are contained in a `Shape[]` that is returned by the `Worksheet.getShapes` method.</span></span> <span data-ttu-id="a885e-121">Puede usar esta matriz para leer valores de la colección o puede tener acceso a objetos específicos desde los métodos del objeto `get*` primario.</span><span class="sxs-lookup"><span data-stu-id="a885e-121">You can use this array to read values from the collection, or you can access specific objects from the parent object's `get*` methods.</span></span>

> [!NOTE]
> <span data-ttu-id="a885e-122">No agregue ni quite manualmente objetos de estas matrices de colección.</span><span class="sxs-lookup"><span data-stu-id="a885e-122">Do not manually add or remove objects from these collection arrays.</span></span> <span data-ttu-id="a885e-123">Utilice los `add` métodos de los objetos primarios y los métodos de los objetos de tipo `delete` colección.</span><span class="sxs-lookup"><span data-stu-id="a885e-123">Use the `add` methods on the parent objects and the `delete` methods on the collection-type objects.</span></span> <span data-ttu-id="a885e-124">Por ejemplo, agregue una [tabla](/javascript/api/office-scripts/excelscript/excelscript.table) a una [hoja de cálculo](/javascript/api/office-scripts/excelscript/excelscript.worksheet) con el método y quite el uso `Worksheet.addTable` `Table` `Table.delete` .</span><span class="sxs-lookup"><span data-stu-id="a885e-124">For example, add a [Table](/javascript/api/office-scripts/excelscript/excelscript.table) to a [Worksheet](/javascript/api/office-scripts/excelscript/excelscript.worksheet) with the `Worksheet.addTable` method and remove the `Table` using `Table.delete`.</span></span>

<span data-ttu-id="a885e-125">El siguiente script registra el tipo de cada forma de la hoja de cálculo actual.</span><span class="sxs-lookup"><span data-stu-id="a885e-125">The following script logs the type of every shape in the current worksheet.</span></span>

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

<span data-ttu-id="a885e-126">El siguiente script elimina la forma más antigua de la hoja de cálculo actual.</span><span class="sxs-lookup"><span data-stu-id="a885e-126">The following script deletes the oldest shape in the current worksheet.</span></span>

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

## <a name="date"></a><span data-ttu-id="a885e-127">Fecha</span><span class="sxs-lookup"><span data-stu-id="a885e-127">Date</span></span>

<span data-ttu-id="a885e-128">El objeto [Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) proporciona una forma estandarizada de trabajar con fechas en el script.</span><span class="sxs-lookup"><span data-stu-id="a885e-128">The [Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) object provides a standardized way to work with dates in your script.</span></span> <span data-ttu-id="a885e-129">`Date.now()` genera un objeto con la fecha y hora actuales, lo que resulta útil al agregar marcas de tiempo a la entrada de datos del script.</span><span class="sxs-lookup"><span data-stu-id="a885e-129">`Date.now()` generates an object with the current date and time, which is useful when adding timestamps to your script's data entry.</span></span>

<span data-ttu-id="a885e-130">El siguiente script agrega la fecha actual a la hoja de cálculo.</span><span class="sxs-lookup"><span data-stu-id="a885e-130">The following script adds the current date to the worksheet.</span></span> <span data-ttu-id="a885e-131">Tenga en cuenta que mediante el uso del `toLocaleDateString` método, Excel reconoce el valor como una fecha y cambia el formato de número de la celda automáticamente.</span><span class="sxs-lookup"><span data-stu-id="a885e-131">Note that by using the `toLocaleDateString` method, Excel recognizes the value as a date and changes the number format of the cell automatically.</span></span>

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

<span data-ttu-id="a885e-132">La sección [Trabajar con fechas](../resources/samples/excel-samples.md#dates) de los ejemplos tiene más scripts relacionados con la fecha.</span><span class="sxs-lookup"><span data-stu-id="a885e-132">The [Work with dates](../resources/samples/excel-samples.md#dates) section of the samples has more date-related scripts.</span></span>

## <a name="math"></a><span data-ttu-id="a885e-133">Matemáticas</span><span class="sxs-lookup"><span data-stu-id="a885e-133">Math</span></span>

<span data-ttu-id="a885e-134">El objeto [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math) proporciona métodos y constantes para operaciones matemáticas comunes.</span><span class="sxs-lookup"><span data-stu-id="a885e-134">The [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math) object provides methods and constants for common mathematical operations.</span></span> <span data-ttu-id="a885e-135">Estos proporcionan muchas funciones también disponibles en Excel, sin necesidad de utilizar el motor de cálculo del libro de trabajo.</span><span class="sxs-lookup"><span data-stu-id="a885e-135">These provide many functions also available in Excel, without the need to use the workbook's calculation engine.</span></span> <span data-ttu-id="a885e-136">Esto evita que el script tenga que consultar el libro, lo que mejora el rendimiento.</span><span class="sxs-lookup"><span data-stu-id="a885e-136">This saves your script from having to query the workbook, which improves performance.</span></span>

<span data-ttu-id="a885e-137">El siguiente script utiliza `Math.min` para buscar y registrar el número más pequeño en el intervalo **A1:D4.**</span><span class="sxs-lookup"><span data-stu-id="a885e-137">The following script uses `Math.min` to find and log the smallest number in the **A1:D4** range.</span></span> <span data-ttu-id="a885e-138">Tenga en cuenta que en este ejemplo se supone que todo el intervalo solo contiene números, no cadenas.</span><span class="sxs-lookup"><span data-stu-id="a885e-138">Note that this sample assumes the entire range contains only numbers, not strings.</span></span>

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

## <a name="use-of-external-javascript-libraries-is-not-supported"></a><span data-ttu-id="a885e-139">No se admite el uso de bibliotecas JavaScript externas</span><span class="sxs-lookup"><span data-stu-id="a885e-139">Use of external JavaScript libraries is not supported</span></span>

<span data-ttu-id="a885e-140">Office Los scripts no admiten el uso de bibliotecas externas de terceros.</span><span class="sxs-lookup"><span data-stu-id="a885e-140">Office Scripts don't support the use of external, third-party libraries.</span></span> <span data-ttu-id="a885e-141">El script solo puede utilizar los objetos JavaScript integrados y las API de scripts de Office.</span><span class="sxs-lookup"><span data-stu-id="a885e-141">Your script can only use the built-in JavaScript objects and the Office Scripts APIs.</span></span>

## <a name="see-also"></a><span data-ttu-id="a885e-142">Vea también</span><span class="sxs-lookup"><span data-stu-id="a885e-142">See also</span></span>

- [<span data-ttu-id="a885e-143">Objetos incorporados estándar</span><span class="sxs-lookup"><span data-stu-id="a885e-143">Standard built-in objects</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects)
- [<span data-ttu-id="a885e-144">Office Entorno del Editor de código de scripts</span><span class="sxs-lookup"><span data-stu-id="a885e-144">Office Scripts Code Editor environment</span></span>](../overview/code-editor-environment.md)
