---
title: Usar objetos integrados de JavaScript en los scripts de Office
description: Cómo llamar a las API de JavaScript integradas desde un script de Office en Excel en la Web.
ms.date: 04/08/2020
localization_priority: Normal
ms.openlocfilehash: 54cadb6e9ce60e631488bbe7de00c29a6db35eb7
ms.sourcegitcommit: b13dedb5ee2048f0a244aa2294bf2c38697cb62c
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 04/10/2020
ms.locfileid: "43215262"
---
# <a name="using-built-in-javascript-objects-in-office-scripts"></a><span data-ttu-id="8d64c-103">Usar objetos integrados de JavaScript en los scripts de Office</span><span class="sxs-lookup"><span data-stu-id="8d64c-103">Using built-in JavaScript objects in Office Scripts</span></span>

<span data-ttu-id="8d64c-104">JavaScript proporciona varios objetos integrados que puede usar en los scripts de Office, independientemente de si está creando scripts en JavaScript o [TypeScript](../overview/code-editor-environment.md) (un superconjunto de JavaScript).</span><span class="sxs-lookup"><span data-stu-id="8d64c-104">JavaScript provides several built-in objects that you can use in your Office Scripts, regardless of whether you're scripting in JavaScript or [TypeScript](../overview/code-editor-environment.md) (a superset of JavaScript).</span></span> <span data-ttu-id="8d64c-105">En este artículo se describe cómo se pueden usar algunos de los objetos de JavaScript integrados en los scripts de Office para Excel en la Web.</span><span class="sxs-lookup"><span data-stu-id="8d64c-105">This article describes how you can use some of the built-in JavaScript objects in Office Scripts for Excel on the web.</span></span>

> [!NOTE]
> <span data-ttu-id="8d64c-106">Para obtener una lista completa de todos los objetos de JavaScript integrados, consulte el artículo de [objetos integrados estándar](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects) de Mozilla.</span><span class="sxs-lookup"><span data-stu-id="8d64c-106">For a complete list of all built-in JavaScript objects, see Mozilla's [Standard built-in objects](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects) article.</span></span>

## <a name="array"></a><span data-ttu-id="8d64c-107">Matriz</span><span class="sxs-lookup"><span data-stu-id="8d64c-107">Array</span></span>

<span data-ttu-id="8d64c-108">El objeto [array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) proporciona una forma estandarizada de trabajar con matrices en el script.</span><span class="sxs-lookup"><span data-stu-id="8d64c-108">The [Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) object provides a standardized way to work with arrays in your script.</span></span> <span data-ttu-id="8d64c-109">Aunque las matrices son construcciones estándar de JavaScript, están relacionadas con los scripts de Office de dos maneras principales: Ranges y Collections.</span><span class="sxs-lookup"><span data-stu-id="8d64c-109">While arrays are standard JavaScript constructs, they relate to Office Scripts in two major ways: ranges and collections.</span></span>

### <a name="working-with-ranges"></a><span data-ttu-id="8d64c-110">Trabajar con rangos</span><span class="sxs-lookup"><span data-stu-id="8d64c-110">Working with ranges</span></span>

<span data-ttu-id="8d64c-111">Los rangos contienen varias matrices bidimensionales que se asignan directamente a las celdas de ese intervalo.</span><span class="sxs-lookup"><span data-stu-id="8d64c-111">Ranges contain several two-dimensional arrays that directly map to the cells in that range.</span></span> <span data-ttu-id="8d64c-112">Entre ellas se incluyen propiedades `values`como `formulas`, y `numberFormat`.</span><span class="sxs-lookup"><span data-stu-id="8d64c-112">These include properties such as `values`, `formulas`, and `numberFormat`.</span></span> <span data-ttu-id="8d64c-113">Las propiedades de tipo de matriz deben [cargarse](scripting-fundamentals.md#sync-and-load) como cualquier otra propiedad.</span><span class="sxs-lookup"><span data-stu-id="8d64c-113">Array-type properties must be [loaded](scripting-fundamentals.md#sync-and-load) like any other properties.</span></span>

<span data-ttu-id="8d64c-114">La siguiente secuencia de comandos busca el intervalo **a1: D4** para cualquier formato de número que contenga un "$".</span><span class="sxs-lookup"><span data-stu-id="8d64c-114">The following script searches the **A1:D4** range for any number format containing a "$".</span></span> <span data-ttu-id="8d64c-115">La secuencia de comandos establece el color de relleno de esas celdas en "amarillo".</span><span class="sxs-lookup"><span data-stu-id="8d64c-115">The script sets the fill color in those cells to "yellow".</span></span>

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

### <a name="working-with-collections"></a><span data-ttu-id="8d64c-116">Trabajar con colecciones</span><span class="sxs-lookup"><span data-stu-id="8d64c-116">Working with collections</span></span>

<span data-ttu-id="8d64c-117">Muchos objetos de Excel están incluidos en una colección.</span><span class="sxs-lookup"><span data-stu-id="8d64c-117">Many Excel objects are contained in a collection.</span></span> <span data-ttu-id="8d64c-118">Por ejemplo, todas las [formas](/javascript/api/office-scripts/excel/excel.shape) de una hoja de cálculo están contenidas en `Worksheet.shapes` [ShapeCollection](/javascript/api/office-scripts/excel/excel.shapecollection) (como la propiedad).</span><span class="sxs-lookup"><span data-stu-id="8d64c-118">For example, all [Shapes](/javascript/api/office-scripts/excel/excel.shape) in a worksheet are contained in a [ShapeCollection](/javascript/api/office-scripts/excel/excel.shapecollection) (as the `Worksheet.shapes` property).</span></span> <span data-ttu-id="8d64c-119">Cada `*Collection` objeto contiene una `items` propiedad, que es una matriz que almacena los objetos dentro de dicha colección.</span><span class="sxs-lookup"><span data-stu-id="8d64c-119">Each `*Collection` object contains an `items` property, which is an array that stores the objects inside that collection.</span></span> <span data-ttu-id="8d64c-120">Esto puede tratarse como una matriz de JavaScript normal, pero los elementos de la colección deben cargarse primero.</span><span class="sxs-lookup"><span data-stu-id="8d64c-120">This can be treated like a normal JavaScript array, but the items in the collection have to first be loaded.</span></span> <span data-ttu-id="8d64c-121">Si necesita trabajar con una propiedad en cada objeto de la colección, use una instrucción de carga jerárquica (`items/propertyName`).</span><span class="sxs-lookup"><span data-stu-id="8d64c-121">If you need to work with a property on every object in the collection, use a hierarchal load statement (`items/propertyName`).</span></span>

<span data-ttu-id="8d64c-122">La siguiente secuencia de comandos registra el tipo de cada forma de la hoja de cálculo actual.</span><span class="sxs-lookup"><span data-stu-id="8d64c-122">The following script logs the type of every shape in the current worksheet.</span></span>

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

<span data-ttu-id="8d64c-123">Puede cargar objetos individuales de una colección mediante los `getItem` métodos o `getItemAt` .</span><span class="sxs-lookup"><span data-stu-id="8d64c-123">You can load individual objects from a collection using the `getItem` or `getItemAt` methods.</span></span> <span data-ttu-id="8d64c-124">`getItem`Obtiene un objeto mediante un identificador único como un nombre (a menudo, los nombres se especifican en el script).</span><span class="sxs-lookup"><span data-stu-id="8d64c-124">`getItem` gets an object by using a unique identifier like a name (such names are often specified by your script).</span></span> <span data-ttu-id="8d64c-125">`getItemAt`Obtiene un objeto mediante su índice en la colección.</span><span class="sxs-lookup"><span data-stu-id="8d64c-125">`getItemAt` gets an object by using its index in the collection.</span></span> <span data-ttu-id="8d64c-126">Cada llamada debe ir seguida de un `await context.sync();` comando para que se pueda usar el objeto.</span><span class="sxs-lookup"><span data-stu-id="8d64c-126">Either call must be followed by a `await context.sync();` command before the object can be used.</span></span>

<span data-ttu-id="8d64c-127">La siguiente secuencia de comandos elimina la forma más antigua de la hoja de cálculo actual.</span><span class="sxs-lookup"><span data-stu-id="8d64c-127">The following script deletes the oldest shape in the current worksheet.</span></span>

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

## <a name="date"></a><span data-ttu-id="8d64c-128">Fecha</span><span class="sxs-lookup"><span data-stu-id="8d64c-128">Date</span></span>

<span data-ttu-id="8d64c-129">El objeto [Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) proporciona una forma estandarizada de trabajar con fechas en el script.</span><span class="sxs-lookup"><span data-stu-id="8d64c-129">The [Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) object provides a standardized way to work with dates in your script.</span></span> <span data-ttu-id="8d64c-130">`Date.now()`genera un objeto con la fecha y hora actuales, lo que resulta útil cuando se agregan marcas de tiempo a la entrada de datos del script.</span><span class="sxs-lookup"><span data-stu-id="8d64c-130">`Date.now()` generates an object with the current date and time, which is useful when adding timestamps to your script's data entry.</span></span>

<span data-ttu-id="8d64c-131">La siguiente secuencia de comandos agrega la fecha actual a la hoja de cálculo.</span><span class="sxs-lookup"><span data-stu-id="8d64c-131">The following script adds the current date to the worksheet.</span></span> <span data-ttu-id="8d64c-132">Tenga en cuenta que, `toLocaleDateString` al usar el método, Excel reconoce el valor como una fecha y cambia automáticamente el formato de número de la celda.</span><span class="sxs-lookup"><span data-stu-id="8d64c-132">Note that by using the `toLocaleDateString` method, Excel recognizes the value as a date and changes the number format of the cell automatically.</span></span>

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

<span data-ttu-id="8d64c-133">La sección [work with](../resources/excel-samples.md#work-with-dates) Dates de los ejemplos tiene más scripts relacionados con la fecha.</span><span class="sxs-lookup"><span data-stu-id="8d64c-133">The [Work with dates](../resources/excel-samples.md#work-with-dates) section of the samples has more Date-related scripts.</span></span>

## <a name="math"></a><span data-ttu-id="8d64c-134">Matemáticas</span><span class="sxs-lookup"><span data-stu-id="8d64c-134">Math</span></span>

<span data-ttu-id="8d64c-135">El objeto [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math) proporciona métodos y constantes para las operaciones matemáticas comunes.</span><span class="sxs-lookup"><span data-stu-id="8d64c-135">The [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math) object provides methods and constants for common mathematical operations.</span></span> <span data-ttu-id="8d64c-136">Estos proporcionan muchas funciones que también están disponibles en Excel, sin necesidad de usar el motor de cálculo del libro.</span><span class="sxs-lookup"><span data-stu-id="8d64c-136">These provide many functions also available in Excel, without the need to use the workbook's calculation engine.</span></span> <span data-ttu-id="8d64c-137">Esto evita que el script tenga que consultar el libro, lo que mejora el rendimiento.</span><span class="sxs-lookup"><span data-stu-id="8d64c-137">This saves your script from having to query the workbook, which improves performance.</span></span>

<span data-ttu-id="8d64c-138">El siguiente script usa `Math.min` para buscar y registrar el número menor del intervalo de **a1: D4** .</span><span class="sxs-lookup"><span data-stu-id="8d64c-138">The following script uses `Math.min` to find and log the smallest number in the **A1:D4** range.</span></span> <span data-ttu-id="8d64c-139">Tenga en cuenta que en este ejemplo se supone que el rango completo contiene sólo números, no cadenas.</span><span class="sxs-lookup"><span data-stu-id="8d64c-139">Note that this sample assumes the entire range contains only numbers, not strings.</span></span>

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

## <a name="use-of-external-javascript-libraries-is-not-supported"></a><span data-ttu-id="8d64c-140">No se admite el uso de bibliotecas de JavaScript externas</span><span class="sxs-lookup"><span data-stu-id="8d64c-140">Use of external JavaScript libraries is not supported</span></span>

<span data-ttu-id="8d64c-141">Los scripts de Office no admiten el uso de bibliotecas externas de terceros.</span><span class="sxs-lookup"><span data-stu-id="8d64c-141">Office Scripts don't support the use of external, third-party libraries.</span></span> <span data-ttu-id="8d64c-142">El script solo puede usar los objetos de JavaScript integrados y las API de scripts de Office.</span><span class="sxs-lookup"><span data-stu-id="8d64c-142">Your script can only use the built-in JavaScript objects and the Office Scripts APIs.</span></span>

## <a name="see-also"></a><span data-ttu-id="8d64c-143">Vea también</span><span class="sxs-lookup"><span data-stu-id="8d64c-143">See also</span></span>

- [<span data-ttu-id="8d64c-144">Objetos integrados estándar</span><span class="sxs-lookup"><span data-stu-id="8d64c-144">Standard built-in objects</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects)
- [<span data-ttu-id="8d64c-145">Entorno de editor de código de scripts de Office</span><span class="sxs-lookup"><span data-stu-id="8d64c-145">Office Scripts Code Editor environment</span></span>](../overview/code-editor-environment.md)
