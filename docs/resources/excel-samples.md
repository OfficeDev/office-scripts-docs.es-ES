---
title: Scripts de ejemplo para scripts de Office en Excel en la web
description: Una colección de ejemplos de código para usar con scripts de Office en Excel en la Web.
ms.date: 04/06/2020
localization_priority: Normal
ms.openlocfilehash: abf6b87b63ad027cca8ee5c947b687f54815409c
ms.sourcegitcommit: 0b2232c4c228b14d501edb8bb489fe0e84748b42
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 04/08/2020
ms.locfileid: "43191007"
---
# <a name="sample-scripts-for-office-scripts-in-excel-on-the-web-preview"></a><span data-ttu-id="28774-103">Scripts de ejemplo para scripts de Office en Excel en la web (vista previa)</span><span class="sxs-lookup"><span data-stu-id="28774-103">Sample scripts for Office Scripts in Excel on the web (preview)</span></span>

<span data-ttu-id="28774-104">Los siguientes ejemplos son scripts sencillos que puede probar en sus propios libros.</span><span class="sxs-lookup"><span data-stu-id="28774-104">The following samples are simple scripts for you to try on your own workbooks.</span></span> <span data-ttu-id="28774-105">Para usarlas en Excel en la web:</span><span class="sxs-lookup"><span data-stu-id="28774-105">To use them in Excel on the web:</span></span>

1. <span data-ttu-id="28774-106">Abra la pestaña **Automatizar**.</span><span class="sxs-lookup"><span data-stu-id="28774-106">Open the **Automate** tab.</span></span>
2. <span data-ttu-id="28774-107">Presione el **Editor de código**.</span><span class="sxs-lookup"><span data-stu-id="28774-107">Press **Code Editor**.</span></span>
3. <span data-ttu-id="28774-108">Presione **nueva secuencia de comandos** en el panel de tareas del editor de código.</span><span class="sxs-lookup"><span data-stu-id="28774-108">Press **New Script** in the Code Editor's task pane.</span></span>
4. <span data-ttu-id="28774-109">Reemplace todo el script por el ejemplo de su elección.</span><span class="sxs-lookup"><span data-stu-id="28774-109">Replace the entire script with the sample of your choice.</span></span>
5. <span data-ttu-id="28774-110">Presione **Ejecutar** en el panel de tareas del editor de código.</span><span class="sxs-lookup"><span data-stu-id="28774-110">Press **Run** in the Code Editor's task pane.</span></span>

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="scripting-basics"></a><span data-ttu-id="28774-111">Conceptos básicos de scripting</span><span class="sxs-lookup"><span data-stu-id="28774-111">Scripting basics</span></span>

<span data-ttu-id="28774-112">Estos ejemplos muestran bloques de creación fundamentales para los scripts de Office.</span><span class="sxs-lookup"><span data-stu-id="28774-112">These samples demonstrate fundamental building blocks for Office Scripts.</span></span> <span data-ttu-id="28774-113">Agréguelos a los scripts para ampliar la solución y resolver problemas comunes.</span><span class="sxs-lookup"><span data-stu-id="28774-113">Add these to your scripts to extend your solution and solve common problems.</span></span>

### <a name="read-and-log-one-cell"></a><span data-ttu-id="28774-114">Leer e iniciar sesión en una celda</span><span class="sxs-lookup"><span data-stu-id="28774-114">Read and log one cell</span></span>

<span data-ttu-id="28774-115">En este ejemplo se lee el valor de **a1** y se imprime en la consola.</span><span class="sxs-lookup"><span data-stu-id="28774-115">This sample reads the value of **A1** and prints it to the console.</span></span>

``` TypeScript
async function main(context: Excel.RequestContext) {
  // Get the current worksheet.
  let selectedSheet = context.workbook.worksheets.getActiveWorksheet();

  // Get the value of cell A1.
  let range = selectedSheet.getRange("A1");
  range.load("values");
  await context.sync();

  // Print the value of A1.
  console.log(range.values);
}
```

### <a name="work-with-dates"></a><span data-ttu-id="28774-116">Trabajar con fechas</span><span class="sxs-lookup"><span data-stu-id="28774-116">Work with dates</span></span>

<span data-ttu-id="28774-117">Los ejemplos de esta sección muestran cómo usar el objeto [Date](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date) de JavaScript.</span><span class="sxs-lookup"><span data-stu-id="28774-117">The samples in this section show how to use the JavaScript [Date](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date) object.</span></span>

<span data-ttu-id="28774-118">En el ejemplo siguiente se obtiene la fecha y hora actuales y, a continuación, se escriben los valores en dos celdas de la hoja de cálculo activa.</span><span class="sxs-lookup"><span data-stu-id="28774-118">The following sample gets the current date and time and then writes those values to two cells in the active worksheet.</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the cells at A1 and B1.
  let dateRange = context.workbook.worksheets.getActiveWorksheet().getRange("A1");
  let timeRange = context.workbook.worksheets.getActiveWorksheet().getRange("B1");

  // Get the current date and time with the JavaScript Date object.
  let date = new Date(Date.now());

  // Add the date string to A1.
  dateRange.values = [[date.toLocaleDateString()]];
  
  // Add the time string to B1.
  timeRange.values = [[date.toLocaleTimeString()]];
}
```

<span data-ttu-id="28774-119">El siguiente ejemplo lee una fecha que está almacenada en Excel y la convierte en un objeto Date de JavaScript.</span><span class="sxs-lookup"><span data-stu-id="28774-119">The next sample reads a date that's stored in Excel and translates it to a JavaScript Date object.</span></span> <span data-ttu-id="28774-120">Usa el [número de serie numérico de la fecha](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) como entrada para la fecha de JavaScript.</span><span class="sxs-lookup"><span data-stu-id="28774-120">It uses the [date's numeric serial number](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) as input for the JavaScript Date.</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
  // Read a date at cell A1 from Excel.
  let dateRange = context.workbook.worksheets.getActiveWorksheet().getRange("A1");
  dateRange.load("values");
  await context.sync();

  // Convert the Excel date to a JavaScript Date object.
  let excelDateValue = dateRange.values[0][0];
  let javaScriptDate = new Date(Math.round((excelDateValue - 25569) * 86400 * 1000));
  console.log(javaScriptDate);
}
```

## <a name="display-data"></a><span data-ttu-id="28774-121">Mostrar datos</span><span class="sxs-lookup"><span data-stu-id="28774-121">Display data</span></span>

<span data-ttu-id="28774-122">En estos ejemplos se muestra cómo trabajar con los datos de la hoja de cálculo y proporcionar a los usuarios una vista o organización mejor.</span><span class="sxs-lookup"><span data-stu-id="28774-122">These samples demonstrate how to work with worksheet data and provide users with a better view or organization.</span></span>

### <a name="apply-conditional-formatting"></a><span data-ttu-id="28774-123">Aplicar formato condicional</span><span class="sxs-lookup"><span data-stu-id="28774-123">Apply conditional formatting</span></span>

<span data-ttu-id="28774-124">En este ejemplo se aplica formato condicional al intervalo que se usa actualmente en la hoja de cálculo.</span><span class="sxs-lookup"><span data-stu-id="28774-124">This sample applies conditional formatting to the currently used range in the worksheet.</span></span> <span data-ttu-id="28774-125">El formato condicional es un relleno verde para el 10% de los valores principales.</span><span class="sxs-lookup"><span data-stu-id="28774-125">The conditional formatting is a green fill for the top 10% of values.</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the current worksheet.
  let selectedSheet = context.workbook.worksheets.getActiveWorksheet();

  // Get the used range in the worksheet.
  let range = selectedSheet.getUsedRange();

  // Set the fill color to green for the top 10% of values in the range.
  let conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.topBottom);
  conditionalFormat.topBottom.format.fill.color = "green";
  conditionalFormat.topBottom.rule = {
    rank: 10, // The percentage threshold.
    type: Excel.ConditionalTopBottomCriterionType.topPercent // The type of the top/bottom condition.
  };
}
```

### <a name="create-a-sorted-table"></a><span data-ttu-id="28774-126">Crear una tabla ordenada</span><span class="sxs-lookup"><span data-stu-id="28774-126">Create a sorted table</span></span>

<span data-ttu-id="28774-127">En este ejemplo se crea una tabla a partir del rango usado de la hoja de cálculo actual y, a continuación, se ordena basándose en la primera columna.</span><span class="sxs-lookup"><span data-stu-id="28774-127">This sample creates a table from the current worksheet's used range, then sorts it based on the first column.</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the current worksheet.
  let selectedSheet = context.workbook.worksheets.getActiveWorksheet();

  // Create a table with the used cells.
  let usedRange = selectedSheet.getUsedRange();
  let newTable = selectedSheet.tables.add(usedRange, true);

  // Sort the table using the first column.
  newTable.sort.apply([{ key: 0, ascending: true }]);
}
```

## <a name="collaboration"></a><span data-ttu-id="28774-128">Colaboración</span><span class="sxs-lookup"><span data-stu-id="28774-128">Collaboration</span></span>

<span data-ttu-id="28774-129">En estos ejemplos se muestra cómo trabajar con las características relacionadas con la colaboración de Excel, como los comentarios.</span><span class="sxs-lookup"><span data-stu-id="28774-129">These samples demonstrate how to work with collaboration-related features of Excel, such as comments.</span></span>

### <a name="delete-resolved-comments"></a><span data-ttu-id="28774-130">Eliminar comentarios resueltos</span><span class="sxs-lookup"><span data-stu-id="28774-130">Delete resolved comments</span></span>

<span data-ttu-id="28774-131">Este ejemplo elimina todos los comentarios resueltos de la hoja de cálculo actual.</span><span class="sxs-lookup"><span data-stu-id="28774-131">This sample deletes all resolved comments from the current worksheet.</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the current worksheet.
  let selectedSheet = context.workbook.worksheets.getActiveWorksheet();

  // Get the comments on this worksheet.
  let comments = selectedSheet.comments;
  comments.load("items/resolved");
  await context.sync();

  // Delete the resolved comments.
  comments.items.forEach((comment) => {
      if (comment.resolved) {
          comment.delete();
      }
  });
}
```

## <a name="scenario-samples"></a><span data-ttu-id="28774-132">Ejemplos de escenario</span><span class="sxs-lookup"><span data-stu-id="28774-132">Scenario samples</span></span>

<span data-ttu-id="28774-133">Para obtener ejemplos que muestren soluciones de gran tamaño para el mundo real, visite ejemplos [de escenarios de Office scripts](scenarios/sample-scenario-overview.md).</span><span class="sxs-lookup"><span data-stu-id="28774-133">For samples showcasing larger, real-world solutions, visit [Sample scenarios for Office Scripts](scenarios/sample-scenario-overview.md).</span></span>

## <a name="suggest-new-samples"></a><span data-ttu-id="28774-134">Sugerir nuevos ejemplos</span><span class="sxs-lookup"><span data-stu-id="28774-134">Suggest new samples</span></span>

<span data-ttu-id="28774-135">Agradecemos las sugerencias para los nuevos ejemplos.</span><span class="sxs-lookup"><span data-stu-id="28774-135">We welcome suggestions for new samples.</span></span> <span data-ttu-id="28774-136">Si hay un escenario común que ayudaría a otros programadores de scripts, indíquenos en la sección Comentarios a continuación.</span><span class="sxs-lookup"><span data-stu-id="28774-136">If there is a common scenario that would help other script developers, please tell us in the feedback section below.</span></span>
