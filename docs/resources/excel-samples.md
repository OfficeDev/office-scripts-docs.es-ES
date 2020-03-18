---
title: Scripts de ejemplo para scripts de Office en Excel en la web
description: Una colección de ejemplos de código para usar con scripts de Office en Excel en la Web.
ms.date: 02/19/2020
localization_priority: Normal
ms.openlocfilehash: abb4064dfde8b644035e725832e481e6463e979e
ms.sourcegitcommit: b075eed5a6f275274fbbf6d62633219eac416f26
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 03/10/2020
ms.locfileid: "42700416"
---
# <a name="sample-scripts-for-office-scripts-in-excel-on-the-web-preview"></a><span data-ttu-id="71ff8-103">Scripts de ejemplo para scripts de Office en Excel en la web (vista previa)</span><span class="sxs-lookup"><span data-stu-id="71ff8-103">Sample scripts for Office Scripts in Excel on the web (preview)</span></span>

<span data-ttu-id="71ff8-104">Los siguientes ejemplos son scripts sencillos que puede probar en sus propios libros.</span><span class="sxs-lookup"><span data-stu-id="71ff8-104">The following samples are simple scripts for you to try on your own workbooks.</span></span> <span data-ttu-id="71ff8-105">Para usarlas en Excel en la web:</span><span class="sxs-lookup"><span data-stu-id="71ff8-105">To use them in Excel on the web:</span></span>

1. <span data-ttu-id="71ff8-106">Abra la ficha **automatizar** .</span><span class="sxs-lookup"><span data-stu-id="71ff8-106">Open the **Automate** tab.</span></span>
2. <span data-ttu-id="71ff8-107">Presione el **Editor de código**.</span><span class="sxs-lookup"><span data-stu-id="71ff8-107">Press **Code Editor**.</span></span>
3. <span data-ttu-id="71ff8-108">Presione **nueva secuencia de comandos** en el panel de tareas del editor de código.</span><span class="sxs-lookup"><span data-stu-id="71ff8-108">Press **New Script** in the Code Editor's task pane.</span></span>
4. <span data-ttu-id="71ff8-109">Reemplace todo el script por el ejemplo de su elección.</span><span class="sxs-lookup"><span data-stu-id="71ff8-109">Replace the entire script with the sample of your choice.</span></span>
5. <span data-ttu-id="71ff8-110">Presione **Ejecutar** en el panel de tareas del editor de código.</span><span class="sxs-lookup"><span data-stu-id="71ff8-110">Press **Run** in the Code Editor's task pane.</span></span>

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="scripting-basics"></a><span data-ttu-id="71ff8-111">Conceptos básicos de scripting</span><span class="sxs-lookup"><span data-stu-id="71ff8-111">Scripting basics</span></span>

<span data-ttu-id="71ff8-112">Estos ejemplos muestran bloques de creación fundamentales para los scripts de Office.</span><span class="sxs-lookup"><span data-stu-id="71ff8-112">These samples demonstrate fundamental building blocks for Office Scripts.</span></span> <span data-ttu-id="71ff8-113">Agréguelos a los scripts para ampliar la solución y resolver problemas comunes.</span><span class="sxs-lookup"><span data-stu-id="71ff8-113">Add these to your scripts to extend your solution and solve common problems.</span></span>

### <a name="read-and-log-one-cell"></a><span data-ttu-id="71ff8-114">Leer e iniciar sesión en una celda</span><span class="sxs-lookup"><span data-stu-id="71ff8-114">Read and log one cell</span></span>

<span data-ttu-id="71ff8-115">En este ejemplo se lee el valor de **a1** y se imprime en la consola.</span><span class="sxs-lookup"><span data-stu-id="71ff8-115">This sample reads the value of **A1** and prints it to the console.</span></span>

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

### <a name="work-with-dates"></a><span data-ttu-id="71ff8-116">Trabajar con fechas</span><span class="sxs-lookup"><span data-stu-id="71ff8-116">Work with dates</span></span>

<span data-ttu-id="71ff8-117">En este ejemplo se utiliza el objeto [Date](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date) de JavaScript para obtener la fecha y hora actuales y, a continuación, se escriben los valores en dos celdas de la hoja de cálculo activa.</span><span class="sxs-lookup"><span data-stu-id="71ff8-117">This sample uses the JavaScript [Date](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date) object to get the current date and time and then writes those values to two cells in the active worksheet.</span></span>

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

## <a name="display-data"></a><span data-ttu-id="71ff8-118">Mostrar datos</span><span class="sxs-lookup"><span data-stu-id="71ff8-118">Display data</span></span>

<span data-ttu-id="71ff8-119">En estos ejemplos se muestra cómo trabajar con los datos de la hoja de cálculo y proporcionar a los usuarios una vista o organización mejor.</span><span class="sxs-lookup"><span data-stu-id="71ff8-119">These samples demonstrate how to work with worksheet data and provide users with a better view or organization.</span></span>

### <a name="apply-conditional-formatting"></a><span data-ttu-id="71ff8-120">Aplicar formato condicional</span><span class="sxs-lookup"><span data-stu-id="71ff8-120">Apply conditional formatting</span></span>

<span data-ttu-id="71ff8-121">En este ejemplo se aplica formato condicional al intervalo que se usa actualmente en la hoja de cálculo.</span><span class="sxs-lookup"><span data-stu-id="71ff8-121">This sample applies conditional formatting to the currently used range in the worksheet.</span></span> <span data-ttu-id="71ff8-122">El formato condicional es un relleno verde para el 10% de los valores principales.</span><span class="sxs-lookup"><span data-stu-id="71ff8-122">The conditional formatting is a green fill for the top 10% of values.</span></span>

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

### <a name="create-a-sorted-table"></a><span data-ttu-id="71ff8-123">Crear una tabla ordenada</span><span class="sxs-lookup"><span data-stu-id="71ff8-123">Create a sorted table</span></span>

<span data-ttu-id="71ff8-124">En este ejemplo se crea una tabla a partir del rango usado de la hoja de cálculo actual y, a continuación, se ordena basándose en la primera columna.</span><span class="sxs-lookup"><span data-stu-id="71ff8-124">This sample creates a table from the current worksheet's used range, then sorts it based on the first column.</span></span>

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

## <a name="collaboration"></a><span data-ttu-id="71ff8-125">Colaboración</span><span class="sxs-lookup"><span data-stu-id="71ff8-125">Collaboration</span></span>

<span data-ttu-id="71ff8-126">En estos ejemplos se muestra cómo trabajar con las características relacionadas con la colaboración de Excel, como los comentarios.</span><span class="sxs-lookup"><span data-stu-id="71ff8-126">These samples demonstrate how to work with collaboration-related features of Excel, such as comments.</span></span>

### <a name="delete-resolved-comments"></a><span data-ttu-id="71ff8-127">Eliminar comentarios resueltos</span><span class="sxs-lookup"><span data-stu-id="71ff8-127">Delete resolved comments</span></span>

<span data-ttu-id="71ff8-128">Este ejemplo elimina todos los comentarios resueltos de la hoja de cálculo actual.</span><span class="sxs-lookup"><span data-stu-id="71ff8-128">This sample deletes all resolved comments from the current worksheet.</span></span>

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

## <a name="scenario-samples"></a><span data-ttu-id="71ff8-129">Ejemplos de escenario</span><span class="sxs-lookup"><span data-stu-id="71ff8-129">Scenario samples</span></span>

<span data-ttu-id="71ff8-130">Para obtener ejemplos que muestren soluciones de gran tamaño para el mundo real, visite ejemplos [de escenarios de Office scripts](scenarios/sample-scenario-overview.md).</span><span class="sxs-lookup"><span data-stu-id="71ff8-130">For samples showcasing larger, real-world solutions, visit [Sample scenarios for Office Scripts](scenarios/sample-scenario-overview.md).</span></span>

## <a name="suggest-new-samples"></a><span data-ttu-id="71ff8-131">Sugerir nuevos ejemplos</span><span class="sxs-lookup"><span data-stu-id="71ff8-131">Suggest new samples</span></span>

<span data-ttu-id="71ff8-132">Agradecemos las sugerencias para los nuevos ejemplos.</span><span class="sxs-lookup"><span data-stu-id="71ff8-132">We welcome suggestions for new samples.</span></span> <span data-ttu-id="71ff8-133">Si hay un escenario común que ayudaría a otros programadores de scripts, indíquenos en la sección Comentarios a continuación.</span><span class="sxs-lookup"><span data-stu-id="71ff8-133">If there is a common scenario that would help other script developers, please tell us in the feedback section below.</span></span>
