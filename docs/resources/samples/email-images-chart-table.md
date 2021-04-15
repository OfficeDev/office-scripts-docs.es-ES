---
title: Enviar por correo electrónico las imágenes de un gráfico y tabla de Excel
description: Obtenga información sobre cómo usar Scripts de Office y Power Automate para extraer y enviar por correo electrónico las imágenes de un gráfico y tabla de Excel.
ms.date: 04/01/2021
localization_priority: Normal
ms.openlocfilehash: de3cf16537cb12db45d4d465d367d797d053afc4
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 04/14/2021
ms.locfileid: "51754813"
---
# <a name="use-office-scripts-and-power-automate-to-email-images-of-a-chart-and-table"></a><span data-ttu-id="96828-103">Usar Scripts de Office y Power Automate para enviar por correo electrónico imágenes de un gráfico y una tabla</span><span class="sxs-lookup"><span data-stu-id="96828-103">Use Office Scripts and Power Automate to email images of a chart and table</span></span>

<span data-ttu-id="96828-104">En este ejemplo se usan scripts de Office y Power Automate para crear un gráfico.</span><span class="sxs-lookup"><span data-stu-id="96828-104">This sample uses Office Scripts and Power Automate to create a chart.</span></span> <span data-ttu-id="96828-105">A continuación, envía un correo electrónico a las imágenes del gráfico y su tabla base.</span><span class="sxs-lookup"><span data-stu-id="96828-105">It then emails images of the chart and its base table.</span></span>

## <a name="example-scenario"></a><span data-ttu-id="96828-106">Escenario de ejemplo</span><span class="sxs-lookup"><span data-stu-id="96828-106">Example scenario</span></span>

* <span data-ttu-id="96828-107">Calcule para obtener los resultados más recientes.</span><span class="sxs-lookup"><span data-stu-id="96828-107">Calculate to get latest results.</span></span>
* <span data-ttu-id="96828-108">Crear gráfico.</span><span class="sxs-lookup"><span data-stu-id="96828-108">Create chart.</span></span>
* <span data-ttu-id="96828-109">Obtener imágenes de gráfico y tabla.</span><span class="sxs-lookup"><span data-stu-id="96828-109">Get chart and table images.</span></span>
* <span data-ttu-id="96828-110">Envíe un correo electrónico a las imágenes con Power Automate.</span><span class="sxs-lookup"><span data-stu-id="96828-110">Email the images with Power Automate.</span></span>

<span data-ttu-id="96828-111">_Datos de entrada_</span><span class="sxs-lookup"><span data-stu-id="96828-111">_Input data_</span></span>

:::image type="content" source="../../images/input-data.png" alt-text="Hoja de cálculo que muestra una tabla de datos de entrada.":::

<span data-ttu-id="96828-113">_Gráfico de salida_</span><span class="sxs-lookup"><span data-stu-id="96828-113">_Output chart_</span></span>

:::image type="content" source="../../images/chart-created.png" alt-text="Gráfico de columnas creado que muestra el importe debido por cliente.":::

<span data-ttu-id="96828-115">_Correo electrónico que se recibió a través del flujo de Power Automate_</span><span class="sxs-lookup"><span data-stu-id="96828-115">_Email that was received through Power Automate flow_</span></span>

:::image type="content" source="../../images/email-received.png" alt-text="El correo electrónico enviado por el flujo que muestra el gráfico de Excel incrustado en el cuerpo.":::

## <a name="solution"></a><span data-ttu-id="96828-117">Solución</span><span class="sxs-lookup"><span data-stu-id="96828-117">Solution</span></span>

<span data-ttu-id="96828-118">Esta solución tiene dos partes:</span><span class="sxs-lookup"><span data-stu-id="96828-118">This solution has two parts:</span></span>

1. [<span data-ttu-id="96828-119">Un script de Office para calcular y extraer gráfico y tabla de Excel</span><span class="sxs-lookup"><span data-stu-id="96828-119">An Office Script to calculate and extract Excel chart and table</span></span>](#sample-code-calculate-and-extract-excel-chart-and-table)
1. <span data-ttu-id="96828-120">Flujo de Power Automate para invocar el script y enviar por correo electrónico los resultados.</span><span class="sxs-lookup"><span data-stu-id="96828-120">A Power Automate flow to invoke the script and email the results.</span></span> <span data-ttu-id="96828-121">Para obtener un ejemplo sobre cómo hacerlo, consulte [Create an automated workflow with Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).</span><span class="sxs-lookup"><span data-stu-id="96828-121">For an example on how to do this, see [Create an automated workflow with Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).</span></span>

## <a name="sample-code-calculate-and-extract-excel-chart-and-table"></a><span data-ttu-id="96828-122">Código de ejemplo: calcular y extraer gráfico y tabla de Excel</span><span class="sxs-lookup"><span data-stu-id="96828-122">Sample code: Calculate and extract Excel chart and table</span></span>

<span data-ttu-id="96828-123">El siguiente script calcula y extrae un gráfico y una tabla de Excel.</span><span class="sxs-lookup"><span data-stu-id="96828-123">The following script calculates and extracts an Excel chart and table.</span></span>

<span data-ttu-id="96828-124">Descargue el archivo de <a href="email-chart-table.xlsx">email-chart-table.xlsx</a> y ústelo con este script para probarlo usted mismo.</span><span class="sxs-lookup"><span data-stu-id="96828-124">Download the sample file <a href="email-chart-table.xlsx">email-chart-table.xlsx</a> and use it with this script to try it out yourself!</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): ReportImages {

  workbook.getApplication().calculate(ExcelScript.CalculationType.full);
  
  let sheet1 = workbook.getWorksheet("Sheet1");
  const table = workbook.getWorksheet('InvoiceAmounts').getTables()[0];
  const rows = table.getRange().getTexts();

  const selectColumns = rows.map((row) => {
    return [row[2], row[5]];
  });
  table.setShowTotals(true);
  selectColumns.splice(selectColumns.length-1, 1);
  console.log(selectColumns);

  workbook.getWorksheet('ChartSheet')?.delete();
  const chartSheet = workbook.addWorksheet('ChartSheet');
  const targetRange = updateRange(chartSheet, selectColumns);

  // Insert chart on sheet 'Sheet1'.
  let chart_2 = chartSheet.addChart(ExcelScript.ChartType.columnClustered, targetRange);
  chart_2.setPosition('D1');
  const chartImage = chart_2.getImage();
  const tableImage = table.getRange().getImage();
  return {
    chartImage,
    tableImage
  }
}

function updateRange(sheet: ExcelScript.Worksheet, data: string[][]): ExcelScript.Range {
  const targetRange = sheet.getRange('A1').getResizedRange(data.length-1, data[0].length-1);
  targetRange.setValues(data);
  return targetRange;
}

interface ReportImages {
  chartImage: string
  tableImage: string
}
```

## <a name="training-video-extract-and-email-images-of-chart-and-table"></a><span data-ttu-id="96828-125">Vídeo de aprendizaje: extraer y enviar por correo electrónico imágenes de gráfico y tabla</span><span class="sxs-lookup"><span data-stu-id="96828-125">Training video: Extract and email images of chart and table</span></span>

<span data-ttu-id="96828-126">[![Ver vídeo paso a paso sobre cómo extraer y enviar por correo electrónico imágenes de gráfico y tabla](../../images/charts-image-vid.jpg)](https://youtu.be/152GJyqc-Kw "Vídeo paso a paso sobre cómo extraer y enviar por correo electrónico imágenes de gráfico y tabla")</span><span class="sxs-lookup"><span data-stu-id="96828-126">[![Watch step-by-step video on how to extract and email images of chart and table](../../images/charts-image-vid.jpg)](https://youtu.be/152GJyqc-Kw "Step-by-step video on how to extract and email images of chart and table")</span></span>
