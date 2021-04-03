---
title: Enviar por correo electrónico las imágenes de un gráfico y tabla de Excel
description: Obtenga información sobre cómo usar Scripts de Office y Power Automate para extraer y enviar por correo electrónico las imágenes de un gráfico y tabla de Excel.
ms.date: 04/01/2021
localization_priority: Normal
ms.openlocfilehash: 7eb12526f97d72de31acdc3c9a4228c670875e2b
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571576"
---
# <a name="use-office-scripts-and-power-automate-to-email-images-of-a-chart-and-table"></a><span data-ttu-id="48108-103">Usar Scripts de Office y Power Automate para enviar por correo electrónico imágenes de un gráfico y una tabla</span><span class="sxs-lookup"><span data-stu-id="48108-103">Use Office Scripts and Power Automate to email images of a chart and table</span></span>

<span data-ttu-id="48108-104">En este ejemplo se usan scripts de Office y Power Automate para crear un gráfico.</span><span class="sxs-lookup"><span data-stu-id="48108-104">This sample uses Office Scripts and Power Automate to create a chart.</span></span> <span data-ttu-id="48108-105">A continuación, envía un correo electrónico a las imágenes del gráfico y su tabla base.</span><span class="sxs-lookup"><span data-stu-id="48108-105">It then emails images of the chart and its base table.</span></span>

## <a name="example-scenario"></a><span data-ttu-id="48108-106">Escenario de ejemplo</span><span class="sxs-lookup"><span data-stu-id="48108-106">Example scenario</span></span>

* <span data-ttu-id="48108-107">Calcule para obtener los resultados más recientes.</span><span class="sxs-lookup"><span data-stu-id="48108-107">Calculate to get latest results.</span></span>
* <span data-ttu-id="48108-108">Crear gráfico.</span><span class="sxs-lookup"><span data-stu-id="48108-108">Create chart.</span></span>
* <span data-ttu-id="48108-109">Obtener imágenes de gráfico y tabla.</span><span class="sxs-lookup"><span data-stu-id="48108-109">Get chart and table images.</span></span>
* <span data-ttu-id="48108-110">Envíe un correo electrónico a las imágenes con Power Automate.</span><span class="sxs-lookup"><span data-stu-id="48108-110">Email the images with Power Automate.</span></span>

<span data-ttu-id="48108-111">_Datos de entrada_</span><span class="sxs-lookup"><span data-stu-id="48108-111">_Input data_</span></span>

![Datos de entrada](../../images/input-data.png)

<span data-ttu-id="48108-113">_Gráfico de salida_</span><span class="sxs-lookup"><span data-stu-id="48108-113">_Output chart_</span></span>

![Gráfico creado](../../images/chart-created.png)

<span data-ttu-id="48108-115">_Correo electrónico que se recibió a través del flujo de Power Automate_</span><span class="sxs-lookup"><span data-stu-id="48108-115">_Email that was received through Power Automate flow_</span></span>

![Correo electrónico recibido](../../images/email-received.png)

## <a name="solution"></a><span data-ttu-id="48108-117">Solución</span><span class="sxs-lookup"><span data-stu-id="48108-117">Solution</span></span>

<span data-ttu-id="48108-118">Esta solución tiene dos partes:</span><span class="sxs-lookup"><span data-stu-id="48108-118">This solution has two parts:</span></span>

1. [<span data-ttu-id="48108-119">Un script de Office para calcular y extraer gráfico y tabla de Excel</span><span class="sxs-lookup"><span data-stu-id="48108-119">An Office Script to calculate and extract Excel chart and table</span></span>](#sample-code-calculate-and-extract-excel-chart-and-table)
1. <span data-ttu-id="48108-120">Flujo de Power Automate para invocar el script y enviar por correo electrónico los resultados.</span><span class="sxs-lookup"><span data-stu-id="48108-120">A Power Automate flow to invoke the script and email the results.</span></span> <span data-ttu-id="48108-121">Para obtener un ejemplo sobre cómo hacerlo, consulte [Create an automated workflow with Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).</span><span class="sxs-lookup"><span data-stu-id="48108-121">For an example on how to do this, see [Create an automated workflow with Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).</span></span>

## <a name="sample-code-calculate-and-extract-excel-chart-and-table"></a><span data-ttu-id="48108-122">Código de ejemplo: calcular y extraer gráfico y tabla de Excel</span><span class="sxs-lookup"><span data-stu-id="48108-122">Sample code: Calculate and extract Excel chart and table</span></span>

<span data-ttu-id="48108-123">El siguiente script calcula y extrae un gráfico y una tabla de Excel.</span><span class="sxs-lookup"><span data-stu-id="48108-123">The following script calculates and extracts an Excel chart and table.</span></span>

<span data-ttu-id="48108-124">Descargue el archivo de <a href="email-chart-table.xlsx">email-chart-table.xlsx</a> y ústelo con este script para probarlo usted mismo.</span><span class="sxs-lookup"><span data-stu-id="48108-124">Download the sample file <a href="email-chart-table.xlsx">email-chart-table.xlsx</a> and use it with this script to try it out yourself!</span></span>

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

## <a name="training-video-extract-and-email-images-of-chart-and-table"></a><span data-ttu-id="48108-125">Vídeo de aprendizaje: extraer y enviar por correo electrónico imágenes de gráfico y tabla</span><span class="sxs-lookup"><span data-stu-id="48108-125">Training video: Extract and email images of chart and table</span></span>

<span data-ttu-id="48108-126">[![Ver vídeo paso a paso sobre cómo extraer y enviar por correo electrónico imágenes de gráfico y tabla](../../images/charts-image-vid.jpg)](https://youtu.be/152GJyqc-Kw "Vídeo paso a paso sobre cómo extraer y enviar por correo electrónico imágenes de gráfico y tabla")</span><span class="sxs-lookup"><span data-stu-id="48108-126">[![Watch step-by-step video on how to extract and email images of chart and table](../../images/charts-image-vid.jpg)](https://youtu.be/152GJyqc-Kw "Step-by-step video on how to extract and email images of chart and table")</span></span>
