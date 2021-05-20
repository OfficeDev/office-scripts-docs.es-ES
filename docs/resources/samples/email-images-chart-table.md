---
title: Envíe por correo electrónico las imágenes de un gráfico y una tabla de Excel
description: Aprenda a usar scripts y Power Automate Office para extraer y enviar por correo electrónico las imágenes de un gráfico y una tabla Excel.
ms.date: 05/06/2021
localization_priority: Normal
ms.openlocfilehash: 54b6b67a0f211f2dc6c881bab17ff23220619e6e
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545782"
---
# <a name="use-office-scripts-and-power-automate-to-email-images-of-a-chart-and-table"></a><span data-ttu-id="95b9b-103">Utilice Office Scripts y Power Automate para enviar imágenes por correo electrónico de un gráfico y una tabla</span><span class="sxs-lookup"><span data-stu-id="95b9b-103">Use Office Scripts and Power Automate to email images of a chart and table</span></span>

<span data-ttu-id="95b9b-104">En este ejemplo se usa Office scripts y Power Automate para crear un gráfico.</span><span class="sxs-lookup"><span data-stu-id="95b9b-104">This sample uses Office Scripts and Power Automate to create a chart.</span></span> <span data-ttu-id="95b9b-105">A continuación, envía por correo electrónico imágenes del gráfico y su tabla base.</span><span class="sxs-lookup"><span data-stu-id="95b9b-105">It then emails images of the chart and its base table.</span></span>

## <a name="example-scenario"></a><span data-ttu-id="95b9b-106">Ejemplo ficticio</span><span class="sxs-lookup"><span data-stu-id="95b9b-106">Example scenario</span></span>

* <span data-ttu-id="95b9b-107">Calcule para obtener los últimos resultados.</span><span class="sxs-lookup"><span data-stu-id="95b9b-107">Calculate to get latest results.</span></span>
* <span data-ttu-id="95b9b-108">Crear gráfico.</span><span class="sxs-lookup"><span data-stu-id="95b9b-108">Create chart.</span></span>
* <span data-ttu-id="95b9b-109">Obtenga imágenes de gráficos y tablas.</span><span class="sxs-lookup"><span data-stu-id="95b9b-109">Get chart and table images.</span></span>
* <span data-ttu-id="95b9b-110">Envíe por correo electrónico las imágenes con Power Automate.</span><span class="sxs-lookup"><span data-stu-id="95b9b-110">Email the images with Power Automate.</span></span>

<span data-ttu-id="95b9b-111">_Datos de entrada_</span><span class="sxs-lookup"><span data-stu-id="95b9b-111">_Input data_</span></span>

:::image type="content" source="../../images/input-data.png" alt-text="Una hoja de trabajo que muestra una tabla de datos de entrada":::

<span data-ttu-id="95b9b-113">_Gráfico de salida_</span><span class="sxs-lookup"><span data-stu-id="95b9b-113">_Output chart_</span></span>

:::image type="content" source="../../images/chart-created.png" alt-text="El gráfico de columnas creado que muestra el monto adeudado por el cliente":::

<span data-ttu-id="95b9b-115">_Correo electrónico que se recibió a través de Power Automate flujo_</span><span class="sxs-lookup"><span data-stu-id="95b9b-115">_Email that was received through Power Automate flow_</span></span>

:::image type="content" source="../../images/email-received.png" alt-text="El correo electrónico enviado por el flujo que muestra el gráfico de Excel incrustado en el cuerpo":::

## <a name="solution"></a><span data-ttu-id="95b9b-117">Solución</span><span class="sxs-lookup"><span data-stu-id="95b9b-117">Solution</span></span>

<span data-ttu-id="95b9b-118">Esta solución tiene dos partes:</span><span class="sxs-lookup"><span data-stu-id="95b9b-118">This solution has two parts:</span></span>

1. [<span data-ttu-id="95b9b-119">Un script Office para calcular y extraer Excel gráfico y tabla</span><span class="sxs-lookup"><span data-stu-id="95b9b-119">An Office Script to calculate and extract Excel chart and table</span></span>](#sample-code-calculate-and-extract-excel-chart-and-table)
1. <span data-ttu-id="95b9b-120">Un flujo Power Automate para invocar el script y enviar por correo electrónico los resultados.</span><span class="sxs-lookup"><span data-stu-id="95b9b-120">A Power Automate flow to invoke the script and email the results.</span></span> <span data-ttu-id="95b9b-121">Para obtener un ejemplo sobre cómo hacerlo, consulte [Crear un flujo de trabajo automatizado con Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).</span><span class="sxs-lookup"><span data-stu-id="95b9b-121">For an example on how to do this, see [Create an automated workflow with Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).</span></span>

## <a name="sample-code-calculate-and-extract-excel-chart-and-table"></a><span data-ttu-id="95b9b-122">Código de muestra: calcule y extraiga Excel gráfico y tabla</span><span class="sxs-lookup"><span data-stu-id="95b9b-122">Sample code: Calculate and extract Excel chart and table</span></span>

<span data-ttu-id="95b9b-123">El siguiente script calcula y extrae un gráfico y una tabla Excel.</span><span class="sxs-lookup"><span data-stu-id="95b9b-123">The following script calculates and extracts an Excel chart and table.</span></span>

<span data-ttu-id="95b9b-124">Descargar el archivo de ejemplo <a href="email-chart-table.xlsx">email-chart-table.xlsx</a> y utilizarlo con este script para probarlo usted mismo!</span><span class="sxs-lookup"><span data-stu-id="95b9b-124">Download the sample file <a href="email-chart-table.xlsx">email-chart-table.xlsx</a> and use it with this script to try it out yourself!</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): ReportImages {
  // Recalculate the workbook to ensure all tables and charts are updated.
  workbook.getApplication().calculate(ExcelScript.CalculationType.full);
  
  // Get the data from the "InvoiceAmounts" table.
  let sheet1 = workbook.getWorksheet("Sheet1");
  const table = workbook.getWorksheet('InvoiceAmounts').getTables()[0];
  const rows = table.getRange().getTexts();

  // Get only the "Customer Name" and "Amount due" columns, then remove the "Total" row.
  const selectColumns = rows.map((row) => {
    return [row[2], row[5]];
  });
  table.setShowTotals(true);
  selectColumns.splice(selectColumns.length-1, 1);
  console.log(selectColumns);

  // Delete the "ChartSheet" worksheet if it's present, then recreate it.
  workbook.getWorksheet('ChartSheet')?.delete();
  const chartSheet = workbook.addWorksheet('ChartSheet');

  // Add the selected data to the new worksheet.
  const targetRange = chartSheet.getRange('A1').getResizedRange(selectColumns.length-1, selectColumns[0].length-1);
  targetRange.setValues(selectColumns);

  // Insert the chart on sheet 'ChartSheet' at cell "D1".
  let chart_2 = chartSheet.addChart(ExcelScript.ChartType.columnClustered, targetRange);
  chart_2.setPosition('D1');

  // Get images of the chart and table, then return them for a Power Automate flow.
  const chartImage = chart_2.getImage();
  const tableImage = table.getRange().getImage();
  return {chartImage, tableImage};
}

// The interface for table and chart images.
interface ReportImages {
  chartImage: string
  tableImage: string
}
```

## <a name="power-automate-flow-email-the-chart-and-table-images"></a><span data-ttu-id="95b9b-125">flujo Power Automate: Envíe un correo electrónico al gráfico y a las imágenes de la tabla</span><span class="sxs-lookup"><span data-stu-id="95b9b-125">Power Automate flow: Email the chart and table images</span></span>

<span data-ttu-id="95b9b-126">Este flujo ejecuta el script y envía correos electrónicos a las imágenes devueltas.</span><span class="sxs-lookup"><span data-stu-id="95b9b-126">This flow runs the script and emails the returned images.</span></span>

1. <span data-ttu-id="95b9b-127">Cree un nuevo **flujo de nube instantánea.**</span><span class="sxs-lookup"><span data-stu-id="95b9b-127">Create a new **Instant cloud flow**.</span></span>
1. <span data-ttu-id="95b9b-128">Seleccione **Activar manualmente un flujo** y pulse **Crear**.</span><span class="sxs-lookup"><span data-stu-id="95b9b-128">Select **Manually trigger a flow** and press **Create**.</span></span>
1. <span data-ttu-id="95b9b-129">Agregue un **nuevo paso** que use el conector Excel en **línea (empresa)** con la acción **Ejecutar script.**</span><span class="sxs-lookup"><span data-stu-id="95b9b-129">Add a **New step** that uses the **Excel Online (Business)** connector with the **Run script** action.</span></span> <span data-ttu-id="95b9b-130">Utilice los siguientes valores para la acción:</span><span class="sxs-lookup"><span data-stu-id="95b9b-130">Use the following values for the action:</span></span>
    * <span data-ttu-id="95b9b-131">**Ubicación**: OneDrive para la Empresa</span><span class="sxs-lookup"><span data-stu-id="95b9b-131">**Location**: OneDrive for Business</span></span>
    * <span data-ttu-id="95b9b-132">**Biblioteca de documentos**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="95b9b-132">**Document Library**: OneDrive</span></span>
    * <span data-ttu-id="95b9b-133">**Archivo**: Su libro de trabajo ([seleccionado con el selector de archivos](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))</span><span class="sxs-lookup"><span data-stu-id="95b9b-133">**File**: Your workbook ([selected with the file chooser](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))</span></span>
    * <span data-ttu-id="95b9b-134">**Script**: Su nombre de guión</span><span class="sxs-lookup"><span data-stu-id="95b9b-134">**Script**: Your script name</span></span>

    :::image type="content" source="../../images/email-chart-sample-flow-1.png" alt-text="El conector Excel online (business) completado en Power Automate":::
1. <span data-ttu-id="95b9b-136">Este ejemplo utiliza Outlook como cliente de correo electrónico.</span><span class="sxs-lookup"><span data-stu-id="95b9b-136">This sample uses Outlook as the email client.</span></span> <span data-ttu-id="95b9b-137">Puede usar cualquier conector de correo electrónico Power Automate admite, pero el resto de los pasos suponen que eligió Outlook.</span><span class="sxs-lookup"><span data-stu-id="95b9b-137">You could use any email connector Power Automate supports, but the rest of the steps assume that you chose Outlook.</span></span> <span data-ttu-id="95b9b-138">Agregue un **nuevo paso** que use el conector **de Office 365 Outlook** y la acción Enviar y enviar correo **electrónico (V2).**</span><span class="sxs-lookup"><span data-stu-id="95b9b-138">Add a **New step** that uses the **Office 365 Outlook** connector and the **Send and email (V2)** action.</span></span> <span data-ttu-id="95b9b-139">Utilice los siguientes valores para la acción:</span><span class="sxs-lookup"><span data-stu-id="95b9b-139">Use the following values for the action:</span></span>
    * <span data-ttu-id="95b9b-140">**Para**: Su cuenta de correo electrónico de prueba (o correo electrónico personal)</span><span class="sxs-lookup"><span data-stu-id="95b9b-140">**To**: Your test email account (or personal email)</span></span>
    * <span data-ttu-id="95b9b-141">**Asunto**: Revise los datos del informe</span><span class="sxs-lookup"><span data-stu-id="95b9b-141">**Subject**: Please Review Report Data</span></span>
    * <span data-ttu-id="95b9b-142">Para el campo **Cuerpo,** seleccione "Vista de código" ( `</>` ) e introduzca lo siguiente:</span><span class="sxs-lookup"><span data-stu-id="95b9b-142">For the **Body** field, select "Code View" (`</>`) and enter the following:</span></span>

    ```HTML
    <p>Please review the following report data:<br>
    <br>
    Chart:<br>
    <br>
    <img src="data:image/png;base64,@{outputs('Run_script')?['body/result/chartImage']}"/>
    <br>
    Data:<br>
    <br>
    <img src="data:image/png;base64,@{outputs('Run_script')?['body/result/tableImage']}"/>
    <br>
    </p>
    ```

    :::image type="content" source="../../images/email-chart-sample-flow-2.png" alt-text="El conector de Office 365 Outlook completado en Power Automate":::
1. <span data-ttu-id="95b9b-144">Guarde el flujo y pruébalo.</span><span class="sxs-lookup"><span data-stu-id="95b9b-144">Save the flow and try it out.</span></span>

## <a name="training-video-extract-and-email-images-of-chart-and-table"></a><span data-ttu-id="95b9b-145">Vídeo de entrenamiento: Extraer e imágenes de correo electrónico de gráficos y tablas</span><span class="sxs-lookup"><span data-stu-id="95b9b-145">Training video: Extract and email images of chart and table</span></span>

<span data-ttu-id="95b9b-146">[Mira a Sudhi Ramamurthy caminar a través de esta muestra en YouTube.](https://youtu.be/152GJyqc-Kw)</span><span class="sxs-lookup"><span data-stu-id="95b9b-146">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/152GJyqc-Kw).</span></span>
