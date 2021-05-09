---
title: Enviar por correo electrónico las imágenes de un Excel gráfico y tabla
description: Obtenga información sobre cómo usar Office scripts y Power Automate para extraer y enviar por correo electrónico las imágenes de un Excel gráfico y tabla.
ms.date: 05/06/2021
localization_priority: Normal
ms.openlocfilehash: f8b52cbf8c19b93c5fc4288fe97775a25e922ab9
ms.sourcegitcommit: 763d341857bcb209b2f2c278a82fdb63d0e18f0a
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 05/08/2021
ms.locfileid: "52285860"
---
# <a name="use-office-scripts-and-power-automate-to-email-images-of-a-chart-and-table"></a><span data-ttu-id="05186-103">Usar Office scripts y Power Automate para enviar por correo electrónico imágenes de un gráfico y una tabla</span><span class="sxs-lookup"><span data-stu-id="05186-103">Use Office Scripts and Power Automate to email images of a chart and table</span></span>

<span data-ttu-id="05186-104">En este ejemplo se Office scripts y Power Automate para crear un gráfico.</span><span class="sxs-lookup"><span data-stu-id="05186-104">This sample uses Office Scripts and Power Automate to create a chart.</span></span> <span data-ttu-id="05186-105">A continuación, envía un correo electrónico a las imágenes del gráfico y su tabla base.</span><span class="sxs-lookup"><span data-stu-id="05186-105">It then emails images of the chart and its base table.</span></span>

## <a name="example-scenario"></a><span data-ttu-id="05186-106">Ejemplo ficticio</span><span class="sxs-lookup"><span data-stu-id="05186-106">Example scenario</span></span>

* <span data-ttu-id="05186-107">Calcule para obtener los resultados más recientes.</span><span class="sxs-lookup"><span data-stu-id="05186-107">Calculate to get latest results.</span></span>
* <span data-ttu-id="05186-108">Crear gráfico.</span><span class="sxs-lookup"><span data-stu-id="05186-108">Create chart.</span></span>
* <span data-ttu-id="05186-109">Obtener imágenes de gráfico y tabla.</span><span class="sxs-lookup"><span data-stu-id="05186-109">Get chart and table images.</span></span>
* <span data-ttu-id="05186-110">Envíe por correo electrónico las imágenes Power Automate.</span><span class="sxs-lookup"><span data-stu-id="05186-110">Email the images with Power Automate.</span></span>

<span data-ttu-id="05186-111">_Datos de entrada_</span><span class="sxs-lookup"><span data-stu-id="05186-111">_Input data_</span></span>

:::image type="content" source="../../images/input-data.png" alt-text="Una hoja de cálculo que muestra una tabla de datos de entrada":::

<span data-ttu-id="05186-113">_Gráfico de salida_</span><span class="sxs-lookup"><span data-stu-id="05186-113">_Output chart_</span></span>

:::image type="content" source="../../images/chart-created.png" alt-text="Gráfico de columnas creado que muestra la cantidad adeudada por el cliente":::

<span data-ttu-id="05186-115">_Correo electrónico que se recibió a través Power Automate flujo_</span><span class="sxs-lookup"><span data-stu-id="05186-115">_Email that was received through Power Automate flow_</span></span>

:::image type="content" source="../../images/email-received.png" alt-text="El correo electrónico enviado por el flujo que muestra Excel gráfico incrustado en el cuerpo":::

## <a name="solution"></a><span data-ttu-id="05186-117">Solución</span><span class="sxs-lookup"><span data-stu-id="05186-117">Solution</span></span>

<span data-ttu-id="05186-118">Esta solución tiene dos partes:</span><span class="sxs-lookup"><span data-stu-id="05186-118">This solution has two parts:</span></span>

1. [<span data-ttu-id="05186-119">Un script Office para calcular y extraer Excel gráfico y tabla</span><span class="sxs-lookup"><span data-stu-id="05186-119">An Office Script to calculate and extract Excel chart and table</span></span>](#sample-code-calculate-and-extract-excel-chart-and-table)
1. <span data-ttu-id="05186-120">Flujo Power Automate para invocar el script y enviar por correo electrónico los resultados.</span><span class="sxs-lookup"><span data-stu-id="05186-120">A Power Automate flow to invoke the script and email the results.</span></span> <span data-ttu-id="05186-121">Para obtener un ejemplo sobre cómo hacerlo, vea [Create an automated workflow with Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).</span><span class="sxs-lookup"><span data-stu-id="05186-121">For an example on how to do this, see [Create an automated workflow with Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).</span></span>

## <a name="sample-code-calculate-and-extract-excel-chart-and-table"></a><span data-ttu-id="05186-122">Código de ejemplo: calcular y extraer Excel gráfico y tabla</span><span class="sxs-lookup"><span data-stu-id="05186-122">Sample code: Calculate and extract Excel chart and table</span></span>

<span data-ttu-id="05186-123">El siguiente script calcula y extrae un Excel gráfico y tabla.</span><span class="sxs-lookup"><span data-stu-id="05186-123">The following script calculates and extracts an Excel chart and table.</span></span>

<span data-ttu-id="05186-124">Descargue el archivo de <a href="email-chart-table.xlsx">email-chart-table.xlsx</a> y ústelo con este script para probarlo usted mismo.</span><span class="sxs-lookup"><span data-stu-id="05186-124">Download the sample file <a href="email-chart-table.xlsx">email-chart-table.xlsx</a> and use it with this script to try it out yourself!</span></span>

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

## <a name="power-automate-flow-email-the-chart-and-table-images"></a><span data-ttu-id="05186-125">Power Automate: enviar por correo electrónico las imágenes del gráfico y de la tabla</span><span class="sxs-lookup"><span data-stu-id="05186-125">Power Automate flow: Email the chart and table images</span></span>

<span data-ttu-id="05186-126">Este flujo ejecuta el script y envía correos electrónicos a las imágenes devueltas.</span><span class="sxs-lookup"><span data-stu-id="05186-126">This flow runs the script and emails the returned images.</span></span>

1. <span data-ttu-id="05186-127">Crear un nuevo **flujo de nube instantánea**.</span><span class="sxs-lookup"><span data-stu-id="05186-127">Create a new **Instant cloud flow**.</span></span>
1. <span data-ttu-id="05186-128">Seleccione **Desencadenar manualmente un flujo y** presione **Crear**.</span><span class="sxs-lookup"><span data-stu-id="05186-128">Select **Manually trigger a flow** and press **Create**.</span></span>
1. <span data-ttu-id="05186-129">Agregue un **nuevo paso** que use el conector Excel **online (empresa)** con la **acción Ejecutar script (versión** preliminar).</span><span class="sxs-lookup"><span data-stu-id="05186-129">Add a **New step** that uses the **Excel Online (Business)** connector with the **Run script (preview)** action.</span></span> <span data-ttu-id="05186-130">Use los siguientes valores para la acción:</span><span class="sxs-lookup"><span data-stu-id="05186-130">Use the following values for the action:</span></span>
    * <span data-ttu-id="05186-131">**Ubicación**: OneDrive para la Empresa</span><span class="sxs-lookup"><span data-stu-id="05186-131">**Location**: OneDrive for Business</span></span>
    * <span data-ttu-id="05186-132">**Biblioteca de documentos**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="05186-132">**Document Library**: OneDrive</span></span>
    * <span data-ttu-id="05186-133">**Archivo:** el libro ([seleccionado con el seleccionador de archivos](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))</span><span class="sxs-lookup"><span data-stu-id="05186-133">**File**: Your workbook ([selected with the file chooser](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))</span></span>
    * <span data-ttu-id="05186-134">**Script:** el nombre del script</span><span class="sxs-lookup"><span data-stu-id="05186-134">**Script**: Your script name</span></span>

    :::image type="content" source="../../images/email-chart-sample-flow-1.png" alt-text="El conector Excel online (empresa) completado en Power Automate":::
1. <span data-ttu-id="05186-136">En este ejemplo se Outlook como cliente de correo electrónico.</span><span class="sxs-lookup"><span data-stu-id="05186-136">This sample uses Outlook as the email client.</span></span> <span data-ttu-id="05186-137">Puede usar cualquier conector de correo Power Automate admite, pero el resto de los pasos supone que eligió Outlook.</span><span class="sxs-lookup"><span data-stu-id="05186-137">You could use any email connector Power Automate supports, but the rest of the steps assume that you chose Outlook.</span></span> <span data-ttu-id="05186-138">Agregue un **nuevo paso** que use el **conector Office 365 Outlook** y la acción Enviar y correo electrónico **(V2).**</span><span class="sxs-lookup"><span data-stu-id="05186-138">Add a **New step** that uses the **Office 365 Outlook** connector and the **Send and email (V2)** action.</span></span> <span data-ttu-id="05186-139">Use los siguientes valores para la acción:</span><span class="sxs-lookup"><span data-stu-id="05186-139">Use the following values for the action:</span></span>
    * <span data-ttu-id="05186-140">**To**: Su cuenta de correo electrónico de prueba (o correo electrónico personal)</span><span class="sxs-lookup"><span data-stu-id="05186-140">**To**: Your test email account (or personal email)</span></span>
    * <span data-ttu-id="05186-141">**Asunto**: Revise los datos del informe</span><span class="sxs-lookup"><span data-stu-id="05186-141">**Subject**: Please Review Report Data</span></span>
    * <span data-ttu-id="05186-142">Para el **campo Cuerpo,** seleccione "Vista de código" ( `</>` ) y escriba lo siguiente:</span><span class="sxs-lookup"><span data-stu-id="05186-142">For the **Body** field, select "Code View" (`</>`) and enter the following:</span></span>

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

    :::image type="content" source="../../images/email-chart-sample-flow-2.png" alt-text="El conector Office 365 Outlook completado en Power Automate":::
1. <span data-ttu-id="05186-144">Guarde el flujo y pruébalo.</span><span class="sxs-lookup"><span data-stu-id="05186-144">Save the flow and try it out.</span></span>

## <a name="training-video-extract-and-email-images-of-chart-and-table"></a><span data-ttu-id="05186-145">Vídeo de aprendizaje: extraer y enviar por correo electrónico imágenes de gráfico y tabla</span><span class="sxs-lookup"><span data-stu-id="05186-145">Training video: Extract and email images of chart and table</span></span>

<span data-ttu-id="05186-146">[Vea el recorrido de Sudhi Ramamurthy por este ejemplo en YouTube](https://youtu.be/152GJyqc-Kw).</span><span class="sxs-lookup"><span data-stu-id="05186-146">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/152GJyqc-Kw).</span></span>
