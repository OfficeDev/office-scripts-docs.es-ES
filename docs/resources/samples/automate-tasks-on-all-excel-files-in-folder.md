---
title: Ejecutar un script en todos los archivos de Excel de una carpeta
description: Obtenga información sobre cómo ejecutar un script en todos los Excel archivos de una carpeta en OneDrive para la Empresa.
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: bf9c0c486dacced5c3017b267ea65dfd215a5197
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 07/07/2021
ms.locfileid: "53313900"
---
# <a name="run-a-script-on-all-excel-files-in-a-folder"></a><span data-ttu-id="2e7b5-103">Ejecutar un script en todos los archivos de Excel de una carpeta</span><span class="sxs-lookup"><span data-stu-id="2e7b5-103">Run a script on all Excel files in a folder</span></span>

<span data-ttu-id="2e7b5-104">Este proyecto realiza un conjunto de tareas de automatización en todos los archivos situados en una carpeta de OneDrive para la Empresa.</span><span class="sxs-lookup"><span data-stu-id="2e7b5-104">This project performs a set of automation tasks on all files situated in a folder on OneDrive for Business.</span></span> <span data-ttu-id="2e7b5-105">También se puede usar en una SharePoint carpeta.</span><span class="sxs-lookup"><span data-stu-id="2e7b5-105">It could also be used on a SharePoint folder.</span></span>
<span data-ttu-id="2e7b5-106">Realiza cálculos en los archivos Excel, agrega formato e inserta un comentario que @mentions [compañero.](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7)</span><span class="sxs-lookup"><span data-stu-id="2e7b5-106">It performs calculations on the Excel files, adds formatting, and inserts a comment that [@mentions](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) a colleague.</span></span>

## <a name="sample-excel-files"></a><span data-ttu-id="2e7b5-107">Archivos Excel ejemplo</span><span class="sxs-lookup"><span data-stu-id="2e7b5-107">Sample Excel files</span></span>

<span data-ttu-id="2e7b5-108">Descargue <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/highlight-alert-excel-files.zip?raw=true">highlight-alert-excel-files.zip</a> todos los libros que necesitará para este ejemplo.</span><span class="sxs-lookup"><span data-stu-id="2e7b5-108">Download <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/highlight-alert-excel-files.zip?raw=true">highlight-alert-excel-files.zip</a> for all the workbooks you'll need for this sample.</span></span> <span data-ttu-id="2e7b5-109">Extraiga esos archivos en una carpeta titulada **Ventas**.</span><span class="sxs-lookup"><span data-stu-id="2e7b5-109">Extract those files to a folder titled **Sales**.</span></span> <span data-ttu-id="2e7b5-110">Agregue el siguiente script a la colección de scripts para probar el ejemplo usted mismo.</span><span class="sxs-lookup"><span data-stu-id="2e7b5-110">Add the following script to your script collection to try the sample yourself!</span></span>

## <a name="sample-code-add-formatting-and-insert-comment"></a><span data-ttu-id="2e7b5-111">Código de ejemplo: Agregar formato e insertar comentario</span><span class="sxs-lookup"><span data-stu-id="2e7b5-111">Sample code: Add formatting and insert comment</span></span>

<span data-ttu-id="2e7b5-112">Este es el script que se ejecuta en cada libro individual.</span><span class="sxs-lookup"><span data-stu-id="2e7b5-112">This is the script that runs on each individual workbook.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the table named "Table1" in the workbook.
  let table1 = workbook.getTable("Table1");

  // If the table is empty, end the script.
  const rowCount = table1.getRowCount();
  if (rowCount === 0) {
    return;
  }

  // Force the workbook to be completely recalculated.
  workbook.getApplication().calculate(ExcelScript.CalculationType.full);

  // Get the "Amount Due" column from the table.
  const amountDueColumn = table1.getColumnByName('Amount Due');
  const amountDueValues = amountDueColumn.getRangeBetweenHeaderAndTotal().getValues();

  // Find the highest amount that's due.
  let highestValue = amountDueValues[0][0];
  let row = 0;
  for (let i = 1; i < amountDueValues.length; i++) {
    if (amountDueValues[i][0] > highestValue) {
      highestValue = amountDueValues[i][0];
      row = i;
    }
  }

  let highestAmountDue = table1.getColumn("Amount due").getRangeBetweenHeaderAndTotal().getRow(row);

  // Set the fill color to yellow for the cell with the highest value in the "Amount Due" column.
  highestAmountDue
    .getFormat()
    .getFill()
    .setColor("FFFF00");

  // Insert an @mention comment in the cell.
  workbook.addComment(highestAmountDue, {
    mentions: [{
      email: "AdeleV@M365x904181.OnMicrosoft.com",
      id: 0,
      name: "Adele Vance"
    }],
    richContent: "<at id=\"0\">Adele Vance</at> Please review this amount"
  }, ExcelScript.ContentType.mention);
}
```

## <a name="power-automate-flow-run-the-script-on-every-workbook-in-the-folder"></a><span data-ttu-id="2e7b5-113">Power Automate: ejecute el script en todos los libros de la carpeta</span><span class="sxs-lookup"><span data-stu-id="2e7b5-113">Power Automate flow: Run the script on every workbook in the folder</span></span>

<span data-ttu-id="2e7b5-114">Este flujo ejecuta el script en todos los libros de la carpeta "Ventas".</span><span class="sxs-lookup"><span data-stu-id="2e7b5-114">This flow runs the script on every workbook in the "Sales" folder.</span></span>

1. <span data-ttu-id="2e7b5-115">Crear un nuevo **flujo de nube instantánea**.</span><span class="sxs-lookup"><span data-stu-id="2e7b5-115">Create a new **Instant cloud flow**.</span></span>
1. <span data-ttu-id="2e7b5-116">Elija **Desencadenar manualmente un flujo y** seleccione **Crear**.</span><span class="sxs-lookup"><span data-stu-id="2e7b5-116">Choose **Manually trigger a flow** and select **Create**.</span></span>
1. <span data-ttu-id="2e7b5-117">Agregue un **nuevo paso que** use el conector **OneDrive para la Empresa** y los archivos de lista en la acción **de carpeta.**</span><span class="sxs-lookup"><span data-stu-id="2e7b5-117">Add a **New step** that uses the **OneDrive for Business** connector and the **List files in folder** action.</span></span>

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-1.png" alt-text="El conector OneDrive para la Empresa completo en Power Automate.":::
1. <span data-ttu-id="2e7b5-119">Seleccione la carpeta "Ventas" con los libros extraídos.</span><span class="sxs-lookup"><span data-stu-id="2e7b5-119">Select the "Sales" folder with the extracted workbooks.</span></span>
1. <span data-ttu-id="2e7b5-120">Para asegurarse de que solo están seleccionados los libros, elija **Nuevo paso** y, a continuación, **seleccione Condición** y establezca los siguientes valores:</span><span class="sxs-lookup"><span data-stu-id="2e7b5-120">To ensure only workbooks are selected, choose **New step**, then select **Condition** and set the following values:</span></span>
    1. <span data-ttu-id="2e7b5-121">**Name** (el OneDrive de nombre de archivo)</span><span class="sxs-lookup"><span data-stu-id="2e7b5-121">**Name** (the OneDrive file name value)</span></span>
    1. <span data-ttu-id="2e7b5-122">"termina con"</span><span class="sxs-lookup"><span data-stu-id="2e7b5-122">"ends with"</span></span>
    1. <span data-ttu-id="2e7b5-123">"xlsx".</span><span class="sxs-lookup"><span data-stu-id="2e7b5-123">"xlsx".</span></span>

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-2.png" alt-text="El Power Automate condición que aplica acciones posteriores a cada archivo.":::
1. <span data-ttu-id="2e7b5-125">En la **rama If yes,** agregue **el conector Excel Online (Empresa)** con la **acción Ejecutar script.**</span><span class="sxs-lookup"><span data-stu-id="2e7b5-125">Under the **If yes** branch, add the **Excel Online (Business)** connector with the **Run script** action.</span></span> <span data-ttu-id="2e7b5-126">Use los siguientes valores para la acción:</span><span class="sxs-lookup"><span data-stu-id="2e7b5-126">Use the following values for the action:</span></span>
    1. <span data-ttu-id="2e7b5-127">**Ubicación**: OneDrive para la Empresa</span><span class="sxs-lookup"><span data-stu-id="2e7b5-127">**Location**: OneDrive for Business</span></span>
    1. <span data-ttu-id="2e7b5-128">**Biblioteca de documentos**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="2e7b5-128">**Document Library**: OneDrive</span></span>
    1. <span data-ttu-id="2e7b5-129">**Archivo:** **Identificador** (el valor OneDrive id. de archivo)</span><span class="sxs-lookup"><span data-stu-id="2e7b5-129">**File**: **Id** (the OneDrive file ID value)</span></span>
    1. <span data-ttu-id="2e7b5-130">**Script:** el nombre del script</span><span class="sxs-lookup"><span data-stu-id="2e7b5-130">**Script**: Your script name</span></span>

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-3.png" alt-text="El conector Excel Online (Empresa) completado en Power Automate.":::
1. <span data-ttu-id="2e7b5-132">Guarde el flujo y pruébalo. Use el **botón Probar** en la página del editor de flujo o ejecute el flujo a través de la pestaña **Mis flujos.** Asegúrese de permitir el acceso cuando se le pida.</span><span class="sxs-lookup"><span data-stu-id="2e7b5-132">Save the flow and try it out. Use the **Test** button on the flow editor page or run the flow through your **My flows** tab. Be sure to allow access when prompted.</span></span>

## <a name="training-video-run-a-script-on-all-excel-files-in-a-folder"></a><span data-ttu-id="2e7b5-133">Vídeo de aprendizaje: ejecutar un script en todos Excel archivos de una carpeta</span><span class="sxs-lookup"><span data-stu-id="2e7b5-133">Training video: Run a script on all Excel files in a folder</span></span>

<span data-ttu-id="2e7b5-134">[Vea el recorrido de Sudhi Ramamurthy por este ejemplo en YouTube](https://youtu.be/xMg711o7k6w).</span><span class="sxs-lookup"><span data-stu-id="2e7b5-134">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/xMg711o7k6w).</span></span>
