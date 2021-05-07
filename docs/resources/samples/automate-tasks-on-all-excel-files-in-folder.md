---
title: Ejecutar un script en todos los archivos de Excel de una carpeta
description: Obtenga información sobre cómo ejecutar un script en todos los Excel archivos de una carpeta en OneDrive para la Empresa.
ms.date: 04/28/2021
localization_priority: Normal
ms.openlocfilehash: a6b869e2b346635e2b28fa7c6273c1a86a5bc5c5
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232630"
---
# <a name="run-a-script-on-all-excel-files-in-a-folder"></a><span data-ttu-id="62904-103">Ejecutar un script en todos los archivos de Excel de una carpeta</span><span class="sxs-lookup"><span data-stu-id="62904-103">Run a script on all Excel files in a folder</span></span>

<span data-ttu-id="62904-104">Este proyecto realiza un conjunto de tareas de automatización en todos los archivos situados en una carpeta de OneDrive para la Empresa.</span><span class="sxs-lookup"><span data-stu-id="62904-104">This project performs a set of automation tasks on all files situated in a folder on OneDrive for Business.</span></span> <span data-ttu-id="62904-105">También se puede usar en una SharePoint carpeta.</span><span class="sxs-lookup"><span data-stu-id="62904-105">It could also be used on a SharePoint folder.</span></span>
<span data-ttu-id="62904-106">Realiza cálculos en los archivos Excel, agrega formato e inserta un comentario que @mentions [compañero.](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7)</span><span class="sxs-lookup"><span data-stu-id="62904-106">It performs calculations on the Excel files, adds formatting, and inserts a comment that [@mentions](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) a colleague.</span></span>

<span data-ttu-id="62904-107">Descargue el archivo <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/highlight-alert-excel-files.zip?raw=true">highlight-alert-excel-files.zip</a>, extraiga los archivos en una carpeta titulada **Ventas** usada en este ejemplo y pruébalo usted mismo.</span><span class="sxs-lookup"><span data-stu-id="62904-107">Download the file <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/highlight-alert-excel-files.zip?raw=true">highlight-alert-excel-files.zip</a>, extract the files to a folder titled **Sales** used in this sample, and try it out yourself!</span></span>

## <a name="sample-code-add-formatting-and-insert-comment"></a><span data-ttu-id="62904-108">Código de ejemplo: Agregar formato e insertar comentario</span><span class="sxs-lookup"><span data-stu-id="62904-108">Sample code: Add formatting and insert comment</span></span>

<span data-ttu-id="62904-109">Este es el script que se ejecuta en cada libro individual.</span><span class="sxs-lookup"><span data-stu-id="62904-109">This is the script that runs on each individual workbook.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  let table1 = workbook.getTable("Table1");
  const rowCount = table1.getRowCount();
  if (rowCount === 0) {
    return;
  }
  workbook.getApplication().calculate(ExcelScript.CalculationType.full);

  const amountDueCol = table1.getColumnByName('Amount Due');
  const amountDueValues = amountDueCol.getRangeBetweenHeaderAndTotal().getValues();

  let highestValue = amountDueValues[0][0];
  let row = 0;
  for (let i = 1; i < amountDueValues.length; i++) {
    if (amountDueValues[i][0] > highestValue) {
      highestValue = amountDueValues[i][0];
      row = i;
    }
  }
  // Set fill color to FFFF00 for range in table Table1 cell in row 0 on column "Amount due".
  table1.getColumn("Amount due")
    .getRangeBetweenHeaderAndTotal()
    .getRow(row)
    .getFormat()
    .getFill()
    .setColor("FFFF00");
  let selectedSheet = workbook.getActiveWorksheet();
  // Insert comment at cell InvoiceAmounts!F2.
  workbook.addComment(table1.getColumn("Amount due")
    .getRangeBetweenHeaderAndTotal()
    .getRow(row), {
    mentions: [{
      email: "AdeleV@M365x904181.OnMicrosoft.com",
      id: 0,
      name: "Adele Vance"
    }],
    richContent: "<at id=\"0\">Adele Vance</at> Please review this amount"
  }, ExcelScript.ContentType.mention);
}
```

## <a name="power-automate-flow-run-the-script-on-every-workbook-in-the-folder"></a><span data-ttu-id="62904-110">Power Automate: ejecute el script en todos los libros de la carpeta</span><span class="sxs-lookup"><span data-stu-id="62904-110">Power Automate flow: Run the script on every workbook in the folder</span></span>

<span data-ttu-id="62904-111">Este flujo ejecuta el script en todos los libros de la carpeta "Ventas".</span><span class="sxs-lookup"><span data-stu-id="62904-111">This flow runs the script on every workbook in the "Sales" folder.</span></span>

1. <span data-ttu-id="62904-112">Crear un nuevo **flujo de nube instantánea**.</span><span class="sxs-lookup"><span data-stu-id="62904-112">Create a new **Instant cloud flow**.</span></span>
1. <span data-ttu-id="62904-113">Seleccione **Desencadenar manualmente un flujo y** presione **Crear**.</span><span class="sxs-lookup"><span data-stu-id="62904-113">Select **Manually trigger a flow** and press **Create**.</span></span>
1. <span data-ttu-id="62904-114">Agregue un **nuevo paso que** use el conector **OneDrive para la Empresa** y los archivos de lista en la acción **de carpeta.**</span><span class="sxs-lookup"><span data-stu-id="62904-114">Add a **New step** that uses the **OneDrive for Business** connector and the **List files in folder** action.</span></span>

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-1.png" alt-text="El conector OneDrive para la Empresa en Power Automate":::
1. <span data-ttu-id="62904-116">Seleccione la carpeta "Ventas" con los libros extraídos.</span><span class="sxs-lookup"><span data-stu-id="62904-116">Select the "Sales" folder with the extracted workbooks.</span></span>
1. <span data-ttu-id="62904-117">Para asegurarse de que solo están seleccionados los libros, elija **Nuevo paso** y, a continuación, **seleccione Condición** y establezca los siguientes valores:</span><span class="sxs-lookup"><span data-stu-id="62904-117">To ensure only workbooks are selected, choose **New step**, then select **Condition** and set the following values:</span></span>
    1. <span data-ttu-id="62904-118">**Name** (el OneDrive de nombre de archivo)</span><span class="sxs-lookup"><span data-stu-id="62904-118">**Name** (the OneDrive file name value)</span></span>
    1. <span data-ttu-id="62904-119">"termina con"</span><span class="sxs-lookup"><span data-stu-id="62904-119">"ends with"</span></span>
    1. <span data-ttu-id="62904-120">"xlsx".</span><span class="sxs-lookup"><span data-stu-id="62904-120">"xlsx".</span></span>

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-2.png" alt-text="El Power Automate condición que aplica acciones posteriores a cada archivo":::
1. <span data-ttu-id="62904-122">En la **rama If yes,** agregue **el conector Excel Online (Empresa)** con la acción **Ejecutar script (versión** preliminar).</span><span class="sxs-lookup"><span data-stu-id="62904-122">Under the **If yes** branch, add the **Excel Online (Business)** connector with the **Run script (preview)** action.</span></span> <span data-ttu-id="62904-123">Use los siguientes valores para la acción:</span><span class="sxs-lookup"><span data-stu-id="62904-123">Use the following values for the action:</span></span>
    1. <span data-ttu-id="62904-124">**Ubicación**: OneDrive para la Empresa</span><span class="sxs-lookup"><span data-stu-id="62904-124">**Location**: OneDrive for Business</span></span>
    1. <span data-ttu-id="62904-125">**Biblioteca de documentos**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="62904-125">**Document Library**: OneDrive</span></span>
    1. <span data-ttu-id="62904-126">**Archivo:** **Identificador** (el valor OneDrive id. de archivo)</span><span class="sxs-lookup"><span data-stu-id="62904-126">**File**: **Id** (the OneDrive file ID value)</span></span>
    1. <span data-ttu-id="62904-127">**Script:** el nombre del script</span><span class="sxs-lookup"><span data-stu-id="62904-127">**Script**: Your script name</span></span>

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-3.png" alt-text="El conector Excel online (empresa) completado en Power Automate":::
1. <span data-ttu-id="62904-129">Guarde el flujo y pruébalo.</span><span class="sxs-lookup"><span data-stu-id="62904-129">Save the flow and try it out.</span></span>

## <a name="training-video-run-a-script-on-all-excel-files-in-a-folder"></a><span data-ttu-id="62904-130">Vídeo de aprendizaje: ejecutar un script en todos Excel archivos de una carpeta</span><span class="sxs-lookup"><span data-stu-id="62904-130">Training video: Run a script on all Excel files in a folder</span></span>

<span data-ttu-id="62904-131">[Vea el recorrido de Sudhi Ramamurthy por este ejemplo en YouTube](https://youtu.be/xMg711o7k6w).</span><span class="sxs-lookup"><span data-stu-id="62904-131">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/xMg711o7k6w).</span></span>
