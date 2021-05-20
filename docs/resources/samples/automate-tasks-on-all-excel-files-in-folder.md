---
title: Ejecutar un script en todos los archivos de Excel de una carpeta
description: Obtén información sobre cómo ejecutar un script en todos los archivos de Excel de una carpeta de OneDrive para la Empresa.
ms.date: 05/03/2021
localization_priority: Normal
ms.openlocfilehash: fb9a4deb01b52ef031cb1ba3400bd6f10de9d9f5
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545797"
---
# <a name="run-a-script-on-all-excel-files-in-a-folder"></a><span data-ttu-id="64e1b-103">Ejecutar un script en todos los archivos de Excel de una carpeta</span><span class="sxs-lookup"><span data-stu-id="64e1b-103">Run a script on all Excel files in a folder</span></span>

<span data-ttu-id="64e1b-104">Este proyecto realiza un conjunto de tareas de automatización en todos los archivos situados en una carpeta de OneDrive para la Empresa.</span><span class="sxs-lookup"><span data-stu-id="64e1b-104">This project performs a set of automation tasks on all files situated in a folder on OneDrive for Business.</span></span> <span data-ttu-id="64e1b-105">También se puede utilizar en una carpeta de SharePoint.</span><span class="sxs-lookup"><span data-stu-id="64e1b-105">It could also be used on a SharePoint folder.</span></span>
<span data-ttu-id="64e1b-106">Realiza cálculos en los archivos Excel, agrega formato e inserta un comentario que [@mentions](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) un colega.</span><span class="sxs-lookup"><span data-stu-id="64e1b-106">It performs calculations on the Excel files, adds formatting, and inserts a comment that [@mentions](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) a colleague.</span></span>

<span data-ttu-id="64e1b-107">Descargar el archivo <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/highlight-alert-excel-files.zip?raw=true">highlight-alert-excel-files.zip</a>, extraer los archivos en una carpeta titulada **Ventas** utilizadas en este ejemplo, y probarlo usted mismo!</span><span class="sxs-lookup"><span data-stu-id="64e1b-107">Download the file <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/highlight-alert-excel-files.zip?raw=true">highlight-alert-excel-files.zip</a>, extract the files to a folder titled **Sales** used in this sample, and try it out yourself!</span></span>

## <a name="sample-code-add-formatting-and-insert-comment"></a><span data-ttu-id="64e1b-108">Código de ejemplo: Agregue el formato y inserte el comentario</span><span class="sxs-lookup"><span data-stu-id="64e1b-108">Sample code: Add formatting and insert comment</span></span>

<span data-ttu-id="64e1b-109">Este es el script que se ejecuta en cada libro individual.</span><span class="sxs-lookup"><span data-stu-id="64e1b-109">This is the script that runs on each individual workbook.</span></span>

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

## <a name="power-automate-flow-run-the-script-on-every-workbook-in-the-folder"></a><span data-ttu-id="64e1b-110">flujo de Power Automate: ejecute el script en cada libro de trabajo de la carpeta</span><span class="sxs-lookup"><span data-stu-id="64e1b-110">Power Automate flow: Run the script on every workbook in the folder</span></span>

<span data-ttu-id="64e1b-111">Este flujo ejecuta el script en cada libro de trabajo de la carpeta "Ventas".</span><span class="sxs-lookup"><span data-stu-id="64e1b-111">This flow runs the script on every workbook in the "Sales" folder.</span></span>

1. <span data-ttu-id="64e1b-112">Cree un nuevo **flujo de nube instantánea.**</span><span class="sxs-lookup"><span data-stu-id="64e1b-112">Create a new **Instant cloud flow**.</span></span>
1. <span data-ttu-id="64e1b-113">Seleccione **Activar manualmente un flujo** y pulse **Crear**.</span><span class="sxs-lookup"><span data-stu-id="64e1b-113">Select **Manually trigger a flow** and press **Create**.</span></span>
1. <span data-ttu-id="64e1b-114">Agregue un **nuevo paso** que use el conector **de OneDrive para la Empresa** y la acción **Listar archivos en carpeta.**</span><span class="sxs-lookup"><span data-stu-id="64e1b-114">Add a **New step** that uses the **OneDrive for Business** connector and the **List files in folder** action.</span></span>

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-1.png" alt-text="El conector de OneDrive para la Empresa completado en Power Automate":::
1. <span data-ttu-id="64e1b-116">Seleccione la carpeta "Ventas" con los libros extraídos.</span><span class="sxs-lookup"><span data-stu-id="64e1b-116">Select the "Sales" folder with the extracted workbooks.</span></span>
1. <span data-ttu-id="64e1b-117">Para asegurarse de que solo se seleccionan libros de trabajo, elija **Nuevo paso** y, a continuación, seleccione **Condición** y establezca los siguientes valores:</span><span class="sxs-lookup"><span data-stu-id="64e1b-117">To ensure only workbooks are selected, choose **New step**, then select **Condition** and set the following values:</span></span>
    1. <span data-ttu-id="64e1b-118">**Nombre** (el valor OneDrive nombre de archivo)</span><span class="sxs-lookup"><span data-stu-id="64e1b-118">**Name** (the OneDrive file name value)</span></span>
    1. <span data-ttu-id="64e1b-119">"Termina con"</span><span class="sxs-lookup"><span data-stu-id="64e1b-119">"ends with"</span></span>
    1. <span data-ttu-id="64e1b-120">"xlsx".</span><span class="sxs-lookup"><span data-stu-id="64e1b-120">"xlsx".</span></span>

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-2.png" alt-text="El bloque de condición Power Automate que aplica acciones posteriores a cada archivo":::
1. <span data-ttu-id="64e1b-122">En la rama **Si sí,** agregue el conector **Excel en línea (empresa)** con la acción **Ejecutar script.**</span><span class="sxs-lookup"><span data-stu-id="64e1b-122">Under the **If yes** branch, add the **Excel Online (Business)** connector with the **Run script** action.</span></span> <span data-ttu-id="64e1b-123">Utilice los siguientes valores para la acción:</span><span class="sxs-lookup"><span data-stu-id="64e1b-123">Use the following values for the action:</span></span>
    1. <span data-ttu-id="64e1b-124">**Ubicación**: OneDrive para la Empresa</span><span class="sxs-lookup"><span data-stu-id="64e1b-124">**Location**: OneDrive for Business</span></span>
    1. <span data-ttu-id="64e1b-125">**Biblioteca de documentos**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="64e1b-125">**Document Library**: OneDrive</span></span>
    1. <span data-ttu-id="64e1b-126">**Archivo**: **Id** (el valor de ID de archivo OneDrive)</span><span class="sxs-lookup"><span data-stu-id="64e1b-126">**File**: **Id** (the OneDrive file ID value)</span></span>
    1. <span data-ttu-id="64e1b-127">**Script**: Su nombre de guión</span><span class="sxs-lookup"><span data-stu-id="64e1b-127">**Script**: Your script name</span></span>

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-3.png" alt-text="El conector Excel online (business) completado en Power Automate":::
1. <span data-ttu-id="64e1b-129">Guarde el flujo y pruébalo.</span><span class="sxs-lookup"><span data-stu-id="64e1b-129">Save the flow and try it out.</span></span>

## <a name="training-video-run-a-script-on-all-excel-files-in-a-folder"></a><span data-ttu-id="64e1b-130">Vídeo de entrenamiento: ejecute un script en todos los archivos Excel de una carpeta</span><span class="sxs-lookup"><span data-stu-id="64e1b-130">Training video: Run a script on all Excel files in a folder</span></span>

<span data-ttu-id="64e1b-131">[Mira a Sudhi Ramamurthy caminar a través de esta muestra en YouTube.](https://youtu.be/xMg711o7k6w)</span><span class="sxs-lookup"><span data-stu-id="64e1b-131">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/xMg711o7k6w).</span></span>
