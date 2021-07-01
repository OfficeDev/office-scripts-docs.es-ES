---
title: Archivos de Excel referencias cruzadas con Power Automate
description: Obtenga información sobre cómo usar Office scripts y Power Automate para hacer referencia cruzada y dar formato a un Excel archivo.
ms.date: 06/25/2021
localization_priority: Normal
ms.openlocfilehash: 89c4a5fa5dcff21681fa20cd4118447d39d9b6da
ms.sourcegitcommit: a063b3faf6c1b7c294bd6a73e46845b352f2a22d
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 06/29/2021
ms.locfileid: "53202878"
---
# <a name="cross-reference-excel-files-with-power-automate"></a><span data-ttu-id="a12aa-103">Archivos de Excel referencias cruzadas con Power Automate</span><span class="sxs-lookup"><span data-stu-id="a12aa-103">Cross-reference Excel files with Power Automate</span></span>

<span data-ttu-id="a12aa-104">Esta solución muestra cómo comparar datos en dos Excel para encontrar discrepancias.</span><span class="sxs-lookup"><span data-stu-id="a12aa-104">This solution shows how to compare data across two Excel files to find discrepancies.</span></span> <span data-ttu-id="a12aa-105">Usa scripts Office para analizar datos y Power Automate para comunicarse entre los libros.</span><span class="sxs-lookup"><span data-stu-id="a12aa-105">It uses Office Scripts to analyze data and Power Automate to communicate between the workbooks.</span></span>

## <a name="example-scenario"></a><span data-ttu-id="a12aa-106">Ejemplo ficticio</span><span class="sxs-lookup"><span data-stu-id="a12aa-106">Example scenario</span></span>

<span data-ttu-id="a12aa-107">Es un coordinador de eventos que está programando oradores para las próximas conferencias.</span><span class="sxs-lookup"><span data-stu-id="a12aa-107">You're an event coordinator who is scheduling speakers for upcoming conferences.</span></span> <span data-ttu-id="a12aa-108">Los datos del evento se mantienen en una hoja de cálculo y los registros de altavoces en otra.</span><span class="sxs-lookup"><span data-stu-id="a12aa-108">You keep the event data in one spreadsheet and the speaker registrations in another.</span></span> <span data-ttu-id="a12aa-109">Para asegurarse de que los dos libros se mantienen sincronizados, use un flujo con Office scripts para resaltar los posibles problemas.</span><span class="sxs-lookup"><span data-stu-id="a12aa-109">To ensure the two workbooks are kept in sync, you use a flow with Office Scripts to highlight any potential problems.</span></span>

## <a name="sample-excel-files"></a><span data-ttu-id="a12aa-110">Archivos Excel ejemplo</span><span class="sxs-lookup"><span data-stu-id="a12aa-110">Sample Excel files</span></span>

<span data-ttu-id="a12aa-111">Descarga los siguientes archivos usados en esta solución para probarlos tú mismo.</span><span class="sxs-lookup"><span data-stu-id="a12aa-111">Download the following files used in this solution to try it out yourself!</span></span>

1. <span data-ttu-id="a12aa-112"><a href="event-data.xlsx">event-data.xlsx</a></span><span class="sxs-lookup"><span data-stu-id="a12aa-112"><a href="event-data.xlsx">event-data.xlsx</a></span></span>
1. <span data-ttu-id="a12aa-113"><a href="speaker-registrations.xlsx">speaker-registrations.xlsx</a></span><span class="sxs-lookup"><span data-stu-id="a12aa-113"><a href="speaker-registrations.xlsx">speaker-registrations.xlsx</a></span></span>

## <a name="sample-code-get-event-data"></a><span data-ttu-id="a12aa-114">Código de ejemplo: Obtener datos de eventos</span><span class="sxs-lookup"><span data-stu-id="a12aa-114">Sample code: Get event data</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): string {
  // Get the first table in the "Keys" worksheet.
  let table = workbook.getWorksheet('Keys').getTables()[0];

  // Get the rows in the event table.
  let range = table.getRangeBetweenHeaderAndTotal();
  let rows = range.getValues();

  // Save each row as an EventData object. This lets them be passed through Power Automate.
  let records: EventData[] = [];
  for (let row of rows) {
    let [eventId, date, location, capacity] = row;
    records.push({
      eventId: eventId as string,
      date: date as number,
      location: location as string,
      capacity: capacity as number
    })
  }

  // Log the event data to the console and return it for a flow.
  let stringResult = JSON.stringify(records);
  console.log(stringResult);
  return stringResult;
}

// An interface representing a row of event data.
interface EventData {
  eventId: string
  date: number
  location: string
  capacity: number
}
```

## <a name="sample-code-validate-speaker-registrations"></a><span data-ttu-id="a12aa-115">Código de ejemplo: Validar registros de orador</span><span class="sxs-lookup"><span data-stu-id="a12aa-115">Sample code: Validate speaker registrations</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook, keys: string): string {
  // Get the first table in the "Transactions" worksheet.
  let table = workbook.getWorksheet('Transactions').getTables()[0];

  // Clear the existing formatting in the table.
  let range = table.getRangeBetweenHeaderAndTotal();
  range.clear(ExcelScript.ClearApplyTo.formats);

  // Compare the data in the table to the keys passed into the script.
  let keysObject = JSON.parse(keys) as EventData[];
  let speakerSlotsRemaining = keysObject.map(value => value.capacity);
  let overallMatch = true;

  // Iterate over every row looking for differences from the other worksheet.
  let rows = range.getValues();
  for (let i = 0; i < rows.length; i++) {
    let row = rows[i];
    let [eventId, date, location, capacity] = row;
    let match = false;

    // Look at each key provided for a matching Event ID.
    for (let keyIndex = 0; keyIndex < keysObject.length; keyIndex++) {
      let event = keysObject[keyIndex];
      if (event.eventId === eventId) {
        match = true;
        speakerSlotsRemaining[keyIndex]--;
        // If there's a match on the event ID, look for things that don't match and highlight them.
        if (event.date !== date) {
          overallMatch = false;
          range.getCell(i, 1).getFormat()
            .getFill()
            .setColor("FFFF00");
        }
        if (event.location !== location) {
          overallMatch = false;
          range.getCell(i, 2).getFormat()
            .getFill()
            .setColor("FFFF00");
        }

        break;
      }
    }

    // If no matching Event ID is found, highlight the Event ID's cell.
    if (!match) {
      overallMatch = false;
      range.getCell(i, 0).getFormat()
        .getFill()
        .setColor("FFFF00");
    }
  }

  

  // Choose a message to send to the user.
  let returnString = "All the data is in the right order.";
  if (overallMatch === false) {
    returnString = "Mismatch found. Data requires your review.";
  } else if (speakerSlotsRemaining.find(remaining => remaining < 0)){
    returnString = "Event potentially overbooked. Please review."
  }

  console.log("Returning: " + returnString);
  return returnString;
}

// An interface representing a row of event data.
interface EventData {
  eventId: string
  date: number
  location: string
  capacity: number
}
```

## <a name="power-automate-flow-check-for-inconsistencies-across-the-workbooks"></a><span data-ttu-id="a12aa-116">Power Automate de datos: compruebe si hay incoherencias en los libros</span><span class="sxs-lookup"><span data-stu-id="a12aa-116">Power Automate flow: Check for inconsistencies across the workbooks</span></span>

<span data-ttu-id="a12aa-117">Este flujo extrae la información del evento del primer libro y usa los datos para validar el segundo libro.</span><span class="sxs-lookup"><span data-stu-id="a12aa-117">This flow extracts the event information from the first workbook and uses that data to validate the second workbook.</span></span>

1. <span data-ttu-id="a12aa-118">Inicie sesión [Power Automate](https://flow.microsoft.com) y cree un nuevo **flujo de nube instantánea.**</span><span class="sxs-lookup"><span data-stu-id="a12aa-118">Sign into [Power Automate](https://flow.microsoft.com) and create a new **Instant cloud flow**.</span></span>
1. <span data-ttu-id="a12aa-119">Seleccione **Desencadenar manualmente un flujo y** presione **Crear**.</span><span class="sxs-lookup"><span data-stu-id="a12aa-119">Select **Manually trigger a flow** and press **Create**.</span></span>
1. <span data-ttu-id="a12aa-120">Agregue un **paso Nuevo** que use el conector Excel **Online (Empresa)** con la **acción Ejecutar script.**</span><span class="sxs-lookup"><span data-stu-id="a12aa-120">Add a **New step** that uses the **Excel Online (Business)** connector with the **Run script** action.</span></span> <span data-ttu-id="a12aa-121">Use los siguientes valores para la acción:</span><span class="sxs-lookup"><span data-stu-id="a12aa-121">Use the following values for the action:</span></span>
    * <span data-ttu-id="a12aa-122">**Ubicación**: OneDrive para la Empresa</span><span class="sxs-lookup"><span data-stu-id="a12aa-122">**Location**: OneDrive for Business</span></span>
    * <span data-ttu-id="a12aa-123">**Biblioteca de documentos**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="a12aa-123">**Document Library**: OneDrive</span></span>
    * <span data-ttu-id="a12aa-124">**Archivo**: event-data.xlsx ([seleccionado con el seleccionador de archivos](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))</span><span class="sxs-lookup"><span data-stu-id="a12aa-124">**File**: event-data.xlsx ([selected with the file chooser](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))</span></span>
    * <span data-ttu-id="a12aa-125">**Script**: Obtener datos de eventos</span><span class="sxs-lookup"><span data-stu-id="a12aa-125">**Script**: Get event data</span></span>

    :::image type="content" source="../../images/cross-reference-flow-1.png" alt-text="El conector Excel online (empresa) para el primer script de Power Automate.":::

1. <span data-ttu-id="a12aa-127">Agregue un segundo **paso Nuevo** que use el conector Excel **Online (Empresa)** con la **acción Ejecutar script.**</span><span class="sxs-lookup"><span data-stu-id="a12aa-127">Add a second **New step** that uses the **Excel Online (Business)** connector with the **Run script** action.</span></span> <span data-ttu-id="a12aa-128">Use los siguientes valores para la acción:</span><span class="sxs-lookup"><span data-stu-id="a12aa-128">Use the following values for the action:</span></span>
    * <span data-ttu-id="a12aa-129">**Ubicación**: OneDrive para la Empresa</span><span class="sxs-lookup"><span data-stu-id="a12aa-129">**Location**: OneDrive for Business</span></span>
    * <span data-ttu-id="a12aa-130">**Biblioteca de documentos**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="a12aa-130">**Document Library**: OneDrive</span></span>
    * <span data-ttu-id="a12aa-131">**Archivo**: speaker-registration.xlsx ([seleccionado con el seleccionador de archivos](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))</span><span class="sxs-lookup"><span data-stu-id="a12aa-131">**File**: speaker-registration.xlsx ([selected with the file chooser](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))</span></span>
    * <span data-ttu-id="a12aa-132">**Script**: Validar el registro de altavoces</span><span class="sxs-lookup"><span data-stu-id="a12aa-132">**Script**: Validate speaker registration</span></span>

    :::image type="content" source="../../images/cross-reference-flow-2.png" alt-text="El conector Excel online (empresa) para el segundo script de Power Automate.":::
1. <span data-ttu-id="a12aa-134">En este ejemplo se Outlook como cliente de correo electrónico.</span><span class="sxs-lookup"><span data-stu-id="a12aa-134">This sample uses Outlook as the email client.</span></span> <span data-ttu-id="a12aa-135">Puede usar cualquier conector de correo electrónico Power Automate admite.</span><span class="sxs-lookup"><span data-stu-id="a12aa-135">You could use any email connector Power Automate supports.</span></span> <span data-ttu-id="a12aa-136">Agregue un **nuevo paso** que use el **conector Office 365 Outlook** y la acción Enviar y correo electrónico **(V2).**</span><span class="sxs-lookup"><span data-stu-id="a12aa-136">Add a **New step** that uses the **Office 365 Outlook** connector and the **Send and email (V2)** action.</span></span> <span data-ttu-id="a12aa-137">Use los siguientes valores para la acción:</span><span class="sxs-lookup"><span data-stu-id="a12aa-137">Use the following values for the action:</span></span>
    * <span data-ttu-id="a12aa-138">**To**: Su cuenta de correo electrónico de prueba (o correo electrónico personal)</span><span class="sxs-lookup"><span data-stu-id="a12aa-138">**To**: Your test email account (or personal email)</span></span>
    * <span data-ttu-id="a12aa-139">**Asunto**: Resultados de validación de eventos</span><span class="sxs-lookup"><span data-stu-id="a12aa-139">**Subject**: Event validation results</span></span>
    * <span data-ttu-id="a12aa-140">**Body**: result (_dynamic content from Run script **2**_)</span><span class="sxs-lookup"><span data-stu-id="a12aa-140">**Body**: result (_dynamic content from **Run script 2**_)</span></span>

    :::image type="content" source="../../images/cross-reference-flow-3.png" alt-text="El conector Office 365 Outlook completado en Power Automate.":::
1. <span data-ttu-id="a12aa-142">Guarde el flujo y, a continuación, **seleccione Probar** para probarlo. Debe recibir un correo electrónico que diga "Error encontrado.</span><span class="sxs-lookup"><span data-stu-id="a12aa-142">Save the flow, then select **Test** to try it out. You should receive an email saying "Mismatch found.</span></span> <span data-ttu-id="a12aa-143">Los datos requieren su revisión".</span><span class="sxs-lookup"><span data-stu-id="a12aa-143">Data requires your review."</span></span> <span data-ttu-id="a12aa-144">Esto indica que hay diferencias entre las filas de **speaker-registrations.xlsx** y las filas de **event-data.xlsx**.</span><span class="sxs-lookup"><span data-stu-id="a12aa-144">This indicates there are differences between rows in **speaker-registrations.xlsx** and rows in **event-data.xlsx**.</span></span> <span data-ttu-id="a12aa-145">Abra **speaker-registrations.xlsx** para ver varias celdas resaltadas donde hay posibles problemas con las listas de registro de orador.</span><span class="sxs-lookup"><span data-stu-id="a12aa-145">Open **speaker-registrations.xlsx** to see several highlighted cells where there are potential problems with the speaker registration listings.</span></span>
