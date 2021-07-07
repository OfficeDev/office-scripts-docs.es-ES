---
title: Archivos de Excel referencias cruzadas con Power Automate
description: Obtenga información sobre cómo usar Office scripts y Power Automate para hacer referencia cruzada y dar formato a un Excel archivo.
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: 0776ce49cacecfa15339cc7c0cd4866daad789ff
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 07/07/2021
ms.locfileid: "53313963"
---
# <a name="cross-reference-excel-files-with-power-automate"></a><span data-ttu-id="e51d4-103">Archivos de Excel referencias cruzadas con Power Automate</span><span class="sxs-lookup"><span data-stu-id="e51d4-103">Cross-reference Excel files with Power Automate</span></span>

<span data-ttu-id="e51d4-104">Esta solución muestra cómo comparar datos en dos Excel para encontrar discrepancias.</span><span class="sxs-lookup"><span data-stu-id="e51d4-104">This solution shows how to compare data across two Excel files to find discrepancies.</span></span> <span data-ttu-id="e51d4-105">Usa scripts Office para analizar datos y Power Automate para comunicarse entre los libros.</span><span class="sxs-lookup"><span data-stu-id="e51d4-105">It uses Office Scripts to analyze data and Power Automate to communicate between the workbooks.</span></span>

## <a name="example-scenario"></a><span data-ttu-id="e51d4-106">Ejemplo ficticio</span><span class="sxs-lookup"><span data-stu-id="e51d4-106">Example scenario</span></span>

<span data-ttu-id="e51d4-107">Es un coordinador de eventos que está programando oradores para las próximas conferencias.</span><span class="sxs-lookup"><span data-stu-id="e51d4-107">You're an event coordinator who is scheduling speakers for upcoming conferences.</span></span> <span data-ttu-id="e51d4-108">Los datos del evento se mantienen en una hoja de cálculo y los registros de altavoces en otra.</span><span class="sxs-lookup"><span data-stu-id="e51d4-108">You keep the event data in one spreadsheet and the speaker registrations in another.</span></span> <span data-ttu-id="e51d4-109">Para asegurarse de que los dos libros se mantienen sincronizados, use un flujo con Office scripts para resaltar los posibles problemas.</span><span class="sxs-lookup"><span data-stu-id="e51d4-109">To ensure the two workbooks are kept in sync, you use a flow with Office Scripts to highlight any potential problems.</span></span>

## <a name="sample-excel-files"></a><span data-ttu-id="e51d4-110">Archivos Excel ejemplo</span><span class="sxs-lookup"><span data-stu-id="e51d4-110">Sample Excel files</span></span>

<span data-ttu-id="e51d4-111">Descargue los siguientes archivos para obtener libros listos para usar para el ejemplo.</span><span class="sxs-lookup"><span data-stu-id="e51d4-111">Download the following files to get ready-to-use workbooks for the sample.</span></span>

1. <span data-ttu-id="e51d4-112"><a href="event-data.xlsx">event-data.xlsx</a></span><span class="sxs-lookup"><span data-stu-id="e51d4-112"><a href="event-data.xlsx">event-data.xlsx</a></span></span>
1. <span data-ttu-id="e51d4-113"><a href="speaker-registrations.xlsx">speaker-registrations.xlsx</a></span><span class="sxs-lookup"><span data-stu-id="e51d4-113"><a href="speaker-registrations.xlsx">speaker-registrations.xlsx</a></span></span>

<span data-ttu-id="e51d4-114">Agregue los siguientes scripts para probar el ejemplo usted mismo.</span><span class="sxs-lookup"><span data-stu-id="e51d4-114">Add the following scripts to try the sample yourself!</span></span>

## <a name="sample-code-get-event-data"></a><span data-ttu-id="e51d4-115">Código de ejemplo: Obtener datos de eventos</span><span class="sxs-lookup"><span data-stu-id="e51d4-115">Sample code: Get event data</span></span>

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

## <a name="sample-code-validate-speaker-registrations"></a><span data-ttu-id="e51d4-116">Código de ejemplo: Validar registros de orador</span><span class="sxs-lookup"><span data-stu-id="e51d4-116">Sample code: Validate speaker registrations</span></span>

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

## <a name="power-automate-flow-check-for-inconsistencies-across-the-workbooks"></a><span data-ttu-id="e51d4-117">Power Automate de datos: compruebe si hay incoherencias en los libros</span><span class="sxs-lookup"><span data-stu-id="e51d4-117">Power Automate flow: Check for inconsistencies across the workbooks</span></span>

<span data-ttu-id="e51d4-118">Este flujo extrae la información del evento del primer libro y usa los datos para validar el segundo libro.</span><span class="sxs-lookup"><span data-stu-id="e51d4-118">This flow extracts the event information from the first workbook and uses that data to validate the second workbook.</span></span>

1. <span data-ttu-id="e51d4-119">Inicie sesión [Power Automate](https://flow.microsoft.com) y cree un nuevo **flujo de nube instantánea.**</span><span class="sxs-lookup"><span data-stu-id="e51d4-119">Sign into [Power Automate](https://flow.microsoft.com) and create a new **Instant cloud flow**.</span></span>
1. <span data-ttu-id="e51d4-120">Elija **Desencadenar manualmente un flujo y** seleccione **Crear**.</span><span class="sxs-lookup"><span data-stu-id="e51d4-120">Choose **Manually trigger a flow** and select **Create**.</span></span>
1. <span data-ttu-id="e51d4-121">Agregue un **paso Nuevo** que use el conector Excel **Online (Empresa)** con la **acción Ejecutar script.**</span><span class="sxs-lookup"><span data-stu-id="e51d4-121">Add a **New step** that uses the **Excel Online (Business)** connector with the **Run script** action.</span></span> <span data-ttu-id="e51d4-122">Use los siguientes valores para la acción:</span><span class="sxs-lookup"><span data-stu-id="e51d4-122">Use the following values for the action:</span></span>
    * <span data-ttu-id="e51d4-123">**Ubicación**: OneDrive para la Empresa</span><span class="sxs-lookup"><span data-stu-id="e51d4-123">**Location**: OneDrive for Business</span></span>
    * <span data-ttu-id="e51d4-124">**Biblioteca de documentos**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="e51d4-124">**Document Library**: OneDrive</span></span>
    * <span data-ttu-id="e51d4-125">**Archivo**: event-data.xlsx ([seleccionado con el seleccionador de archivos](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))</span><span class="sxs-lookup"><span data-stu-id="e51d4-125">**File**: event-data.xlsx ([selected with the file chooser](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))</span></span>
    * <span data-ttu-id="e51d4-126">**Script**: Obtener datos de eventos</span><span class="sxs-lookup"><span data-stu-id="e51d4-126">**Script**: Get event data</span></span>

    :::image type="content" source="../../images/cross-reference-flow-1.png" alt-text="El conector Excel online (empresa) para el primer script de Power Automate.":::

1. <span data-ttu-id="e51d4-128">Agregue un segundo **paso Nuevo** que use el conector Excel **Online (Empresa)** con la **acción Ejecutar script.**</span><span class="sxs-lookup"><span data-stu-id="e51d4-128">Add a second **New step** that uses the **Excel Online (Business)** connector with the **Run script** action.</span></span> <span data-ttu-id="e51d4-129">Use los siguientes valores para la acción:</span><span class="sxs-lookup"><span data-stu-id="e51d4-129">Use the following values for the action:</span></span>
    * <span data-ttu-id="e51d4-130">**Ubicación**: OneDrive para la Empresa</span><span class="sxs-lookup"><span data-stu-id="e51d4-130">**Location**: OneDrive for Business</span></span>
    * <span data-ttu-id="e51d4-131">**Biblioteca de documentos**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="e51d4-131">**Document Library**: OneDrive</span></span>
    * <span data-ttu-id="e51d4-132">**Archivo**: speaker-registration.xlsx ([seleccionado con el seleccionador de archivos](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))</span><span class="sxs-lookup"><span data-stu-id="e51d4-132">**File**: speaker-registration.xlsx ([selected with the file chooser](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))</span></span>
    * <span data-ttu-id="e51d4-133">**Script**: Validar el registro de altavoces</span><span class="sxs-lookup"><span data-stu-id="e51d4-133">**Script**: Validate speaker registration</span></span>

    :::image type="content" source="../../images/cross-reference-flow-2.png" alt-text="El conector Excel online (empresa) para el segundo script de Power Automate.":::
1. <span data-ttu-id="e51d4-135">En este ejemplo se Outlook como cliente de correo electrónico.</span><span class="sxs-lookup"><span data-stu-id="e51d4-135">This sample uses Outlook as the email client.</span></span> <span data-ttu-id="e51d4-136">Puede usar cualquier conector de correo electrónico Power Automate admite.</span><span class="sxs-lookup"><span data-stu-id="e51d4-136">You could use any email connector Power Automate supports.</span></span> <span data-ttu-id="e51d4-137">Agregue un **nuevo paso** que use el **conector Office 365 Outlook** y la acción Enviar y correo electrónico **(V2).**</span><span class="sxs-lookup"><span data-stu-id="e51d4-137">Add a **New step** that uses the **Office 365 Outlook** connector and the **Send and email (V2)** action.</span></span> <span data-ttu-id="e51d4-138">Use los siguientes valores para la acción:</span><span class="sxs-lookup"><span data-stu-id="e51d4-138">Use the following values for the action:</span></span>
    * <span data-ttu-id="e51d4-139">**To**: Su cuenta de correo electrónico de prueba (o correo electrónico personal)</span><span class="sxs-lookup"><span data-stu-id="e51d4-139">**To**: Your test email account (or personal email)</span></span>
    * <span data-ttu-id="e51d4-140">**Asunto**: Resultados de validación de eventos</span><span class="sxs-lookup"><span data-stu-id="e51d4-140">**Subject**: Event validation results</span></span>
    * <span data-ttu-id="e51d4-141">**Body**: result (_dynamic content from Run script **2**_)</span><span class="sxs-lookup"><span data-stu-id="e51d4-141">**Body**: result (_dynamic content from **Run script 2**_)</span></span>

    :::image type="content" source="../../images/cross-reference-flow-3.png" alt-text="El conector Office 365 Outlook completado en Power Automate.":::
1. <span data-ttu-id="e51d4-143">Guarde el flujo.</span><span class="sxs-lookup"><span data-stu-id="e51d4-143">Save the flow.</span></span> <span data-ttu-id="e51d4-144">Use el **botón Probar** en la página del editor de flujo o ejecute el flujo a través de la pestaña **Mis flujos.** Asegúrese de permitir el acceso cuando se le pida.</span><span class="sxs-lookup"><span data-stu-id="e51d4-144">Use the **Test** button on the flow editor page or run the flow through your **My flows** tab. Be sure to allow access when prompted.</span></span>
1. <span data-ttu-id="e51d4-145">Debe recibir un correo electrónico que diga "Error encontrado.</span><span class="sxs-lookup"><span data-stu-id="e51d4-145">You should receive an email saying "Mismatch found.</span></span> <span data-ttu-id="e51d4-146">Los datos requieren su revisión".</span><span class="sxs-lookup"><span data-stu-id="e51d4-146">Data requires your review."</span></span> <span data-ttu-id="e51d4-147">Esto indica que hay diferencias entre las filas de **speaker-registrations.xlsx** y las filas de **event-data.xlsx**.</span><span class="sxs-lookup"><span data-stu-id="e51d4-147">This indicates there are differences between rows in **speaker-registrations.xlsx** and rows in **event-data.xlsx**.</span></span> <span data-ttu-id="e51d4-148">Abra **speaker-registrations.xlsx** para ver varias celdas resaltadas donde hay posibles problemas con las listas de registro de orador.</span><span class="sxs-lookup"><span data-stu-id="e51d4-148">Open **speaker-registrations.xlsx** to see several highlighted cells where there are potential problems with the speaker registration listings.</span></span>
