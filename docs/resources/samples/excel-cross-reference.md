---
title: Archivos de Excel de referencia cruzada con Power Automate
description: Obtenga información sobre cómo usar scripts de Office y Power Automate para hacer referencia cruzada y dar formato a un archivo Excel.
ms.date: 06/06/2022
ms.localizationpriority: medium
ms.openlocfilehash: 02c06b6376d3726b3e1b44255df14aa64be196ea
ms.sourcegitcommit: f5fc9146d5c096e3a580a3fa8f9714147c548df4
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 06/12/2022
ms.locfileid: "66038675"
---
# <a name="cross-reference-excel-files-with-power-automate"></a>Archivos de Excel de referencia cruzada con Power Automate

Esta solución muestra cómo comparar datos entre dos archivos Excel para encontrar discrepancias. Usa scripts de Office para analizar datos y Power Automate para comunicarse entre los libros.

## <a name="example-scenario"></a>Escenario de ejemplo

Usted es un coordinador de eventos que está programando oradores para las próximas conferencias. Los datos de eventos se conservan en una hoja de cálculo y en los registros del hablante en otra. Para asegurarse de que los dos libros se mantienen sincronizados, use un flujo con scripts de Office para resaltar los posibles problemas.

## <a name="sample-excel-files"></a>Archivos de Excel de ejemplo

Descargue los siguientes archivos para obtener libros listos para usar para el ejemplo.

1. <a href="event-data.xlsx">event-data.xlsx</a>
1. <a href="speaker-registrations.xlsx">speaker-registrations.xlsx</a>

Agregue los siguientes scripts para probar el ejemplo usted mismo.

## <a name="sample-code-get-event-data"></a>Código de ejemplo: Obtención de datos de eventos

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

## <a name="sample-code-validate-speaker-registrations"></a>Código de ejemplo: Validación de registros de orador

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

## <a name="power-automate-flow-check-for-inconsistencies-across-the-workbooks"></a>flujo de Power Automate: comprobar si hay incoherencias en los libros

Este flujo extrae la información del evento del primer libro y usa esos datos para validar el segundo libro.

1. Inicie sesión en [Power Automate](https://flow.microsoft.com) y cree un flujo **de nube instantáneo**.
1. Elija **Desencadenar manualmente un flujo** y seleccione **Crear**.
1. Agregue un **nuevo paso** que use el conector **Excel Online (Empresa)** con la acción **Ejecutar script**. Use los siguientes valores para la acción.
    * **Ubicación**: OneDrive para la Empresa
    * **Biblioteca de documentos**: OneDrive
    * **Archivo**: event-data.xlsx ([seleccionado con el selector de archivos](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))
    * **Script**: Obtención de datos de eventos

    :::image type="content" source="../../images/cross-reference-flow-1.png" alt-text="El conector Excel Online (Business) completado para el primer script de Power Automate.":::

1. Agregue un segundo **paso Nuevo** que use el conector **Excel Online (Empresa)** con la acción **Ejecutar script**. Esto usa los valores devueltos del script **Get event data (Obtener datos de evento** ) como entrada para el script **validate event data (Validar datos de eventos** ). Use los siguientes valores para la acción.
    * **Ubicación**: OneDrive para la Empresa
    * **Biblioteca de documentos**: OneDrive
    * **Archivo**: speaker-registration.xlsx ([seleccionado con el selector de archivos](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))
    * **Script**: Validación del registro del hablante
    * **keys**: result (_contenido dinámico del **script de ejecución**_)

    :::image type="content" source="../../images/cross-reference-flow-2.png" alt-text="El conector Excel Online (Business) completado para el segundo script de Power Automate.":::
1. En este ejemplo se usa Outlook como cliente de correo electrónico. Puede usar cualquier conector de correo electrónico Power Automate admita. Agregue un **nuevo paso** que use el conector **de Office 365 Outlook** y la acción **Enviar y enviar correo electrónico (V2).** Esto usa los valores devueltos del script **Validate speaker registration (Validar el script de registro del hablante** ) como contenido del cuerpo del correo electrónico. Use los siguientes valores para la acción.
    * **Para**: Su cuenta de correo electrónico de prueba (o correo electrónico personal)
    * **Asunto**: Resultados de validación de eventos
    * **Cuerpo**: resultado (_contenido dinámico del **script de ejecución 2**_)

    :::image type="content" source="../../images/cross-reference-flow-3.png" alt-text="Conector de Office 365 Outlook completado en Power Automate.":::
1. Guarde el flujo. Use el botón **Probar** de la página del editor de flujo o ejecute el flujo a través de la pestaña **Mis flujos** . Asegúrese de permitir el acceso cuando se le solicite.
1. Debería recibir un correo electrónico que dice "Error de coincidencia encontrado. Los datos requieren su revisión". Esto indica que hay diferencias entre las filas de **speaker-registrations.xlsx** y las filas de **event-data.xlsx**. Abra **speaker-registrations.xlsx** para ver varias celdas resaltadas donde hay posibles problemas con las listas de registro del hablante.
