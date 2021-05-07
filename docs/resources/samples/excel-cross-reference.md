---
title: Referencia cruzada y formato de un Excel archivo
description: Obtenga información sobre cómo usar Office scripts y Power Automate para hacer referencia cruzada y dar formato a un Excel archivo.
ms.date: 04/28/2021
localization_priority: Normal
ms.openlocfilehash: 858fe561c1a82f471bc3c0f43d81e457fb02b627
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232385"
---
# <a name="cross-reference-and-format-an-excel-file"></a><span data-ttu-id="4228d-103">Referencia cruzada y formato de un Excel archivo</span><span class="sxs-lookup"><span data-stu-id="4228d-103">Cross-reference and format an Excel file</span></span>

<span data-ttu-id="4228d-104">Esta solución muestra cómo se puede hacer referencia Excel y dar formato a dos archivos de Office scripts y Power Automate.</span><span class="sxs-lookup"><span data-stu-id="4228d-104">This solution shows how two Excel files can be cross-referenced and formatted using Office Scripts and Power Automate.</span></span>

<span data-ttu-id="4228d-105">El proyecto logra lo siguiente:</span><span class="sxs-lookup"><span data-stu-id="4228d-105">The project achieves the following:</span></span>

1. <span data-ttu-id="4228d-106">Extrae datos de eventos de <a href="events.xlsx">events.xlsx</a> mediante una acción ejecutar script.</span><span class="sxs-lookup"><span data-stu-id="4228d-106">Extracts event data from <a href="events.xlsx">events.xlsx</a> using one Run script action.</span></span>
1. <span data-ttu-id="4228d-107">Pasa estos datos al segundo archivo Excel que contiene datos de transacción de eventos y los usa para realizar la validación básica de los datos y el formato de datos que faltan o incorrectos mediante scripts de Office.</span><span class="sxs-lookup"><span data-stu-id="4228d-107">Passes that data to the second Excel file containing event transaction data and uses that data to do basic validation of data and formatting of missing or incorrect data using Office Scripts.</span></span>
1. <span data-ttu-id="4228d-108">Envía el resultado por correo electrónico a un revisor.</span><span class="sxs-lookup"><span data-stu-id="4228d-108">Emails the result to a reviewer.</span></span>

<span data-ttu-id="4228d-109">Para obtener más información, vea [Cross Reference and formatting two Excel files using Office Scripts](https://powerusers.microsoft.com/t5/Power-Automate-Cookbook/Cross-Reference-and-formatting-two-Excel-files-using-Office/td-p/728535).</span><span class="sxs-lookup"><span data-stu-id="4228d-109">For further details, see [Cross Reference and formatting two Excel files using Office Scripts](https://powerusers.microsoft.com/t5/Power-Automate-Cookbook/Cross-Reference-and-formatting-two-Excel-files-using-Office/td-p/728535).</span></span>

## <a name="sample-excel-files"></a><span data-ttu-id="4228d-110">Archivos Excel ejemplo</span><span class="sxs-lookup"><span data-stu-id="4228d-110">Sample Excel files</span></span>

<span data-ttu-id="4228d-111">Descarga los siguientes archivos usados en esta solución para probarlos tú mismo.</span><span class="sxs-lookup"><span data-stu-id="4228d-111">Download the following files used in this solution to try it out yourself!</span></span>

1. <span data-ttu-id="4228d-112"><a href="events.xlsx">events.xlsx</a></span><span class="sxs-lookup"><span data-stu-id="4228d-112"><a href="events.xlsx">events.xlsx</a></span></span>
1. <span data-ttu-id="4228d-113"><a href="event-transactions.xlsx">event-transactions.xlsx</a></span><span class="sxs-lookup"><span data-stu-id="4228d-113"><a href="event-transactions.xlsx">event-transactions.xlsx</a></span></span>

## <a name="sample-code-get-event-data"></a><span data-ttu-id="4228d-114">Código de ejemplo: Obtener datos de eventos</span><span class="sxs-lookup"><span data-stu-id="4228d-114">Sample code: Get event data</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): EventData[] {
    let table = workbook.getWorksheet('Keys').getTables()[0];
    let range = table.getRangeBetweenHeaderAndTotal();
    let rows = range.getValues();
    let records: EventData[] = [];
    for (let row of rows) {
        let [event, date, location, capacity] = row;
        records.push({
            event: event as string,
            date: date as number, 
            location: location as string,
            capacity: capacity as number
        })
    }
    console.log(JSON.stringify(records))
    return records;
}

interface EventData {
    event: string
    date: number
    location: string
    capacity: number
}
```

## <a name="sample-code-validate-event-transactions"></a><span data-ttu-id="4228d-115">Código de ejemplo: Validar transacciones de eventos</span><span class="sxs-lookup"><span data-stu-id="4228d-115">Sample code: Validate event transactions</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook, keys: string): string {
    let table = workbook.getWorksheet('Transactions').getTables()[0];
    let range = table.getRangeBetweenHeaderAndTotal();
    range.clear(ExcelScript.ClearApplyTo.formats);
  
    let overallMatch = true;
  
    table.getColumnByName('Date').getRangeBetweenHeaderAndTotal().setNumberFormatLocal("yyyy-mm-dd;@");
    table.getColumnByName('Capacity').getRangeBetweenHeaderAndTotal().getFormat()
      .setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
    let rows = range.getValues();
    let keysObject = JSON.parse(keys) as EventData[];
    for (let i=0; i < rows.length; i++){
      let row = rows[i];
      let [event, date, location, capacity] = row;
      let match = false;
      for (let keyObject of keysObject){
        if (keyObject.event === event) {
          match = true;
          if (keyObject.date !== date) {
            overallMatch = false;
            range.getCell(i, 1).getFormat()
              .getFill()
              .setColor("FFFF00");
          }
          if (keyObject.location !== location) {
            overallMatch = false;
            range.getCell(i, 2).getFormat()
              .getFill()
              .setColor("FFFF00");
          }
          if (keyObject.capacity !== capacity) {
            overallMatch = false;
            range.getCell(i, 3).getFormat()
              .getFill()
              .setColor("FFFF00");
          }   
          break;             
        }
      }
      if (!match) {
        overallMatch = false;
        range.getCell(i, 0).getFormat()
          .getFill()
          .setColor("FFFF00");      
      }
  
    }
    let returnString = "All the data is in the right order.";
    if (overallMatch === false) {
      returnString = "Mismatch found. Data requires your review.";
    }
    console.log("Returning: " + returnString);
    return returnString;
}

interface EventData {
event: string
date: number
location: string
capacity: number
}
```

## <a name="training-video-cross-reference-and-format-an-excel-file"></a><span data-ttu-id="4228d-116">Vídeo de aprendizaje: referencia cruzada y formato de un Excel archivo</span><span class="sxs-lookup"><span data-stu-id="4228d-116">Training video: Cross-reference and format an Excel file</span></span>

<span data-ttu-id="4228d-117">[Vea el recorrido de Sudhi Ramamurthy por este ejemplo en YouTube](https://youtu.be/dVwqBf483qo").</span><span class="sxs-lookup"><span data-stu-id="4228d-117">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/dVwqBf483qo").</span></span>
