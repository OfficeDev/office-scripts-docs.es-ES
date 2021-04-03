---
title: Referencia cruzada y formato de un archivo de Excel
description: Obtenga información sobre cómo usar Scripts de Office y Power Automate para hacer referencia cruzada y dar formato a un archivo de Excel.
ms.date: 03/18/2021
localization_priority: Normal
ms.openlocfilehash: 287de604733b7e6a126d0c81cb4e23351e558c61
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571622"
---
# <a name="cross-reference-and-format-an-excel-file"></a>Referencia cruzada y formato de un archivo de Excel

Esta solución muestra cómo se puede hacer referencia a dos archivos de Excel y dar formato con scripts de Office y Power Automate.

El proyecto logra lo siguiente:

1. Extrae datos de eventos de <a href="events.xlsx">events.xlsx</a> mediante una acción ejecutar script.
1. Pasa estos datos al segundo archivo de Excel que contiene datos de transacción de eventos y los usa para realizar la validación básica de los datos y el formato de datos que faltan o incorrectos mediante scripts de Office.
1. Envía el resultado por correo electrónico a un revisor.

Para obtener más información, vea [Cross Reference and formatting two Excel files using Office Scripts](https://powerusers.microsoft.com/t5/Power-Automate-Cookbook/Cross-Reference-and-formatting-two-Excel-files-using-Office/td-p/728535).

## <a name="sample-excel-files"></a>Archivos de Excel de ejemplo

Descarga los siguientes archivos usados en esta solución para probarlos tú mismo.

1. <a href="events.xlsx">events.xlsx</a>
1. <a href="event-transactions.xlsx">event-transactions.xlsx</a>

## <a name="sample-code-get-event-data"></a>Código de ejemplo: Obtener datos de eventos

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

## <a name="sample-code-validate-event-transactions"></a>Código de ejemplo: Validar transacciones de eventos

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

## <a name="training-video-cross-reference-and-format-an-excel-file"></a>Vídeo de aprendizaje: referencia cruzada y formato de un archivo de Excel

[![Ver vídeo paso a paso sobre cómo hacer referencias cruzadas y dar formato a un archivo de Excel](../../images/cross-ref-tables-vid.jpg)](https://youtu.be/dVwqBf483qo "Vídeo paso a paso sobre cómo hacer referencias cruzadas y dar formato a un archivo de Excel")
