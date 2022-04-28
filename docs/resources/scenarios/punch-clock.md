---
title: 'escenario de ejemplo de scripts de Office: botón de reloj de perforación'
description: En este ejemplo se agrega un botón de reloj perforado y se permite que un usuario entre y salga del reloj con la hora actual.
ms.date: 04/26/2022
ms.localizationpriority: medium
ms.openlocfilehash: de56fb651d6f6088620678cfd72ce662875eafa7
ms.sourcegitcommit: e6428a5214fa38aef036a952a0e3c09dbf6e4d3e
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 04/28/2022
ms.locfileid: "65109315"
---
# <a name="office-scripts-sample-scenario-punch-clock-button"></a>escenario de ejemplo de scripts de Office: botón de reloj de perforación

La idea de escenario y el script utilizados en este ejemplo fueron aportados por el miembro de la comunidad de scripts de Office [Brian Gonzalez](https://github.com/b-gonzalez).

En este escenario, creará una hoja de horas para un empleado que le permite grabar sus horas de inicio y finalización con la pulsación de un [botón](../../develop/script-buttons.md). En función de lo que se haya grabado anteriormente, al presionar el botón se iniciará su día (entrada del reloj) o finalizará su día (salida del reloj). El ejemplo funciona tanto para Excel en la Web como para Windows.

:::image type="content" source="../../images/punch-clock-sample-3.png" alt-text="Una tabla con tres columnas ('Clock In', 'Clock Out' y 'Duration') y un botón con la etiqueta 'Punch clock' en el libro.":::

## <a name="setup-instructions"></a>Instrucciones de instalación

1. Descargue <a href="punch-clock-sample.xlsx">punch-clock-sample.xlsx</a> en su OneDrive.

    :::image type="content" source="../../images/punch-clock-sample-1.png" alt-text="Tabla con tres columnas: &quot;Clock In&quot;, &quot;Clock Out&quot; y &quot;Duration&quot;.":::

1. Abra el libro en Excel en la Web.

1. En la pestaña **Automatizar** , seleccione **Nuevo script** y pegue el siguiente script en el editor.

    ```typescript
    /**
     * This script records either the start or end time of a shift, 
     * depending on what is filled out in the table. 
     * It is intended to be used with a Script Button.
     */
    function main(workbook: ExcelScript.Workbook) {
      // Get the first table in the timesheet.
      const timeSheet = workbook.getWorksheet("MyTimeSheet");
      const timeTable = timeSheet.getTables()[0];
    
      // Get the appropriate table columns.
      const clockInColumn = timeTable.getColumnByName("Clock In");
      const clockOutColumn = timeTable.getColumnByName("Clock Out");
      const durationColumn = timeTable.getColumnByName("Duration");
    
      // Get the last rows for the Clock In and Clock Out columns.
      let clockInLastRow = clockInColumn.getRangeBetweenHeaderAndTotal().getLastRow();
      let clockOutLastRow = clockOutColumn.getRangeBetweenHeaderAndTotal().getLastRow();
    
      // Get the current date to use as the start or end time.
      let date: Date = new Date();
    
      // Add the current time to a column based on the state of the table.
      if (clockInLastRow.getValue() as string === "") {
        // If the Clock In column has an empty value in the table, add a start time.
        clockInLastRow.setValue(date.toLocaleString());
      } else if (clockOutLastRow.getValue() as string === "") {
        // If the Clock Out column has an empty value in the table, 
        // add an end time and calculate the shift duration.
        clockOutLastRow.setValue(date.toLocaleString());
        const clockInTime = new Date(clockInLastRow.getValue() as string);
        const clockOutTime  = new Date(clockOutLastRow.getValue() as string);
        const clockDuration = Math.abs((clockOutTime.getTime() - clockInTime.getTime()));
    
        let durationString = getDurationMessage(clockDuration);
        durationColumn.getRangeBetweenHeaderAndTotal().getLastRow().setValue(durationString);
      } else {
        // If both columns are full, add a new row, then add a start time.
        timeTable.addRow()
        clockInLastRow.getOffsetRange(1, 0).setValue(date.toLocaleString());
      }
    }
    
    /**
     * A function to write a time duration as a string.
     */
    function getDurationMessage(delta: number) {
      // Adapted from here:
      // https://stackoverflow.com/questions/13903897/javascript-return-number-of-days-hours-minutes-seconds-between-two-dates
    
      delta = delta / 1000;
      let durationString = "";
    
      let days = Math.floor(delta / 86400);
      delta -= days * 86400;
    
      let hours = Math.floor(delta / 3600) % 24;
      delta -= hours * 3600;
    
      let minutes = Math.floor(delta / 60) % 60;
    
      if (days >= 1) {
        durationString += days;
        durationString += (days > 1 ? " days" : " day");
    
        if (hours >= 1 && minutes >= 1) {
          durationString += ", ";
        }
        else if (hours >= 1 || minutes > 1) {
          durationString += " and ";
        }
      }
    
      if (hours >= 1) {
        durationString += hours;
        durationString += (hours > 1 ? " hours" : " hour");
        if (minutes >= 1) {
          durationString += " and ";
        }
      }
    
      if (minutes >= 1) {
        durationString += minutes;
        durationString += (minutes > 1 ? " minutes" : " minute");
      }
    
      return durationString;
    }
    ```

1. Cambie el nombre del script a "Punch clock".

1. Guarde el script.

1. En el libro, seleccione la celda **E2**.

1. Agregar un botón de script. Vaya al menú **Más opciones (...)** de la página **Detalles del script** y seleccione **el botón Agregar**.

    :::image type="content" source="../../images/punch-clock-sample-2.png" alt-text="El menú &quot;Más opciones&quot; y el botón &quot;Agregar botón&quot;.":::

1. Guarde el libro.

## <a name="run-the-script"></a>Ejecución del script

Presione el botón **Punch clock (Reloj de perforación** ) para ejecutar el script. Registra la hora actual en "Entrada de reloj" o "Salida del reloj", dependiendo de lo que se haya escrito anteriormente.

:::image type="content" source="../../images/punch-clock-sample-3.png" alt-text="La tabla y el botón &quot;Punch clock&quot; del libro.":::

> [!NOTE]
> La duración solo se registra si es más de un minuto. Edite manualmente el tiempo "Clock In" para probar duraciones más grandes.
