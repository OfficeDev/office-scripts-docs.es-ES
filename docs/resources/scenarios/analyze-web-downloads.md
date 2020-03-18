---
title: 'Escenario de ejemplo de scripts de Office: analizar descargas Web'
description: Un ejemplo que toma datos de tráfico de Internet sin procesar en un libro de Excel y determina la ubicación del origen antes de organizar dicha información en una tabla.
ms.date: 02/20/2020
localization_priority: Normal
ms.openlocfilehash: 9ee12c8d4ca7c191168e3734d7cd9eadc333c165
ms.sourcegitcommit: b075eed5a6f275274fbbf6d62633219eac416f26
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 03/10/2020
ms.locfileid: "42700410"
---
# <a name="office-scripts-sample-scenario-analyze-web-downloads"></a>Escenario de ejemplo de scripts de Office: analizar descargas Web

En este escenario, usted tiene la tarea de analizar los informes de descarga desde el sitio web de su empresa. El objetivo de este análisis es determinar si el tráfico web procede de Estados Unidos o de otros lugares del mundo.

Los compañeros cargan los datos sin procesar en el libro. El conjunto de datos de cada semana tiene su propia hoja de cálculo. También hay una hoja de cálculo de **Resumen** con una tabla y un gráfico que muestra tendencias semanales sobre la semana.

Desarrollará un script que analiza los datos de descarga semanal de la hoja de cálculo activa. Analizará la dirección IP asociada con cada descarga y determinará si llegó o no con nosotros. La respuesta se insertará en la hoja de cálculo como un valor booleano ("TRUE" o "FALSE") y se aplicará el formato condicional a esas celdas. Los resultados de la ubicación de la dirección IP se totalizarán en la hoja de cálculo y se copiarán en la tabla de resumen.

## <a name="scripting-skills-covered"></a>Habilidades de scripting cubiertas

- Análisis de texto
- Subfunciones en scripts
- Formato condicional
- Tablas

## <a name="demo-video"></a>Vídeo de demostración

Este ejemplo se ha demodo como parte de la llamada de la comunidad de desarrolladores de complementos de Office para febrero de 2020.

> [!VIDEO https://www.youtube.com/embed/vPEqbb7t6-Y?start=154]

## <a name="setup-instructions"></a>Instrucciones de instalación

1. Descargue <a href="analyze-web-downloads.xlsx">Analyze-web-downloads. xlsx</a> en su OneDrive.

2. Abra el libro con Excel para la Web.

3. En la ficha **automatizar** , abra el **Editor de código**.

4. En el panel de tareas **Editor de código** , presione **nueva secuencia** de comandos y pegue el siguiente script en el editor.

    ```TypeScript
      async function main(context: Excel.RequestContext) {
        let currentWorksheet = context.workbook.worksheets
          .getActiveWorksheet();
        // Get the values of the active range of the active worksheet.
        let logRange = currentWorksheet.getUsedRange().load("values");

        // Get the Summary worksheet and table.
        let summaryWorksheet = context.workbook.worksheets.getItem("Summary");
        let summaryTable = context.workbook.tables.getItem("Table1");

        // Get the range that will contain TRUE/FALSE if the IP address is from the United States (US).
        let isUSColumn = logRange
          .getLastColumn()
          .getOffsetRange(0, 1)
          .load("address");

        // Get the values of all the US IP addresses.
        let ipRange = context.workbook.worksheets
          .getItem("USIPAddresses")
          .getUsedRange()
          .load("values");
        await context.sync();

        // Remove the first row.
        let topRow = logRange.values.shift();

        // Create a new array to contain the boolean representing if this is a US IP address.
        let newCol = [[]];

        // Go through each row in worksheet and add Boolean.
        for (let i = 0; i < logRange.values.length; i++) {
          let curRowIP = logRange.values[i][1];
          if (findIP(ipRange.values, ipAddressToInteger(curRowIP)) > 0) {
            newCol.push([true]);
          } else {
            newCol.push([false]);
          }
        }

        // Remove the empty column header and add proper heading.
        newCol.shift();
        newCol.unshift(["Is US IP"]);

        // Write the result to the spreadsheet.
        isUSColumn.values = newCol;
        addSummaryData();
        applyConditionalFormatting();
        currentWorksheet.getUsedRange().format.autofitColumns();

        // Get the calculated summary data.
        let summaryRange = currentWorksheet.getRange("J2:M2").load("values");
        await context.sync();

        // Add the corresponding row to the summary table.
        summaryTable.rows.add(null, summaryRange.values);

        // Function to apply conditional formatting to the new column.
        function applyConditionalFormatting() {
          // Add conditional formatting to the new column.
          let conditionalFormatTrue = isUSColumn.conditionalFormats.add(
            Excel.ConditionalFormatType.cellValue
          );
          let conditionalFormatFalse = isUSColumn.conditionalFormats.add(
            Excel.ConditionalFormatType.cellValue
          );
          // Set TRUE to light blue and FALSE to light orange.
          conditionalFormatTrue.cellValue.format.fill.color = "#8FA8DB";
          conditionalFormatTrue.cellValue.rule = {
            formula1: "=TRUE",
            operator: "EqualTo"
          };
          conditionalFormatFalse.cellValue.format.fill.color = "#F8CCAD";
          conditionalFormatFalse.cellValue.rule = {
            formula1: "=FALSE",
            operator: "EqualTo"
          };
        }

        // Adds the summary data to the current sheet and to the summary table.
        function addSummaryData() {
          // Add a summary row and table.
          let summaryHeader = [["Year", "Week", "US", "Other"]];
          let countTrueFormula =
            "=COUNTIF(" + isUSColumn.address + ', "=TRUE")/' + (newCol.length - 1);
          let countFalseFormula =
            "=COUNTIF(" + isUSColumn.address + ', "=FALSE")/' + (newCol.length - 1);

          let summaryContent = [
            [
              '=TEXT(A2,"YYYY")',
              '=TEXTJOIN(" ", FALSE, "Wk", WEEKNUM(A2))',
              countTrueFormula,
              countFalseFormula
            ]
          ];
          let summaryHeaderRow = context.workbook.worksheets
            .getActiveWorksheet()
            .getRange("J1:M1");
          let summaryContentRow = context.workbook.worksheets
            .getActiveWorksheet()
            .getRange("J2:M2");
          summaryHeaderRow.values = summaryHeader;
          summaryContentRow.values = summaryContent;
          let formats = [[".000", ".000"]];
          summaryContentRow
            .getOffsetRange(0, 2)
            .getResizedRange(0, -2).numberFormat = formats;
        }
      }

      // Translate an IP address into an integer.
      function ipAddressToInteger(ipAddress: string) {
        // Split the IP address into octets.
        let octets = ipAddress.split(".");

        // Create a number for each octet and do the math to create the integer value of the IP address.
        let fullNum =
          // Define an arbitrary number for the last octet.
          111 +
          parseInt(octets[2]) * 256 +
          parseInt(octets[1]) * 65536 +
          parseInt(octets[0]) * 16777216;
        return fullNum;
      }

      // Return the row number where the ip address is found.
      function findIP(ipLookupTable: number[][], n: number) {
        for (let i = 0; i < ipLookupTable.length; i++) {
          if (ipLookupTable[i][0] <= n && ipLookupTable[i][1] >= n) {
            return i;
          }
        }
        return -1;
      }
    ```

5. Cambie el nombre del script para **analizar las descargas Web** y guardarlas.

## <a name="running-the-script"></a>Ejecución del script

Navegue a cualquiera de las hojas de cálculo de la **semana\* ** y ejecute el script **analizar descargas Web** . El script aplicará el formato condicional y el etiquetado de ubicación en la hoja actual. También se actualizará la hoja de cálculo de **Resumen** .

### <a name="before-running-the-script"></a>Antes de ejecutar el script

![Una hoja de cálculo que muestra datos de tráfico web sin formato.](../../images/scenario-analyze-web-downloads-before.png)

### <a name="after-running-the-script"></a>Después de ejecutar el script

![Una hoja de cálculo que muestra información de ubicación IP con formato con las filas de tráfico web anteriores.](../../images/scenario-analyze-web-downloads-after.png)

![La tabla de Resumen y el gráfico que resume las hojas de cálculo en las que se ha ejecutado el script.](../../images/scenario-analyze-web-downloads-table.png)
