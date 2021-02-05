---
title: 'Escenario de ejemplo de scripts de Office: datos de nivel de agua de Gráfico de NOAA'
description: Ejemplo que recupera datos JSON de una base de datos NOAA y los usa para crear un gráfico.
ms.date: 01/11/2021
localization_priority: Normal
ms.openlocfilehash: 5b0b4e3675cbe053368f63123d819f0dab626e60
ms.sourcegitcommit: 7580dcb8f2f97974c2a9cce25ea30d6526730e28
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 01/14/2021
ms.locfileid: "49867880"
---
# <a name="office-scripts-sample-scenario-fetch-and-graph-water-level-data-from-noaa"></a>Escenario de ejemplo de scripts de Office: obtener y representar gráficos de datos de nivel de agua de NOAA

En este escenario, debe trazar el nivel de agua en la estación seattle de administración nacional oceánica y de la administración de [connacionales.](https://tidesandcurrents.noaa.gov/stationhome.html?id=9447130) Usará datos externos para rellenar una hoja de cálculo y crear un gráfico.

Desarrollará un script que usa el comando para consultar la base de datos de fechas y finales de `fetch` [NOAA.](https://tidesandcurrents.noaa.gov/) De esta forma, se registrará el nivel de agua en un intervalo de tiempo determinado. La información se devolverá como JSON, por lo que parte del script lo traducirá en valores de intervalo. Una vez que los datos están en la hoja de cálculo, se usarán para crear un gráfico.

## <a name="scripting-skills-covered"></a>Habilidades de scripting cubiertas

- Llamadas a API externas ( `fetch` )
- Análisis JSON
- Gráficos

## <a name="setup-instructions"></a>Instrucciones de configuración

1. Abra el libro con Excel en la Web.

1. En la **pestaña** Automatizar, seleccione **Todos los scripts.**

1. En el **panel de** tareas Editor de código, seleccione **Nuevo script** y pegue el siguiente script en el editor.

    ```typescript
    /**
     * Gets data from the National Oceanic and Atmospheric Administration's Tides and Currents database. 
     * That data is used to make a chart.
     */
    async function main(workbook: ExcelScript.Workbook): Promise<void> {
      // Get the current sheet.
      let currentSheet = workbook.getActiveWorksheet();
    
      // Create selection of parameters for the fetch URL.
      // More information on the NOAA APIs is found here: 
      // https://api.tidesandcurrents.noaa.gov/api/prod/
      const option = "water_level";
      const startDate = "20201225"; /* yyyymmdd date format */
      const endDate = "20201227";
      const station = "9447130"; /* Seattle */
    
      // Construct the URL for the fetch call.
      const strQuery = `https://api.tidesandcurrents.noaa.gov/api/prod/datagetter?product=${option}&begin_date=${startDate}&end_date=${endDate}&datum=MLLW&station=${station}&units=english&time_zone=gmt&application=NOS.COOPS.TAC.WL&format=json`;
    
      console.log(strQuery);
    
      // Resolve the Promises returned by the fetch operation.
      const response = await fetch(strQuery);
      const rawJson = await response.json();
    
      // Translate the raw JSON into a usable state.
      const stringifiedJson = JSON.stringify(rawJson);
      const noaaData = JSON.parse(stringifiedJson);
    
      // Create table headers and format them to stand out.
      let headers = [["Time", "Level"]];
      let headerRange = currentSheet.getRange("A1:B1");
      headerRange.setValues(headers);
      headerRange.getFormat().getFill().setColor("#4472C4");
      headerRange.getFormat().getFont().setColor("white");
    
      // Insert all the data in rows from JSON.
      let noaaDataCount = noaaData.data.length;
      let dataToEnter = [[], []]
      for (let i = 0; i < noaaDataCount; i++) {
        let currentDataPiece = noaaData.data[i];
        dataToEnter[i] = [currentDataPiece.t, currentDataPiece.v];
      }
    
      let dataRange = currentSheet.getRange("A2:B" + String(noaaDataCount + 1)); /* +1 to account for the title row */
      dataRange.setValues(dataToEnter);
      
      // Format the "Time" column for timestamps.
      dataRange.getColumn(0).setNumberFormatLocal("[$-en-US]mm/dd/yyyy hh:mm AM/PM;@");
    
      // Create and format a chart with the level data.
      let chart = currentSheet.addChart(ExcelScript.ChartType.xyscatterSmooth,dataRange);
      chart.getTitle().setText("Water Level - Seattle");
      chart.setTop(0);
      chart.setLeft(300);
      chart.setWidth(500);
      chart.setHeight(300);
      chart.getAxes().getValueAxis().setShowDisplayUnitLabel(false);
      chart.getAxes().getCategoryAxis().setTextOrientation(60);
      chart.getLegend().setVisible(false);

      // Add a comment with the data attribution.
      currentSheet.addComment(
        "A1", 
        `This data was taken from the National Oceanic and Atmospheric Administration's Tides and Currents database on ${new Date(Date.now())}.`
      );
    }
    ```

1. Cambie el nombre del script al **gráfico de nivel de agua de NOAA** y guárdelo.

## <a name="running-the-script"></a>Ejecución del script

En cualquier hoja de cálculo, ejecute el script gráfico de nivel **de agua de NOAA.** El script recupera los datos de nivel de agua desde el 25 de diciembre de 2020 hasta el 27 de diciembre de 2020. Las variables al principio del script se pueden cambiar para usar `const` fechas diferentes u obtener información de estación diferente. La [API de CO-OPS para la recuperación de](https://api.tidesandcurrents.noaa.gov/api/prod/) datos describe cómo obtener todos estos datos.

### <a name="after-running-the-script"></a>Después de ejecutar el script

![La hoja de cálculo después de ejecutar el script muestra algunos datos de nivel de agua y un gráfico.](../../images/scenario-noaa-water-level-after.png)