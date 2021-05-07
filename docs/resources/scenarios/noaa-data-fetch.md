---
title: 'Office Escenario de ejemplo de scripts: Graph datos de nivel de agua de NOAA'
description: Ejemplo que captura datos JSON de una base de datos NOAA y los usa para crear un gráfico.
ms.date: 04/26/2021
localization_priority: Normal
ms.openlocfilehash: d35af59d9eed1abc9f3844834c92752ed80de80f
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232693"
---
# <a name="office-scripts-sample-scenario-fetch-and-graph-water-level-data-from-noaa"></a><span data-ttu-id="74186-103">Office Escenario de ejemplo de scripts: capturar y representar datos de nivel de agua de NOAA</span><span class="sxs-lookup"><span data-stu-id="74186-103">Office Scripts sample scenario: Fetch and graph water-level data from NOAA</span></span>

<span data-ttu-id="74186-104">En este escenario, debe trazar el nivel del agua en la estación Seattle de la Administración Nacional Oceánica [y Atmosférico.](https://tidesandcurrents.noaa.gov/stationhome.html?id=9447130)</span><span class="sxs-lookup"><span data-stu-id="74186-104">In this scenario, you need to plot the water level at the [National Oceanic and Atmospheric Administration's Seattle station](https://tidesandcurrents.noaa.gov/stationhome.html?id=9447130).</span></span> <span data-ttu-id="74186-105">Usará datos externos para rellenar una hoja de cálculo y crear un gráfico.</span><span class="sxs-lookup"><span data-stu-id="74186-105">You'll use external data to populate a spreadsheet and create a chart.</span></span>

<span data-ttu-id="74186-106">Desarrollará un script que usa el comando para consultar la base de datos de corrientes y `fetch` [mareas de NOAA.](https://tidesandcurrents.noaa.gov/)</span><span class="sxs-lookup"><span data-stu-id="74186-106">You'll develop a script that uses the `fetch` command to query the [NOAA Tides and Currents database](https://tidesandcurrents.noaa.gov/).</span></span> <span data-ttu-id="74186-107">Esto hará que el nivel de agua se grabe en un intervalo de tiempo determinado.</span><span class="sxs-lookup"><span data-stu-id="74186-107">That will get the water level recorded across a given time span.</span></span> <span data-ttu-id="74186-108">La información se devolverá como JSON, por lo que parte del script lo traducirá en valores de intervalo.</span><span class="sxs-lookup"><span data-stu-id="74186-108">The information will be returned as JSON, so part of the script will translate that into range values.</span></span> <span data-ttu-id="74186-109">Una vez que los datos están en la hoja de cálculo, se usarán para crear un gráfico.</span><span class="sxs-lookup"><span data-stu-id="74186-109">Once the data is in the spreadsheet, it will be used to make a chart.</span></span>

## <a name="scripting-skills-covered"></a><span data-ttu-id="74186-110">Habilidades de scripting cubiertas</span><span class="sxs-lookup"><span data-stu-id="74186-110">Scripting skills covered</span></span>

- <span data-ttu-id="74186-111">Llamadas DE API externas ( `fetch` )</span><span class="sxs-lookup"><span data-stu-id="74186-111">External API calls (`fetch`)</span></span>
- <span data-ttu-id="74186-112">Análisis JSON</span><span class="sxs-lookup"><span data-stu-id="74186-112">JSON parsing</span></span>
- <span data-ttu-id="74186-113">Gráficos</span><span class="sxs-lookup"><span data-stu-id="74186-113">Charts</span></span>

## <a name="setup-instructions"></a><span data-ttu-id="74186-114">Instrucciones de configuración</span><span class="sxs-lookup"><span data-stu-id="74186-114">Setup instructions</span></span>

1. <span data-ttu-id="74186-115">Abra el libro con Excel en la Web.</span><span class="sxs-lookup"><span data-stu-id="74186-115">Open the workbook with Excel on the web.</span></span>

1. <span data-ttu-id="74186-116">En la **pestaña Automatizar,** seleccione **Todos los scripts**.</span><span class="sxs-lookup"><span data-stu-id="74186-116">Under the **Automate** tab, select **All Scripts**.</span></span>

1. <span data-ttu-id="74186-117">En el **panel de tareas Editor** de código, seleccione Nuevo **script** y pegue el siguiente script en el editor.</span><span class="sxs-lookup"><span data-stu-id="74186-117">In the **Code Editor** task pane, select **New Script** and paste the following script into the editor.</span></span>

    ```TypeScript
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
      const rawJson: string = await response.json();
    
      // Translate the raw JSON into a usable state.
      const stringifiedJson = JSON.stringify(rawJson);
    
      // Note that we're only taking the data part of the JSON and excluding the metadata.
      const noaaData: NOAAData[] = JSON.parse(stringifiedJson).data;
    
      // Create table headers and format them to stand out.
      let headers = [["Time", "Level"]];
      let headerRange = currentSheet.getRange("A1:B1");
      headerRange.setValues(headers);
      headerRange.getFormat().getFill().setColor("#4472C4");
      headerRange.getFormat().getFont().setColor("white");
    
      // Insert all the data in rows from JSON.
      let noaaDataCount = noaaData.length;
      let dataToEnter = [[], []]
      for (let i = 0; i < noaaDataCount; i++) {
        let currentDataPiece = noaaData[i];
        dataToEnter[i] = [currentDataPiece.t, currentDataPiece.v];
      }
    
      let dataRange = currentSheet.getRange("A2:B" + String(noaaDataCount + 1)); /* +1 to account for the title row */
      dataRange.setValues(dataToEnter);
    
      // Format the "Time" column for timestamps.
      dataRange.getColumn(0).setNumberFormatLocal("[$-en-US]mm/dd/yyyy hh:mm AM/PM;@");
    
      // Create and format a chart with the level data.
      let chart = currentSheet.addChart(ExcelScript.ChartType.xyscatterSmooth, dataRange);
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
    
      /**
       * An interface to wrap the parts of the JSON we need.
       * These properties must match the names used in the JSON.
       */ 
      interface NOAAData {
        t: string; // Time
        v: number; // Level
      }
    }
    ```

1. <span data-ttu-id="74186-118">Cambie el nombre del script a **NoaA Water Level Chart** y guárdelo.</span><span class="sxs-lookup"><span data-stu-id="74186-118">Rename the script to **NOAA Water Level Chart** and save it.</span></span>

## <a name="running-the-script"></a><span data-ttu-id="74186-119">Ejecución del script</span><span class="sxs-lookup"><span data-stu-id="74186-119">Running the script</span></span>

<span data-ttu-id="74186-120">En cualquier hoja de cálculo, ejecute el script gráfico de nivel de agua **de NOAA.**</span><span class="sxs-lookup"><span data-stu-id="74186-120">On any worksheet, run the **NOAA Water Level Chart** script.</span></span> <span data-ttu-id="74186-121">El script captura los datos de nivel de agua del 25 de diciembre de 2020 al 27 de diciembre de 2020.</span><span class="sxs-lookup"><span data-stu-id="74186-121">The script fetches the water level data from December 25, 2020 to December 27, 2020.</span></span> <span data-ttu-id="74186-122">Las variables al principio del script se pueden cambiar para usar diferentes `const` fechas u obtener información de estación diferente.</span><span class="sxs-lookup"><span data-stu-id="74186-122">The `const` variables at the beginning of the script can be changed to use different dates or get different station information.</span></span> <span data-ttu-id="74186-123">La [API de CO-OPS para la recuperación de](https://api.tidesandcurrents.noaa.gov/api/prod/) datos describe cómo obtener todos estos datos.</span><span class="sxs-lookup"><span data-stu-id="74186-123">The [CO-OPS API For Data Retrieval](https://api.tidesandcurrents.noaa.gov/api/prod/) describes how to get all this data.</span></span>

### <a name="after-running-the-script"></a><span data-ttu-id="74186-124">Después de ejecutar el script</span><span class="sxs-lookup"><span data-stu-id="74186-124">After running the script</span></span>

:::image type="content" source="../../images/scenario-noaa-water-level-after.png" alt-text="La hoja de cálculo después de ejecutar el script muestra algunos datos de nivel de agua y un gráfico":::
