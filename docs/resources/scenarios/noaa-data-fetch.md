---
title: 'Escenario de ejemplo de scripts de Office: gráfico de datos de nivel de agua de NOAA'
description: Ejemplo que recupera datos JSON de una base de datos NOAA y los usa para crear un gráfico.
ms.date: 01/05/2021
localization_priority: Normal
ms.openlocfilehash: d2afcd05125ea66c028d8e21bcc878371c20fcc3
ms.sourcegitcommit: 30c4b731dc8d18fca5aa74ce59e18a4a63eb4ffc
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 01/08/2021
ms.locfileid: "49784191"
---
# <a name="office-scripts-sample-scenario-graph-water-level-data-from-noaa"></a><span data-ttu-id="7bf48-103">Escenario de ejemplo de scripts de Office: gráfico de datos de nivel de agua de NOAA</span><span class="sxs-lookup"><span data-stu-id="7bf48-103">Office Scripts sample scenario: Graph water-level data from NOAA</span></span>

<span data-ttu-id="7bf48-104">En este escenario, debe trazar el nivel de agua en la estación seattle de administración nacional oceánica y de la administración de [connacionales.](https://tidesandcurrents.noaa.gov/stationhome.html?id=9447130)</span><span class="sxs-lookup"><span data-stu-id="7bf48-104">In this scenario, you need to plot the water level at the [National Oceanic and Atmospheric Administration's Seattle station](https://tidesandcurrents.noaa.gov/stationhome.html?id=9447130).</span></span> <span data-ttu-id="7bf48-105">Usará datos externos para rellenar una hoja de cálculo y crear un gráfico.</span><span class="sxs-lookup"><span data-stu-id="7bf48-105">You'll use external data to populate a spreadsheet and create a chart.</span></span>

<span data-ttu-id="7bf48-106">Desarrollará un script que usa el comando para consultar la base de datos de fechas y finales de `fetch` [NOAA.](https://tidesandcurrents.noaa.gov/)</span><span class="sxs-lookup"><span data-stu-id="7bf48-106">You'll develop a script that uses the `fetch` command to query the [NOAA Tides and Currents database](https://tidesandcurrents.noaa.gov/).</span></span> <span data-ttu-id="7bf48-107">Esto hará que el nivel de agua se grabe en un intervalo de tiempo determinado.</span><span class="sxs-lookup"><span data-stu-id="7bf48-107">That will get the water level recorded across a given time span.</span></span> <span data-ttu-id="7bf48-108">La información se devolverá como JSON, por lo que parte del script lo traducirá en valores de intervalo.</span><span class="sxs-lookup"><span data-stu-id="7bf48-108">The information will be returned as JSON, so part of the script will translate that into range values.</span></span> <span data-ttu-id="7bf48-109">Una vez que los datos están en la hoja de cálculo, se usarán para crear un gráfico.</span><span class="sxs-lookup"><span data-stu-id="7bf48-109">Once the data is in the spreadsheet, it will be used to make a chart.</span></span>

## <a name="scripting-skills-covered"></a><span data-ttu-id="7bf48-110">Habilidades de scripting cubiertas</span><span class="sxs-lookup"><span data-stu-id="7bf48-110">Scripting skills covered</span></span>

- <span data-ttu-id="7bf48-111">Llamadas a API externas ( `fetch` )</span><span class="sxs-lookup"><span data-stu-id="7bf48-111">External API calls (`fetch`)</span></span>
- <span data-ttu-id="7bf48-112">Análisis JSON</span><span class="sxs-lookup"><span data-stu-id="7bf48-112">JSON parsing</span></span>
- <span data-ttu-id="7bf48-113">Gráficos</span><span class="sxs-lookup"><span data-stu-id="7bf48-113">Charts</span></span>

## <a name="setup-instructions"></a><span data-ttu-id="7bf48-114">Instrucciones de configuración</span><span class="sxs-lookup"><span data-stu-id="7bf48-114">Setup instructions</span></span>

1. <span data-ttu-id="7bf48-115">Abra el libro con Excel en la Web.</span><span class="sxs-lookup"><span data-stu-id="7bf48-115">Open the workbook with Excel on the web.</span></span>

1. <span data-ttu-id="7bf48-116">En la **pestaña Automatizar,** seleccione **Todos los scripts.**</span><span class="sxs-lookup"><span data-stu-id="7bf48-116">Under the **Automate** tab, select **All Scripts**.</span></span>

1. <span data-ttu-id="7bf48-117">En el **panel de** tareas Editor de código, seleccione **Nuevo script** y pegue el siguiente script en el editor.</span><span class="sxs-lookup"><span data-stu-id="7bf48-117">In the **Code Editor** task pane, select **New Script** and paste the following script into the editor.</span></span>

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

1. <span data-ttu-id="7bf48-118">Cambie el nombre del script al **gráfico de nivel de agua de NOAA** y guárdelo.</span><span class="sxs-lookup"><span data-stu-id="7bf48-118">Rename the script to **NOAA Water Level Chart** and save it.</span></span>

## <a name="running-the-script"></a><span data-ttu-id="7bf48-119">Ejecución del script</span><span class="sxs-lookup"><span data-stu-id="7bf48-119">Running the script</span></span>

<span data-ttu-id="7bf48-120">En cualquier hoja de cálculo, ejecute el script gráfico de nivel **de agua de NOAA.**</span><span class="sxs-lookup"><span data-stu-id="7bf48-120">On any worksheet, run the **NOAA Water Level Chart** script.</span></span> <span data-ttu-id="7bf48-121">El script captura los datos de nivel de agua del 25 de diciembre de 2020 al 27 de diciembre de 2020.</span><span class="sxs-lookup"><span data-stu-id="7bf48-121">The script fetches the water level data from December 25, 2020 to December 27, 2020.</span></span> <span data-ttu-id="7bf48-122">Las variables al principio del script se pueden cambiar para usar `const` fechas diferentes u obtener información de estación diferente.</span><span class="sxs-lookup"><span data-stu-id="7bf48-122">The `const` variables at the beginning of the script can be changed to use different dates or get different station information.</span></span> <span data-ttu-id="7bf48-123">La [API de CO-OPS para la recuperación de](https://api.tidesandcurrents.noaa.gov/api/prod/) datos describe cómo obtener todos estos datos.</span><span class="sxs-lookup"><span data-stu-id="7bf48-123">The [CO-OPS API For Data Retrieval](https://api.tidesandcurrents.noaa.gov/api/prod/) describes how to get all this data.</span></span>

### <a name="after-running-the-script"></a><span data-ttu-id="7bf48-124">Después de ejecutar el script</span><span class="sxs-lookup"><span data-stu-id="7bf48-124">After running the script</span></span>

![La hoja de cálculo después de ejecutar el script muestra algunos datos de nivel de agua y un gráfico.](../../images/scenario-noaa-water-level-after.png)
