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
# <a name="office-scripts-sample-scenario-analyze-web-downloads"></a><span data-ttu-id="08181-103">Escenario de ejemplo de scripts de Office: analizar descargas Web</span><span class="sxs-lookup"><span data-stu-id="08181-103">Office Scripts sample scenario: Analyze web downloads</span></span>

<span data-ttu-id="08181-104">En este escenario, usted tiene la tarea de analizar los informes de descarga desde el sitio web de su empresa.</span><span class="sxs-lookup"><span data-stu-id="08181-104">In this scenario, you're tasked with analyzing download reports from your company's website.</span></span> <span data-ttu-id="08181-105">El objetivo de este análisis es determinar si el tráfico web procede de Estados Unidos o de otros lugares del mundo.</span><span class="sxs-lookup"><span data-stu-id="08181-105">The goal of this analysis is to determine if the web traffic is coming from the United States or elsewhere in the world.</span></span>

<span data-ttu-id="08181-106">Los compañeros cargan los datos sin procesar en el libro.</span><span class="sxs-lookup"><span data-stu-id="08181-106">Your colleagues upload the raw data to your workbook.</span></span> <span data-ttu-id="08181-107">El conjunto de datos de cada semana tiene su propia hoja de cálculo.</span><span class="sxs-lookup"><span data-stu-id="08181-107">Each week's set of data has its own worksheet.</span></span> <span data-ttu-id="08181-108">También hay una hoja de cálculo de **Resumen** con una tabla y un gráfico que muestra tendencias semanales sobre la semana.</span><span class="sxs-lookup"><span data-stu-id="08181-108">There is also the **Summary** worksheet with a table and chart that shows week-over-week trends.</span></span>

<span data-ttu-id="08181-109">Desarrollará un script que analiza los datos de descarga semanal de la hoja de cálculo activa.</span><span class="sxs-lookup"><span data-stu-id="08181-109">You'll develop a script that analyzes weekly downloads data in the active worksheet.</span></span> <span data-ttu-id="08181-110">Analizará la dirección IP asociada con cada descarga y determinará si llegó o no con nosotros.</span><span class="sxs-lookup"><span data-stu-id="08181-110">It will parse the IP address associated with each download and determine whether or not it came from the US.</span></span> <span data-ttu-id="08181-111">La respuesta se insertará en la hoja de cálculo como un valor booleano ("TRUE" o "FALSE") y se aplicará el formato condicional a esas celdas.</span><span class="sxs-lookup"><span data-stu-id="08181-111">The answer will be inserted in the worksheet as a boolean value ("TRUE" or "FALSE") and conditional formatting will be applied to those cells.</span></span> <span data-ttu-id="08181-112">Los resultados de la ubicación de la dirección IP se totalizarán en la hoja de cálculo y se copiarán en la tabla de resumen.</span><span class="sxs-lookup"><span data-stu-id="08181-112">The IP address location results will be totaled on the worksheet and copied to the summary table.</span></span>

## <a name="scripting-skills-covered"></a><span data-ttu-id="08181-113">Habilidades de scripting cubiertas</span><span class="sxs-lookup"><span data-stu-id="08181-113">Scripting skills covered</span></span>

- <span data-ttu-id="08181-114">Análisis de texto</span><span class="sxs-lookup"><span data-stu-id="08181-114">Text parsing</span></span>
- <span data-ttu-id="08181-115">Subfunciones en scripts</span><span class="sxs-lookup"><span data-stu-id="08181-115">Subfunctions in scripts</span></span>
- <span data-ttu-id="08181-116">Formato condicional</span><span class="sxs-lookup"><span data-stu-id="08181-116">Conditional formatting</span></span>
- <span data-ttu-id="08181-117">Tablas</span><span class="sxs-lookup"><span data-stu-id="08181-117">Tables</span></span>

## <a name="demo-video"></a><span data-ttu-id="08181-118">Vídeo de demostración</span><span class="sxs-lookup"><span data-stu-id="08181-118">Demo video</span></span>

<span data-ttu-id="08181-119">Este ejemplo se ha demodo como parte de la llamada de la comunidad de desarrolladores de complementos de Office para febrero de 2020.</span><span class="sxs-lookup"><span data-stu-id="08181-119">This sample was demoed as part of the Office Add-ins developer community call for February 2020.</span></span>

> [!VIDEO https://www.youtube.com/embed/vPEqbb7t6-Y?start=154]

## <a name="setup-instructions"></a><span data-ttu-id="08181-120">Instrucciones de instalación</span><span class="sxs-lookup"><span data-stu-id="08181-120">Setup instructions</span></span>

1. <span data-ttu-id="08181-121">Descargue <a href="analyze-web-downloads.xlsx">Analyze-web-downloads. xlsx</a> en su OneDrive.</span><span class="sxs-lookup"><span data-stu-id="08181-121">Download <a href="analyze-web-downloads.xlsx">analyze-web-downloads.xlsx</a> to your OneDrive.</span></span>

2. <span data-ttu-id="08181-122">Abra el libro con Excel para la Web.</span><span class="sxs-lookup"><span data-stu-id="08181-122">Open the workbook with Excel for the web.</span></span>

3. <span data-ttu-id="08181-123">En la ficha **automatizar** , abra el **Editor de código**.</span><span class="sxs-lookup"><span data-stu-id="08181-123">Under the **Automate** tab, open the **Code Editor**.</span></span>

4. <span data-ttu-id="08181-124">En el panel de tareas **Editor de código** , presione **nueva secuencia** de comandos y pegue el siguiente script en el editor.</span><span class="sxs-lookup"><span data-stu-id="08181-124">In the **Code Editor** task pane, press **New Script** and paste the following script into the editor.</span></span>

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

5. <span data-ttu-id="08181-125">Cambie el nombre del script para **analizar las descargas Web** y guardarlas.</span><span class="sxs-lookup"><span data-stu-id="08181-125">Rename the script to **Analyze Web Downloads** and save it.</span></span>

## <a name="running-the-script"></a><span data-ttu-id="08181-126">Ejecución del script</span><span class="sxs-lookup"><span data-stu-id="08181-126">Running the script</span></span>

<span data-ttu-id="08181-127">Navegue a cualquiera de las hojas de cálculo de la \*\*semana\* \*\* y ejecute el script **analizar descargas Web** .</span><span class="sxs-lookup"><span data-stu-id="08181-127">Navigate to any of the **Week\*\*** worksheets and run the **Analyze Web Downloads** script.</span></span> <span data-ttu-id="08181-128">El script aplicará el formato condicional y el etiquetado de ubicación en la hoja actual.</span><span class="sxs-lookup"><span data-stu-id="08181-128">The script will apply the conditional formatting and location labelling on the current sheet.</span></span> <span data-ttu-id="08181-129">También se actualizará la hoja de cálculo de **Resumen** .</span><span class="sxs-lookup"><span data-stu-id="08181-129">It will also update the **Summary** worksheet.</span></span>

### <a name="before-running-the-script"></a><span data-ttu-id="08181-130">Antes de ejecutar el script</span><span class="sxs-lookup"><span data-stu-id="08181-130">Before running the script</span></span>

![Una hoja de cálculo que muestra datos de tráfico web sin formato.](../../images/scenario-analyze-web-downloads-before.png)

### <a name="after-running-the-script"></a><span data-ttu-id="08181-132">Después de ejecutar el script</span><span class="sxs-lookup"><span data-stu-id="08181-132">After running the script</span></span>

![Una hoja de cálculo que muestra información de ubicación IP con formato con las filas de tráfico web anteriores.](../../images/scenario-analyze-web-downloads-after.png)

![La tabla de Resumen y el gráfico que resume las hojas de cálculo en las que se ha ejecutado el script.](../../images/scenario-analyze-web-downloads-table.png)
