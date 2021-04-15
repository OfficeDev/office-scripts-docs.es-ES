---
title: 'Escenario de ejemplo scripts de Office: analizar descargas web'
description: Ejemplo que toma datos de tráfico de Internet sin procesar en un libro de Excel y determina la ubicación de origen, antes de organizar esa información en una tabla.
ms.date: 12/17/2020
localization_priority: Normal
ms.openlocfilehash: e351cd6c4a12e83a07a2f4ce5678d7aa10625118
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 04/14/2021
ms.locfileid: "51755038"
---
# <a name="office-scripts-sample-scenario-analyze-web-downloads"></a><span data-ttu-id="620a8-103">Escenario de ejemplo scripts de Office: analizar descargas web</span><span class="sxs-lookup"><span data-stu-id="620a8-103">Office Scripts sample scenario: Analyze web downloads</span></span>

<span data-ttu-id="620a8-104">En este escenario, tiene la tarea de analizar los informes de descarga desde el sitio web de su empresa.</span><span class="sxs-lookup"><span data-stu-id="620a8-104">In this scenario, you're tasked with analyzing download reports from your company's website.</span></span> <span data-ttu-id="620a8-105">El objetivo de este análisis es determinar si el tráfico web viene de Estados Unidos o de otra parte del mundo.</span><span class="sxs-lookup"><span data-stu-id="620a8-105">The goal of this analysis is to determine if the web traffic is coming from the United States or elsewhere in the world.</span></span>

<span data-ttu-id="620a8-106">Sus compañeros cargan los datos sin procesar en el libro.</span><span class="sxs-lookup"><span data-stu-id="620a8-106">Your colleagues upload the raw data to your workbook.</span></span> <span data-ttu-id="620a8-107">El conjunto de datos de cada semana tiene su propia hoja de cálculo.</span><span class="sxs-lookup"><span data-stu-id="620a8-107">Each week's set of data has its own worksheet.</span></span> <span data-ttu-id="620a8-108">También hay la hoja de cálculo **Resumen** con una tabla y un gráfico que muestra las tendencias semana a semana.</span><span class="sxs-lookup"><span data-stu-id="620a8-108">There is also the **Summary** worksheet with a table and chart that shows week-over-week trends.</span></span>

<span data-ttu-id="620a8-109">Desarrollará un script que analice los datos de descargas semanales en la hoja de cálculo activa.</span><span class="sxs-lookup"><span data-stu-id="620a8-109">You'll develop a script that analyzes weekly downloads data in the active worksheet.</span></span> <span data-ttu-id="620a8-110">Analizará la dirección IP asociada a cada descarga y determinará si provenía o no de Estados Unidos.</span><span class="sxs-lookup"><span data-stu-id="620a8-110">It will parse the IP address associated with each download and determine whether or not it came from the US.</span></span> <span data-ttu-id="620a8-111">La respuesta se insertará en la hoja de cálculo como un valor booleano ("TRUE" o "FALSE") y el formato condicional se aplicará a esas celdas.</span><span class="sxs-lookup"><span data-stu-id="620a8-111">The answer will be inserted in the worksheet as a boolean value ("TRUE" or "FALSE") and conditional formatting will be applied to those cells.</span></span> <span data-ttu-id="620a8-112">Los resultados de la ubicación de la dirección IP se completarán en la hoja de cálculo y se copiarán en la tabla de resumen.</span><span class="sxs-lookup"><span data-stu-id="620a8-112">The IP address location results will be totaled on the worksheet and copied to the summary table.</span></span>

## <a name="scripting-skills-covered"></a><span data-ttu-id="620a8-113">Habilidades de scripting cubiertas</span><span class="sxs-lookup"><span data-stu-id="620a8-113">Scripting skills covered</span></span>

- <span data-ttu-id="620a8-114">Análisis de texto</span><span class="sxs-lookup"><span data-stu-id="620a8-114">Text parsing</span></span>
- <span data-ttu-id="620a8-115">Subfunciones en scripts</span><span class="sxs-lookup"><span data-stu-id="620a8-115">Subfunctions in scripts</span></span>
- <span data-ttu-id="620a8-116">Formato condicional</span><span class="sxs-lookup"><span data-stu-id="620a8-116">Conditional formatting</span></span>
- <span data-ttu-id="620a8-117">Tablas</span><span class="sxs-lookup"><span data-stu-id="620a8-117">Tables</span></span>

## <a name="demo-video"></a><span data-ttu-id="620a8-118">Vídeo de demostración</span><span class="sxs-lookup"><span data-stu-id="620a8-118">Demo video</span></span>

<span data-ttu-id="620a8-119">Este ejemplo se ha degradado como parte de la llamada de la comunidad de desarrolladores de complementos de Office para febrero de 2020.</span><span class="sxs-lookup"><span data-stu-id="620a8-119">This sample was demoed as part of the Office Add-ins developer community call for February 2020.</span></span>

> [!VIDEO https://www.youtube.com/embed/vPEqbb7t6-Y?start=154]

> [!NOTE]
> <span data-ttu-id="620a8-120">El código que se muestra en este vídeo usa un modelo de API anterior (las API asincrónicas de [Scripts de Office).](../../develop/excel-async-model.md)</span><span class="sxs-lookup"><span data-stu-id="620a8-120">The code shown in this video uses an older API model (the [Office Scripts Async APIs](../../develop/excel-async-model.md)).</span></span> <span data-ttu-id="620a8-121">El ejemplo presentado en esta página se ha actualizado, pero el código tiene un aspecto un poco diferente de la grabación.</span><span class="sxs-lookup"><span data-stu-id="620a8-121">The sample presented on this page has been updated, but the code looks a little different from the recording.</span></span> <span data-ttu-id="620a8-122">Los cambios no afectan al comportamiento del script ni al otro contenido de la demostración del moderador.</span><span class="sxs-lookup"><span data-stu-id="620a8-122">The changes don't affect the behavior of the script or the other content in the presenter's demo.</span></span>

## <a name="setup-instructions"></a><span data-ttu-id="620a8-123">Instrucciones de configuración</span><span class="sxs-lookup"><span data-stu-id="620a8-123">Setup instructions</span></span>

1. <span data-ttu-id="620a8-124">Descarga <a href="analyze-web-downloads.xlsx">analyze-web-downloads.xlsx</a> a tu OneDrive.</span><span class="sxs-lookup"><span data-stu-id="620a8-124">Download <a href="analyze-web-downloads.xlsx">analyze-web-downloads.xlsx</a> to your OneDrive.</span></span>

2. <span data-ttu-id="620a8-125">Abra el libro con Excel para la web.</span><span class="sxs-lookup"><span data-stu-id="620a8-125">Open the workbook with Excel for the web.</span></span>

3. <span data-ttu-id="620a8-126">En la **pestaña Automatizar,** abra **Todos los scripts**.</span><span class="sxs-lookup"><span data-stu-id="620a8-126">Under the **Automate** tab, open **All Scripts**.</span></span>

4. <span data-ttu-id="620a8-127">En el **panel de tareas Editor** de código, presione Nuevo **script** y pegue el siguiente script en el editor.</span><span class="sxs-lookup"><span data-stu-id="620a8-127">In the **Code Editor** task pane, press **New Script** and paste the following script into the editor.</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
      /* Get the Summary worksheet and table.
        * End the script early if either object is not in the workbook.
        */
      let summaryWorksheet = workbook.getWorksheet("Summary");
      if (!summaryWorksheet) {
        console.log("The script expects a worksheet named \"Summary\". Please download the correct template and try again.");
        return;
      }
      let summaryTable = summaryWorksheet.getTable("Table1");
      if (!summaryTable) {
        console.log("The script expects a summary table named \"Table1\". Please download the correct template and try again.");
        return;
      }
  
      // Get the current worksheet.
      let currentWorksheet = workbook.getActiveWorksheet();
      if (currentWorksheet.getName().toLocaleLowerCase().indexOf("week") !== 0) {
        console.log("Please switch worksheet to one of the weekly data sheets and try again.")
        return;
      }
  
      // Get the values of the active range of the active worksheet.
      let logRange = currentWorksheet.getUsedRange();
  
      if (logRange.getColumnCount() !== 8) {
        console.log(`Verify that you are on the correct worksheet. Either the week's data has been already processed or the content is incorrect. The following columns are expected: ${[
            "Time Stamp", "IP Address", "kilobytes", "user agent code", "milliseconds", "Request", "Results", "Referrer"
        ]}`);
        return;
      }
      // Get the range that will contain TRUE/FALSE if the IP address is from the United States (US).
      let isUSColumn = logRange
          .getLastColumn()
          .getOffsetRange(0, 1);
  
      // Get the values of all the US IP addresses.
      let ipRange = workbook.getWorksheet("USIPAddresses").getUsedRange();
      let ipRangeValues = ipRange.getValues() as number[][];
      let logRangeValues = logRange.getValues() as string[][];
      // Remove the first row.
      let topRow = logRangeValues.shift();
      console.log(`Analyzing ${logRangeValues.length} entries.`);
  
      // Create a new array to contain the boolean representing if this is a US IP address.
      let newCol = [];
  
      // Go through each row in worksheet and add Boolean.
      for (let i = 0; i < logRangeValues.length; i++) {
        let curRowIP = logRangeValues[i][1];
        if (findIP(ipRangeValues, ipAddressToInteger(curRowIP)) > 0) {
          newCol.push([true]);
        } else {
          newCol.push([false]);
        }
      }
  
      // Remove the empty column header and add proper heading.
      newCol = [["Is US IP"], ...newCol];
  
      // Write the result to the spreadsheet.
      console.log(`Adding column to indicate whether IP belongs to US region or not at address: ${isUSColumn.getAddress()}`);
      console.log(newCol.length);
      console.log(newCol);
      isUSColumn.setValues(newCol);
  
      // Call the local function to add summary data to the worksheet.
      addSummaryData();
  
      // Call the local function to apply conditional formatting.
      applyConditionalFormatting(isUSColumn);
  
      // Autofit columns.
      currentWorksheet.getUsedRange().getFormat().autofitColumns();
  
      // Get the calculated summary data.
      let summaryRangeValues = currentWorksheet.getRange("J2:M2").getValues();
  
      // Add the corresponding row to the summary table.
      summaryTable.addRow(null, summaryRangeValues[0]);
      console.log("Complete.");
      return;
  
      /**
       * A function to add summary data on the worksheet.
        */
      function addSummaryData() {
        // Add a summary row and table.
        let summaryHeader = [["Year", "Week", "US", "Other"]];
        let countTrueFormula =
            "=COUNTIF(" + isUSColumn.getAddress() + ', "=TRUE")/' + (newCol.length - 1);
        let countFalseFormula =
            "=COUNTIF(" + isUSColumn.getAddress() + ', "=FALSE")/' + (newCol.length - 1);

        let summaryContent = [
          [
            '=TEXT(A2,"YYYY")',
            '=TEXTJOIN(" ", FALSE, "Wk", WEEKNUM(A2))',
            countTrueFormula,
            countFalseFormula
          ]
        ];
        let summaryHeaderRow = currentWorksheet.getRange("J1:M1");
        let summaryContentRow = currentWorksheet.getRange("J2:M2");
        console.log("2");

        summaryHeaderRow.setValues(summaryHeader);
        console.log("3");

        summaryContentRow.setValues(summaryContent);
        console.log("4");

        let formats = [[".000", ".000"]];
        summaryContentRow
            .getOffsetRange(0, 2)
            .getResizedRange(0, -2).setNumberFormats(formats);
      }
    }
    /**
     * Apply conditional formatting based on TRUE/FALSE values of the Is US IP column.
     */
    function applyConditionalFormatting(isUSColumn: ExcelScript.Range) {
      // Add conditional formatting to the new column.
      let conditionalFormatTrue = isUSColumn.addConditionalFormat(
          ExcelScript.ConditionalFormatType.cellValue
      );
      let conditionalFormatFalse = isUSColumn.addConditionalFormat(
          ExcelScript.ConditionalFormatType.cellValue
      );
      // Set TRUE to light blue and FALSE to light orange.
      conditionalFormatTrue.getCellValue().getFormat().getFill().setColor("#8FA8DB");
      conditionalFormatTrue.getCellValue().setRule({
          formula1: "=TRUE",
          operator: ExcelScript.ConditionalCellValueOperator.equalTo
      });
      conditionalFormatFalse.getCellValue().getFormat().getFill().setColor("#F8CCAD");
      conditionalFormatFalse.getCellValue().setRule({
          formula1: "=FALSE",
          operator: ExcelScript.ConditionalCellValueOperator.equalTo
      });
    }
    /**
     * Translate an IP address into an integer.
     * @param ipAddress: IP address to verify.
     */
    function ipAddressToInteger(ipAddress: string): number {
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
    /**
     * Return the row number where the ip address is found.
     * @param ipLookupTable IP look-up table.
     * @param n IP address to number value.  
     */
    function findIP(ipLookupTable: number[][], n: number): number {
      for (let i = 0; i < ipLookupTable.length; i++) {
        if (ipLookupTable[i][0] <= n && ipLookupTable[i][1] >= n) {
          return i;
        }
      }
      return -1;
    }
    ```

5. <span data-ttu-id="620a8-128">Cambie el nombre del script a **Analizar descargas web** y guárdelo.</span><span class="sxs-lookup"><span data-stu-id="620a8-128">Rename the script to **Analyze Web Downloads** and save it.</span></span>

## <a name="running-the-script"></a><span data-ttu-id="620a8-129">Ejecución del script</span><span class="sxs-lookup"><span data-stu-id="620a8-129">Running the script</span></span>

<span data-ttu-id="620a8-130">Navegue a cualquiera de las hojas **de \* \*** cálculo de semana y ejecute el script **Analizar descargas web.**</span><span class="sxs-lookup"><span data-stu-id="620a8-130">Navigate to any of the **Week\*\*** worksheets and run the **Analyze Web Downloads** script.</span></span> <span data-ttu-id="620a8-131">El script aplicará el formato condicional y el etiquetado de ubicación en la hoja actual.</span><span class="sxs-lookup"><span data-stu-id="620a8-131">The script will apply the conditional formatting and location labelling on the current sheet.</span></span> <span data-ttu-id="620a8-132">También actualizará la hoja **de cálculo Resumen.**</span><span class="sxs-lookup"><span data-stu-id="620a8-132">It will also update the **Summary** worksheet.</span></span>

### <a name="before-running-the-script"></a><span data-ttu-id="620a8-133">Antes de ejecutar el script</span><span class="sxs-lookup"><span data-stu-id="620a8-133">Before running the script</span></span>

:::image type="content" source="../../images/scenario-analyze-web-downloads-before.png" alt-text="Hoja de cálculo que muestra datos de tráfico web sin procesar.":::

### <a name="after-running-the-script"></a><span data-ttu-id="620a8-135">Después de ejecutar el script</span><span class="sxs-lookup"><span data-stu-id="620a8-135">After running the script</span></span>

:::image type="content" source="../../images/scenario-analyze-web-downloads-after.png" alt-text="Hoja de cálculo que muestra información de ubicación IP con formato con las filas de tráfico web anteriores.":::

:::image type="content" source="../../images/scenario-analyze-web-downloads-table.png" alt-text="Tabla de resumen y gráfico que resume las hojas de cálculo en las que se ha ejecutado el script.":::
