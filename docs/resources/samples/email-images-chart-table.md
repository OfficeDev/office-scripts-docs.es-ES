---
title: Envíe por correo electrónico las imágenes de un gráfico y una tabla de Excel
description: Aprenda a usar scripts y Power Automate Office para extraer y enviar por correo electrónico las imágenes de un gráfico y una tabla Excel.
ms.date: 05/06/2021
localization_priority: Normal
ms.openlocfilehash: 54b6b67a0f211f2dc6c881bab17ff23220619e6e
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545782"
---
# <a name="use-office-scripts-and-power-automate-to-email-images-of-a-chart-and-table"></a>Utilice Office Scripts y Power Automate para enviar imágenes por correo electrónico de un gráfico y una tabla

En este ejemplo se usa Office scripts y Power Automate para crear un gráfico. A continuación, envía por correo electrónico imágenes del gráfico y su tabla base.

## <a name="example-scenario"></a>Ejemplo ficticio

* Calcule para obtener los últimos resultados.
* Crear gráfico.
* Obtenga imágenes de gráficos y tablas.
* Envíe por correo electrónico las imágenes con Power Automate.

_Datos de entrada_

:::image type="content" source="../../images/input-data.png" alt-text="Una hoja de trabajo que muestra una tabla de datos de entrada":::

_Gráfico de salida_

:::image type="content" source="../../images/chart-created.png" alt-text="El gráfico de columnas creado que muestra el monto adeudado por el cliente":::

_Correo electrónico que se recibió a través de Power Automate flujo_

:::image type="content" source="../../images/email-received.png" alt-text="El correo electrónico enviado por el flujo que muestra el gráfico de Excel incrustado en el cuerpo":::

## <a name="solution"></a>Solución

Esta solución tiene dos partes:

1. [Un script Office para calcular y extraer Excel gráfico y tabla](#sample-code-calculate-and-extract-excel-chart-and-table)
1. Un flujo Power Automate para invocar el script y enviar por correo electrónico los resultados. Para obtener un ejemplo sobre cómo hacerlo, consulte [Crear un flujo de trabajo automatizado con Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).

## <a name="sample-code-calculate-and-extract-excel-chart-and-table"></a>Código de muestra: calcule y extraiga Excel gráfico y tabla

El siguiente script calcula y extrae un gráfico y una tabla Excel.

Descargar el archivo de ejemplo <a href="email-chart-table.xlsx">email-chart-table.xlsx</a> y utilizarlo con este script para probarlo usted mismo!

```TypeScript
function main(workbook: ExcelScript.Workbook): ReportImages {
  // Recalculate the workbook to ensure all tables and charts are updated.
  workbook.getApplication().calculate(ExcelScript.CalculationType.full);
  
  // Get the data from the "InvoiceAmounts" table.
  let sheet1 = workbook.getWorksheet("Sheet1");
  const table = workbook.getWorksheet('InvoiceAmounts').getTables()[0];
  const rows = table.getRange().getTexts();

  // Get only the "Customer Name" and "Amount due" columns, then remove the "Total" row.
  const selectColumns = rows.map((row) => {
    return [row[2], row[5]];
  });
  table.setShowTotals(true);
  selectColumns.splice(selectColumns.length-1, 1);
  console.log(selectColumns);

  // Delete the "ChartSheet" worksheet if it's present, then recreate it.
  workbook.getWorksheet('ChartSheet')?.delete();
  const chartSheet = workbook.addWorksheet('ChartSheet');

  // Add the selected data to the new worksheet.
  const targetRange = chartSheet.getRange('A1').getResizedRange(selectColumns.length-1, selectColumns[0].length-1);
  targetRange.setValues(selectColumns);

  // Insert the chart on sheet 'ChartSheet' at cell "D1".
  let chart_2 = chartSheet.addChart(ExcelScript.ChartType.columnClustered, targetRange);
  chart_2.setPosition('D1');

  // Get images of the chart and table, then return them for a Power Automate flow.
  const chartImage = chart_2.getImage();
  const tableImage = table.getRange().getImage();
  return {chartImage, tableImage};
}

// The interface for table and chart images.
interface ReportImages {
  chartImage: string
  tableImage: string
}
```

## <a name="power-automate-flow-email-the-chart-and-table-images"></a>flujo Power Automate: Envíe un correo electrónico al gráfico y a las imágenes de la tabla

Este flujo ejecuta el script y envía correos electrónicos a las imágenes devueltas.

1. Cree un nuevo **flujo de nube instantánea.**
1. Seleccione **Activar manualmente un flujo** y pulse **Crear**.
1. Agregue un **nuevo paso** que use el conector Excel en **línea (empresa)** con la acción **Ejecutar script.** Utilice los siguientes valores para la acción:
    * **Ubicación**: OneDrive para la Empresa
    * **Biblioteca de documentos**: OneDrive
    * **Archivo**: Su libro de trabajo ([seleccionado con el selector de archivos](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))
    * **Script**: Su nombre de guión

    :::image type="content" source="../../images/email-chart-sample-flow-1.png" alt-text="El conector Excel online (business) completado en Power Automate":::
1. Este ejemplo utiliza Outlook como cliente de correo electrónico. Puede usar cualquier conector de correo electrónico Power Automate admite, pero el resto de los pasos suponen que eligió Outlook. Agregue un **nuevo paso** que use el conector **de Office 365 Outlook** y la acción Enviar y enviar correo **electrónico (V2).** Utilice los siguientes valores para la acción:
    * **Para**: Su cuenta de correo electrónico de prueba (o correo electrónico personal)
    * **Asunto**: Revise los datos del informe
    * Para el campo **Cuerpo,** seleccione "Vista de código" ( `</>` ) e introduzca lo siguiente:

    ```HTML
    <p>Please review the following report data:<br>
    <br>
    Chart:<br>
    <br>
    <img src="data:image/png;base64,@{outputs('Run_script')?['body/result/chartImage']}"/>
    <br>
    Data:<br>
    <br>
    <img src="data:image/png;base64,@{outputs('Run_script')?['body/result/tableImage']}"/>
    <br>
    </p>
    ```

    :::image type="content" source="../../images/email-chart-sample-flow-2.png" alt-text="El conector de Office 365 Outlook completado en Power Automate":::
1. Guarde el flujo y pruébalo.

## <a name="training-video-extract-and-email-images-of-chart-and-table"></a>Vídeo de entrenamiento: Extraer e imágenes de correo electrónico de gráficos y tablas

[Mira a Sudhi Ramamurthy caminar a través de esta muestra en YouTube.](https://youtu.be/152GJyqc-Kw)
