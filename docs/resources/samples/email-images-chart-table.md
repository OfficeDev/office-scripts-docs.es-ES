---
title: Enviar por correo electrónico las imágenes de Excel gráfico y tabla
description: Obtenga información sobre cómo usar Office scripts y Power Automate para extraer y enviar por correo electrónico las imágenes de un Excel gráfico y tabla.
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: 2930a70a5bed4eb49f33f315460ae32f40b5a2f2
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585509"
---
# <a name="use-office-scripts-and-power-automate-to-email-images-of-a-chart-and-table"></a>Usar Office scripts y Power Automate para enviar por correo electrónico imágenes de un gráfico y una tabla

En este ejemplo se Office scripts y Power Automate para crear un gráfico. A continuación, envía un correo electrónico a las imágenes del gráfico y su tabla base.

## <a name="example-scenario"></a>Escenario de ejemplo

* Calcule para obtener los resultados más recientes.
* Crear gráfico.
* Obtener imágenes de gráfico y tabla.
* Envíe un correo electrónico a las imágenes Power Automate.

_Datos de entrada_

:::image type="content" source="../../images/input-data.png" alt-text="Hoja de cálculo que muestra una tabla de datos de entrada.":::

_Gráfico de salida_

:::image type="content" source="../../images/chart-created.png" alt-text="Gráfico de columnas creado que muestra el importe debido por cliente.":::

_Correo electrónico que se recibió a través Power Automate flujo_

:::image type="content" source="../../images/email-received.png" alt-text="El correo electrónico enviado por el flujo que muestra Excel gráfico incrustado en el cuerpo.":::

## <a name="solution"></a>Solución

Esta solución tiene dos partes:

1. [Un script Office para calcular y extraer Excel gráfico y tabla](#sample-code-calculate-and-extract-excel-chart-and-table)
1. Flujo Power Automate para invocar el script y enviar por correo electrónico los resultados. Para obtener un ejemplo sobre cómo hacerlo, vea [Create an automated workflow with Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).

## <a name="sample-excel-file"></a>Archivo Excel ejemplo

Descargue <a href="email-chart-table.xlsx">email-chart-table.xlsx</a> para un libro listo para usar. Agregue el siguiente script para probar el ejemplo usted mismo.

## <a name="sample-code-calculate-and-extract-excel-chart-and-table"></a>Código de ejemplo: calcular y extraer Excel gráfico y tabla

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

## <a name="power-automate-flow-email-the-chart-and-table-images"></a>Power Automate: enviar por correo electrónico las imágenes del gráfico y de la tabla

Este flujo ejecuta el script y envía correos electrónicos a las imágenes devueltas.

1. Cree un nuevo **flujo de nube instantánea**.
1. Elija **Desencadenar manualmente un flujo y** seleccione **Crear**.
1. Agregue un **paso Nuevo** que use el **conector Excel Online (Empresa)** con la **acción Ejecutar script**. Use los siguientes valores para la acción.
    * **Ubicación**: OneDrive para la Empresa
    * **Biblioteca de documentos**: OneDrive
    * **Archivo**: el libro ([seleccionado con el seleccionador de archivos](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))
    * **Script**: el nombre del script

    :::image type="content" source="../../images/email-chart-sample-flow-1.png" alt-text="El conector Excel online (empresa) completado en Power Automate.":::
1. En este ejemplo se Outlook como cliente de correo electrónico. Puede usar cualquier conector de correo Power Automate admite, pero el resto de los pasos supone que eligió Outlook. Agregue un **nuevo paso** que use el **conector Office 365 Outlook** y la **acción Enviar y correo electrónico (V2**). Use los siguientes valores para la acción.
    * **Para**: Su cuenta de correo electrónico de prueba (o correo electrónico personal)
    * **Asunto**: Revise los datos del informe
    * Para el **campo Cuerpo** , seleccione "Vista de código" (`</>`) y escriba lo siguiente:

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

    :::image type="content" source="../../images/email-chart-sample-flow-2.png" alt-text="El conector Office 365 Outlook completo en Power Automate.":::
1. Guarde el flujo y pruébalo. Use el **botón Probar** en la página del editor de flujo o ejecute el flujo a través de la **pestaña Mis flujos** . Asegúrese de permitir el acceso cuando se le pida.

## <a name="training-video-extract-and-email-images-of-chart-and-table"></a>Vídeo de aprendizaje: extraer y enviar por correo electrónico imágenes de gráfico y tabla

[Vea el recorrido de Sudhi Ramamurthy a través de esta muestra en YouTube](https://youtu.be/152GJyqc-Kw).
