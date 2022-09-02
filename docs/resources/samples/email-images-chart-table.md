---
title: Email las imágenes de un gráfico y una tabla de Excel
description: Obtenga información sobre cómo usar scripts de Office y Power Automate para extraer y enviar por correo electrónico las imágenes de un gráfico y una tabla de Excel.
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: dbf9135723a735321c99991d94f4b4387d800702
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572468"
---
# <a name="use-office-scripts-and-power-automate-to-email-images-of-a-chart-and-table"></a>Uso de scripts de Office y Power Automate para enviar imágenes por correo electrónico de un gráfico y una tabla

En este ejemplo se usan scripts de Office y Power Automate para crear un gráfico. A continuación, envía mensajes de correo electrónico a las imágenes del gráfico y su tabla base.

## <a name="example-scenario"></a>Escenario de ejemplo

* Calcule para obtener los resultados más recientes.
* Crear gráfico.
* Obtener imágenes de gráficos y tablas.
* Email las imágenes con Power Automate.

_Datos de entrada_

:::image type="content" source="../../images/input-data.png" alt-text="Hoja de cálculo que muestra una tabla de datos de entrada.":::

_Gráfico de salida_

:::image type="content" source="../../images/chart-created.png" alt-text="Gráfico de columnas creado que muestra la cantidad vencida por el cliente.":::

_Email que se recibió a través del flujo de Power Automate_

:::image type="content" source="../../images/email-received.png" alt-text="Correo electrónico enviado por el flujo que muestra el gráfico de Excel incrustado en el cuerpo.":::

## <a name="solution"></a>Solución

Esta solución tiene dos partes:

1. [Un script de Office para calcular y extraer gráficos y tablas de Excel](#sample-code-calculate-and-extract-excel-chart-and-table)
1. Un flujo de Power Automate para invocar el script y enviar por correo electrónico los resultados. Para obtener un ejemplo sobre cómo hacerlo, consulte [Creación de un flujo de trabajo automatizado con Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).

## <a name="sample-excel-file"></a>Archivo de Excel de ejemplo

Descargue [email-chart-table.xlsx](email-chart-table.xlsx) de un libro listo para usar. Agregue el siguiente script para probar el ejemplo usted mismo.

## <a name="sample-code-calculate-and-extract-excel-chart-and-table"></a>Código de ejemplo: calcular y extraer gráficos y tablas de Excel

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

## <a name="power-automate-flow-email-the-chart-and-table-images"></a>Flujo de Power Automate: Email las imágenes de gráfico y tabla

Este flujo ejecuta el script y envía correos electrónicos a las imágenes devueltas.

1. Cree un flujo **de nube instantáneo**.
1. Elija **Desencadenar manualmente un flujo** y seleccione **Crear**.
1. Agregue un **nuevo paso** que use el conector **de Excel Online (Empresa)** con la acción **Ejecutar script** . Use los siguientes valores para la acción.
    * **Ubicación**: OneDrive para la Empresa
    * **Biblioteca de documentos**: OneDrive
    * **Archivo**: libro ([seleccionado con el selector de archivos](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))
    * **Script**: nombre del script

    :::image type="content" source="../../images/email-chart-sample-flow-1.png" alt-text="Conector de Excel Online (Empresa) completado en Power Automate.":::
1. En este ejemplo se usa Outlook como cliente de correo electrónico. Puede usar cualquier conector de correo electrónico compatible con Power Automate, pero en el resto de los pasos se supone que eligió Outlook. Agregue un **nuevo paso** que use el conector **Office 365 Outlook** y la acción **Enviar y enviar correo electrónico (V2).** Use los siguientes valores para la acción.
    * **Para**: Su cuenta de correo electrónico de prueba (o correo electrónico personal)
    * **Asunto**: Revise los datos del informe
    * En el campo **Cuerpo** , seleccione "Vista de código" (`</>`) y escriba lo siguiente:

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

    :::image type="content" source="../../images/email-chart-sample-flow-2.png" alt-text="El conector de Outlook Office 365 completado en Power Automate.":::
1. Guarde el flujo y pruébelo. Use el botón **Probar** de la página del editor de flujo o ejecute el flujo a través de la pestaña **Mis flujos** . Asegúrese de permitir el acceso cuando se le solicite.

## <a name="training-video-extract-and-email-images-of-chart-and-table"></a>Vídeo de entrenamiento: Extracción y correo electrónico de imágenes de gráfico y tabla

[Vea cómo Sudhi Ramamurthy recorre este ejemplo en YouTube](https://youtu.be/152GJyqc-Kw).
