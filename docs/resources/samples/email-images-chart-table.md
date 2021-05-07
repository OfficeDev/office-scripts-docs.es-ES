---
title: Enviar por correo electrónico las imágenes de un Excel gráfico y tabla
description: Obtenga información sobre cómo usar Office scripts y Power Automate para extraer y enviar por correo electrónico las imágenes de un Excel gráfico y tabla.
ms.date: 04/28/2021
localization_priority: Normal
ms.openlocfilehash: b49b6670562d117bb3dd6dcf894c54432bc5ceaa
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232595"
---
# <a name="use-office-scripts-and-power-automate-to-email-images-of-a-chart-and-table"></a>Usar Office scripts y Power Automate para enviar por correo electrónico imágenes de un gráfico y una tabla

En este ejemplo se Office scripts y Power Automate para crear un gráfico. A continuación, envía un correo electrónico a las imágenes del gráfico y su tabla base.

## <a name="example-scenario"></a>Ejemplo ficticio

* Calcule para obtener los resultados más recientes.
* Crear gráfico.
* Obtener imágenes de gráfico y tabla.
* Envíe por correo electrónico las imágenes Power Automate.

_Datos de entrada_

:::image type="content" source="../../images/input-data.png" alt-text="Una hoja de cálculo que muestra una tabla de datos de entrada":::

_Gráfico de salida_

:::image type="content" source="../../images/chart-created.png" alt-text="Gráfico de columnas creado que muestra la cantidad adeudada por el cliente":::

_Correo electrónico que se recibió a través Power Automate flujo_

:::image type="content" source="../../images/email-received.png" alt-text="El correo electrónico enviado por el flujo que muestra Excel gráfico incrustado en el cuerpo":::

## <a name="solution"></a>Solución

Esta solución tiene dos partes:

1. [Un script Office para calcular y extraer Excel gráfico y tabla](#sample-code-calculate-and-extract-excel-chart-and-table)
1. Flujo Power Automate para invocar el script y enviar por correo electrónico los resultados. Para obtener un ejemplo sobre cómo hacerlo, vea [Create an automated workflow with Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).

## <a name="sample-code-calculate-and-extract-excel-chart-and-table"></a>Código de ejemplo: calcular y extraer Excel gráfico y tabla

El siguiente script calcula y extrae un Excel gráfico y tabla.

Descargue el archivo de <a href="email-chart-table.xlsx">email-chart-table.xlsx</a> y ústelo con este script para probarlo usted mismo.

```TypeScript
function main(workbook: ExcelScript.Workbook): ReportImages {

  workbook.getApplication().calculate(ExcelScript.CalculationType.full);
  
  let sheet1 = workbook.getWorksheet("Sheet1");
  const table = workbook.getWorksheet('InvoiceAmounts').getTables()[0];
  const rows = table.getRange().getTexts();

  const selectColumns = rows.map((row) => {
    return [row[2], row[5]];
  });
  table.setShowTotals(true);
  selectColumns.splice(selectColumns.length-1, 1);
  console.log(selectColumns);

  workbook.getWorksheet('ChartSheet')?.delete();
  const chartSheet = workbook.addWorksheet('ChartSheet');
  const targetRange = updateRange(chartSheet, selectColumns);

  // Insert chart on sheet 'Sheet1'.
  let chart_2 = chartSheet.addChart(ExcelScript.ChartType.columnClustered, targetRange);
  chart_2.setPosition('D1');
  const chartImage = chart_2.getImage();
  const tableImage = table.getRange().getImage();
  return {
    chartImage,
    tableImage
  }
}

function updateRange(sheet: ExcelScript.Worksheet, data: string[][]): ExcelScript.Range {
  const targetRange = sheet.getRange('A1').getResizedRange(data.length-1, data[0].length-1);
  targetRange.setValues(data);
  return targetRange;
}

interface ReportImages {
  chartImage: string
  tableImage: string
}
```

## <a name="power-automate-flow-email-the-chart-and-table-images"></a>Power Automate: enviar por correo electrónico las imágenes del gráfico y de la tabla

Este flujo ejecuta el script y envía correos electrónicos a las imágenes devueltas.

1. Crear un nuevo **flujo de nube instantánea**.
1. Seleccione **Desencadenar manualmente un flujo y** presione **Crear**.
1. Agregue un **nuevo paso** que use el conector Excel **online (empresa)** con la **acción Ejecutar script (versión** preliminar). Use los siguientes valores para la acción:
    * **Ubicación**: OneDrive para la Empresa
    * **Biblioteca de documentos**: OneDrive
    * **Archivo:** el libro ([seleccionado con el seleccionador de archivos](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))
    * **Script:** el nombre del script

    :::image type="content" source="../../images/email-chart-sample-flow-1.png" alt-text="El conector Excel online (empresa) completado en Power Automate":::
1. En este ejemplo se Outlook como cliente de correo electrónico. Puede usar cualquier conector de correo Power Automate admite, pero el resto de los pasos supone que eligió Outlook. Agregue un **nuevo paso** que use el **conector Office 365 Outlook** y la acción Enviar y correo electrónico **(V2).** Use los siguientes valores para la acción:
    * **To**: Su cuenta de correo electrónico de prueba (o correo electrónico personal)
    * **Asunto**: Revise los datos del informe
    * Para el **campo Cuerpo,** seleccione "Vista de código" ( `</>` ) y escriba lo siguiente:

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

    :::image type="content" source="../../images/email-chart-sample-flow-2.png" alt-text="El conector Office 365 Outlook completado en Power Automate":::
1. Guarde el flujo y pruébalo.

## <a name="training-video-extract-and-email-images-of-chart-and-table"></a>Vídeo de aprendizaje: extraer y enviar por correo electrónico imágenes de gráfico y tabla

[Vea el recorrido de Sudhi Ramamurthy por este ejemplo en YouTube](https://youtu.be/152GJyqc-Kw).
