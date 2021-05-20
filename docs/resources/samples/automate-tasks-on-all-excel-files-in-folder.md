---
title: Ejecutar un script en todos los archivos de Excel de una carpeta
description: Obtén información sobre cómo ejecutar un script en todos los archivos de Excel de una carpeta de OneDrive para la Empresa.
ms.date: 05/03/2021
localization_priority: Normal
ms.openlocfilehash: fb9a4deb01b52ef031cb1ba3400bd6f10de9d9f5
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545797"
---
# <a name="run-a-script-on-all-excel-files-in-a-folder"></a>Ejecutar un script en todos los archivos de Excel de una carpeta

Este proyecto realiza un conjunto de tareas de automatización en todos los archivos situados en una carpeta de OneDrive para la Empresa. También se puede utilizar en una carpeta de SharePoint.
Realiza cálculos en los archivos Excel, agrega formato e inserta un comentario que [@mentions](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) un colega.

Descargar el archivo <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/highlight-alert-excel-files.zip?raw=true">highlight-alert-excel-files.zip</a>, extraer los archivos en una carpeta titulada **Ventas** utilizadas en este ejemplo, y probarlo usted mismo!

## <a name="sample-code-add-formatting-and-insert-comment"></a>Código de ejemplo: Agregue el formato y inserte el comentario

Este es el script que se ejecuta en cada libro individual.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the table named "Table1" in the workbook.
  let table1 = workbook.getTable("Table1");

  // If the table is empty, end the script.
  const rowCount = table1.getRowCount();
  if (rowCount === 0) {
    return;
  }

  // Force the workbook to be completely recalculated.
  workbook.getApplication().calculate(ExcelScript.CalculationType.full);

  // Get the "Amount Due" column from the table.
  const amountDueColumn = table1.getColumnByName('Amount Due');
  const amountDueValues = amountDueColumn.getRangeBetweenHeaderAndTotal().getValues();

  // Find the highest amount that's due.
  let highestValue = amountDueValues[0][0];
  let row = 0;
  for (let i = 1; i < amountDueValues.length; i++) {
    if (amountDueValues[i][0] > highestValue) {
      highestValue = amountDueValues[i][0];
      row = i;
    }
  }

  let highestAmountDue = table1.getColumn("Amount due").getRangeBetweenHeaderAndTotal().getRow(row);

  // Set the fill color to yellow for the cell with the highest value in the "Amount Due" column.
  highestAmountDue
    .getFormat()
    .getFill()
    .setColor("FFFF00");

  // Insert an @mention comment in the cell.
  workbook.addComment(highestAmountDue, {
    mentions: [{
      email: "AdeleV@M365x904181.OnMicrosoft.com",
      id: 0,
      name: "Adele Vance"
    }],
    richContent: "<at id=\"0\">Adele Vance</at> Please review this amount"
  }, ExcelScript.ContentType.mention);
}
```

## <a name="power-automate-flow-run-the-script-on-every-workbook-in-the-folder"></a>flujo de Power Automate: ejecute el script en cada libro de trabajo de la carpeta

Este flujo ejecuta el script en cada libro de trabajo de la carpeta "Ventas".

1. Cree un nuevo **flujo de nube instantánea.**
1. Seleccione **Activar manualmente un flujo** y pulse **Crear**.
1. Agregue un **nuevo paso** que use el conector **de OneDrive para la Empresa** y la acción **Listar archivos en carpeta.**

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-1.png" alt-text="El conector de OneDrive para la Empresa completado en Power Automate":::
1. Seleccione la carpeta "Ventas" con los libros extraídos.
1. Para asegurarse de que solo se seleccionan libros de trabajo, elija **Nuevo paso** y, a continuación, seleccione **Condición** y establezca los siguientes valores:
    1. **Nombre** (el valor OneDrive nombre de archivo)
    1. "Termina con"
    1. "xlsx".

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-2.png" alt-text="El bloque de condición Power Automate que aplica acciones posteriores a cada archivo":::
1. En la rama **Si sí,** agregue el conector **Excel en línea (empresa)** con la acción **Ejecutar script.** Utilice los siguientes valores para la acción:
    1. **Ubicación**: OneDrive para la Empresa
    1. **Biblioteca de documentos**: OneDrive
    1. **Archivo**: **Id** (el valor de ID de archivo OneDrive)
    1. **Script**: Su nombre de guión

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-3.png" alt-text="El conector Excel online (business) completado en Power Automate":::
1. Guarde el flujo y pruébalo.

## <a name="training-video-run-a-script-on-all-excel-files-in-a-folder"></a>Vídeo de entrenamiento: ejecute un script en todos los archivos Excel de una carpeta

[Mira a Sudhi Ramamurthy caminar a través de esta muestra en YouTube.](https://youtu.be/xMg711o7k6w)
