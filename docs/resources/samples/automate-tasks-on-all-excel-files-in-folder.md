---
title: Ejecutar un script en todos los archivos de Excel de una carpeta
description: Obtenga información sobre cómo ejecutar un script en todos los Excel archivos de una carpeta en OneDrive para la Empresa.
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

Este proyecto realiza un conjunto de tareas de automatización en todos los archivos situados en una carpeta de OneDrive para la Empresa. También se puede usar en una SharePoint carpeta.
Realiza cálculos en los archivos Excel, agrega formato e inserta un comentario que @mentions [compañero.](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7)

Descargue el archivo <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/highlight-alert-excel-files.zip?raw=true">highlight-alert-excel-files.zip</a>, extraiga los archivos en una carpeta titulada **Ventas** usada en este ejemplo y pruébalo usted mismo.

## <a name="sample-code-add-formatting-and-insert-comment"></a>Código de ejemplo: Agregar formato e insertar comentario

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

## <a name="power-automate-flow-run-the-script-on-every-workbook-in-the-folder"></a>Power Automate: ejecute el script en todos los libros de la carpeta

Este flujo ejecuta el script en todos los libros de la carpeta "Ventas".

1. Crear un nuevo **flujo de nube instantánea**.
1. Seleccione **Desencadenar manualmente un flujo y** presione **Crear**.
1. Agregue un **nuevo paso que** use el conector **OneDrive para la Empresa** y los archivos de lista en la acción **de carpeta.**

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-1.png" alt-text="El conector OneDrive para la Empresa en Power Automate":::
1. Seleccione la carpeta "Ventas" con los libros extraídos.
1. Para asegurarse de que solo están seleccionados los libros, elija **Nuevo paso** y, a continuación, **seleccione Condición** y establezca los siguientes valores:
    1. **Name** (el OneDrive de nombre de archivo)
    1. "termina con"
    1. "xlsx".

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-2.png" alt-text="El Power Automate condición que aplica acciones posteriores a cada archivo":::
1. En la **rama If yes,** agregue **el conector Excel Online (Empresa)** con la **acción Ejecutar script.** Use los siguientes valores para la acción:
    1. **Ubicación**: OneDrive para la Empresa
    1. **Biblioteca de documentos**: OneDrive
    1. **Archivo:** **Identificador** (el valor OneDrive id. de archivo)
    1. **Script:** el nombre del script

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-3.png" alt-text="El conector Excel online (empresa) completado en Power Automate":::
1. Guarde el flujo y pruébalo.

## <a name="training-video-run-a-script-on-all-excel-files-in-a-folder"></a>Vídeo de aprendizaje: ejecutar un script en todos Excel archivos de una carpeta

[Vea el recorrido de Sudhi Ramamurthy por este ejemplo en YouTube](https://youtu.be/xMg711o7k6w).
