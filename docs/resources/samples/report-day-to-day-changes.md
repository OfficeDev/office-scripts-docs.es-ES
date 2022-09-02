---
title: Registrar los cambios diarios en Excel y notificarlos con un flujo de Power Automate
description: Aprenda a usar scripts de Office y Power Automate para realizar un seguimiento de los cambios de valor en un libro
ms.date: 08/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: 083ca08573db060aa4788aea58fc67e50d004a4b
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572663"
---
# <a name="record-day-to-day-changes-in-excel-and-report-them-with-a-power-automate-flow"></a>Registrar los cambios diarios en Excel y notificarlos con un flujo de Power Automate

Power Automate y los scripts de Office se combinan para controlar tareas repetitivas por usted. En este ejemplo, se le encarga grabar una sola lectura numérica en un libro todos los días y notificar el cambio desde ayer. Creará un flujo para obtener esa lectura, registrarla en el libro e informar del cambio a través de un correo electrónico.

## <a name="sample-excel-file"></a>Archivo de Excel de ejemplo

Descargue [daily-readings.xlsx](daily-readings.xlsx) de un libro listo para usar. Agregue el siguiente script para probar el ejemplo usted mismo.

## <a name="sample-code-record-and-report-daily-readings"></a>Código de ejemplo: registro e informe de lecturas diarias

```TypeScript
function main(workbook: ExcelScript.Workbook, newData: string): string {
  // Get the table by its name.
  const table = workbook.getTable("ReadingTable");

  // Read the current last entry in the Reading column.
  const readingColumn = table.getColumnByName("Reading");
  const readingColumnValues = readingColumn.getRange().getValues();
  const previousValue = readingColumnValues[readingColumnValues.length - 1][0] as number;

  // Add a row with the date, new value, and a formula calculating the difference.
  const currentDate = new Date(Date.now()).toLocaleDateString();
  const newRow = [currentDate, newData, "=[@Reading]-OFFSET([@Reading],-1,0)"];
  table.addRow(-1, newRow,);

  // Return the difference between the newData and the previous entry.
  const difference = Number.parseFloat(newData) - previousValue;
  console.log(difference);
  return difference;
}
```

## <a name="sample-flow-report-day-to-day-changes"></a>Flujo de ejemplo: Notificar cambios diarios

Siga estos pasos para crear un flujo [de Power Automate](https://powerautomate.microsoft.com/) para el ejemplo.

1. Cree un nuevo **flujo de nube programada**.
1. Programe el flujo para que se repita cada **1 día**.

    :::image type="content" source="../../images/day-to-day-changes-flow-1.png" alt-text="El paso de creación de flujo que muestra que se repetirá todos los días.":::
1. Seleccione **Crear**.
1. En un flujo real, agregará un paso que obtiene los datos. Los datos pueden proceder de otro libro, una tarjeta adaptable de Teams o cualquier otro origen. Para probar el ejemplo, realice un número de prueba. Agregue un nuevo paso con la acción **Inicializar variable** . Asígnele los siguientes valores.
    1. **Nombre**: entrada
    1. **Tipo**: Integer
    1. **Valor**: 190000

    :::image type="content" source="../../images/day-to-day-changes-flow-2.png" alt-text="Acción Inicializar variable con los valores especificados.":::
1. Agregue un nuevo paso con el conector **de Excel Online (Empresa)** con la acción **Ejecutar script** . Use los siguientes valores para la acción.
    1. **Ubicación**: OneDrive para la Empresa
    1. **Biblioteca de documentos**: OneDrive
    1. **Archivo**: daily-readings.xlsx *(elegido mediante el explorador de archivos)*
    1. **Script**: nombre del script
    1. **newData**: entrada *(contenido dinámico)*

    :::image type="content" source="../../images/day-to-day-changes-flow-3.png" alt-text="La acción Ejecutar script con los valores especificados.":::
1. El script devuelve la diferencia de lectura diaria como contenido dinámico denominado "result". Para el ejemplo, puede enviar la información por correo electrónico a usted mismo. Cree un nuevo paso que use el conector de **Outlook** con la acción **Enviar un correo electrónico (V2) (** o cualquier cliente de correo electrónico que prefiera). Use los siguientes valores para completar la acción.
    1. **Para**: Su dirección de correo electrónico
    1. **Asunto**: Cambio de lectura diario
    1. **Cuerpo**: resultado "Diferencia de ayer" *(contenido dinámico de Excel)*

    :::image type="content" source="../../images/day-to-day-changes-flow-4.png" alt-text="El conector de Outlook completado en Power Automate.":::
1. Guarde el flujo y pruébelo. Use el botón **Probar** de la página del editor de flujo. Asegúrese de permitir el acceso cuando se le solicite.
