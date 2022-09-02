---
title: Agregar comentarios en Excel
description: Obtenga información sobre cómo usar scripts de Office para agregar comentarios en una hoja de cálculo.
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: 90f072805e6798a4f9d6e74889ccca15610c87bd
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572496"
---
# <a name="add-comments-in-excel"></a>Agregar comentarios en Excel

En este ejemplo se muestra cómo agregar comentarios a una celda, incluido [@mentioning](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) un compañero.

## <a name="example-scenario"></a>Escenario de ejemplo

* El responsable del equipo mantiene la programación de turnos. El responsable del equipo asigna un identificador de empleado al registro de turnos.
* El jefe del equipo desea notificar al empleado. Al agregar un comentario que @mentions el empleado, el empleado se envía por correo electrónico con un mensaje personalizado de la hoja de cálculo.
* Posteriormente, el empleado puede ver el libro y responder al comentario a su conveniencia.

## <a name="solution"></a>Solución

1. El script extrae información del empleado de la hoja de cálculo del empleado.
1. A continuación, el script agrega un comentario (incluido el correo electrónico del empleado correspondiente) a la celda adecuada del registro de turnos.
1. Los comentarios existentes en la celda se quitan antes de agregar el nuevo comentario.

## <a name="sample-excel-file"></a>Archivo de Excel de ejemplo

Descargue [excel-comments.xlsx](excel-comments.xlsx) de un libro listo para usar. Agregue el siguiente script para probar el ejemplo usted mismo.

## <a name="sample-code-add-comments"></a>Código de ejemplo: Agregar comentarios

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the list of employees.
  const employees = workbook.getWorksheet('Employees').getUsedRange().getTexts();
  console.log(employees); 
  
  // Get the schedule information from the schedule table.
  const scheduleSheet = workbook.getWorksheet('Schedule');
  const table = scheduleSheet.getTables()[0];
  const range = table.getRangeBetweenHeaderAndTotal();
  const scheduleData = range.getTexts();

  // Look through the schedule for a matching employee.
  for (let i = 0; i < scheduleData.length; i++) {
    let employeeId = scheduleData[i][3];

    // Compare the employee ID in the schedule against the employee information table.
    let employeeInfo = employees.find(employeeRow => employeeRow[0] === employeeId);
    if (employeeInfo) {
      console.log("Found a match " + employeeInfo);
      let adminNotes = scheduleData[i][4];

      // Look for and delete old comments, so we avoid conflicts.
      let comment = workbook.getCommentByCell(range.getCell(i, 5));
      if (comment) {
        comment.delete();
      }

      // Add a comment using the admin notes as the text.
      workbook.addComment(range.getCell(i,5), {
        mentions: [{
          email: employeeInfo[1],
          id: 0, // This ID maps this mention to the `id=0` text in the comment.
          name: employeeInfo[2]
        }],
        richContent: `<at id=\"0\">${employeeInfo[2]}</at> ${adminNotes}`
      }, ExcelScript.ContentType.mention);        
      
    } else {
      console.log("No match for: " + employeeId);
    }
  }
}
```

## <a name="training-video-add-comments"></a>Vídeo de entrenamiento: Adición de comentarios

[Vea cómo Sudhi Ramamurthy recorre este ejemplo en YouTube](https://youtu.be/CpR78nkaOFw).
