---
title: Agregar comentarios en Excel
description: Obtenga información sobre cómo usar Office scripts para agregar comentarios en una hoja de cálculo.
ms.date: 05/03/2021
localization_priority: Normal
ms.openlocfilehash: e5e5d17c076eceaf06fddeea1a67d31ee3581f31
ms.sourcegitcommit: 763d341857bcb209b2f2c278a82fdb63d0e18f0a
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 05/08/2021
ms.locfileid: "52285937"
---
# <a name="add-comments-in-excel"></a>Agregar comentarios en Excel

En este ejemplo se muestra cómo agregar comentarios a una celda [@mentioning](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) un compañero.

## <a name="example-scenario"></a>Ejemplo ficticio

* El líder del equipo mantiene la programación de turnos. El responsable del equipo asigna un identificador de empleado al registro de turno.
* El líder del equipo desea notificar al empleado. Al agregar un comentario que @mentions empleado, el empleado se envía por correo electrónico con un mensaje personalizado de la hoja de cálculo.
* Posteriormente, el empleado puede ver el libro y responder al comentario a su conveniencia.

## <a name="solution"></a>Solución

1. El script extrae información de los empleados de la hoja de cálculo del empleado.
1. A continuación, el script agrega un comentario (incluido el correo electrónico de empleado relevante) a la celda correspondiente en el registro de turno.
1. Los comentarios existentes en la celda se quitan antes de agregar el nuevo comentario.

## <a name="sample-code-add-comments"></a>Código de ejemplo: Agregar comentarios

Descarga el archivo <a href="excel-comments.xlsx">excel-comments.xlsx</a> usado en este ejemplo y pruébalo tú mismo.

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

## <a name="training-video-add-comments"></a>Vídeo de aprendizaje: Agregar comentarios

[Vea el recorrido de Sudhi Ramamurthy por este ejemplo en YouTube](https://youtu.be/CpR78nkaOFw).
