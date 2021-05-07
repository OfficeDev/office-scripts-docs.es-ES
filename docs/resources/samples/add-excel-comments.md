---
title: Agregar comentarios en Excel
description: Obtenga información sobre cómo usar Office scripts para agregar comentarios en una hoja de cálculo.
ms.date: 04/28/2021
localization_priority: Normal
ms.openlocfilehash: d592b37c3af8e475c81e8650dda44921fee7aeaf
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232511"
---
# <a name="add-comments-in-excel"></a><span data-ttu-id="c10d6-103">Agregar comentarios en Excel</span><span class="sxs-lookup"><span data-stu-id="c10d6-103">Add comments in Excel</span></span>

<span data-ttu-id="c10d6-104">En este ejemplo se muestra cómo agregar comentarios a una celda [@mentioning](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) un compañero.</span><span class="sxs-lookup"><span data-stu-id="c10d6-104">This sample shows how to add comments to a cell including [@mentioning](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) a colleague.</span></span>

## <a name="example-scenario"></a><span data-ttu-id="c10d6-105">Ejemplo ficticio</span><span class="sxs-lookup"><span data-stu-id="c10d6-105">Example scenario</span></span>

* <span data-ttu-id="c10d6-106">El líder del equipo mantiene la programación de turnos.</span><span class="sxs-lookup"><span data-stu-id="c10d6-106">The team lead maintains the shift schedule.</span></span> <span data-ttu-id="c10d6-107">El responsable del equipo asigna un identificador de empleado al registro de turno.</span><span class="sxs-lookup"><span data-stu-id="c10d6-107">The team lead assigns an employee ID to the shift record.</span></span>
* <span data-ttu-id="c10d6-108">El líder del equipo desea notificar al empleado.</span><span class="sxs-lookup"><span data-stu-id="c10d6-108">The team lead wishes to notify the employee.</span></span> <span data-ttu-id="c10d6-109">Al agregar un comentario que @mentions empleado, el empleado se envía por correo electrónico con un mensaje personalizado de la hoja de cálculo.</span><span class="sxs-lookup"><span data-stu-id="c10d6-109">By adding a comment that @mentions the employee, the employee is emailed with a custom message from the worksheet.</span></span>
* <span data-ttu-id="c10d6-110">Posteriormente, el empleado puede ver el libro y responder al comentario a su conveniencia.</span><span class="sxs-lookup"><span data-stu-id="c10d6-110">Subsequently, the employee can view the workbook and respond to the comment at their convenience.</span></span>

## <a name="solution"></a><span data-ttu-id="c10d6-111">Solución</span><span class="sxs-lookup"><span data-stu-id="c10d6-111">Solution</span></span>

1. <span data-ttu-id="c10d6-112">El script extrae información de los empleados de la hoja de cálculo del empleado.</span><span class="sxs-lookup"><span data-stu-id="c10d6-112">The script extracts employee information from the employee worksheet.</span></span>
1. <span data-ttu-id="c10d6-113">A continuación, el script agrega un comentario (incluido el correo electrónico de empleado relevante) a la celda correspondiente en el registro de turno.</span><span class="sxs-lookup"><span data-stu-id="c10d6-113">The script then adds a comment (including the relevant employee email) to the appropriate cell in the shift record.</span></span>
1. <span data-ttu-id="c10d6-114">Los comentarios existentes en la celda se quitan antes de agregar el nuevo comentario.</span><span class="sxs-lookup"><span data-stu-id="c10d6-114">Existing comments in the cell are removed before adding the new comment.</span></span>

## <a name="sample-code-add-comments"></a><span data-ttu-id="c10d6-115">Código de ejemplo: Agregar comentarios</span><span class="sxs-lookup"><span data-stu-id="c10d6-115">Sample code: Add comments</span></span>

<span data-ttu-id="c10d6-116">Descarga el archivo <a href="excel-comments.xlsx">excel-comments.xlsx</a> usado en este ejemplo y pruébalo tú mismo.</span><span class="sxs-lookup"><span data-stu-id="c10d6-116">Download the file <a href="excel-comments.xlsx">excel-comments.xlsx</a> used in this sample and try it out yourself!</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    const employees = workbook.getWorksheet('Employees').getUsedRange().getTexts();
    console.log(employees); 

    const scheduleSheet = workbook.getWorksheet('Schedule');
    const table = scheduleSheet.getTables()[0];
    const range = table.getRangeBetweenHeaderAndTotal();
    const scheduleData = range.getTexts();

    for (let i=0; i < scheduleData.length; i++) {
      let eId = scheduleData[i][3];

      let employeeInfo = employees.find(e => e[0] === eId);
      if (employeeInfo) {
        console.log("Found a match " + employeeInfo);
        let adminNotes = scheduleData[i][4];
        try { 
          let comment = workbook.getCommentByCell(range.getCell(i, 5));
          comment.delete();
        } catch {
            console.log("Ignore if there is no existing comment in the cell");
        }
        workbook.addComment(range.getCell(i,5), {
          mentions: [{
            email: employeeInfo[1],
            id: 0,
            name: employeeInfo[2]
          }],
          richContent: `<at id=\"0\">${employeeInfo[2]}</at> ${adminNotes}`
        }, ExcelScript.ContentType.mention);        
        
      } else {
        console.log("No match for: " + eId);
      }
    }
    return;
}
```

## <a name="training-video-add-comments"></a><span data-ttu-id="c10d6-117">Vídeo de aprendizaje: Agregar comentarios</span><span class="sxs-lookup"><span data-stu-id="c10d6-117">Training video: Add comments</span></span>

<span data-ttu-id="c10d6-118">[Vea el recorrido de Sudhi Ramamurthy por este ejemplo en YouTube](https://youtu.be/CpR78nkaOFw).</span><span class="sxs-lookup"><span data-stu-id="c10d6-118">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/CpR78nkaOFw).</span></span>
