---
title: Agregar comentarios en Excel
description: Obtenga información sobre cómo usar scripts de Office para agregar comentarios en una hoja de cálculo.
ms.date: 03/29/2021
localization_priority: Normal
ms.openlocfilehash: aaaf26df6973bd081290b0fbb67edecad8627e53
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571628"
---
# <a name="add-comments-in-excel"></a><span data-ttu-id="c58ea-103">Agregar comentarios en Excel</span><span class="sxs-lookup"><span data-stu-id="c58ea-103">Add comments in Excel</span></span>

<span data-ttu-id="c58ea-104">En este ejemplo se muestra cómo agregar comentarios a una celda [@mentioning](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) un compañero.</span><span class="sxs-lookup"><span data-stu-id="c58ea-104">This sample shows how to add comments to a cell including [@mentioning](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) a colleague.</span></span>

## <a name="example-scenario"></a><span data-ttu-id="c58ea-105">Escenario de ejemplo</span><span class="sxs-lookup"><span data-stu-id="c58ea-105">Example scenario</span></span>

* <span data-ttu-id="c58ea-106">El líder del equipo mantiene la programación de turnos.</span><span class="sxs-lookup"><span data-stu-id="c58ea-106">The team lead maintains the shift schedule.</span></span> <span data-ttu-id="c58ea-107">El responsable del equipo asigna un identificador de empleado al registro de turno.</span><span class="sxs-lookup"><span data-stu-id="c58ea-107">The team lead assigns an employee ID to the shift record.</span></span>
* <span data-ttu-id="c58ea-108">El líder del equipo desea notificar al empleado.</span><span class="sxs-lookup"><span data-stu-id="c58ea-108">The team lead wishes to notify the employee.</span></span> <span data-ttu-id="c58ea-109">Al agregar un comentario que @mentions empleado, el empleado se envía por correo electrónico con un mensaje personalizado de la hoja de cálculo.</span><span class="sxs-lookup"><span data-stu-id="c58ea-109">By adding a comment that @mentions the employee, the employee is emailed with a custom message from the worksheet.</span></span>
* <span data-ttu-id="c58ea-110">Posteriormente, el empleado puede ver el libro y responder al comentario a su conveniencia.</span><span class="sxs-lookup"><span data-stu-id="c58ea-110">Subsequently, the employee can view the workbook and respond to the comment at their convenience.</span></span>

## <a name="solution"></a><span data-ttu-id="c58ea-111">Solución</span><span class="sxs-lookup"><span data-stu-id="c58ea-111">Solution</span></span>

1. <span data-ttu-id="c58ea-112">El script extrae información de los empleados de la hoja de cálculo del empleado.</span><span class="sxs-lookup"><span data-stu-id="c58ea-112">The script extracts employee information from the employee worksheet.</span></span>
1. <span data-ttu-id="c58ea-113">A continuación, el script agrega un comentario (incluido el correo electrónico de empleado relevante) a la celda correspondiente en el registro de turno.</span><span class="sxs-lookup"><span data-stu-id="c58ea-113">The script then adds a comment (including the relevant employee email) to the appropriate cell in the shift record.</span></span>
1. <span data-ttu-id="c58ea-114">Los comentarios existentes en la celda se quitan antes de agregar el nuevo comentario.</span><span class="sxs-lookup"><span data-stu-id="c58ea-114">Existing comments in the cell are removed before adding the new comment.</span></span>

## <a name="sample-code-add-comments"></a><span data-ttu-id="c58ea-115">Código de ejemplo: Agregar comentarios</span><span class="sxs-lookup"><span data-stu-id="c58ea-115">Sample code: Add comments</span></span>

<span data-ttu-id="c58ea-116">Descarga el archivo <a href="excel-comments.xlsx">excel-comments.xlsx</a> usado en este ejemplo y pruébalo tú mismo.</span><span class="sxs-lookup"><span data-stu-id="c58ea-116">Download the file <a href="excel-comments.xlsx">excel-comments.xlsx</a> used in this sample and try it out yourself!</span></span>

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

## <a name="training-video-add-comments"></a><span data-ttu-id="c58ea-117">Vídeo de aprendizaje: Agregar comentarios</span><span class="sxs-lookup"><span data-stu-id="c58ea-117">Training video: Add comments</span></span>

<span data-ttu-id="c58ea-118">[![Ver vídeo paso a paso sobre cómo agregar comentarios en un archivo de Excel](../../images/comments-vid.jpg)](https://youtu.be/CpR78nkaOFw "Vídeo paso a paso sobre cómo agregar comentarios en un archivo de Excel")</span><span class="sxs-lookup"><span data-stu-id="c58ea-118">[![Watch step-by-step video on how to add comments in an Excel file](../../images/comments-vid.jpg)](https://youtu.be/CpR78nkaOFw "Step-by-step video on how to add comments in an Excel file")</span></span>
