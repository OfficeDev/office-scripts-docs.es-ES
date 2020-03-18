---
title: 'Escenario de ejemplo de scripts de Office: Calculadora de calificaciones'
description: Un ejemplo que determina el porcentaje y las calificaciones de una clase de alumnos.
ms.date: 02/20/2020
localization_priority: Normal
ms.openlocfilehash: 0db6f7c116594f7655bfc0adc8f5a79dbbf2a0af
ms.sourcegitcommit: b075eed5a6f275274fbbf6d62633219eac416f26
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 03/10/2020
ms.locfileid: "42700403"
---
# <a name="office-scripts-sample-scenario-grade-calculator"></a><span data-ttu-id="b1217-103">Escenario de ejemplo de scripts de Office: Calculadora de calificaciones</span><span class="sxs-lookup"><span data-stu-id="b1217-103">Office Scripts sample scenario: Grade calculator</span></span>

<span data-ttu-id="b1217-104">En este escenario, usted es un instructor que rellene las calificaciones de fin de período de cada estudiante.</span><span class="sxs-lookup"><span data-stu-id="b1217-104">In this scenario, you're an instructor tallying every student's end-of-term grades.</span></span> <span data-ttu-id="b1217-105">Ha estado especificando los resultados de sus asignaciones y pruebas a medida que avanza.</span><span class="sxs-lookup"><span data-stu-id="b1217-105">You've been entering the scores for their assignments and tests as you go.</span></span> <span data-ttu-id="b1217-106">Ahora, es el momento de determinar los Fates de los alumnos.</span><span class="sxs-lookup"><span data-stu-id="b1217-106">Now, it is time to determine the students' fates.</span></span>

<span data-ttu-id="b1217-107">Desarrollará un script que totaliza las calificaciones para cada categoría de punto.</span><span class="sxs-lookup"><span data-stu-id="b1217-107">You'll develop a script that totals the grades for each point category.</span></span> <span data-ttu-id="b1217-108">A continuación, asignará una letra de calificación a cada estudiante en función del total.</span><span class="sxs-lookup"><span data-stu-id="b1217-108">It will then assign a letter grade to each student based on the total.</span></span> <span data-ttu-id="b1217-109">Para ayudar a garantizar la precisión, agregará un par de comprobaciones para ver si alguna puntuación individual es demasiado baja o alta.</span><span class="sxs-lookup"><span data-stu-id="b1217-109">To help ensure accuracy, you'll add a couple checks to see if any individual scores are too low or high.</span></span> <span data-ttu-id="b1217-110">Si la puntuación de un estudiante es menor que cero o mayor que el valor de punto posible, el script marcará la celda con un relleno rojo y no hará un total de los puntos del estudiante.</span><span class="sxs-lookup"><span data-stu-id="b1217-110">If a student's score is less than zero or more than the possible point value, the script will flag the cell with a red fill and not total that student's points.</span></span> <span data-ttu-id="b1217-111">Esto será una indicación clara de los registros que debe comprobar dos veces.</span><span class="sxs-lookup"><span data-stu-id="b1217-111">This will be a clear indication of which records you need to double-check.</span></span> <span data-ttu-id="b1217-112">También agregará formato básico a las calificaciones para que pueda ver rápidamente la parte superior e inferior de la clase.</span><span class="sxs-lookup"><span data-stu-id="b1217-112">You'll also add some basic formatting to the grades so you can quickly view the top and bottom of the class.</span></span>

## <a name="scripting-skills-covered"></a><span data-ttu-id="b1217-113">Habilidades de scripting cubiertas</span><span class="sxs-lookup"><span data-stu-id="b1217-113">Scripting skills covered</span></span>

- <span data-ttu-id="b1217-114">Formato de celda</span><span class="sxs-lookup"><span data-stu-id="b1217-114">Cell formatting</span></span>
- <span data-ttu-id="b1217-115">Comprobación de errores</span><span class="sxs-lookup"><span data-stu-id="b1217-115">Error checking</span></span>
- <span data-ttu-id="b1217-116">Expresiones regulares</span><span class="sxs-lookup"><span data-stu-id="b1217-116">Regular expressions</span></span>

## <a name="setup-instructions"></a><span data-ttu-id="b1217-117">Instrucciones de instalación</span><span class="sxs-lookup"><span data-stu-id="b1217-117">Setup instructions</span></span>

1. <span data-ttu-id="b1217-118">Descargue <a href="grade-calculator.xlsx">grade-Calculator. xlsx</a> en su OneDrive.</span><span class="sxs-lookup"><span data-stu-id="b1217-118">Download <a href="grade-calculator.xlsx">grade-calculator.xlsx</a> to your OneDrive.</span></span>

2. <span data-ttu-id="b1217-119">Abra el libro con Excel para la Web.</span><span class="sxs-lookup"><span data-stu-id="b1217-119">Open the workbook with Excel for the web.</span></span>

3. <span data-ttu-id="b1217-120">En la ficha **automatizar** , abra el **Editor de código**.</span><span class="sxs-lookup"><span data-stu-id="b1217-120">Under the **Automate** tab, open the **Code Editor**.</span></span>

4. <span data-ttu-id="b1217-121">En el panel de tareas **Editor de código** , presione **nueva secuencia** de comandos y pegue el siguiente script en el editor.</span><span class="sxs-lookup"><span data-stu-id="b1217-121">In the **Code Editor** task pane, press **New Script** and paste the following script into the editor.</span></span>

    ```TypeScript
    async function main(context: Excel.RequestContext) {
      // Get the number of student record rows.
      let sheet = context.workbook.worksheets.getActiveWorksheet();
      let studentsRange = sheet.getUsedRange().load("values, rowCount");
      await context.sync();
      console.log("Total students: " + (studentsRange.rowCount - 1));

      // Clean up any formatting from previous runs of the script.
      studentsRange.clear(Excel.ClearApplyTo.formats);
      studentsRange.getColumn(4).getCell(0, 0).getRowsBelow(studentsRange.rowCount - 1).clear(Excel.ClearApplyTo.all);
      studentsRange.getColumn(5).getCell(0, 0).getRowsBelow(studentsRange.rowCount - 1).clear(Excel.ClearApplyTo.all);
      await context.sync();

      // Parse the headers for the maximum possible scores for each category.
      // The format is `category (score)`.
      let assignmentsMax = studentsRange.values[0][1].match(/\d+/)[0];
      let midTermMax = studentsRange.values[0][2].match(/\d+/)[0];
      let finalsMax = studentsRange.values[0][3].match(/\d+/)[0];
      console.log("Assignments max score:" + assignmentsMax);
      console.log("Mid-term max score: " + midTermMax);
      console.log("Final max score: " + finalsMax);

      // Look at every student row.
      for (let i = 1; i < studentsRange.values.length; i++) {
        let row = studentsRange.values[i];
        let total = row[1] + row[2] + row[3];
        let valid = true;

        // Look for any records that are too low or too high.
        if (row[1] < 0 || row[1] > assignmentsMax) {
          studentsRange.getCell(i, 1).format.fill.color = "Red";
          valid = false;
        }
        if (row[2] < 0 || row[2] > midTermMax) {
          studentsRange.getCell(i, 2).format.fill.color = "Red";
          valid = false;
        }
        if (row[3] < 0 || row[3] > finalsMax) {
          studentsRange.getCell(i, 3).format.fill.color = "Red";
          valid = false;
        }

        // If the scores are valid, total that student's points and assign them a letter grade.
        if (valid) {
          let grade: string;
          switch (true) {
            case total < 60:
              grade = "E";
              break;
            case total < 70:
              grade = "D";
              break;
            case total < 80:
              grade = "C";
              break;
            case total < 90:
              grade = "B";
              break;
            default:
              grade = "A";
              break;
          }

          studentsRange.getCell(i, 4).values = [[total]];
          studentsRange.getCell(i, 5).values = [[grade]];

          // Highlight excellent students and those in need of attention.
          if (grade === "A") {
            studentsRange.getCell(i, 5).format.fill.color = "Green";
          } else if (grade === "E" || grade === "D") {
            studentsRange.getCell(i, 5).format.fill.color = "Orange";
          }
        }
      }

      studentsRange.getColumn(5).format.horizontalAlignment = "Center";
    }
    ```

5. <span data-ttu-id="b1217-122">Cambie el nombre del script a **Evaluation Calculator** y guárdelo.</span><span class="sxs-lookup"><span data-stu-id="b1217-122">Rename the script to **Grade Calculator** and save it.</span></span>

## <a name="running-the-script"></a><span data-ttu-id="b1217-123">Ejecución del script</span><span class="sxs-lookup"><span data-stu-id="b1217-123">Running the script</span></span>

<span data-ttu-id="b1217-124">Ejecutar el script de la **calculadora de calificación** en la única hoja de cálculo.</span><span class="sxs-lookup"><span data-stu-id="b1217-124">Run the **Grade Calculator** script on the only worksheet.</span></span> <span data-ttu-id="b1217-125">El script totalizará las calificaciones y asignará a cada alumno una carta de calificación.</span><span class="sxs-lookup"><span data-stu-id="b1217-125">The script will total the grades and assign each student a letter grade.</span></span> <span data-ttu-id="b1217-126">Si alguna de las calificaciones tiene más puntos de los que merece la asignación o la prueba, la calificación infractora se marcará como roja y no se calculará el total.</span><span class="sxs-lookup"><span data-stu-id="b1217-126">If any individual grades have more points than the assignment or test is worth, then the offending grade is marked red and the total is not calculated.</span></span>

### <a name="before-running-the-script"></a><span data-ttu-id="b1217-127">Antes de ejecutar el script</span><span class="sxs-lookup"><span data-stu-id="b1217-127">Before running the script</span></span>

![Hoja de cálculo que muestra las filas de los resultados de los alumnos.](../../images/scenario-grade-calculator-before.png)

### <a name="after-running-the-script"></a><span data-ttu-id="b1217-129">Después de ejecutar el script</span><span class="sxs-lookup"><span data-stu-id="b1217-129">After running the script</span></span>

![Una hoja de cálculo que muestra los datos de puntuación del alumno con celdas no válidas en los totales de rojo para las filas de alumnos válidas.](../../images/scenario-grade-calculator-after.png)
