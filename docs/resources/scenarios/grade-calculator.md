---
title: 'Escenario de ejemplo scripts de Office: calculadora de calificaciones'
description: Muestra que determina el porcentaje y las calificaciones de letras de una clase de alumnos.
ms.date: 12/17/2020
localization_priority: Normal
ms.openlocfilehash: b8c45ad405c06a943c75e76391c1160ecb1bd18e
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 04/14/2021
ms.locfileid: "51755031"
---
# <a name="office-scripts-sample-scenario-grade-calculator"></a><span data-ttu-id="c16a8-103">Escenario de ejemplo scripts de Office: calculadora de calificaciones</span><span class="sxs-lookup"><span data-stu-id="c16a8-103">Office Scripts sample scenario: Grade calculator</span></span>

<span data-ttu-id="c16a8-104">En este escenario, eres un instructor que cuenta las calificaciones de fin de curso de cada alumno.</span><span class="sxs-lookup"><span data-stu-id="c16a8-104">In this scenario, you're an instructor tallying every student's end-of-term grades.</span></span> <span data-ttu-id="c16a8-105">Has estado especificando las puntuaciones de sus tareas y pruebas a medida que vas.</span><span class="sxs-lookup"><span data-stu-id="c16a8-105">You've been entering the scores for their assignments and tests as you go.</span></span> <span data-ttu-id="c16a8-106">Ahora, es el momento de determinar el destino de los alumnos.</span><span class="sxs-lookup"><span data-stu-id="c16a8-106">Now, it is time to determine the students' fates.</span></span>

<span data-ttu-id="c16a8-107">Desarrollará un script que totale las calificaciones de cada categoría de puntos.</span><span class="sxs-lookup"><span data-stu-id="c16a8-107">You'll develop a script that totals the grades for each point category.</span></span> <span data-ttu-id="c16a8-108">A continuación, asignará una nota de letra a cada alumno en función del total.</span><span class="sxs-lookup"><span data-stu-id="c16a8-108">It will then assign a letter grade to each student based on the total.</span></span> <span data-ttu-id="c16a8-109">Para garantizar la precisión, agregarás un par de comprobaciones para ver si las puntuaciones individuales son demasiado bajas o altas.</span><span class="sxs-lookup"><span data-stu-id="c16a8-109">To help ensure accuracy, you'll add a couple checks to see if any individual scores are too low or high.</span></span> <span data-ttu-id="c16a8-110">Si la puntuación de un alumno es menor que cero o más que el valor de punto posible, el script marcará la celda con un relleno rojo y no total de los puntos del alumno.</span><span class="sxs-lookup"><span data-stu-id="c16a8-110">If a student's score is less than zero or more than the possible point value, the script will flag the cell with a red fill and not total that student's points.</span></span> <span data-ttu-id="c16a8-111">Esta será una indicación clara de los registros que necesita comprobar.</span><span class="sxs-lookup"><span data-stu-id="c16a8-111">This will be a clear indication of which records you need to double-check.</span></span> <span data-ttu-id="c16a8-112">También agregará algún formato básico a las calificaciones para que pueda ver rápidamente la parte superior e inferior de la clase.</span><span class="sxs-lookup"><span data-stu-id="c16a8-112">You'll also add some basic formatting to the grades so you can quickly view the top and bottom of the class.</span></span>

## <a name="scripting-skills-covered"></a><span data-ttu-id="c16a8-113">Habilidades de scripting cubiertas</span><span class="sxs-lookup"><span data-stu-id="c16a8-113">Scripting skills covered</span></span>

- <span data-ttu-id="c16a8-114">Formato de celda</span><span class="sxs-lookup"><span data-stu-id="c16a8-114">Cell formatting</span></span>
- <span data-ttu-id="c16a8-115">Comprobación de errores</span><span class="sxs-lookup"><span data-stu-id="c16a8-115">Error checking</span></span>
- <span data-ttu-id="c16a8-116">Expresiones regulares</span><span class="sxs-lookup"><span data-stu-id="c16a8-116">Regular expressions</span></span>
- <span data-ttu-id="c16a8-117">Formato condicional</span><span class="sxs-lookup"><span data-stu-id="c16a8-117">Conditional formatting</span></span>

## <a name="setup-instructions"></a><span data-ttu-id="c16a8-118">Instrucciones de configuración</span><span class="sxs-lookup"><span data-stu-id="c16a8-118">Setup instructions</span></span>

1. <span data-ttu-id="c16a8-119">Descarga <a href="grade-calculator.xlsx">grade-calculator.xlsx</a> a tu OneDrive.</span><span class="sxs-lookup"><span data-stu-id="c16a8-119">Download <a href="grade-calculator.xlsx">grade-calculator.xlsx</a> to your OneDrive.</span></span>

2. <span data-ttu-id="c16a8-120">Abra el libro con Excel para la web.</span><span class="sxs-lookup"><span data-stu-id="c16a8-120">Open the workbook with Excel for the web.</span></span>

3. <span data-ttu-id="c16a8-121">En la **pestaña Automatizar,** abra **Todos los scripts**.</span><span class="sxs-lookup"><span data-stu-id="c16a8-121">Under the **Automate** tab, open **All Scripts**.</span></span>

4. <span data-ttu-id="c16a8-122">En el **panel de tareas Editor** de código, presione Nuevo **script** y pegue el siguiente script en el editor.</span><span class="sxs-lookup"><span data-stu-id="c16a8-122">In the **Code Editor** task pane, press **New Script** and paste the following script into the editor.</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
      // Get the worksheet and validate the data.
      let studentsRange = workbook.getActiveWorksheet().getUsedRange();
      if (studentsRange.getColumnCount() !== 6) {
        throw new Error(`The required columns are not present. Expected column headers: "Student ID | Assignment score | Mid-term | Final | Total | Grade"`);
      }

      let studentData = studentsRange.getValues();

      // Clear the total and grade columns.
      studentsRange.getColumn(4).getCell(1, 0).getAbsoluteResizedRange(studentData.length - 1, 2).clear();

      // Clear all conditional formatting.
      workbook.getActiveWorksheet().getUsedRange().clearAllConditionalFormats();

      // Use regular expressions to read the max score from the assignment, mid-term, and final scores columns.
      let maxScores: string[] = [];
      const assignmentMaxMatches = (studentData[0][1] as string).match(/\d+/);
      const midtermMaxMatches = (studentData[0][2] as string).match(/\d+/);
      const finalMaxMatches = (studentData[0][3] as string).match(/\d+/);

      // Check the matches happened before proceeding.
      if (!(assignmentMaxMatches && midtermMaxMatches && finalMaxMatches)) {
        throw new Error(`The scores are not present in the column headers. Expected format: "Assignments (n)|Mid-term (n)|Final (n)"`);
      }

      // Use the first (and only) match from the regular expressions as the max scores.
      maxScores = [assignmentMaxMatches[0], midtermMaxMatches[0], finalMaxMatches[0]];

      // Set conditional formatting for each of the assignment, mid-term, and final scores columns.
      maxScores.forEach((score, i) => {
        let range = studentsRange.getColumn(i + 1).getCell(0, 0).getRowsBelow(studentData.length - 1);
        setCellValueConditionalFormatting(
          score,
          range,
          "#9C0006",
          "#FFC7CE",
          ExcelScript.ConditionalCellValueOperator.greaterThan
        )
      });

      // Store the current range information to avoid calling the workbook in the loop.
      let studentsRangeFormulas = studentsRange.getColumn(4).getFormulasR1C1();
      let studentsRangeValues = studentsRange.getColumn(5).getValues();

      /* Iterate over each of the student rows and compute the total score and letter grade.
      * Note that iterator starts at index 1 to skip first (header) row.
      */
      for (let i = 1; i < studentData.length; i++) {
        // If any of the scores are invalid, skip processing it.
        if (studentData[i][1] > maxScores[0] ||
          studentData[i][2] > maxScores[1] ||
          studentData[i][3] > maxScores[2]) {
          continue;
        }
        const total = (studentData[i][1] as number) + (studentData[i][2] as number) + (studentData[i][3] as number);
        let grade: string;
        switch (true) {
          case total < 60:
            grade = "F";
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
    
        // Set total score formula.
        studentsRangeFormulas[i][0] = '=RC[-2]+RC[-1]';
        // Set grade cell.
        studentsRangeValues[i][0] = grade;
      }

      // Set the formulas and values outside the loop.
      studentsRange.getColumn(4).setFormulasR1C1(studentsRangeFormulas);
      studentsRange.getColumn(5).setValues(studentsRangeValues);

      // Put a conditional formatting on the grade column.
      let totalRange = studentsRange.getColumn(5).getCell(0, 0).getRowsBelow(studentData.length - 1);
      setCellValueConditionalFormatting(
        "A",
        totalRange,
        "#001600",
        "#C6EFCE",
        ExcelScript.ConditionalCellValueOperator.equalTo
      );
      ["D", "F"].forEach((grade) => {
        setCellValueConditionalFormatting(
          grade,
          totalRange,
          "#443300",
          "#FFEE22",
          ExcelScript.ConditionalCellValueOperator.equalTo
        );
      })
      // Center the grade column.
      studentsRange.getColumn(5).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
    }

    /**
     * Helper function to apply conditional formatting.
     * @param value Cell value to use in conditional formatting formula1.
     * @param range Target range.
     * @param fontColor Font color to use.
     * @param fillColor Fill color to use.
     * @param operator Operator to use in conditional formatting.
     */
    function setCellValueConditionalFormatting(
      value: string,
      range: ExcelScript.Range,
      fontColor: string,
      fillColor: string,
      operator: ExcelScript.ConditionalCellValueOperator) {
      // Determine the formula1 based on the type of value parameter.
      let formula1: string;
      if (isNaN(Number(value))) {
        // For cell value equalTo rule, use this format: formula1: "=\"A\"",
        formula1 = `=\"${value}\"`;
      } else {
        // For number input (greater-than or less-than rules), just append '='.
        formula1 = `=${value}`;
      }

      // Apply conditional formatting.
      let conditionalFormatting: ExcelScript.ConditionalFormat;
      conditionalFormatting = range.addConditionalFormat(ExcelScript.ConditionalFormatType.cellValue);
      conditionalFormatting.getCellValue().getFormat().getFont().setColor(fontColor);
      conditionalFormatting.getCellValue().getFormat().getFill().setColor(fillColor);
      conditionalFormatting.getCellValue().setRule({ formula1, operator });
    }
    ```

5. <span data-ttu-id="c16a8-123">Cambie el nombre del script a **Calculadora de calificaciones** y guárdelo.</span><span class="sxs-lookup"><span data-stu-id="c16a8-123">Rename the script to **Grade Calculator** and save it.</span></span>

## <a name="running-the-script"></a><span data-ttu-id="c16a8-124">Ejecución del script</span><span class="sxs-lookup"><span data-stu-id="c16a8-124">Running the script</span></span>

<span data-ttu-id="c16a8-125">Ejecute el script **calculadora de** calificaciones en la única hoja de cálculo.</span><span class="sxs-lookup"><span data-stu-id="c16a8-125">Run the **Grade Calculator** script on the only worksheet.</span></span> <span data-ttu-id="c16a8-126">El script totalirá las calificaciones y asignará a cada alumno una nota de letra.</span><span class="sxs-lookup"><span data-stu-id="c16a8-126">The script will total the grades and assign each student a letter grade.</span></span> <span data-ttu-id="c16a8-127">Si las calificaciones individuales tienen más puntos que la asignación o la prueba, la calificación infractora se marca en rojo y el total no se calcula.</span><span class="sxs-lookup"><span data-stu-id="c16a8-127">If any individual grades have more points than the assignment or test is worth, then the offending grade is marked red and the total is not calculated.</span></span> <span data-ttu-id="c16a8-128">Además, las calificaciones "A" se resaltan en verde, mientras que las calificaciones "D" y "F" se resaltan en amarillo.</span><span class="sxs-lookup"><span data-stu-id="c16a8-128">Also, any 'A' grades are highlighted in green, while 'D' and 'F' grades are highlighted in yellow.</span></span>

### <a name="before-running-the-script"></a><span data-ttu-id="c16a8-129">Antes de ejecutar el script</span><span class="sxs-lookup"><span data-stu-id="c16a8-129">Before running the script</span></span>

:::image type="content" source="../../images/scenario-grade-calculator-before.png" alt-text="Hoja de cálculo que muestra filas de puntuaciones para los alumnos.":::

### <a name="after-running-the-script"></a><span data-ttu-id="c16a8-131">Después de ejecutar el script</span><span class="sxs-lookup"><span data-stu-id="c16a8-131">After running the script</span></span>

:::image type="content" source="../../images/scenario-grade-calculator-after.png" alt-text="Hoja de cálculo que muestra los datos de puntuación del alumno con celdas no válidas en totales rojos para filas de alumnos válidas.":::
