---
title: 'Escenario de ejemplo de Scripts de Office: calculadora de calificaciones'
description: Muestra que determina el porcentaje y las calificaciones de letra para una clase de alumnos.
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: 7dda3ebe84dc3edd10998cbe2c4cd0806da11411
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572531"
---
# <a name="office-scripts-sample-scenario-grade-calculator"></a>Escenario de ejemplo de Scripts de Office: calculadora de calificaciones

En este escenario, es un instructor que analiza las calificaciones de fin de curso de todos los alumnos. Ha estado escribiendo las puntuaciones de sus asignaciones y pruebas a medida que avanza. Ahora, es el momento de determinar el destino de los estudiantes.

Desarrollará un script que totale las calificaciones de cada categoría de punto. A continuación, asignará una calificación de letra a cada alumno en función del total. Para ayudar a garantizar la precisión, agregará un par de comprobaciones para ver si alguna puntuación individual es demasiado baja o alta. Si la puntuación de un alumno es menor que cero o más que el valor de punto posible, el script marcará la celda con un relleno rojo y no sumará los puntos de ese alumno. Esto será una indicación clara de qué registros debe comprobar. También agregará algún formato básico a las calificaciones para que pueda ver rápidamente la parte superior e inferior de la clase.

## <a name="scripting-skills-covered"></a>Aptitudes de scripting cubiertas

- Formato de celda
- Comprobación de errores
- Expresiones regulares
- Formato condicional

## <a name="setup-instructions"></a>Instrucciones de instalación

1. Descargue [grade-calculator.xlsx](grade-calculator.xlsx) en Su OneDrive.

1. Abra el libro con Excel para la Web.

1. En la pestaña **Automatizar** , seleccione **Nuevo script** y pegue el siguiente script en el editor.

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

1. Cambie el nombre del script a **Calculadora de calificaciones** y guárdelo.

## <a name="running-the-script"></a>Ejecución del script

Ejecute el script **calculadora de calificaciones** en la única hoja de cálculo. El script sumará las calificaciones y asignará a cada alumno una calificación de letra. Si alguna calificación individual tiene más puntos que la asignación o prueba vale la pena, la calificación infractora se marca como roja y el total no se calcula. Además, los grados "A" se resaltan en verde, mientras que los grados "D" y "F" están resaltados en amarillo.

### <a name="before-running-the-script"></a>Antes de ejecutar el script

:::image type="content" source="../../images/scenario-grade-calculator-before.png" alt-text="Hoja de cálculo que muestra filas de puntuaciones para los alumnos.":::

### <a name="after-running-the-script"></a>Después de ejecutar el script

:::image type="content" source="../../images/scenario-grade-calculator-after.png" alt-text="Hoja de cálculo que muestra los datos de puntuación de alumnos con celdas no válidas en totales rojos para filas de alumnos válidas.":::
