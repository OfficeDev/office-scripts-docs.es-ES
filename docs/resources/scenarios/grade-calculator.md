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
# <a name="office-scripts-sample-scenario-grade-calculator"></a>Escenario de ejemplo de scripts de Office: Calculadora de calificaciones

En este escenario, usted es un instructor que rellene las calificaciones de fin de período de cada estudiante. Ha estado especificando los resultados de sus asignaciones y pruebas a medida que avanza. Ahora, es el momento de determinar los Fates de los alumnos.

Desarrollará un script que totaliza las calificaciones para cada categoría de punto. A continuación, asignará una letra de calificación a cada estudiante en función del total. Para ayudar a garantizar la precisión, agregará un par de comprobaciones para ver si alguna puntuación individual es demasiado baja o alta. Si la puntuación de un estudiante es menor que cero o mayor que el valor de punto posible, el script marcará la celda con un relleno rojo y no hará un total de los puntos del estudiante. Esto será una indicación clara de los registros que debe comprobar dos veces. También agregará formato básico a las calificaciones para que pueda ver rápidamente la parte superior e inferior de la clase.

## <a name="scripting-skills-covered"></a>Habilidades de scripting cubiertas

- Formato de celda
- Comprobación de errores
- Expresiones regulares

## <a name="setup-instructions"></a>Instrucciones de instalación

1. Descargue <a href="grade-calculator.xlsx">grade-Calculator. xlsx</a> en su OneDrive.

2. Abra el libro con Excel para la Web.

3. En la ficha **automatizar** , abra el **Editor de código**.

4. En el panel de tareas **Editor de código** , presione **nueva secuencia** de comandos y pegue el siguiente script en el editor.

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

5. Cambie el nombre del script a **Evaluation Calculator** y guárdelo.

## <a name="running-the-script"></a>Ejecución del script

Ejecutar el script de la **calculadora de calificación** en la única hoja de cálculo. El script totalizará las calificaciones y asignará a cada alumno una carta de calificación. Si alguna de las calificaciones tiene más puntos de los que merece la asignación o la prueba, la calificación infractora se marcará como roja y no se calculará el total.

### <a name="before-running-the-script"></a>Antes de ejecutar el script

![Hoja de cálculo que muestra las filas de los resultados de los alumnos.](../../images/scenario-grade-calculator-before.png)

### <a name="after-running-the-script"></a>Después de ejecutar el script

![Una hoja de cálculo que muestra los datos de puntuación del alumno con celdas no válidas en los totales de rojo para las filas de alumnos válidas.](../../images/scenario-grade-calculator-after.png)
