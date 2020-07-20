---
title: Mejorar el rendimiento de los scripts de Office
description: Cree scripts más rápidos mediante la comprensión de la comunicación entre el libro de Excel y el script.
ms.date: 06/15/2020
localization_priority: Normal
ms.openlocfilehash: 4d5b7c70f14e3fc598b95a6226e3ef8caf89f651
ms.sourcegitcommit: aec3c971c6640429f89b6bb99d2c95ea06725599
ms.translationtype: Auto
ms.contentlocale: es-ES
ms.lasthandoff: 06/25/2020
ms.locfileid: "44878901"
---
# <a name="improve-the-performance-of-your-office-scripts"></a>Mejorar el rendimiento de los scripts de Office

El propósito de las secuencias de comandos de Office es automatizar la serie de tareas que se suele realizar para ahorrar tiempo. Un script lento puede sentir como que no acelera el flujo de trabajo. La mayoría de las veces, el script será perfectamente correcto y se ejecutará como se esperaba. Sin embargo, hay algunos escenarios evitables que pueden afectar al rendimiento.

La causa más común de un script lento es una comunicación excesiva con el libro. El script se ejecuta en el equipo local, mientras que el libro está en la nube. En determinados momentos, el script sincroniza sus datos locales con el del libro. Esto significa que las operaciones de escritura (como `workbook.addWorksheet()` ) solo se aplican al libro cuando se produce esta sincronización en segundo plano. Del mismo modo, las operaciones de lectura (como `myRange.getValues()` ) solo obtienen datos del libro para el script en esos momentos. En cualquier caso, el script recopila información antes de que actúe en los datos. Por ejemplo, el siguiente código registra con precisión el número de filas en el rango usado.

```TypeScript
let usedRange = workbook.getActiveWorksheet().getUsedRange();
let rowCount = usedRange.getRowCount();
// The script will read the range and row count from
// the workbook before logging the information.
console.log(rowCount);
```

Las API de scripts de Office garantizan que los datos del libro o script sean precisos y actualizados cuando sea necesario. No tiene que preocuparse por estas sincronizaciones para que el script se ejecute correctamente. Sin embargo, un conocimiento de esta comunicación de script a nube puede ayudarle a evitar llamadas de red no necesarias.

## <a name="performance-optimizations"></a>Optimizaciones de rendimiento

Puede aplicar técnicas sencillas para ayudar a reducir la comunicación a la nube. Los siguientes modelos ayudan a acelerar los scripts.

- Leer los datos del libro una vez en lugar de repetidamente en un bucle.
- Quite las instrucciones innecesarias `console.log` .
- Evite usar bloques try/catch.

### <a name="read-workbook-data-outside-of-a-loop"></a>Leer datos de un libro fuera de un bucle

Cualquier método que obtenga datos del libro puede desencadenar una llamada de red. En lugar de realizar repetidamente la misma llamada, debe guardar los datos de forma local siempre que sea posible. Esto es especialmente cierto al tratar con bucles.

Considere un script para obtener el número de números negativos en el rango usado de una hoja de cálculo. El script debe recorrer en iteración todas las celdas del rango usado. Para ello, necesita el intervalo, el número de filas y el número de columnas. Debe almacenarlos como variables locales antes de iniciar el bucle. De lo contrario, cada iteración del bucle forzará una devolución al libro.

```TypeScript
/**
 * This script provides the count of negative numbers that are present
 * in the used range of the current worksheet.
 */
function main(workbook: ExcelScript.Workbook) {
  // Get the working range.
  let usedRange = workbook.getActiveWorksheet().getUsedRange();

  // Save the values locally to avoid repeatedly asking the workbook.
  let usedRangeValues = usedRange.getValues();

  // Start the negative number counter.
  let negativeCount = 0;

  // Iterate over the entire range looking for negative numbers.
  for (let i = 0; i < usedRangeValues.length; i++) {
    for (let j = 0; j < usedRangeValues[i].length; j++) {
      if (usedRangeValues[i][j] < 0) {
        negativeCount++;
      }
    }
  }

  // Log the negative number count to the console.
  console.log(negativeCount);
}
```

> [!NOTE]
> Como experimento, intente reemplazar `usedRangeValues` el bucle por `usedRange.getValues()` . Es posible que note que el script tarda mucho más tiempo en ejecutarse cuando se trata con rangos grandes.

### <a name="remove-unnecessary-consolelog-statements"></a>Quitar instrucciones innecesarias `console.log`

El registro de consola es una herramienta vital para [la depuración de scripts](../testing/troubleshooting.md). Sin embargo, sí obliga a que la secuencia de comandos se sincronice con el libro para asegurarse de que la información registrada está actualizada. Considere la posibilidad de quitar instrucciones de registro innecesarias (como las que se usan para las pruebas) antes de compartir el script. Esto normalmente no provocará un problema de rendimiento perceptible, a menos que la `console.log()` instrucción esté en un bucle.

### <a name="avoid-using-trycatch-blocks"></a>Evitar el uso de bloques try/catch

No se recomienda usar [ `try` / `catch` bloques](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) como parte del flujo de control esperado de un script. La mayoría de los errores se pueden evitar comprobando los objetos devueltos del libro. Por ejemplo, el script siguiente comprueba que la tabla devuelta por el libro existe antes de intentar agregar una fila.

```TypeScript
/**
 * This script adds a row to "MyTable", if that table is present.
 */
function main(workbook: ExcelScript.Workbook) {
  let table = workbook.getTable("MyTable");

  // Check if the table exists.
  if (table) {
    // Add the row.
    table.addRow(-1, ["2012", "Yes", "Maybe"]);
  } else {
    // Report the missing table.
    console.log("MyTable not found.");
  }
}
```

## <a name="case-by-case-help"></a>Ayuda caso por caso

A medida que la plataforma de scripts de Office se expande para trabajar con la automatización de la [potencia](https://flow.microsoft.com/), [tarjetas adaptables](https://docs.microsoft.com/adaptive-cards)y otras características de productos cruzados, los detalles de la comunicación del libro y de la secuencia de comandos se vuelven más complejos. Si necesita ayuda para que la secuencia de comandos se ejecute más rápido, póngase en contacto con el [desbordamiento de pila](https://stackoverflow.com/questions/tagged/office-scripts). Asegúrese de etiquetar su pregunta con "Office-scripts" para que los expertos puedan encontrarla y ayudarle.

## <a name="see-also"></a>Vea también

- [Conceptos básicos de los scripts de Office en Excel en la Web](scripting-fundamentals.md)
- [Documentos web de MDN: bucles e iteración](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Loops_and_iteration)
