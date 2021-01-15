---
title: Mejorar el rendimiento de los scripts de Office
description: Cree scripts más rápidos mediante la comprensión de la comunicación entre el libro de Excel y el script.
ms.date: 06/15/2020
localization_priority: Normal
ms.openlocfilehash: ce50a6fd7ad02ddcd2dd304be8b4dd8fa3d0acf3
ms.sourcegitcommit: 7580dcb8f2f97974c2a9cce25ea30d6526730e28
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 01/14/2021
ms.locfileid: "49867873"
---
# <a name="improve-the-performance-of-your-office-scripts"></a>Mejorar el rendimiento de los scripts de Office

El propósito de los scripts de Office es automatizar una serie de tareas que se realizan habitualmente para ahorrar tiempo. Un script lento puede tener la sensación de que no acelera el flujo de trabajo. La mayoría de las veces, el script estará perfectamente bien y se ejecutará según lo esperado. Sin embargo, hay algunos escenarios evitables que pueden afectar al rendimiento.

La razón más común para un script lento es una comunicación excesiva con el libro. El script se ejecuta en el equipo local, mientras que el libro existe en la nube. En determinados momentos, el script sincroniza sus datos locales con los del libro. Esto significa que las operaciones de escritura (como ) solo se aplican al libro cuando se produce esta sincronización en segundo `workbook.addWorksheet()` plano. Del mismo modo, las operaciones de lectura (como ) solo obtienen datos del `myRange.getValues()` libro para el script en esos momentos. En cualquier caso, el script recupera información antes de que actúe en los datos. Por ejemplo, el siguiente código registrará con precisión el número de filas del rango usado.

```TypeScript
let usedRange = workbook.getActiveWorksheet().getUsedRange();
let rowCount = usedRange.getRowCount();
// The script will read the range and row count from
// the workbook before logging the information.
console.log(rowCount);
```

Las API de scripts de Office garantizan que los datos del libro o script sean precisos y actualizados cuando sea necesario. No es necesario preocuparse por estas sincronizaciones para que el script se ejecute correctamente. Sin embargo, un conocimiento de esta comunicación de script a nube puede ayudarle a evitar llamadas de red innecesarios.

## <a name="performance-optimizations"></a>Optimizaciones de rendimiento

Puedes aplicar técnicas sencillas para ayudar a reducir la comunicación a la nube. Los siguientes patrones ayudan a acelerar los scripts.

- Leer los datos del libro una vez en lugar de repetirse en un bucle.
- Quite instrucciones `console.log` innecesarias.
- Evita usar bloques try/catch.

### <a name="read-workbook-data-outside-of-a-loop"></a>Leer datos del libro fuera de un bucle

Cualquier método que obtiene datos del libro puede desencadenar una llamada de red. En lugar de realizar repetidamente la misma llamada, debe guardar los datos localmente siempre que sea posible. Esto es especialmente cierto cuando se trata de bucles.

Considere un script para obtener el recuento de números negativos en el rango usado de una hoja de cálculo. El script debe iterar en todas las celdas del rango usado. Para ello, necesita el rango, el número de filas y el número de columnas. Debe almacenar esas variables como variables locales antes de iniciar el bucle. De lo contrario, cada iteración del bucle forzará un retorno al libro.

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
> Como experimento, pruebe a reemplazar `usedRangeValues` en el bucle con `usedRange.getValues()` . Es posible que observe que el script tarda bastante más en ejecutarse cuando se trata de rangos grandes.

### <a name="remove-unnecessary-consolelog-statements"></a>Quitar instrucciones `console.log` innecesarias

El registro de consola es una herramienta fundamental [para depurar los scripts.](../testing/troubleshooting.md) Sin embargo, obliga al script a sincronizarse con el libro para asegurarse de que la información registrada está actualizada. Considere la posibilidad de quitar instrucciones de registro innecesarias (como las que se usan para las pruebas) antes de compartir el script. Esto normalmente no causará un problema de rendimiento notable, a menos que la `console.log()` instrucción esté en un bucle.

### <a name="avoid-using-trycatch-blocks"></a>Evitar el uso de bloques try/catch

No se recomienda usar [ `try` / `catch` bloques como](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) parte del flujo de control esperado de un script. La mayoría de los errores se pueden evitar comprobando los objetos devueltos desde el libro. Por ejemplo, el siguiente script comprueba que la tabla devuelta por el libro existe antes de intentar agregar una fila.

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

A medida que la plataforma de scripts [](/adaptive-cards)de Office se expande para funcionar con [Power Automate,](https://flow.microsoft.com/)tarjetas adaptables y otras características entre productos, los detalles de la comunicación entre scripts y libros se vuelven más intrincados. Si necesita ayuda para que el script se ejecute más rápido, póngase en contacto con Stack [Overflow.](https://stackoverflow.com/questions/tagged/office-scripts) Asegúrese de etiquetar su pregunta con "scripts de office" para que los expertos puedan encontrarla y ayudar.

## <a name="see-also"></a>Vea también

- [Conceptos básicos de los Scripts de Office en Excel en la web](scripting-fundamentals.md)
- [Documentos web de MDN: bucles e iteración](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Loops_and_iteration)