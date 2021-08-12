---
title: Mejorar el rendimiento de los scripts Office scripts
description: Cree scripts más rápidos al comprender la comunicación entre el Excel y el script.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: e71fe3d0b25eab99d075e54a63600861d285756d98a6d6dd8490c3f4e73fab33
ms.sourcegitcommit: 75f7ed8c2d23a104acc293f8ce29ea580b4fcdc5
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 08/11/2021
ms.locfileid: "57847019"
---
# <a name="improve-the-performance-of-your-office-scripts"></a>Mejorar el rendimiento de los scripts Office scripts

El propósito de Office scripts es automatizar series de tareas que se realizan habitualmente para ahorrar tiempo. Un script lento puede sentir que no acelera el flujo de trabajo. La mayoría de las veces, el script estará perfectamente bien y se ejecutará según lo esperado. Sin embargo, hay algunos escenarios evitables que pueden afectar al rendimiento.

La razón más común para un script lento es una comunicación excesiva con el libro. El script se ejecuta en el equipo local, mientras que el libro existe en la nube. En ciertos momentos, el script sincroniza sus datos locales con los del libro. Esto significa que cualquier operación de escritura (como ) solo se aplica al libro cuando se produce esta sincronización entre `workbook.addWorksheet()` bastidores. Del mismo modo, las operaciones de lectura (como ) solo obtienen datos `myRange.getValues()` del libro para el script en esos momentos. En cualquier caso, el script captura información antes de que actúe en los datos. Por ejemplo, el siguiente código registrará con precisión el número de filas en el intervalo usado.

```TypeScript
let usedRange = workbook.getActiveWorksheet().getUsedRange();
let rowCount = usedRange.getRowCount();
// The script will read the range and row count from
// the workbook before logging the information.
console.log(rowCount);
```

Office Las API de scripts garantizan que los datos del libro o script sean precisos y actualizados cuando sea necesario. No es necesario preocuparse por estas sincronizaciones para que el script se ejecute correctamente. Sin embargo, un conocimiento de esta comunicación de script a nube puede ayudarle a evitar llamadas de red innecesarios.

## <a name="performance-optimizations"></a>Optimizaciones de rendimiento

Puede aplicar técnicas sencillas para ayudar a reducir la comunicación a la nube. Los siguientes patrones ayudan a acelerar los scripts.

- Lea los datos del libro una vez en lugar de repetirlo en un bucle.
- Quite instrucciones `console.log` innecesarias.
- Evite usar bloques try/catch.

### <a name="read-workbook-data-outside-of-a-loop"></a>Leer datos del libro fuera de un bucle

Cualquier método que obtiene datos del libro puede desencadenar una llamada de red. En lugar de realizar la misma llamada repetidamente, debe guardar los datos localmente siempre que sea posible. Esto es especialmente cierto cuando se trata de bucles.

Considere un script para obtener el recuento de números negativos en el rango usado de una hoja de cálculo. El script debe iterar en todas las celdas del intervalo usado. Para ello, necesita el intervalo, el número de filas y el número de columnas. Debe almacenar esas variables como variables locales antes de iniciar el bucle. De lo contrario, cada iteración del bucle forzará un retorno al libro.

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
> Como experimento, intente reemplazar `usedRangeValues` en el bucle por `usedRange.getValues()` . Es posible que observe que el script tarda considerablemente más tiempo en ejecutarse cuando se trata de intervalos grandes.

### <a name="avoid-using-trycatch-blocks-in-or-surrounding-loops"></a>Evitar el uso `try...catch` de bloques en bucles o entornos

No se recomienda usar instrucciones en [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) bucles o bucles circundantes. Este es el mismo motivo por el que debe evitar leer datos en un bucle: cada iteración fuerza al script a sincronizarse con el libro para asegurarse de que no se ha producido ningún error. La mayoría de los errores se pueden evitar comprobando los objetos devueltos desde el libro. Por ejemplo, el siguiente script comprueba que la tabla devuelta por el libro existe antes de intentar agregar una fila.

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

### <a name="remove-unnecessary-consolelog-statements"></a>Quitar instrucciones `console.log` innecesarias

El registro de consola es una herramienta vital [para depurar los scripts.](../testing/troubleshooting.md) Sin embargo, obliga al script a sincronizarse con el libro para asegurarse de que la información registrada está actualizada. Considere la posibilidad de quitar instrucciones de registro innecesarias (como las que se usan para las pruebas) antes de compartir el script. Por lo general, esto no causará un problema de rendimiento notable, a menos que la `console.log()` instrucción esté en un bucle.

## <a name="case-by-case-help"></a>Ayuda caso por caso

A medida que la plataforma de scripts de Office [](/adaptive-cards)se expande para funcionar con [Power Automate,](https://flow.microsoft.com/)tarjetas adaptables y otras características entre productos, los detalles de la comunicación entre scripts y libros se vuelven más complejos. Si necesita ayuda para que el script se ejecute más rápido, póngase en contacto con [Microsoft Q&A](/answers/topics/office-scripts-excel-dev.html). Asegúrese de etiquetar su pregunta con "office-scripts-dev" para que los expertos puedan encontrarlo y ayudarle.

## <a name="see-also"></a>Vea también

- [Conceptos básicos de los Scripts de Office en Excel en la web](scripting-fundamentals.md)
- [Documentos web de MDN: bucles e iteración](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Loops_and_iteration)
