---
title: Mejore el rendimiento de sus scripts de Office
description: Cree scripts más rápidos comprendiéndose la comunicación entre el libro de Excel y el script.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 512e2108cb81cf9ac8ae98980951d5d01b3d2de9
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 05/19/2021
ms.locfileid: "52544994"
---
# <a name="improve-the-performance-of-your-office-scripts"></a>Mejore el rendimiento de sus scripts de Office

El propósito de Office Scripts es automatizar series de tareas que se realizan comúnmente para ahorrarle tiempo. Un script lento puede sentir que no acelera el flujo de trabajo. La mayoría de las veces, su script estará perfectamente bien y se ejecutará como se esperaba. Sin embargo, hay algunos escenarios evitables que pueden afectar al rendimiento.

La razón más común de un script lento es la comunicación excesiva con el libro de trabajo. El script se ejecuta en el equipo local, mientras que el libro existe en la nube. En determinados momentos, el script sincroniza sus datos locales con los del libro. Esto significa que las operaciones de escritura (por `workbook.addWorksheet()` ejemplo) solo se aplican al libro cuando se produce esta sincronización entre bastidores. Del mismo modo, las operaciones de lectura (por `myRange.getValues()` ejemplo) solo obtienen datos del libro para el script en esos momentos. En cualquier caso, el script obtiene información antes de que actúe sobre los datos. Por ejemplo, el código siguiente registrará con precisión el número de filas en el intervalo utilizado.

```TypeScript
let usedRange = workbook.getActiveWorksheet().getUsedRange();
let rowCount = usedRange.getRowCount();
// The script will read the range and row count from
// the workbook before logging the information.
console.log(rowCount);
```

Office Las API de scripts garantizan que los datos del libro o script sean precisos y actualizados cuando sea necesario. No es necesario preocuparse por estas sincronizaciones para que el script se ejecute correctamente. Sin embargo, un conocimiento de esta comunicación de script a nube puede ayudarle a evitar llamadas de red innecesarias.

## <a name="performance-optimizations"></a>Optimizaciones de rendimiento

Puede aplicar técnicas sencillas para ayudar a reducir la comunicación a la nube. Los siguientes patrones ayudan a acelerar los scripts.

- Lea los datos del libro una vez en lugar de repetidamente en un bucle.
- Quitar `console.log` instrucciones innecesarias.
- Evite usar bloques try/catch.

### <a name="read-workbook-data-outside-of-a-loop"></a>Leer datos de libros de trabajo fuera de un bucle

Cualquier método que obtenga datos del libro puede desencadenar una llamada de red. En lugar de realizar repetidamente la misma llamada, debe guardar los datos localmente siempre que sea posible. Esto es especialmente cierto cuando se trata de bucles.

Considere la posibilidad de un script para obtener el recuento de números negativos en el rango utilizado de una hoja de cálculo. El script necesita iterar en todas las celdas del intervalo utilizado. Para ello, necesita el intervalo, el número de filas y el número de columnas. Debe almacenarlas como variables locales antes de iniciar el bucle. De lo contrario, cada iteración del bucle forzará una devolución al libro.

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
> Como experimento, intente reemplazar `usedRangeValues` en el bucle con `usedRange.getValues()` . Es posible que observe que el script tarda considerablemente más en ejecutarse cuando se trata de intervalos grandes.

### <a name="avoid-using-trycatch-blocks-in-or-surrounding-loops"></a>Evite usar `try...catch` bloques en bucles o circundantes

No se recomienda usar [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) instrucciones ni en bucles ni en bucles circundantes. Esto es por la misma razón por la que debe evitar leer datos en un bucle: cada iteración obliga al script a sincronizarse con el libro para asegurarse de que no se ha producido ningún error. La mayoría de los errores se pueden evitar comprobando los objetos devueltos desde el libro. Por ejemplo, el siguiente script comprueba que la tabla devuelta por el libro existe antes de intentar agregar una fila.

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

### <a name="remove-unnecessary-consolelog-statements"></a>Eliminar `console.log` declaraciones innecesarias

El registro de consola es una herramienta vital para [depurar los scripts.](../testing/troubleshooting.md) Sin embargo, obliga al script a sincronizarse con el libro para asegurarse de que la información registrada está actualizada. Considere la posibilidad de quitar instrucciones de registro innecesarias (como las utilizadas para las pruebas) antes de compartir el script. Esto normalmente no causará un problema de rendimiento notable, a menos que la `console.log()` instrucción esté en un bucle.

## <a name="case-by-case-help"></a>Ayuda caso por caso

A medida que la plataforma Office Scripts se expande para trabajar con [Power Automate,](https://flow.microsoft.com/) [tarjetas adaptables](/adaptive-cards)y otras características entre productos, los detalles de la comunicación script-workbook se vuelven más intrincados. Si necesita ayuda para que el script se ejecute más rápido, póngase en contacto con [Microsoft Q&A](/answers/topics/office-scripts-dev.html). Asegúrese de etiquetar su pregunta con "office-scripts-dev" para que los expertos puedan encontrarla y ayudar.

## <a name="see-also"></a>Vea también

- [Conceptos básicos de los Scripts de Office en Excel en la web](scripting-fundamentals.md)
- [Documentos web mdn: bucles e iteración](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Loops_and_iteration)
