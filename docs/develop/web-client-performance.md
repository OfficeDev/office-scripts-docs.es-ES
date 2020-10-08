---
title: Mejorar el rendimiento de los scripts de Office
description: Cree scripts más rápidos mediante la comprensión de la comunicación entre el libro de Excel y el script.
ms.date: 06/15/2020
localization_priority: Normal
ms.openlocfilehash: 4d5b7c70f14e3fc598b95a6226e3ef8caf89f651
ms.sourcegitcommit: aec3c971c6640429f89b6bb99d2c95ea06725599
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 06/25/2020
ms.locfileid: "44878901"
---
# <a name="improve-the-performance-of-your-office-scripts"></a><span data-ttu-id="e89c5-103">Mejorar el rendimiento de los scripts de Office</span><span class="sxs-lookup"><span data-stu-id="e89c5-103">Improve the performance of your Office Scripts</span></span>

<span data-ttu-id="e89c5-104">El propósito de las secuencias de comandos de Office es automatizar la serie de tareas que se suele realizar para ahorrar tiempo.</span><span class="sxs-lookup"><span data-stu-id="e89c5-104">The purpose of Office Scripts is to automate commonly performed series of tasks to save you time.</span></span> <span data-ttu-id="e89c5-105">Un script lento puede sentir como que no acelera el flujo de trabajo.</span><span class="sxs-lookup"><span data-stu-id="e89c5-105">A slow script can feel like it doesn't speed up your workflow.</span></span> <span data-ttu-id="e89c5-106">La mayoría de las veces, el script será perfectamente correcto y se ejecutará como se esperaba.</span><span class="sxs-lookup"><span data-stu-id="e89c5-106">Most of the time, your script will be perfectly fine and run as expected.</span></span> <span data-ttu-id="e89c5-107">Sin embargo, hay algunos escenarios evitables que pueden afectar al rendimiento.</span><span class="sxs-lookup"><span data-stu-id="e89c5-107">However, there are a few, avoidable scenarios that can affect performance.</span></span>

<span data-ttu-id="e89c5-108">La causa más común de un script lento es una comunicación excesiva con el libro.</span><span class="sxs-lookup"><span data-stu-id="e89c5-108">The most common reason for a slow script is excessive communication with the workbook.</span></span> <span data-ttu-id="e89c5-109">El script se ejecuta en el equipo local, mientras que el libro está en la nube.</span><span class="sxs-lookup"><span data-stu-id="e89c5-109">Your script runs on your local machine, while the workbook exists in the cloud.</span></span> <span data-ttu-id="e89c5-110">En determinados momentos, el script sincroniza sus datos locales con el del libro.</span><span class="sxs-lookup"><span data-stu-id="e89c5-110">At certain times, your script synchronizes its local data with that of the workbook.</span></span> <span data-ttu-id="e89c5-111">Esto significa que las operaciones de escritura (como `workbook.addWorksheet()` ) solo se aplican al libro cuando se produce esta sincronización en segundo plano.</span><span class="sxs-lookup"><span data-stu-id="e89c5-111">This means that any write operations (such as `workbook.addWorksheet()`) are only applied to the workbook when this behind-the-scenes synchronization happens.</span></span> <span data-ttu-id="e89c5-112">Del mismo modo, las operaciones de lectura (como `myRange.getValues()` ) solo obtienen datos del libro para el script en esos momentos.</span><span class="sxs-lookup"><span data-stu-id="e89c5-112">Likewise, any read operations (such as `myRange.getValues()`) only get data from the workbook for the script at those times.</span></span> <span data-ttu-id="e89c5-113">En cualquier caso, el script recopila información antes de que actúe en los datos.</span><span class="sxs-lookup"><span data-stu-id="e89c5-113">In either case, the script fetches information before it acts on the data.</span></span> <span data-ttu-id="e89c5-114">Por ejemplo, el siguiente código registra con precisión el número de filas en el rango usado.</span><span class="sxs-lookup"><span data-stu-id="e89c5-114">For example, the following code will accurately log the number of rows in the used range.</span></span>

```TypeScript
let usedRange = workbook.getActiveWorksheet().getUsedRange();
let rowCount = usedRange.getRowCount();
// The script will read the range and row count from
// the workbook before logging the information.
console.log(rowCount);
```

<span data-ttu-id="e89c5-115">Las API de scripts de Office garantizan que los datos del libro o script sean precisos y actualizados cuando sea necesario.</span><span class="sxs-lookup"><span data-stu-id="e89c5-115">Office Scripts APIs ensure any data in the workbook or script is accurate and up-to-date when necessary.</span></span> <span data-ttu-id="e89c5-116">No tiene que preocuparse por estas sincronizaciones para que el script se ejecute correctamente.</span><span class="sxs-lookup"><span data-stu-id="e89c5-116">You don't need to worry about these synchronizations for your script to run correctly.</span></span> <span data-ttu-id="e89c5-117">Sin embargo, un conocimiento de esta comunicación de script a nube puede ayudarle a evitar llamadas de red no necesarias.</span><span class="sxs-lookup"><span data-stu-id="e89c5-117">However, an awareness of this script-to-cloud communication can help you avoid unneeded network calls.</span></span>

## <a name="performance-optimizations"></a><span data-ttu-id="e89c5-118">Optimizaciones de rendimiento</span><span class="sxs-lookup"><span data-stu-id="e89c5-118">Performance optimizations</span></span>

<span data-ttu-id="e89c5-119">Puede aplicar técnicas sencillas para ayudar a reducir la comunicación a la nube.</span><span class="sxs-lookup"><span data-stu-id="e89c5-119">You can apply simple techniques to help reduce the communication to the cloud.</span></span> <span data-ttu-id="e89c5-120">Los siguientes modelos ayudan a acelerar los scripts.</span><span class="sxs-lookup"><span data-stu-id="e89c5-120">The following patterns help speed up your scripts.</span></span>

- <span data-ttu-id="e89c5-121">Leer los datos del libro una vez en lugar de repetidamente en un bucle.</span><span class="sxs-lookup"><span data-stu-id="e89c5-121">Read workbook data once instead of repeatedly in a loop.</span></span>
- <span data-ttu-id="e89c5-122">Quite las instrucciones innecesarias `console.log` .</span><span class="sxs-lookup"><span data-stu-id="e89c5-122">Remove unnecessary `console.log` statements.</span></span>
- <span data-ttu-id="e89c5-123">Evite usar bloques try/catch.</span><span class="sxs-lookup"><span data-stu-id="e89c5-123">Avoid using try/catch blocks.</span></span>

### <a name="read-workbook-data-outside-of-a-loop"></a><span data-ttu-id="e89c5-124">Leer datos de un libro fuera de un bucle</span><span class="sxs-lookup"><span data-stu-id="e89c5-124">Read workbook data outside of a loop</span></span>

<span data-ttu-id="e89c5-125">Cualquier método que obtenga datos del libro puede desencadenar una llamada de red.</span><span class="sxs-lookup"><span data-stu-id="e89c5-125">Any method that gets data from the workbook can trigger a network call.</span></span> <span data-ttu-id="e89c5-126">En lugar de realizar repetidamente la misma llamada, debe guardar los datos de forma local siempre que sea posible.</span><span class="sxs-lookup"><span data-stu-id="e89c5-126">Rather than repeatedly making the same call, you should save data locally whenever possible.</span></span> <span data-ttu-id="e89c5-127">Esto es especialmente cierto al tratar con bucles.</span><span class="sxs-lookup"><span data-stu-id="e89c5-127">This is especially true when dealing with loops.</span></span>

<span data-ttu-id="e89c5-128">Considere un script para obtener el número de números negativos en el rango usado de una hoja de cálculo.</span><span class="sxs-lookup"><span data-stu-id="e89c5-128">Consider a script to get the count of negative numbers in the used range of a worksheet.</span></span> <span data-ttu-id="e89c5-129">El script debe recorrer en iteración todas las celdas del rango usado.</span><span class="sxs-lookup"><span data-stu-id="e89c5-129">The script needs to iterate over every cell in the used range.</span></span> <span data-ttu-id="e89c5-130">Para ello, necesita el intervalo, el número de filas y el número de columnas.</span><span class="sxs-lookup"><span data-stu-id="e89c5-130">To do that, it needs the range, the number of rows, and the number of columns.</span></span> <span data-ttu-id="e89c5-131">Debe almacenarlos como variables locales antes de iniciar el bucle.</span><span class="sxs-lookup"><span data-stu-id="e89c5-131">You should store those as local variables before starting the loop.</span></span> <span data-ttu-id="e89c5-132">De lo contrario, cada iteración del bucle forzará una devolución al libro.</span><span class="sxs-lookup"><span data-stu-id="e89c5-132">Otherwise, each iteration of the loop will force a return to the workbook.</span></span>

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
> <span data-ttu-id="e89c5-133">Como experimento, intente reemplazar `usedRangeValues` el bucle por `usedRange.getValues()` .</span><span class="sxs-lookup"><span data-stu-id="e89c5-133">As an experiment, try replacing `usedRangeValues` in the loop with `usedRange.getValues()`.</span></span> <span data-ttu-id="e89c5-134">Es posible que note que el script tarda mucho más tiempo en ejecutarse cuando se trata con rangos grandes.</span><span class="sxs-lookup"><span data-stu-id="e89c5-134">You may notice the script takes considerably longer to run when dealing with large ranges.</span></span>

### <a name="remove-unnecessary-consolelog-statements"></a><span data-ttu-id="e89c5-135">Quitar instrucciones innecesarias `console.log`</span><span class="sxs-lookup"><span data-stu-id="e89c5-135">Remove unnecessary `console.log` statements</span></span>

<span data-ttu-id="e89c5-136">El registro de consola es una herramienta vital para [la depuración de scripts](../testing/troubleshooting.md).</span><span class="sxs-lookup"><span data-stu-id="e89c5-136">Console logging is a vital tool for [debugging your scripts](../testing/troubleshooting.md).</span></span> <span data-ttu-id="e89c5-137">Sin embargo, sí obliga a que la secuencia de comandos se sincronice con el libro para asegurarse de que la información registrada está actualizada.</span><span class="sxs-lookup"><span data-stu-id="e89c5-137">However, it does force the script to synchronize with the workbook to ensure the logged information is up-to-date.</span></span> <span data-ttu-id="e89c5-138">Considere la posibilidad de quitar instrucciones de registro innecesarias (como las que se usan para las pruebas) antes de compartir el script.</span><span class="sxs-lookup"><span data-stu-id="e89c5-138">Consider removing unnecessary logging statements (such as those used for testing) before sharing your script.</span></span> <span data-ttu-id="e89c5-139">Esto normalmente no provocará un problema de rendimiento perceptible, a menos que la `console.log()` instrucción esté en un bucle.</span><span class="sxs-lookup"><span data-stu-id="e89c5-139">This typically won't cause a noticeable performance issue, unless the `console.log()` statement is in a loop.</span></span>

### <a name="avoid-using-trycatch-blocks"></a><span data-ttu-id="e89c5-140">Evitar el uso de bloques try/catch</span><span class="sxs-lookup"><span data-stu-id="e89c5-140">Avoid using try/catch blocks</span></span>

<span data-ttu-id="e89c5-141">No se recomienda usar [ `try` / `catch` bloques](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) como parte del flujo de control esperado de un script.</span><span class="sxs-lookup"><span data-stu-id="e89c5-141">We don't recommend using [`try`/`catch` blocks](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) as part of a script's expected control flow.</span></span> <span data-ttu-id="e89c5-142">La mayoría de los errores se pueden evitar comprobando los objetos devueltos del libro.</span><span class="sxs-lookup"><span data-stu-id="e89c5-142">Most errors can be avoided by checking objects returned from the workbook.</span></span> <span data-ttu-id="e89c5-143">Por ejemplo, el script siguiente comprueba que la tabla devuelta por el libro existe antes de intentar agregar una fila.</span><span class="sxs-lookup"><span data-stu-id="e89c5-143">For example, the following script checks that the table returned by the workbook exists before trying to add a row.</span></span>

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

## <a name="case-by-case-help"></a><span data-ttu-id="e89c5-144">Ayuda caso por caso</span><span class="sxs-lookup"><span data-stu-id="e89c5-144">Case-by-case help</span></span>

<span data-ttu-id="e89c5-145">A medida que la plataforma de scripts de Office se expande para trabajar con la automatización de la [potencia](https://flow.microsoft.com/), [tarjetas adaptables](https://docs.microsoft.com/adaptive-cards)y otras características de productos cruzados, los detalles de la comunicación del libro y de la secuencia de comandos se vuelven más complejos.</span><span class="sxs-lookup"><span data-stu-id="e89c5-145">As the Office Scripts platform expands to work with [Power Automate](https://flow.microsoft.com/), [Adaptive Cards](https://docs.microsoft.com/adaptive-cards), and other cross-product features, the details of the script-workbook communication become more intricate.</span></span> <span data-ttu-id="e89c5-146">Si necesita ayuda para que la secuencia de comandos se ejecute más rápido, póngase en contacto con el [desbordamiento de pila](https://stackoverflow.com/questions/tagged/office-scripts).</span><span class="sxs-lookup"><span data-stu-id="e89c5-146">If you need help making your script run faster, please reach out through [Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts).</span></span> <span data-ttu-id="e89c5-147">Asegúrese de etiquetar su pregunta con "Office-scripts" para que los expertos puedan encontrarla y ayudarle.</span><span class="sxs-lookup"><span data-stu-id="e89c5-147">Be sure to tag your question with "office-scripts" so experts can find it and help.</span></span>

## <a name="see-also"></a><span data-ttu-id="e89c5-148">Vea también</span><span class="sxs-lookup"><span data-stu-id="e89c5-148">See also</span></span>

- [<span data-ttu-id="e89c5-149">Conceptos básicos de los Scripts de Office en Excel en la web</span><span class="sxs-lookup"><span data-stu-id="e89c5-149">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](scripting-fundamentals.md)
- [<span data-ttu-id="e89c5-150">Documentos web de MDN: bucles e iteración</span><span class="sxs-lookup"><span data-stu-id="e89c5-150">MDN web docs: Loops and iteration</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Loops_and_iteration)
