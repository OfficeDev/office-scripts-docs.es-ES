---
title: Mejorar el rendimiento de los scripts Office scripts
description: Cree scripts más rápidos al comprender la comunicación entre el Excel y el script.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: a5bd879625b9c3bac0caa621dde312f7c961dd5c
ms.sourcegitcommit: 2aaf7dc527cb6c9f1206550b2c5745280503b2a3
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 06/16/2021
ms.locfileid: "52957703"
---
# <a name="improve-the-performance-of-your-office-scripts"></a><span data-ttu-id="5dd7a-103">Mejorar el rendimiento de los scripts Office scripts</span><span class="sxs-lookup"><span data-stu-id="5dd7a-103">Improve the performance of your Office Scripts</span></span>

<span data-ttu-id="5dd7a-104">El propósito de Office scripts es automatizar series de tareas que se realizan habitualmente para ahorrar tiempo.</span><span class="sxs-lookup"><span data-stu-id="5dd7a-104">The purpose of Office Scripts is to automate commonly performed series of tasks to save you time.</span></span> <span data-ttu-id="5dd7a-105">Un script lento puede sentir que no acelera el flujo de trabajo.</span><span class="sxs-lookup"><span data-stu-id="5dd7a-105">A slow script can feel like it doesn't speed up your workflow.</span></span> <span data-ttu-id="5dd7a-106">La mayoría de las veces, el script estará perfectamente bien y se ejecutará según lo esperado.</span><span class="sxs-lookup"><span data-stu-id="5dd7a-106">Most of the time, your script will be perfectly fine and run as expected.</span></span> <span data-ttu-id="5dd7a-107">Sin embargo, hay algunos escenarios evitables que pueden afectar al rendimiento.</span><span class="sxs-lookup"><span data-stu-id="5dd7a-107">However, there are a few, avoidable scenarios that can affect performance.</span></span>

<span data-ttu-id="5dd7a-108">La razón más común para un script lento es una comunicación excesiva con el libro.</span><span class="sxs-lookup"><span data-stu-id="5dd7a-108">The most common reason for a slow script is excessive communication with the workbook.</span></span> <span data-ttu-id="5dd7a-109">El script se ejecuta en el equipo local, mientras que el libro existe en la nube.</span><span class="sxs-lookup"><span data-stu-id="5dd7a-109">Your script runs on your local machine, while the workbook exists in the cloud.</span></span> <span data-ttu-id="5dd7a-110">En ciertos momentos, el script sincroniza sus datos locales con los del libro.</span><span class="sxs-lookup"><span data-stu-id="5dd7a-110">At certain times, your script synchronizes its local data with that of the workbook.</span></span> <span data-ttu-id="5dd7a-111">Esto significa que cualquier operación de escritura (como ) solo se aplica al libro cuando se produce esta sincronización entre `workbook.addWorksheet()` bastidores.</span><span class="sxs-lookup"><span data-stu-id="5dd7a-111">This means that any write operations (such as `workbook.addWorksheet()`) are only applied to the workbook when this behind-the-scenes synchronization happens.</span></span> <span data-ttu-id="5dd7a-112">Del mismo modo, las operaciones de lectura (como ) solo obtienen datos `myRange.getValues()` del libro para el script en esos momentos.</span><span class="sxs-lookup"><span data-stu-id="5dd7a-112">Likewise, any read operations (such as `myRange.getValues()`) only get data from the workbook for the script at those times.</span></span> <span data-ttu-id="5dd7a-113">En cualquier caso, el script captura información antes de que actúe en los datos.</span><span class="sxs-lookup"><span data-stu-id="5dd7a-113">In either case, the script fetches information before it acts on the data.</span></span> <span data-ttu-id="5dd7a-114">Por ejemplo, el siguiente código registrará con precisión el número de filas en el intervalo usado.</span><span class="sxs-lookup"><span data-stu-id="5dd7a-114">For example, the following code will accurately log the number of rows in the used range.</span></span>

```TypeScript
let usedRange = workbook.getActiveWorksheet().getUsedRange();
let rowCount = usedRange.getRowCount();
// The script will read the range and row count from
// the workbook before logging the information.
console.log(rowCount);
```

<span data-ttu-id="5dd7a-115">Office Las API de scripts garantizan que los datos del libro o script sean precisos y actualizados cuando sea necesario.</span><span class="sxs-lookup"><span data-stu-id="5dd7a-115">Office Scripts APIs ensure any data in the workbook or script is accurate and up-to-date when necessary.</span></span> <span data-ttu-id="5dd7a-116">No es necesario preocuparse por estas sincronizaciones para que el script se ejecute correctamente.</span><span class="sxs-lookup"><span data-stu-id="5dd7a-116">You don't need to worry about these synchronizations for your script to run correctly.</span></span> <span data-ttu-id="5dd7a-117">Sin embargo, un conocimiento de esta comunicación de script a nube puede ayudarle a evitar llamadas de red innecesarios.</span><span class="sxs-lookup"><span data-stu-id="5dd7a-117">However, an awareness of this script-to-cloud communication can help you avoid unneeded network calls.</span></span>

## <a name="performance-optimizations"></a><span data-ttu-id="5dd7a-118">Optimizaciones de rendimiento</span><span class="sxs-lookup"><span data-stu-id="5dd7a-118">Performance optimizations</span></span>

<span data-ttu-id="5dd7a-119">Puede aplicar técnicas sencillas para ayudar a reducir la comunicación a la nube.</span><span class="sxs-lookup"><span data-stu-id="5dd7a-119">You can apply simple techniques to help reduce the communication to the cloud.</span></span> <span data-ttu-id="5dd7a-120">Los siguientes patrones ayudan a acelerar los scripts.</span><span class="sxs-lookup"><span data-stu-id="5dd7a-120">The following patterns help speed up your scripts.</span></span>

- <span data-ttu-id="5dd7a-121">Lea los datos del libro una vez en lugar de repetirlo en un bucle.</span><span class="sxs-lookup"><span data-stu-id="5dd7a-121">Read workbook data once instead of repeatedly in a loop.</span></span>
- <span data-ttu-id="5dd7a-122">Quite instrucciones `console.log` innecesarias.</span><span class="sxs-lookup"><span data-stu-id="5dd7a-122">Remove unnecessary `console.log` statements.</span></span>
- <span data-ttu-id="5dd7a-123">Evite usar bloques try/catch.</span><span class="sxs-lookup"><span data-stu-id="5dd7a-123">Avoid using try/catch blocks.</span></span>

### <a name="read-workbook-data-outside-of-a-loop"></a><span data-ttu-id="5dd7a-124">Leer datos del libro fuera de un bucle</span><span class="sxs-lookup"><span data-stu-id="5dd7a-124">Read workbook data outside of a loop</span></span>

<span data-ttu-id="5dd7a-125">Cualquier método que obtiene datos del libro puede desencadenar una llamada de red.</span><span class="sxs-lookup"><span data-stu-id="5dd7a-125">Any method that gets data from the workbook can trigger a network call.</span></span> <span data-ttu-id="5dd7a-126">En lugar de realizar la misma llamada repetidamente, debe guardar los datos localmente siempre que sea posible.</span><span class="sxs-lookup"><span data-stu-id="5dd7a-126">Rather than repeatedly making the same call, you should save data locally whenever possible.</span></span> <span data-ttu-id="5dd7a-127">Esto es especialmente cierto cuando se trata de bucles.</span><span class="sxs-lookup"><span data-stu-id="5dd7a-127">This is especially true when dealing with loops.</span></span>

<span data-ttu-id="5dd7a-128">Considere un script para obtener el recuento de números negativos en el rango usado de una hoja de cálculo.</span><span class="sxs-lookup"><span data-stu-id="5dd7a-128">Consider a script to get the count of negative numbers in the used range of a worksheet.</span></span> <span data-ttu-id="5dd7a-129">El script debe iterar en todas las celdas del intervalo usado.</span><span class="sxs-lookup"><span data-stu-id="5dd7a-129">The script needs to iterate over every cell in the used range.</span></span> <span data-ttu-id="5dd7a-130">Para ello, necesita el intervalo, el número de filas y el número de columnas.</span><span class="sxs-lookup"><span data-stu-id="5dd7a-130">To do that, it needs the range, the number of rows, and the number of columns.</span></span> <span data-ttu-id="5dd7a-131">Debe almacenar esas variables como variables locales antes de iniciar el bucle.</span><span class="sxs-lookup"><span data-stu-id="5dd7a-131">You should store those as local variables before starting the loop.</span></span> <span data-ttu-id="5dd7a-132">De lo contrario, cada iteración del bucle forzará un retorno al libro.</span><span class="sxs-lookup"><span data-stu-id="5dd7a-132">Otherwise, each iteration of the loop will force a return to the workbook.</span></span>

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
> <span data-ttu-id="5dd7a-133">Como experimento, intente reemplazar `usedRangeValues` en el bucle por `usedRange.getValues()` .</span><span class="sxs-lookup"><span data-stu-id="5dd7a-133">As an experiment, try replacing `usedRangeValues` in the loop with `usedRange.getValues()`.</span></span> <span data-ttu-id="5dd7a-134">Es posible que observe que el script tarda considerablemente más tiempo en ejecutarse cuando se trata de intervalos grandes.</span><span class="sxs-lookup"><span data-stu-id="5dd7a-134">You may notice the script takes considerably longer to run when dealing with large ranges.</span></span>

### <a name="avoid-using-trycatch-blocks-in-or-surrounding-loops"></a><span data-ttu-id="5dd7a-135">Evitar el uso `try...catch` de bloques en bucles o entornos</span><span class="sxs-lookup"><span data-stu-id="5dd7a-135">Avoid using `try...catch` blocks in or surrounding loops</span></span>

<span data-ttu-id="5dd7a-136">No se recomienda usar instrucciones en [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) bucles o bucles circundantes.</span><span class="sxs-lookup"><span data-stu-id="5dd7a-136">We don't recommend using [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) statements either in loops or surrounding loops.</span></span> <span data-ttu-id="5dd7a-137">Este es el mismo motivo por el que debe evitar leer datos en un bucle: cada iteración fuerza al script a sincronizarse con el libro para asegurarse de que no se ha producido ningún error.</span><span class="sxs-lookup"><span data-stu-id="5dd7a-137">This is for the same reason you should avoid reading data in a loop: each iteration forces the script to synchronize with the workbook to make sure no error has been thrown.</span></span> <span data-ttu-id="5dd7a-138">La mayoría de los errores se pueden evitar comprobando los objetos devueltos desde el libro.</span><span class="sxs-lookup"><span data-stu-id="5dd7a-138">Most errors can be avoided by checking objects returned from the workbook.</span></span> <span data-ttu-id="5dd7a-139">Por ejemplo, el siguiente script comprueba que la tabla devuelta por el libro existe antes de intentar agregar una fila.</span><span class="sxs-lookup"><span data-stu-id="5dd7a-139">For example, the following script checks that the table returned by the workbook exists before trying to add a row.</span></span>

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

### <a name="remove-unnecessary-consolelog-statements"></a><span data-ttu-id="5dd7a-140">Quitar instrucciones `console.log` innecesarias</span><span class="sxs-lookup"><span data-stu-id="5dd7a-140">Remove unnecessary `console.log` statements</span></span>

<span data-ttu-id="5dd7a-141">El registro de consola es una herramienta vital [para depurar los scripts.](../testing/troubleshooting.md)</span><span class="sxs-lookup"><span data-stu-id="5dd7a-141">Console logging is a vital tool for [debugging your scripts](../testing/troubleshooting.md).</span></span> <span data-ttu-id="5dd7a-142">Sin embargo, obliga al script a sincronizarse con el libro para asegurarse de que la información registrada está actualizada.</span><span class="sxs-lookup"><span data-stu-id="5dd7a-142">However, it does force the script to synchronize with the workbook to ensure the logged information is up-to-date.</span></span> <span data-ttu-id="5dd7a-143">Considere la posibilidad de quitar instrucciones de registro innecesarias (como las que se usan para las pruebas) antes de compartir el script.</span><span class="sxs-lookup"><span data-stu-id="5dd7a-143">Consider removing unnecessary logging statements (such as those used for testing) before sharing your script.</span></span> <span data-ttu-id="5dd7a-144">Por lo general, esto no causará un problema de rendimiento notable, a menos que la `console.log()` instrucción esté en un bucle.</span><span class="sxs-lookup"><span data-stu-id="5dd7a-144">This typically won't cause a noticeable performance issue, unless the `console.log()` statement is in a loop.</span></span>

## <a name="case-by-case-help"></a><span data-ttu-id="5dd7a-145">Ayuda caso por caso</span><span class="sxs-lookup"><span data-stu-id="5dd7a-145">Case-by-case help</span></span>

<span data-ttu-id="5dd7a-146">A medida que la plataforma de scripts de Office [](/adaptive-cards)se expande para funcionar con [Power Automate,](https://flow.microsoft.com/)tarjetas adaptables y otras características entre productos, los detalles de la comunicación entre scripts y libros se vuelven más complejos.</span><span class="sxs-lookup"><span data-stu-id="5dd7a-146">As the Office Scripts platform expands to work with [Power Automate](https://flow.microsoft.com/), [Adaptive Cards](/adaptive-cards), and other cross-product features, the details of the script-workbook communication become more intricate.</span></span> <span data-ttu-id="5dd7a-147">Si necesita ayuda para que el script se ejecute más rápido, póngase en contacto con [Microsoft Q&A](/answers/topics/office-scripts-excel-dev.html).</span><span class="sxs-lookup"><span data-stu-id="5dd7a-147">If you need help making your script run faster, please reach out through [Microsoft Q&A](/answers/topics/office-scripts-excel-dev.html).</span></span> <span data-ttu-id="5dd7a-148">Asegúrese de etiquetar su pregunta con "office-scripts-dev" para que los expertos puedan encontrarlo y ayudarle.</span><span class="sxs-lookup"><span data-stu-id="5dd7a-148">Be sure to tag your question with "office-scripts-dev" so experts can find it and help.</span></span>

## <a name="see-also"></a><span data-ttu-id="5dd7a-149">Vea también</span><span class="sxs-lookup"><span data-stu-id="5dd7a-149">See also</span></span>

- [<span data-ttu-id="5dd7a-150">Conceptos básicos de los Scripts de Office en Excel en la web</span><span class="sxs-lookup"><span data-stu-id="5dd7a-150">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](scripting-fundamentals.md)
- [<span data-ttu-id="5dd7a-151">Documentos web de MDN: bucles e iteración</span><span class="sxs-lookup"><span data-stu-id="5dd7a-151">MDN web docs: Loops and iteration</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Loops_and_iteration)
