---
title: Procedimientos recomendados para Scripts de Office
description: Cómo evitar problemas comunes y escribir scripts Office que puedan controlar datos o entradas inesperadas.
ms.date: 05/10/2021
localization_priority: Normal
ms.openlocfilehash: 0697e6fd1fa8f437a4a585d938254deb5a05f20c
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 05/19/2021
ms.locfileid: "52546034"
---
# <a name="best-practices-in-office-scripts"></a><span data-ttu-id="83970-103">Procedimientos recomendados para Scripts de Office</span><span class="sxs-lookup"><span data-stu-id="83970-103">Best practices in Office Scripts</span></span>

<span data-ttu-id="83970-104">Estos patrones y prácticas están diseñados para ayudar a que los scripts se ejecuten correctamente cada vez.</span><span class="sxs-lookup"><span data-stu-id="83970-104">These patterns and practices are designed to help your scripts run successfully every time.</span></span> <span data-ttu-id="83970-105">Úselos para evitar problemas comunes al empezar a automatizar el flujo Excel de trabajo.</span><span class="sxs-lookup"><span data-stu-id="83970-105">Use them to avoid common pitfalls as you start automating your Excel workflow.</span></span>

## <a name="verify-an-object-is-present"></a><span data-ttu-id="83970-106">Comprobar que un objeto está presente</span><span class="sxs-lookup"><span data-stu-id="83970-106">Verify an object is present</span></span>

<span data-ttu-id="83970-107">Los scripts suelen basarse en una determinada hoja de cálculo o tabla que está presente en el libro.</span><span class="sxs-lookup"><span data-stu-id="83970-107">Scripts often rely on a certain worksheet or table being present in the workbook.</span></span> <span data-ttu-id="83970-108">Sin embargo, pueden cambiar el nombre o quitarse entre las ejecuciones de scripts.</span><span class="sxs-lookup"><span data-stu-id="83970-108">However, they might get renamed or removed between script runs.</span></span> <span data-ttu-id="83970-109">Al comprobar si esas tablas o hojas de cálculo existen antes de llamar a métodos en ellas, puede asegurarse de que el script no termine abruptamente.</span><span class="sxs-lookup"><span data-stu-id="83970-109">By checking if those tables or worksheets exist before calling methods on them, you can make sure the script doesn't end abruptly.</span></span>

<span data-ttu-id="83970-110">El siguiente código de ejemplo comprueba si la hoja de cálculo "Índice" está presente en el libro.</span><span class="sxs-lookup"><span data-stu-id="83970-110">The following sample code checks if the "Index" worksheet is present in the workbook.</span></span> <span data-ttu-id="83970-111">Si la hoja de cálculo está presente, el script obtiene un rango y procede.</span><span class="sxs-lookup"><span data-stu-id="83970-111">If the worksheet is present, the script gets a range and proceeds.</span></span> <span data-ttu-id="83970-112">Si no está presente, el script registra un mensaje de error personalizado.</span><span class="sxs-lookup"><span data-stu-id="83970-112">If it isn't present, the script logs a custom error message.</span></span>

```TypeScript
// Make sure the "Index" worksheet exists before using it.
let indexSheet = workbook.getWorksheet('Index');
if (indexSheet) {
  let range = indexSheet.getRange("A1");
  // Continue using the range...
} else {
  console.log("Index sheet not found.");
}
```

<span data-ttu-id="83970-113">El operador TypeScript `?` comprueba si el objeto existe antes de llamar a un método.</span><span class="sxs-lookup"><span data-stu-id="83970-113">The TypeScript `?` operator checks if the object exists before calling a method.</span></span> <span data-ttu-id="83970-114">Esto puede simplificar el código si no necesita hacer nada especial cuando el objeto no existe.</span><span class="sxs-lookup"><span data-stu-id="83970-114">This can make your code more streamlined if you don't need to do anything special when the object doesn't exist.</span></span>

```TypeScript
// The ? ensures that the delete() API is only called if the object exists.
workbook.getWorksheet('Index')?.delete();
```

## <a name="validate-data-and-workbook-state-first"></a><span data-ttu-id="83970-115">Validar primero los datos y el estado del libro</span><span class="sxs-lookup"><span data-stu-id="83970-115">Validate data and workbook state first</span></span>

<span data-ttu-id="83970-116">Asegúrese de que todas las hojas de cálculo, tablas, formas y otros objetos estén presentes antes de trabajar en los datos.</span><span class="sxs-lookup"><span data-stu-id="83970-116">Make sure all your worksheets, tables, shapes, and other objects are present before working on the data.</span></span> <span data-ttu-id="83970-117">Con el patrón anterior, compruebe si todo está en el libro y coincide con sus expectativas.</span><span class="sxs-lookup"><span data-stu-id="83970-117">Using the previous pattern, check to see if everything is in the workbook and matches your expectations.</span></span> <span data-ttu-id="83970-118">Al hacerlo antes de escribir los datos, se asegura de que el script no deje el libro en estado parcial.</span><span class="sxs-lookup"><span data-stu-id="83970-118">Doing this before any data is written ensures your script doesn't leave the workbook in a partial state.</span></span>

<span data-ttu-id="83970-119">El siguiente script requiere que se presenten dos tablas denominadas "Table1" y "Table2".</span><span class="sxs-lookup"><span data-stu-id="83970-119">The following script requires two tables named "Table1" and "Table2" to be present.</span></span> <span data-ttu-id="83970-120">El script comprueba primero si las tablas están presentes y, a continuación, termina con la instrucción y un `return` mensaje adecuado si no lo están.</span><span class="sxs-lookup"><span data-stu-id="83970-120">The script first checks if the tables are present and then ends with the `return` statement and an appropriate message if they're not.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // These tables must be in the workbook for the script.
  const TargetTableName = 'Table1';
  const SourceTableName = 'Table2';

  // Get the table objects.
  let targetTable = workbook.getTable(TargetTableName);
  let sourceTable = workbook.getTable(SourceTableName);

  // Check if the tables are there.
  if (!targetTable || !sourceTable) {
    console.log(`Required tables missing - Check that both the source (${TargetTableName}) and target (${SourceTableName}) tables are present before running the script.`);
    return;
  }

  // Continue....
}
```

<span data-ttu-id="83970-121">Si la comprobación se está produciendo en una función independiente, debe finalizar el script emitiendo la `return` instrucción de la `main` función.</span><span class="sxs-lookup"><span data-stu-id="83970-121">If the verification is happening in a separate function, you still must end the script by issuing the `return` statement from the `main` function.</span></span> <span data-ttu-id="83970-122">La devolución de la subfunción no finaliza el script.</span><span class="sxs-lookup"><span data-stu-id="83970-122">Returning from the subfunction doesn't end the script.</span></span>

<span data-ttu-id="83970-123">El siguiente script tiene el mismo comportamiento que el anterior.</span><span class="sxs-lookup"><span data-stu-id="83970-123">The following script has the same behavior as the previous one.</span></span> <span data-ttu-id="83970-124">La diferencia es que la `main` función llama a la función para comprobar `inputPresent` todo.</span><span class="sxs-lookup"><span data-stu-id="83970-124">The difference is that the `main` function calls the `inputPresent` function to verify everything.</span></span> <span data-ttu-id="83970-125">`inputPresent` devuelve un valor booleano ( `true` o ) para indicar si todas las entradas necesarias están `false` presentes.</span><span class="sxs-lookup"><span data-stu-id="83970-125">`inputPresent` returns a boolean (`true` or `false`) to indicate whether all required inputs are present.</span></span> <span data-ttu-id="83970-126">La `main` función usa ese valor booleano para decidir si continúa o finaliza el script.</span><span class="sxs-lookup"><span data-stu-id="83970-126">The `main` function uses that boolean to decide on continuing or ending the script.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {

  // Get the table objects.
  if (!inputPresent(workbook)) {
    return;
  }

  // Continue....
}

function inputPresent( workbook: ExcelScript.Workbook): boolean {
  // These tables must be in the workbook for the script.
  const TargetTableName = 'Table1';
  const SourceTableName = 'Table2';

  // Get the table objects.
  let targetTable = workbook.getTable(TargetTableName);
  let sourceTable = workbook.getTable(SourceTableName);

  // Check if the tables are there.
  if (!targetTable || !sourceTable) {
    console.log(`Required tables missing - Check that both the source (${TargetTableName}) and target (${SourceTableName}) tables are present before running the script.`);
    return false;
  }

  return true;
}
```

## <a name="when-to-use-a-throw-statement"></a><span data-ttu-id="83970-127">Cuándo usar una `throw` instrucción</span><span class="sxs-lookup"><span data-stu-id="83970-127">When to use a `throw` statement</span></span>

<span data-ttu-id="83970-128">Una [`throw`](https://developer.mozilla.org/docs/web/javascript/reference/statements/throw) instrucción indica que se ha producido un error inesperado.</span><span class="sxs-lookup"><span data-stu-id="83970-128">A [`throw`](https://developer.mozilla.org/docs/web/javascript/reference/statements/throw) statement indicates an unexpected error has occurred.</span></span> <span data-ttu-id="83970-129">Finaliza el código inmediatamente.</span><span class="sxs-lookup"><span data-stu-id="83970-129">It ends the code immediately.</span></span> <span data-ttu-id="83970-130">En su mayoría, no es necesario desde `throw` el script.</span><span class="sxs-lookup"><span data-stu-id="83970-130">For the most part, you don't need to `throw` from your script.</span></span> <span data-ttu-id="83970-131">Normalmente, el script informa automáticamente al usuario de que el script no se pudo ejecutar debido a un problema.</span><span class="sxs-lookup"><span data-stu-id="83970-131">Usually, the script automatically informs the user that the script failed to run due to an issue.</span></span> <span data-ttu-id="83970-132">En la mayoría de los casos, basta con finalizar el script con un mensaje de error y una `return` instrucción de la `main` función.</span><span class="sxs-lookup"><span data-stu-id="83970-132">In most cases, it's sufficient to end the script with an error message and a `return` statement from the `main` function.</span></span>

<span data-ttu-id="83970-133">Sin embargo, si el script se ejecuta como parte de un flujo Power Automate, es posible que desee detener el flujo para que no continúe.</span><span class="sxs-lookup"><span data-stu-id="83970-133">However, if your script is running as part of a Power Automate flow, you may want to stop the flow from continuing.</span></span> <span data-ttu-id="83970-134">Una `throw` instrucción detiene el script y le indica al flujo que se detenga también.</span><span class="sxs-lookup"><span data-stu-id="83970-134">A `throw` statement stops the script and tells the flow to stop as well.</span></span>

<span data-ttu-id="83970-135">El siguiente script muestra cómo usar la `throw` instrucción en nuestro ejemplo de comprobación de tabla.</span><span class="sxs-lookup"><span data-stu-id="83970-135">The following script shows how to use the `throw` statement in our table checking example.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // These tables must be in the workbook for the script.
  const TargetTableName = 'Table1';
  const SourceTableName = 'Table2';

  // Get the table objects.
  let targetTable = workbook.getTable(TargetTableName);
  let sourceTable = workbook.getTable(SourceTableName);

  // Check if the tables are there.
  if (!targetTable || !sourceTable) {
    // Immediately end the script with an error.
    throw `Required tables missing - Check that both the source (${TargetTableName}) and target (${SourceTableName}) tables are present before running the script.`;
  }
  
```

## <a name="when-to-use-a-trycatch-statement"></a><span data-ttu-id="83970-136">Cuándo usar una `try...catch` instrucción</span><span class="sxs-lookup"><span data-stu-id="83970-136">When to use a `try...catch` statement</span></span>

<span data-ttu-id="83970-137">La instrucción es una forma de detectar si se produce un error en una llamada [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) API y seguir ejecutando el script.</span><span class="sxs-lookup"><span data-stu-id="83970-137">The [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) statement is a way to detect if an API call fails and continue running the script.</span></span>

<span data-ttu-id="83970-138">Tenga en cuenta el siguiente fragmento de código que realiza una actualización de datos de gran tamaño en un intervalo.</span><span class="sxs-lookup"><span data-stu-id="83970-138">Consider the following snippet that performs a large data update on a range.</span></span>

```TypeScript
range.setValues(someLargeValues);
```

<span data-ttu-id="83970-139">Si es mayor que Excel puede controlar `someLargeValues` la web, se produce `setValues()` un error en la llamada.</span><span class="sxs-lookup"><span data-stu-id="83970-139">If `someLargeValues` is larger than Excel for the web can handle, the `setValues()` call fails.</span></span> <span data-ttu-id="83970-140">A continuación, el script también produce un [error en tiempo de ejecución](../testing/troubleshooting.md#runtime-errors).</span><span class="sxs-lookup"><span data-stu-id="83970-140">The script then also fails with a [runtime error](../testing/troubleshooting.md#runtime-errors).</span></span> <span data-ttu-id="83970-141">La `try...catch` instrucción permite que el script reconozca esta condición, sin terminar inmediatamente el script y mostrar el error predeterminado.</span><span class="sxs-lookup"><span data-stu-id="83970-141">The `try...catch` statement lets your script recognize this condition, without immediately ending the script and showing the default error.</span></span>

<span data-ttu-id="83970-142">Un enfoque para proporcionar al usuario de script una mejor experiencia es presentarles un mensaje de error personalizado.</span><span class="sxs-lookup"><span data-stu-id="83970-142">One approach for giving the script user a better experience is to present them a custom error message.</span></span> <span data-ttu-id="83970-143">El siguiente fragmento de código muestra `try...catch` una instrucción que registra más información de error para ayudar mejor al lector.</span><span class="sxs-lookup"><span data-stu-id="83970-143">The following snippet shows a `try...catch` statement logging more error information to better help the reader.</span></span>

```TypeScript
try {
    range.setValues(someLargeValues);
} catch (error) {
    console.log(`The script failed to update the values at location ${range.getAddress()}. Please inspect and run again.`);
    console.log(error);
    return; // End the script (assuming this is in the main function).
}
```

<span data-ttu-id="83970-144">Otro enfoque para tratar los errores es tener un comportamiento de reserva que controle el caso de error.</span><span class="sxs-lookup"><span data-stu-id="83970-144">Another approach to dealing with errors is to have fallback behavior that handles the error case.</span></span> <span data-ttu-id="83970-145">El siguiente fragmento de código usa el bloque para intentar un método alternativo dividir la actualización en partes `catch` más pequeñas y evitar el error.</span><span class="sxs-lookup"><span data-stu-id="83970-145">The following snippet uses the `catch` block to try an alternate method break up the update into smaller pieces and avoid the error.</span></span>

> [!TIP]
> <span data-ttu-id="83970-146">Para obtener un ejemplo completo sobre cómo actualizar un intervalo grande, vea [Escribir un conjunto de datos grande.](../resources/samples/write-large-dataset.md)</span><span class="sxs-lookup"><span data-stu-id="83970-146">For a full example on how to update a large range, see [Write a large dataset](../resources/samples/write-large-dataset.md).</span></span>

```TypeScript
try {
    range.setValues(someLargeValues);
} catch (error) {
    console.log(`The script failed to update the values at location ${range.getAddress()}. Trying a different approach.`);
    handleUpdatesInSmallerBatches(someLargeValues);
}

// Continue...
}
```

> [!NOTE]
> <span data-ttu-id="83970-147">Usar `try...catch` dentro o alrededor de un bucle ralentiza el script.</span><span class="sxs-lookup"><span data-stu-id="83970-147">Using `try...catch` inside or around a loop slows down your script.</span></span> <span data-ttu-id="83970-148">Para obtener más información sobre el rendimiento, vea [Avoid using `try...catch` blocks](web-client-performance.md#avoid-using-trycatch-blocks-in-or-surrounding-loops).</span><span class="sxs-lookup"><span data-stu-id="83970-148">For more performance information, see [Avoid using `try...catch` blocks](web-client-performance.md#avoid-using-trycatch-blocks-in-or-surrounding-loops).</span></span>

## <a name="see-also"></a><span data-ttu-id="83970-149">Consulte también</span><span class="sxs-lookup"><span data-stu-id="83970-149">See also</span></span>

- [<span data-ttu-id="83970-150">Solución de problemas de scripts de Office</span><span class="sxs-lookup"><span data-stu-id="83970-150">Troubleshooting Office Scripts</span></span>](../testing/troubleshooting.md)
- [<span data-ttu-id="83970-151">Solución de problemas de información Power Automate con scripts Office datos</span><span class="sxs-lookup"><span data-stu-id="83970-151">Troubleshooting information for Power Automate with Office Scripts</span></span>](../testing/power-automate-troubleshooting.md)
- [<span data-ttu-id="83970-152">Límites de plataforma con Office scripts</span><span class="sxs-lookup"><span data-stu-id="83970-152">Platform limits with Office Scripts</span></span>](../testing/platform-limits.md)
- [<span data-ttu-id="83970-153">Mejorar el rendimiento de los scripts Office scripts</span><span class="sxs-lookup"><span data-stu-id="83970-153">Improve the performance of your Office Scripts</span></span>](web-client-performance.md)
