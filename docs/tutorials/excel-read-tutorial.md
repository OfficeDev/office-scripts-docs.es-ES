---
title: Leer datos de libros con scripts de Office en Excel en la Web
description: Un tutorial de scripts de Office sobre cómo leer datos de libros y evaluarlos en el script.
ms.date: 01/06/2021
localization_priority: Priority
ms.openlocfilehash: 38a8c64fcf0ed998049c2b95c6d4ede148b6925b
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: HT
ms.contentlocale: es-ES
ms.lasthandoff: 07/07/2021
ms.locfileid: "53313991"
---
# <a name="read-workbook-data-with-office-scripts-in-excel-on-the-web"></a><span data-ttu-id="d7389-103">Leer datos de libros con scripts de Office en Excel en la Web</span><span class="sxs-lookup"><span data-stu-id="d7389-103">Read workbook data with Office Scripts in Excel on the web</span></span>

<span data-ttu-id="d7389-104">Este tutorial le enseña a leer datos de un libro con un script de Office para Excel en la Web.</span><span class="sxs-lookup"><span data-stu-id="d7389-104">This tutorial teaches you how to read data from a workbook with an Office Script for Excel on the web.</span></span> <span data-ttu-id="d7389-105">Escribirá un nuevo script que dé formato a un extracto bancario y normalice los datos en ese extracto.</span><span class="sxs-lookup"><span data-stu-id="d7389-105">You'll be writing a new script that formats a bank statement and normalizes the data in that statement.</span></span> <span data-ttu-id="d7389-106">Como parte de la limpieza de datos, el script leerá valores de las celdas de transacción, aplicará una fórmula simple a cada valor y escribirá la respuesta resultante en el libro.</span><span class="sxs-lookup"><span data-stu-id="d7389-106">As part of that data clean-up, your script will read values from the transaction cells, apply a simple formula to each value, and write the resulting answer to the workbook.</span></span> <span data-ttu-id="d7389-107">La lectura de datos del libro le permite automatizar algunos de los procesos de toma de decisiones en el script.</span><span class="sxs-lookup"><span data-stu-id="d7389-107">Reading data from the workbook lets you automate some of your decision making processes in the script.</span></span>

> [!TIP]
> <span data-ttu-id="d7389-108">Si no está familiarizado con los scripts de Office, le recomendamos que empiece con el tutorial de [Guardar, editar y crear scripts de Office en Excel en la Web](excel-tutorial.md).</span><span class="sxs-lookup"><span data-stu-id="d7389-108">If you are new to Office Scripts, we recommend starting with the [Record, edit, and create Office Scripts in Excel on the web](excel-tutorial.md) tutorial.</span></span> <span data-ttu-id="d7389-109">[Scripts de Office usa TypeScript](../overview/code-editor-environment.md) y este tutorial está diseñado para las personas con conocimientos de nivel intermedio de JavaScript o TypeScript.</span><span class="sxs-lookup"><span data-stu-id="d7389-109">[Office Scripts use TypeScript](../overview/code-editor-environment.md) and this tutorial is intended for people with beginner to intermediate-level knowledge of JavaScript or TypeScript.</span></span> <span data-ttu-id="d7389-110">Si no está familiarizado con JavaScript, le recomendamos que comience con el [Tutorial de JavaScript de Mozilla](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span><span class="sxs-lookup"><span data-stu-id="d7389-110">If you're new to JavaScript, we recommend starting with the [Mozilla JavaScript tutorial](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="d7389-111">Requisitos previos</span><span class="sxs-lookup"><span data-stu-id="d7389-111">Prerequisites</span></span>

[!INCLUDE [Tutorial prerequisites](../includes/tutorial-prerequisites.md)]

## <a name="read-a-cell"></a><span data-ttu-id="d7389-112">Leer una celda</span><span class="sxs-lookup"><span data-stu-id="d7389-112">Read a cell</span></span>

<span data-ttu-id="d7389-113">Los scripts creados con la Grabadora de acciones solo pueden escribir información en el libro.</span><span class="sxs-lookup"><span data-stu-id="d7389-113">Scripts made with the Action Recorder can only write information to the workbook.</span></span> <span data-ttu-id="d7389-114">Con el Editor de código, puede además editar y escribir scripts que lean datos de un libro.</span><span class="sxs-lookup"><span data-stu-id="d7389-114">With the Code Editor, you can edit and make scripts that also read data from a workbook.</span></span>

<span data-ttu-id="d7389-115">Creemos un script que lea datos y actúe en función de lo que lee.</span><span class="sxs-lookup"><span data-stu-id="d7389-115">Let's make a script that reads data and acts based on what was read.</span></span> <span data-ttu-id="d7389-116">Vamos a trabajar con un ejemplo de extracto bancario.</span><span class="sxs-lookup"><span data-stu-id="d7389-116">We're going to work with a sample banking statement.</span></span> <span data-ttu-id="d7389-117">Este ejemplo es una declaración combinada de cuenta corriente y crédito.</span><span class="sxs-lookup"><span data-stu-id="d7389-117">This statement is a combined checking and credit statement.</span></span> <span data-ttu-id="d7389-118">Lamentablemente, los informes de saldo muestran los cambios de forma diferente.</span><span class="sxs-lookup"><span data-stu-id="d7389-118">Unfortunately, they report balance changes differently.</span></span> <span data-ttu-id="d7389-119">La declaración de cuenta corriente muestra los ingresos como crédito positivo y los costes como débito negativo.</span><span class="sxs-lookup"><span data-stu-id="d7389-119">The checking statement gives income as positive credit and costs as negative debit.</span></span> <span data-ttu-id="d7389-120">En cambio, la declaración de crédito funciona de manera contraria.</span><span class="sxs-lookup"><span data-stu-id="d7389-120">The credit statement does the opposite.</span></span>

<span data-ttu-id="d7389-121">En el resto del tutorial, armonizaremos los datos de ambos con un script.</span><span class="sxs-lookup"><span data-stu-id="d7389-121">Over the rest of the tutorial, we will normalize this data using a script.</span></span> <span data-ttu-id="d7389-122">En primer lugar, veamos cómo leer los datos del libro.</span><span class="sxs-lookup"><span data-stu-id="d7389-122">First, let's learn how to read data from the workbook.</span></span>

1. <span data-ttu-id="d7389-123">Cree una nueva hoja de cálculo en el libro que ha usado para el resto del tutorial.</span><span class="sxs-lookup"><span data-stu-id="d7389-123">Create a new worksheet in the workbook you've used for the rest of the tutorial.</span></span>
2. <span data-ttu-id="d7389-124">Copie los siguientes datos y péguelos en la nueva hoja de cálculo, comenzando por la celda **A1**.</span><span class="sxs-lookup"><span data-stu-id="d7389-124">Copy the following data and paste it into the new worksheet, starting at cell **A1**.</span></span>

    |<span data-ttu-id="d7389-125">Fecha</span><span class="sxs-lookup"><span data-stu-id="d7389-125">Date</span></span> |<span data-ttu-id="d7389-126">Cuenta</span><span class="sxs-lookup"><span data-stu-id="d7389-126">Account</span></span> |<span data-ttu-id="d7389-127">Descripción</span><span class="sxs-lookup"><span data-stu-id="d7389-127">Description</span></span> |<span data-ttu-id="d7389-128">Débito</span><span class="sxs-lookup"><span data-stu-id="d7389-128">Debit</span></span> |<span data-ttu-id="d7389-129">Crédito</span><span class="sxs-lookup"><span data-stu-id="d7389-129">Credit</span></span> |
    |:--|:--|:--|:--|:--|
    |<span data-ttu-id="d7389-130">10/10/2019</span><span class="sxs-lookup"><span data-stu-id="d7389-130">10/10/2019</span></span> |<span data-ttu-id="d7389-131">Cuenta corriente</span><span class="sxs-lookup"><span data-stu-id="d7389-131">Checking</span></span> |<span data-ttu-id="d7389-132">Coho Vineyard</span><span class="sxs-lookup"><span data-stu-id="d7389-132">Coho Vineyard</span></span> |<span data-ttu-id="d7389-133">-20,05</span><span class="sxs-lookup"><span data-stu-id="d7389-133">-20.05</span></span> | |
    |<span data-ttu-id="d7389-134">11/10/2019</span><span class="sxs-lookup"><span data-stu-id="d7389-134">10/11/2019</span></span> |<span data-ttu-id="d7389-135">Crédito</span><span class="sxs-lookup"><span data-stu-id="d7389-135">Credit</span></span> |<span data-ttu-id="d7389-136">The Phone Company</span><span class="sxs-lookup"><span data-stu-id="d7389-136">The Phone Company</span></span> |<span data-ttu-id="d7389-137">99,95</span><span class="sxs-lookup"><span data-stu-id="d7389-137">99.95</span></span> | |
    |<span data-ttu-id="d7389-138">13/10/2019</span><span class="sxs-lookup"><span data-stu-id="d7389-138">10/13/2019</span></span> |<span data-ttu-id="d7389-139">Crédito</span><span class="sxs-lookup"><span data-stu-id="d7389-139">Credit</span></span> |<span data-ttu-id="d7389-140">Coho Vineyard</span><span class="sxs-lookup"><span data-stu-id="d7389-140">Coho Vineyard</span></span> |<span data-ttu-id="d7389-141">154,43</span><span class="sxs-lookup"><span data-stu-id="d7389-141">154.43</span></span> | |
    |<span data-ttu-id="d7389-142">15/10/2019</span><span class="sxs-lookup"><span data-stu-id="d7389-142">10/15/2019</span></span> |<span data-ttu-id="d7389-143">Cuenta corriente</span><span class="sxs-lookup"><span data-stu-id="d7389-143">Checking</span></span> |<span data-ttu-id="d7389-144">Depósito externo</span><span class="sxs-lookup"><span data-stu-id="d7389-144">External Deposit</span></span> | |<span data-ttu-id="d7389-145">1000</span><span class="sxs-lookup"><span data-stu-id="d7389-145">1000</span></span> |
    |<span data-ttu-id="d7389-146">20/10/2019</span><span class="sxs-lookup"><span data-stu-id="d7389-146">10/20/2019</span></span> |<span data-ttu-id="d7389-147">Crédito</span><span class="sxs-lookup"><span data-stu-id="d7389-147">Credit</span></span> |<span data-ttu-id="d7389-148">Coho Vineyard - Devolución</span><span class="sxs-lookup"><span data-stu-id="d7389-148">Coho Vineyard - Refund</span></span> | |<span data-ttu-id="d7389-149">- 35,45</span><span class="sxs-lookup"><span data-stu-id="d7389-149">-35.45</span></span> |
    |<span data-ttu-id="d7389-150">25/10/2019</span><span class="sxs-lookup"><span data-stu-id="d7389-150">10/25/2019</span></span> |<span data-ttu-id="d7389-151">Cuenta corriente</span><span class="sxs-lookup"><span data-stu-id="d7389-151">Checking</span></span> |<span data-ttu-id="d7389-152">Best For You Organics Company</span><span class="sxs-lookup"><span data-stu-id="d7389-152">Best For You Organics Company</span></span> | <span data-ttu-id="d7389-153">- 85,64</span><span class="sxs-lookup"><span data-stu-id="d7389-153">-85.64</span></span> | |
    |<span data-ttu-id="d7389-154">01/11/2019</span><span class="sxs-lookup"><span data-stu-id="d7389-154">11/01/2019</span></span> |<span data-ttu-id="d7389-155">Cuenta corriente</span><span class="sxs-lookup"><span data-stu-id="d7389-155">Checking</span></span> |<span data-ttu-id="d7389-156">Depósito externo</span><span class="sxs-lookup"><span data-stu-id="d7389-156">External Deposit</span></span> | |<span data-ttu-id="d7389-157">1000</span><span class="sxs-lookup"><span data-stu-id="d7389-157">1000</span></span> |

3. <span data-ttu-id="d7389-158">Abra **Todos los scripts** y seleccione **Nuevo script**.</span><span class="sxs-lookup"><span data-stu-id="d7389-158">Open **All Scripts** and select **New Script**.</span></span>
4. <span data-ttu-id="d7389-159">Limpiemos un poco el formato.</span><span class="sxs-lookup"><span data-stu-id="d7389-159">Let's clean up the formatting.</span></span> <span data-ttu-id="d7389-160">Este es un documento financiero, así que cambie el formato de número de las columnas **Débito** y **Crédito** para mostrar los valores como cantidades en euros.</span><span class="sxs-lookup"><span data-stu-id="d7389-160">This is a financial document, so let's change the number formatting in the **Debit** and **Credit** columns to show values as dollar amounts.</span></span> <span data-ttu-id="d7389-161">También hay que ajustar el ancho de columna a los datos.</span><span class="sxs-lookup"><span data-stu-id="d7389-161">Let's also fit the column width to the data.</span></span>

    <span data-ttu-id="d7389-162">Reemplace el contenido del script por el siguiente código:</span><span class="sxs-lookup"><span data-stu-id="d7389-162">Replace the script contents with the following code:</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
        // Get the current worksheet.
        let selectedSheet = workbook.getActiveWorksheet();

        // Format the range to display numerical dollar amounts.
        selectedSheet.getRange("D2:E8").setNumberFormat("$#,##0.00");

        // Fit the width of all the used columns to the data.
        selectedSheet.getUsedRange().getFormat().autofitColumns();
    }
    ```

5. <span data-ttu-id="d7389-163">Ahora, leamos un valor de una de las columnas de número.</span><span class="sxs-lookup"><span data-stu-id="d7389-163">Now let's read a value from one of the number columns.</span></span> <span data-ttu-id="d7389-164">Agregue el código siguiente al final del script (antes del `}` de cierre):</span><span class="sxs-lookup"><span data-stu-id="d7389-164">Add the following code to the end of the script (before the closing `}`):</span></span>

    ```TypeScript
    // Get the value of cell D2.
    let range = selectedSheet.getRange("D2");
    console.log(range.getValues());
    ```

6. <span data-ttu-id="d7389-165">Ejecute el script.</span><span class="sxs-lookup"><span data-stu-id="d7389-165">Run the script.</span></span>
7. <span data-ttu-id="d7389-166">Debe ver `[Array[1]]` en la consola.</span><span class="sxs-lookup"><span data-stu-id="d7389-166">You should see `[Array[1]]` in the console.</span></span> <span data-ttu-id="d7389-167">No es un número por que los rangos son matrices bidimensionales de datos.</span><span class="sxs-lookup"><span data-stu-id="d7389-167">This is not a number because ranges are two-dimensional arrays of data.</span></span> <span data-ttu-id="d7389-168">Este rango bidimensional se ha registrado en la consola directamente.</span><span class="sxs-lookup"><span data-stu-id="d7389-168">That two-dimensional range is being logged to the console directly.</span></span> <span data-ttu-id="d7389-169">Afortunadamente, el Editor de código le permite ver el contenido de la matriz.</span><span class="sxs-lookup"><span data-stu-id="d7389-169">Luckily, the Code Editor lets you see the contents of the array.</span></span>
8. <span data-ttu-id="d7389-170">Cuando se registra una matriz bidimensional en la consola, se agrupan los valores de columna en cada fila.</span><span class="sxs-lookup"><span data-stu-id="d7389-170">When a two-dimensional array is logged to the console, it groups column values under each row.</span></span> <span data-ttu-id="d7389-171">Expanda el registro de matriz seleccionando el triángulo azul.</span><span class="sxs-lookup"><span data-stu-id="d7389-171">Expand the array log by selecting the blue triangle.</span></span>
9. <span data-ttu-id="d7389-p110">Expanda el segundo nivel de la matriz seleccinando el triángulo azul recién descubierto. Ahora debería ver esto:</span><span class="sxs-lookup"><span data-stu-id="d7389-p110">Expand the second level of the array by selecting the newly revealed blue triangle. You should now see this:</span></span>

    :::image type="content" source="../images/tutorial-4.png" alt-text="El registro de consola mostrando el resultado &quot;-20,05&quot; anidado en dos matrices":::

## <a name="modify-the-value-of-a-cell"></a><span data-ttu-id="d7389-175">Cambiar el valor de una celda</span><span class="sxs-lookup"><span data-stu-id="d7389-175">Modify the value of a cell</span></span>

<span data-ttu-id="d7389-176">Ahora que podemos leer datos, vamos a usarlos para modificar el libro.</span><span class="sxs-lookup"><span data-stu-id="d7389-176">Now that we can read data, let's use that data to modify the workbook.</span></span> <span data-ttu-id="d7389-177">Haremos que el valor de la celda **D2** sea positivo con la función `Math.abs`.</span><span class="sxs-lookup"><span data-stu-id="d7389-177">We'll make the value of the cell **D2** positive with the `Math.abs` function.</span></span> <span data-ttu-id="d7389-178">El objeto [Math](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/math) contiene varias funciones a las que tienen acceso los scripts.</span><span class="sxs-lookup"><span data-stu-id="d7389-178">The [Math](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/math) object contains many functions to which your scripts have access.</span></span> <span data-ttu-id="d7389-179">Puede encontrar más información sobre `Math` y otros objetos integrados en [Usar objetos integrados de JavaScript en los scripts de Office](../develop/javascript-objects.md).</span><span class="sxs-lookup"><span data-stu-id="d7389-179">More information about `Math` and other built-in objects can be found at [Using built-in JavaScript objects in Office Scripts](../develop/javascript-objects.md).</span></span>

1. <span data-ttu-id="d7389-180">Usaremos métodos `getValue` y `setValue` para cambiar el valor de la celda.</span><span class="sxs-lookup"><span data-stu-id="d7389-180">We'll use `getValue` and `setValue` methods to change the value of the cell.</span></span> <span data-ttu-id="d7389-181">Estos métodos funcionan en una sola celda.</span><span class="sxs-lookup"><span data-stu-id="d7389-181">These methods work on a single cell.</span></span> <span data-ttu-id="d7389-182">Cuando trabaje con rangos de varias celdas, es mejor usar `getValues` y `setValues`.</span><span class="sxs-lookup"><span data-stu-id="d7389-182">When handling multi-cell ranges, you'll want to use `getValues` and `setValues`.</span></span> <span data-ttu-id="d7389-183">Agregue el siguiente código al final del script:</span><span class="sxs-lookup"><span data-stu-id="d7389-183">Add the following code to the end of the script:</span></span>

    ```TypeScript
    // Run the `Math.abs` function with the value at D2 and apply that value back to D2.
    let positiveValue = Math.abs(range.getValue() as number);
    range.setValue(positiveValue);
    ```

    > [!NOTE]
    > <span data-ttu-id="d7389-184">Se está [cambiando](https://www.typescripttutorial.net/typescript-tutorial/type-casting/) el valor devuelto de `range.getValue()` a un `number` con la palabra clave `as`.</span><span class="sxs-lookup"><span data-stu-id="d7389-184">We are [casting](https://www.typescripttutorial.net/typescript-tutorial/type-casting/) the returned value of `range.getValue()` to a `number` by using the `as` keyword.</span></span> <span data-ttu-id="d7389-185">Esto es necesario porque un rango puede estar conformado por cadenas, números o booleanos.</span><span class="sxs-lookup"><span data-stu-id="d7389-185">This is necessary because a range could be strings, numbers, or booleans.</span></span> <span data-ttu-id="d7389-186">En este caso, específicamente necesitamos un número.</span><span class="sxs-lookup"><span data-stu-id="d7389-186">In this instance, we explicitly need a number.</span></span>

2. <span data-ttu-id="d7389-187">El valor de la celda **D2** debería ahora ser positivo.</span><span class="sxs-lookup"><span data-stu-id="d7389-187">The value of cell **D2** should now be positive.</span></span>

## <a name="modify-the-values-of-a-column"></a><span data-ttu-id="d7389-188">Modificar los valores de una columna</span><span class="sxs-lookup"><span data-stu-id="d7389-188">Modify the values of a column</span></span>

<span data-ttu-id="d7389-189">Ahora que sabemos cómo leer y escribir en una sola celda, vamos a aplicar este conocimiento a todas las columnas **Débito** y **Crédito**.</span><span class="sxs-lookup"><span data-stu-id="d7389-189">Now that we know how to read and write to a single cell, let's generalize the script to work on the entire **Debit** and **Credit** columns.</span></span>

1. <span data-ttu-id="d7389-190">Quite el código que afecta a una sola celda (el código de valor absoluto anterior), para que el script tenga el siguiente aspecto:</span><span class="sxs-lookup"><span data-stu-id="d7389-190">Remove the code that affects only a single cell (the previous absolute value code), such that your script now looks like this:</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
        // Get the current worksheet.
        let selectedSheet = workbook.getActiveWorksheet();

        // Format the range to display numerical dollar amounts.
        selectedSheet.getRange("D2:E8").setNumberFormat("$#,##0.00");

        // Fit the width of all the used columns to the data.
        selectedSheet.getUsedRange().getFormat().autofitColumns();
    }
    ```

2. <span data-ttu-id="d7389-191">Agregue un bucle al final del script que itere las filas de las dos últimas columnas.</span><span class="sxs-lookup"><span data-stu-id="d7389-191">Add a loop to the end of the script that iterates through the rows in the last two columns.</span></span> <span data-ttu-id="d7389-192">En cada celda, el script establece el valor absoluto del valor actual iterado.</span><span class="sxs-lookup"><span data-stu-id="d7389-192">For each cell, the script sets the value to the current value's absolute value.</span></span>

    <span data-ttu-id="d7389-193">Tenga en cuenta que la matriz que define las ubicaciones de la celda está basada en cero.</span><span class="sxs-lookup"><span data-stu-id="d7389-193">Note that the array defining cell locations is zero-based.</span></span> <span data-ttu-id="d7389-194">Esto significa que la celda **A1** es `range[0][0]`.</span><span class="sxs-lookup"><span data-stu-id="d7389-194">That means cell **A1** is `range[0][0]`.</span></span>

    ```TypeScript
    // Get the values of the used range.
    let range = selectedSheet.getUsedRange();
    let rangeValues = range.getValues();

    // Iterate over the fourth and fifth columns and set their values to their absolute value.
    let rowCount = range.getRowCount();
    for (let i = 1; i < rowCount; i++) {
        // The column at index 3 is column "4" in the worksheet.
        if (rangeValues[i][3] != 0) {
            let positiveValue = Math.abs(rangeValues[i][3] as number);
            selectedSheet.getCell(i, 3).setValue(positiveValue);
        }

        // The column at index 4 is column "5" in the worksheet.
        if (rangeValues[i][4] != 0) {
            let positiveValue = Math.abs(rangeValues[i][4] as number);
            selectedSheet.getCell(i, 4).setValue(positiveValue);
        }
    }
    ```

    <span data-ttu-id="d7389-195">Esta parte del script realiza varias tareas importantes.</span><span class="sxs-lookup"><span data-stu-id="d7389-195">This portion of the script does several important tasks.</span></span> <span data-ttu-id="d7389-196">En primer lugar, obtiene los valores y cuenta las filas del rango usado.</span><span class="sxs-lookup"><span data-stu-id="d7389-196">First, it gets the values and row count of the used range.</span></span> <span data-ttu-id="d7389-197">Esto nos permite ver los valores y averiguar cuándo detenernos.</span><span class="sxs-lookup"><span data-stu-id="d7389-197">This lets us look at values and know when to stop.</span></span> <span data-ttu-id="d7389-198">En segundo lugar, itera el rango usado, verificando cada celda en las columnas **Débito** y **Crédito**.</span><span class="sxs-lookup"><span data-stu-id="d7389-198">Second, it iterates through the used range, checking each cell in the **Debit** or **Credit** columns.</span></span> <span data-ttu-id="d7389-199">Por último, si el valor de la celda no es 0, se reemplaza por su valor absoluto.</span><span class="sxs-lookup"><span data-stu-id="d7389-199">Finally, if the value in the cell is not 0, it is replaced by its absolute value.</span></span> <span data-ttu-id="d7389-200">Evitamos el cero para no cambiar las celdas en blanco.</span><span class="sxs-lookup"><span data-stu-id="d7389-200">We're avoiding zeroes so we can leave the blank cells as they were.</span></span>

3. <span data-ttu-id="d7389-201">Ejecute el script.</span><span class="sxs-lookup"><span data-stu-id="d7389-201">Run the script.</span></span>

    <span data-ttu-id="d7389-202">Ahora, su declaración bancaria debería tener el siguiente aspecto:</span><span class="sxs-lookup"><span data-stu-id="d7389-202">Your banking statement should now look like this:</span></span>

    :::image type="content" source="../images/tutorial-5.png" alt-text="Una hoja de cálculo donde se muestra un extracto bancario como tabla con formato que solo contiene valores positivos.":::

## <a name="next-steps"></a><span data-ttu-id="d7389-204">Pasos siguientes</span><span class="sxs-lookup"><span data-stu-id="d7389-204">Next steps</span></span>

<span data-ttu-id="d7389-205">Abra el Editor de código y pruebe algunos de nuestros [Ejemplos para scripts de Office en Excel en la Web](../resources/samples/excel-samples.md).</span><span class="sxs-lookup"><span data-stu-id="d7389-205">Open the Code Editor and try out some of our [Sample scripts for Office Scripts in Excel on the web](../resources/samples/excel-samples.md).</span></span> <span data-ttu-id="d7389-206">Para obtener más información sobre la creación de scripts de Office, consulte también [Fundamentos para scripts de Office en Excel en la Web](../develop/scripting-fundamentals.md).</span><span class="sxs-lookup"><span data-stu-id="d7389-206">You can also visit [Scripting Fundamentals for Office Scripts in Excel on the web](../develop/scripting-fundamentals.md) to learn more about creating Office Scripts.</span></span>

<span data-ttu-id="d7389-207">La siguiente serie de tutoriales de Scripts de Office se centra en el uso de Scripts de Office con Power Automate.</span><span class="sxs-lookup"><span data-stu-id="d7389-207">The next series of Office Scripts tutorials focus on using Office Scripts with Power Automate.</span></span> <span data-ttu-id="d7389-208">Obtenga más información sobre las ventajas de combinar las dos plataformas en [Ejecutar Scripts de Office con Power Automate](../develop/power-automate-integration.md) o consulte el tutorial [Llamar a scripts desde un flujo manual de Power Automate](excel-power-automate-manual.md) para crear un flujo de Power Automate que use un script de Office.</span><span class="sxs-lookup"><span data-stu-id="d7389-208">Learn more about the advantages combining the two platforms in [Run Office Scripts with Power Automate](../develop/power-automate-integration.md) or try the [Call scripts from a manual Power Automate flow](excel-power-automate-manual.md) tutorial to create a Power Automate flow that uses an Office Script.</span></span>
