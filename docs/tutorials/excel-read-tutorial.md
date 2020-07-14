---
title: Leer datos de libros con scripts de Office en Excel en la Web
description: Un tutorial de scripts de Office sobre cómo leer datos de libros y evaluarlos en el script.
ms.date: 04/23/2020
localization_priority: Priority
ms.openlocfilehash: 93204184d4b5947b2a67107b1fd73c178a73c32e
ms.sourcegitcommit: aec3c971c6640429f89b6bb99d2c95ea06725599
ms.translationtype: HT
ms.contentlocale: es-ES
ms.lasthandoff: 06/25/2020
ms.locfileid: "44878692"
---
# <a name="read-workbook-data-with-office-scripts-in-excel-on-the-web"></a><span data-ttu-id="5d760-103">Leer datos de libros con scripts de Office en Excel en la Web</span><span class="sxs-lookup"><span data-stu-id="5d760-103">Read workbook data with Office Scripts in Excel on the web</span></span>

<span data-ttu-id="5d760-104">Este tutorial le enseña a leer datos de un libro con un script de Office para Excel en la Web.</span><span class="sxs-lookup"><span data-stu-id="5d760-104">This tutorial teaches you how to read data from a workbook with an Office Script for Excel on the web.</span></span> <span data-ttu-id="5d760-105">A continuación, deberá modificar los datos leídos y volver a colocarlos en el libro.</span><span class="sxs-lookup"><span data-stu-id="5d760-105">You'll then edit the data you read and put it back in the workbook.</span></span>

> [!TIP]
> <span data-ttu-id="5d760-106">Si no está familiarizado con los scripts de Office, le recomendamos que empiece con el tutorial de [Guardar, editar y crear scripts de Office en Excel en la Web](excel-tutorial.md).</span><span class="sxs-lookup"><span data-stu-id="5d760-106">If you are new to Office Scripts, we recommend starting with the [Record, edit, and create Office Scripts in Excel on the web](excel-tutorial.md) tutorial.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="5d760-107">Requisitos previos</span><span class="sxs-lookup"><span data-stu-id="5d760-107">Prerequisites</span></span>

[!INCLUDE [Tutorial prerequisites](../includes/tutorial-prerequisites.md)]

> [!IMPORTANT]
> <span data-ttu-id="5d760-108">Este tutorial está diseñado para las personas con conocimientos a nivel intermedio de JavaScript o TypeScript.</span><span class="sxs-lookup"><span data-stu-id="5d760-108">This tutorial is intended for people with beginner to intermediate-level knowledge of JavaScript or TypeScript.</span></span> <span data-ttu-id="5d760-109">Si no está familiarizado con JavaScript, le recomendamos que revise el [Tutorial de JavaScript de Mozilla](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span><span class="sxs-lookup"><span data-stu-id="5d760-109">If you're new to JavaScript, we recommend reviewing the [Mozilla JavaScript tutorial](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span></span> <span data-ttu-id="5d760-110">Para obtener más información sobre el entorno de los scripts, visite [Scripts de Office en Excel en la Web](../overview/excel.md).</span><span class="sxs-lookup"><span data-stu-id="5d760-110">Visit [Office Scripts in Excel on the web](../overview/excel.md) to learn more about the script environment.</span></span>

## <a name="read-a-cell"></a><span data-ttu-id="5d760-111">Leer una celda</span><span class="sxs-lookup"><span data-stu-id="5d760-111">Read a cell</span></span>

<span data-ttu-id="5d760-112">Los scripts creados con la Grabadora de acciones solo pueden escribir información en el libro.</span><span class="sxs-lookup"><span data-stu-id="5d760-112">Scripts made with the Action Recorder can only write information to the workbook.</span></span> <span data-ttu-id="5d760-113">Con el Editor de código, puede además editar y escribir scripts que lean datos de un libro.</span><span class="sxs-lookup"><span data-stu-id="5d760-113">With the Code Editor, you can edit and make scripts that also read data from a workbook.</span></span>

<span data-ttu-id="5d760-114">Creemos un script que lea datos y actúe en función de lo que lee.</span><span class="sxs-lookup"><span data-stu-id="5d760-114">Let's make a script that reads data and acts based on what was read.</span></span> <span data-ttu-id="5d760-115">Vamos a trabajar con un ejemplo de extracto bancario.</span><span class="sxs-lookup"><span data-stu-id="5d760-115">We're going to work with a sample banking statement.</span></span> <span data-ttu-id="5d760-116">Este ejemplo es una declaración combinada de cuenta corriente y crédito.</span><span class="sxs-lookup"><span data-stu-id="5d760-116">This statement is a combined checking and credit statement.</span></span> <span data-ttu-id="5d760-117">Lamentablemente, los informes de saldo muestran los cambios de forma diferente.</span><span class="sxs-lookup"><span data-stu-id="5d760-117">Unfortunately, they report balance changes differently.</span></span> <span data-ttu-id="5d760-118">La declaración de cuenta corriente muestra los ingresos como crédito positivo y los costes como débito negativo.</span><span class="sxs-lookup"><span data-stu-id="5d760-118">The checking statement gives income as positive credit and costs as negative debit.</span></span> <span data-ttu-id="5d760-119">En cambio, la declaración de crédito funciona de manera contraria.</span><span class="sxs-lookup"><span data-stu-id="5d760-119">The credit statement does the opposite.</span></span>

<span data-ttu-id="5d760-120">En el resto del tutorial, armonizaremos los datos de ambos con un script.</span><span class="sxs-lookup"><span data-stu-id="5d760-120">Over the rest of the tutorial, we will normalize this data using a script.</span></span> <span data-ttu-id="5d760-121">En primer lugar, veamos cómo leer los datos del libro.</span><span class="sxs-lookup"><span data-stu-id="5d760-121">First, let's learn how to read data from the workbook.</span></span>

1. <span data-ttu-id="5d760-122">Cree una nueva hoja de cálculo en el libro que ha usado para el resto del tutorial.</span><span class="sxs-lookup"><span data-stu-id="5d760-122">Create a new worksheet in the workbook you've used for the rest of the tutorial.</span></span>
2. <span data-ttu-id="5d760-123">Copie los siguientes datos y péguelos en la nueva hoja de cálculo, comenzando por la celda **A1**.</span><span class="sxs-lookup"><span data-stu-id="5d760-123">Copy the following data and paste it into the new worksheet, starting at cell **A1**.</span></span>

    |<span data-ttu-id="5d760-124">Fecha</span><span class="sxs-lookup"><span data-stu-id="5d760-124">Date</span></span> |<span data-ttu-id="5d760-125">Cuenta</span><span class="sxs-lookup"><span data-stu-id="5d760-125">Account</span></span> |<span data-ttu-id="5d760-126">Descripción</span><span class="sxs-lookup"><span data-stu-id="5d760-126">Description</span></span> |<span data-ttu-id="5d760-127">Débito</span><span class="sxs-lookup"><span data-stu-id="5d760-127">Debit</span></span> |<span data-ttu-id="5d760-128">Crédito</span><span class="sxs-lookup"><span data-stu-id="5d760-128">Credit</span></span> |
    |:--|:--|:--|:--|:--|
    |<span data-ttu-id="5d760-129">10/10/2019</span><span class="sxs-lookup"><span data-stu-id="5d760-129">10/10/2019</span></span> |<span data-ttu-id="5d760-130">Cuenta corriente</span><span class="sxs-lookup"><span data-stu-id="5d760-130">Checking</span></span> |<span data-ttu-id="5d760-131">Coho Vineyard</span><span class="sxs-lookup"><span data-stu-id="5d760-131">Coho Vineyard</span></span> |<span data-ttu-id="5d760-132">-20,05</span><span class="sxs-lookup"><span data-stu-id="5d760-132">-20.05</span></span> | |
    |<span data-ttu-id="5d760-133">11/10/2019</span><span class="sxs-lookup"><span data-stu-id="5d760-133">10/11/2019</span></span> |<span data-ttu-id="5d760-134">Crédito</span><span class="sxs-lookup"><span data-stu-id="5d760-134">Credit</span></span> |<span data-ttu-id="5d760-135">The Phone Company</span><span class="sxs-lookup"><span data-stu-id="5d760-135">The Phone Company</span></span> |<span data-ttu-id="5d760-136">99,95</span><span class="sxs-lookup"><span data-stu-id="5d760-136">99.95</span></span> | |
    |<span data-ttu-id="5d760-137">13/10/2019</span><span class="sxs-lookup"><span data-stu-id="5d760-137">10/13/2019</span></span> |<span data-ttu-id="5d760-138">Crédito</span><span class="sxs-lookup"><span data-stu-id="5d760-138">Credit</span></span> |<span data-ttu-id="5d760-139">Coho Vineyard</span><span class="sxs-lookup"><span data-stu-id="5d760-139">Coho Vineyard</span></span> |<span data-ttu-id="5d760-140">154,43</span><span class="sxs-lookup"><span data-stu-id="5d760-140">154.43</span></span> | |
    |<span data-ttu-id="5d760-141">15/10/2019</span><span class="sxs-lookup"><span data-stu-id="5d760-141">10/15/2019</span></span> |<span data-ttu-id="5d760-142">Cuenta corriente</span><span class="sxs-lookup"><span data-stu-id="5d760-142">Checking</span></span> |<span data-ttu-id="5d760-143">Depósito externo</span><span class="sxs-lookup"><span data-stu-id="5d760-143">External Deposit</span></span> | |<span data-ttu-id="5d760-144">1000</span><span class="sxs-lookup"><span data-stu-id="5d760-144">1000</span></span> |
    |<span data-ttu-id="5d760-145">20/10/2019</span><span class="sxs-lookup"><span data-stu-id="5d760-145">10/20/2019</span></span> |<span data-ttu-id="5d760-146">Crédito</span><span class="sxs-lookup"><span data-stu-id="5d760-146">Credit</span></span> |<span data-ttu-id="5d760-147">Coho Vineyard - Devolución</span><span class="sxs-lookup"><span data-stu-id="5d760-147">Coho Vineyard - Refund</span></span> | |<span data-ttu-id="5d760-148">- 35,45</span><span class="sxs-lookup"><span data-stu-id="5d760-148">-35.45</span></span> |
    |<span data-ttu-id="5d760-149">25/10/2019</span><span class="sxs-lookup"><span data-stu-id="5d760-149">10/25/2019</span></span> |<span data-ttu-id="5d760-150">Cuenta corriente</span><span class="sxs-lookup"><span data-stu-id="5d760-150">Checking</span></span> |<span data-ttu-id="5d760-151">Best For You Organics Company</span><span class="sxs-lookup"><span data-stu-id="5d760-151">Best For You Organics Company</span></span> | <span data-ttu-id="5d760-152">- 85,64</span><span class="sxs-lookup"><span data-stu-id="5d760-152">-85.64</span></span> | |
    |<span data-ttu-id="5d760-153">01/11/2019</span><span class="sxs-lookup"><span data-stu-id="5d760-153">11/01/2019</span></span> |<span data-ttu-id="5d760-154">Cuenta corriente</span><span class="sxs-lookup"><span data-stu-id="5d760-154">Checking</span></span> |<span data-ttu-id="5d760-155">Depósito externo</span><span class="sxs-lookup"><span data-stu-id="5d760-155">External Deposit</span></span> | |<span data-ttu-id="5d760-156">1000</span><span class="sxs-lookup"><span data-stu-id="5d760-156">1000</span></span> |

3. <span data-ttu-id="5d760-157">Abra el **Editor de código** y seleccione **Nuevo script**.</span><span class="sxs-lookup"><span data-stu-id="5d760-157">Open the **Code Editor** and select **New Script**.</span></span>
4. <span data-ttu-id="5d760-158">Limpiemos un poco el formato.</span><span class="sxs-lookup"><span data-stu-id="5d760-158">Let's clean up the formatting.</span></span> <span data-ttu-id="5d760-159">Este es un documento financiero, así que cambie el formato de número de las columnas **Débito** y **Crédito** para mostrar los valores como cantidades en euros.</span><span class="sxs-lookup"><span data-stu-id="5d760-159">This is a financial document, so let's change the number formatting in the **Debit** and **Credit** columns to show values as dollar amounts.</span></span> <span data-ttu-id="5d760-160">También hay que ajustar el ancho de columna a los datos.</span><span class="sxs-lookup"><span data-stu-id="5d760-160">Let's also fit the column width to the data.</span></span>

    <span data-ttu-id="5d760-161">Reemplace el contenido del script por el siguiente código:</span><span class="sxs-lookup"><span data-stu-id="5d760-161">Replace the script contents with the following code:</span></span>

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

5. <span data-ttu-id="5d760-162">Ahora, leamos un valor de una de las columnas de número.</span><span class="sxs-lookup"><span data-stu-id="5d760-162">Now let's read a value from one of the number columns.</span></span> <span data-ttu-id="5d760-163">Agregue el código siguiente al final del script (antes del `}` de cierre):</span><span class="sxs-lookup"><span data-stu-id="5d760-163">Add the following code to the end of the script (before the closing `}`):</span></span>

    ```TypeScript
    // Get the value of cell D2.
    let range = selectedSheet.getRange("D2");
    console.log(range.getValues());
    ```

6. <span data-ttu-id="5d760-164">Ejecute el script.</span><span class="sxs-lookup"><span data-stu-id="5d760-164">Run the script.</span></span>
7. <span data-ttu-id="5d760-165">Abra la consola.</span><span class="sxs-lookup"><span data-stu-id="5d760-165">Open the console.</span></span> <span data-ttu-id="5d760-166">Vaya al menú de **Puntos suspensivos** y presione **Registros...**.</span><span class="sxs-lookup"><span data-stu-id="5d760-166">Go to the **Ellipses** menu and press **Logs...**.</span></span>
8. <span data-ttu-id="5d760-167">Debe ver `[Array[1]]` en la consola.</span><span class="sxs-lookup"><span data-stu-id="5d760-167">You should see `[Array[1]]` in the console.</span></span> <span data-ttu-id="5d760-168">No es un número por que los rangos son matrices bidimensionales de datos.</span><span class="sxs-lookup"><span data-stu-id="5d760-168">This is not a number because ranges are two-dimensional arrays of data.</span></span> <span data-ttu-id="5d760-169">Este rango bidimensional se ha registrado en la consola directamente.</span><span class="sxs-lookup"><span data-stu-id="5d760-169">That two-dimensional range is being logged to the console directly.</span></span> <span data-ttu-id="5d760-170">Afortunadamente, el Editor de código le permite ver el contenido de la matriz.</span><span class="sxs-lookup"><span data-stu-id="5d760-170">Luckily, the Code Editor does let you see the contents of the array.</span></span>
9. <span data-ttu-id="5d760-171">Cuando se registra una matriz bidimensional en la consola, se agrupan los valores de columna en cada fila.</span><span class="sxs-lookup"><span data-stu-id="5d760-171">When a two-dimensional array is logged to the console, it groups column values under each row.</span></span> <span data-ttu-id="5d760-172">Expanda el registro de matriz pulsando en el triángulo azul.</span><span class="sxs-lookup"><span data-stu-id="5d760-172">Expand the array log by pressing the blue triangle.</span></span>
10. <span data-ttu-id="5d760-173">Expanda el segundo nivel de la matriz pulsando en el triángulo azul que ha descubierto recientemente.</span><span class="sxs-lookup"><span data-stu-id="5d760-173">Expand the second level of the array by pressing the newly revealed blue triangle.</span></span> <span data-ttu-id="5d760-174">Ahora debería ver lo siguiente:</span><span class="sxs-lookup"><span data-stu-id="5d760-174">You should now see this:</span></span>

    ![El registro de consola mostrando el resultado "-20,05" anidado en dos matrices.](../images/tutorial-4.png)

## <a name="modify-the-value-of-a-cell"></a><span data-ttu-id="5d760-176">Cambiar el valor de una celda</span><span class="sxs-lookup"><span data-stu-id="5d760-176">Modify the value of a cell</span></span>

<span data-ttu-id="5d760-177">Ahora que podemos leer datos, vamos a usarlos para modificar el libro.</span><span class="sxs-lookup"><span data-stu-id="5d760-177">Now that we can read data, let's use that data to modify the workbook.</span></span> <span data-ttu-id="5d760-178">Haremos que el valor de la celda **D2** sea positivo con la función `Math.abs`.</span><span class="sxs-lookup"><span data-stu-id="5d760-178">We'll make the value of the cell **D2** positive with the `Math.abs` function.</span></span> <span data-ttu-id="5d760-179">El objeto [Math](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/math) contiene varias funciones a las que tienen acceso los scripts.</span><span class="sxs-lookup"><span data-stu-id="5d760-179">The [Math](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/math) object contains many functions to which your scripts have access.</span></span> <span data-ttu-id="5d760-180">Puede encontrar más información sobre `Math` y otros objetos integrados en [Usar objetos integrados de JavaScript en los scripts de Office](../develop/javascript-objects.md).</span><span class="sxs-lookup"><span data-stu-id="5d760-180">More information about `Math` and other built-in objects can be found at [Using built-in JavaScript objects in Office Scripts](../develop/javascript-objects.md).</span></span>

1. <span data-ttu-id="5d760-181">Agregue el siguiente código al final del script:</span><span class="sxs-lookup"><span data-stu-id="5d760-181">Add the following code to the end of the script:</span></span>

    ```TypeScript
        // Run the `Math.abs` function with the value at D2 and apply that value back to D2.
    let positiveValue = Math.abs(range.getValue());
    range.setValue(positiveValue);
    ```

    <span data-ttu-id="5d760-182">Tenga en cuenta que estamos usando `getValue` y `setValue`.</span><span class="sxs-lookup"><span data-stu-id="5d760-182">Note that we're using `getValue` and `setValue`.</span></span> <span data-ttu-id="5d760-183">Estos métodos funcionan en una sola celda.</span><span class="sxs-lookup"><span data-stu-id="5d760-183">These methods work on a single cell.</span></span> <span data-ttu-id="5d760-184">Cuando trabaje con rangos de varias celdas, es mejor usar `getValues` y `setValues`.</span><span class="sxs-lookup"><span data-stu-id="5d760-184">When handling multi-cell ranges, you'll want to use `getValues` and `setValues`.</span></span>

2. <span data-ttu-id="5d760-185">El valor de la celda **D2** debería ahora ser positivo.</span><span class="sxs-lookup"><span data-stu-id="5d760-185">The value of cell **D2** should now be positive.</span></span>

## <a name="modify-the-values-of-a-column"></a><span data-ttu-id="5d760-186">Modificar los valores de una columna</span><span class="sxs-lookup"><span data-stu-id="5d760-186">Modify the values of a column</span></span>

<span data-ttu-id="5d760-187">Ahora que sabemos cómo leer y escribir en una sola celda, vamos a aplicar este conocimiento a todas las columnas **Débito** y **Crédito**.</span><span class="sxs-lookup"><span data-stu-id="5d760-187">Now that we know how to read and write to a single cell, let's generalize the script to work on the entire **Debit** and **Credit** columns.</span></span>

1. <span data-ttu-id="5d760-188">Quite el código que afecta a una sola celda (el código de valor absoluto anterior), para que el script tenga el siguiente aspecto:</span><span class="sxs-lookup"><span data-stu-id="5d760-188">Remove the code that affects only a single cell (the previous absolute value code), such that your script now looks like this:</span></span>

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

2. <span data-ttu-id="5d760-189">Agregue un bucle al final del script que itere las filas de las dos últimas columnas.</span><span class="sxs-lookup"><span data-stu-id="5d760-189">Add a loop to the end of the script that iterates through the rows in the last two columns.</span></span> <span data-ttu-id="5d760-190">En cada celda, el script establece el valor absoluto del valor actual iterado.</span><span class="sxs-lookup"><span data-stu-id="5d760-190">For each cell, the script sets the value to the current value's absolute value.</span></span>

    <span data-ttu-id="5d760-191">Tenga en cuenta que la matriz que define las ubicaciones de la celda está basada en cero.</span><span class="sxs-lookup"><span data-stu-id="5d760-191">Note that the array defining cell locations is zero-based.</span></span> <span data-ttu-id="5d760-192">Esto significa que la celda **A1** es `range[0][0]`.</span><span class="sxs-lookup"><span data-stu-id="5d760-192">That means cell **A1** is `range[0][0]`.</span></span>

    ```TypeScript
    // Get the values of the used range.
    let range = selectedSheet.getUsedRange();
    let rangeValues = range.getValues();

    // Iterate over the fourth and fifth columns and set their values to their absolute value.
    for (let i = 1; i < range.getRowCount(); i++) {
        // The column at index 3 is column "4" in the worksheet.
        if (rangeValues[i][3] != 0) {
            let positiveValue = Math.abs(rangeValues[i][3]);
            selectedSheet.getCell(i, 3).setValue(positiveValue);
        }

        // The column at index 4 is column "5" in the worksheet.
        if (rangeValues[i][4] != 0) {
            let positiveValue = Math.abs(rangeValues[i][4]);
            selectedSheet.getCell(i, 4).setValue(positiveValue);
        }
    }
    ```

    <span data-ttu-id="5d760-193">Esta parte del script realiza varias tareas importantes.</span><span class="sxs-lookup"><span data-stu-id="5d760-193">This portion of the script does several important tasks.</span></span> <span data-ttu-id="5d760-194">En primer lugar, obtiene los valores y cuenta las filas del rango usado.</span><span class="sxs-lookup"><span data-stu-id="5d760-194">First, it gets the values and row count of the used range.</span></span> <span data-ttu-id="5d760-195">Esto nos permite ver los valores y averiguar cuándo detenernos.</span><span class="sxs-lookup"><span data-stu-id="5d760-195">This lets us look at values and know when to stop.</span></span> <span data-ttu-id="5d760-196">En segundo lugar, itera el rango usado, verificando cada celda en las columnas **Débito** y **Crédito**.</span><span class="sxs-lookup"><span data-stu-id="5d760-196">Second, it iterates through the used range, checking each cell in the **Debit** or **Credit** columns.</span></span> <span data-ttu-id="5d760-197">Por último, si el valor de la celda no es 0, se reemplaza por su valor absoluto.</span><span class="sxs-lookup"><span data-stu-id="5d760-197">Finally, if the value in the cell is not 0, it is replaced by its absolute value.</span></span> <span data-ttu-id="5d760-198">Evitamos el cero para no cambiar las celdas en blanco.</span><span class="sxs-lookup"><span data-stu-id="5d760-198">We're avoiding zeroes so we can leave the blank cells as they were.</span></span>

3. <span data-ttu-id="5d760-199">Ejecute el script.</span><span class="sxs-lookup"><span data-stu-id="5d760-199">Run the script.</span></span>

    <span data-ttu-id="5d760-200">Ahora, su declaración bancaria debería tener el siguiente aspecto:</span><span class="sxs-lookup"><span data-stu-id="5d760-200">Your banking statement should now look like this:</span></span>

    ![Declaración bancaria como tabla con formato que solo contiene valores positivos.](../images/tutorial-5.png)

## <a name="next-steps"></a><span data-ttu-id="5d760-202">Pasos siguientes</span><span class="sxs-lookup"><span data-stu-id="5d760-202">Next steps</span></span>

<span data-ttu-id="5d760-203">Abra el Editor de código y pruebe algunos de nuestros [Ejemplos para scripts de Office en Excel en la Web](../resources/excel-samples.md).</span><span class="sxs-lookup"><span data-stu-id="5d760-203">Open the Code Editor and try out some of our [Sample scripts for Office Scripts in Excel on the web](../resources/excel-samples.md).</span></span> <span data-ttu-id="5d760-204">Para obtener más información sobre la creación de scripts de Office, consulte también [Fundamentos para scripts de Office en Excel en la Web](../develop/scripting-fundamentals.md).</span><span class="sxs-lookup"><span data-stu-id="5d760-204">You can also visit [Scripting Fundamentals for Office Scripts in Excel on the web](../develop/scripting-fundamentals.md) to learn more about creating Office Scripts.</span></span>
