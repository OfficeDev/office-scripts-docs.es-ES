---
title: Grabar, editar y crear scripts de Office en Excel en la Web
description: Un tutorial sobre los conceptos básicos de scripts de Office que incluye la grabación de scripts en la Grabadora de acciones y la escritura de datos en un libro.
ms.date: 05/23/2021
localization_priority: Priority
ms.openlocfilehash: f9290a3d7dc723b72bfb4913171052472f2361c5
ms.sourcegitcommit: 4693c8f79428ec74695328275703af0ba1bfea8f
ms.translationtype: HT
ms.contentlocale: es-ES
ms.lasthandoff: 06/23/2021
ms.locfileid: "53074672"
---
# <a name="record-edit-and-create-office-scripts-in-excel-on-the-web"></a><span data-ttu-id="58e91-103">Grabar, editar y crear scripts de Office en Excel en la Web</span><span class="sxs-lookup"><span data-stu-id="58e91-103">Record, edit, and create Office Scripts in Excel on the web</span></span>

<span data-ttu-id="58e91-104">En este tutorial se le enseñan los conceptos básicos de la grabación, la edición y la escritura de un script de Office para Excel en la Web.</span><span class="sxs-lookup"><span data-stu-id="58e91-104">This tutorial teaches you the basics of recording, editing, and writing an Office Script for Excel on the web.</span></span> <span data-ttu-id="58e91-105">Va a grabar un script que dé formato a una hoja de cálculo con registros de ventas.</span><span class="sxs-lookup"><span data-stu-id="58e91-105">You'll record a script that applies some formatting to a sales record worksheet.</span></span> <span data-ttu-id="58e91-106">A continuación, editará el script grabado para dar más formato, crear una tabla y ordenar la tabla.</span><span class="sxs-lookup"><span data-stu-id="58e91-106">You'll then edit the recorded script to apply more formatting, create a table, and sort that table.</span></span> <span data-ttu-id="58e91-107">Este método de grabar y luego editar le permite ver el código que resulta de las acciones que ha realizado en Excel.</span><span class="sxs-lookup"><span data-stu-id="58e91-107">This record-then-edit pattern is an important tool to see what your Excel actions look like as code.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="58e91-108">Requisitos previos</span><span class="sxs-lookup"><span data-stu-id="58e91-108">Prerequisites</span></span>

[!INCLUDE [Tutorial prerequisites](../includes/tutorial-prerequisites.md)]

> [!IMPORTANT]
> <span data-ttu-id="58e91-109">Este tutorial está diseñado para las personas con conocimientos de nivel intermedio de JavaScript o TypeScript.</span><span class="sxs-lookup"><span data-stu-id="58e91-109">This tutorial is intended for people with beginner to intermediate-level knowledge of JavaScript or TypeScript.</span></span> <span data-ttu-id="58e91-110">Si no está familiarizado con JavaScript, le recomendamos que comience con el [Tutorial de JavaScript de Mozilla](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span><span class="sxs-lookup"><span data-stu-id="58e91-110">If you're new to JavaScript, we recommend starting with the [Mozilla JavaScript tutorial](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span></span> <span data-ttu-id="58e91-111">Para obtener más información sobre el entorno de los scripts, visite [Entorno del Editor de código de Scripts de Office](../overview/code-editor-environment.md).</span><span class="sxs-lookup"><span data-stu-id="58e91-111">Visit [Office Scripts Code Editor environment](../overview/code-editor-environment.md) to learn more about the script environment.</span></span>

## <a name="add-data-and-record-a-basic-script"></a><span data-ttu-id="58e91-112">Agregar datos y grabar un script básico</span><span class="sxs-lookup"><span data-stu-id="58e91-112">Add data and record a basic script</span></span>

<span data-ttu-id="58e91-113">En primer lugar, necesitaremos algunos datos y un pequeño script inicial.</span><span class="sxs-lookup"><span data-stu-id="58e91-113">First, we'll need some data and a small starting script.</span></span>

1. <span data-ttu-id="58e91-114">Cree un libro nuevo en Excel para la Web.</span><span class="sxs-lookup"><span data-stu-id="58e91-114">Create a new workbook in Excel for the Web.</span></span>
2. <span data-ttu-id="58e91-115">Copie los siguientes datos de ventas de frutas y péguelos en la hoja de cálculo, comenzando por la celda **A1**.</span><span class="sxs-lookup"><span data-stu-id="58e91-115">Copy the following fruit sales data and paste it into the worksheet, starting at cell **A1**.</span></span>

    |<span data-ttu-id="58e91-116">Fruta</span><span class="sxs-lookup"><span data-stu-id="58e91-116">Fruit</span></span> |<span data-ttu-id="58e91-117">2018</span><span class="sxs-lookup"><span data-stu-id="58e91-117">2018</span></span> |<span data-ttu-id="58e91-118">2019</span><span class="sxs-lookup"><span data-stu-id="58e91-118">2019</span></span> |
    |:---|:---|:---|
    |<span data-ttu-id="58e91-119">Naranjas</span><span class="sxs-lookup"><span data-stu-id="58e91-119">Oranges</span></span> |<span data-ttu-id="58e91-120">1000</span><span class="sxs-lookup"><span data-stu-id="58e91-120">1000</span></span> |<span data-ttu-id="58e91-121">1200</span><span class="sxs-lookup"><span data-stu-id="58e91-121">1200</span></span> |
    |<span data-ttu-id="58e91-122">Limones</span><span class="sxs-lookup"><span data-stu-id="58e91-122">Lemons</span></span> |<span data-ttu-id="58e91-123">800</span><span class="sxs-lookup"><span data-stu-id="58e91-123">800</span></span> |<span data-ttu-id="58e91-124">900</span><span class="sxs-lookup"><span data-stu-id="58e91-124">900</span></span> |
    |<span data-ttu-id="58e91-125">Limas</span><span class="sxs-lookup"><span data-stu-id="58e91-125">Limes</span></span> |<span data-ttu-id="58e91-126">600</span><span class="sxs-lookup"><span data-stu-id="58e91-126">600</span></span> |<span data-ttu-id="58e91-127">500</span><span class="sxs-lookup"><span data-stu-id="58e91-127">500</span></span> |
    |<span data-ttu-id="58e91-128">Pomelos</span><span class="sxs-lookup"><span data-stu-id="58e91-128">Grapefruits</span></span> |<span data-ttu-id="58e91-129">900</span><span class="sxs-lookup"><span data-stu-id="58e91-129">900</span></span> |<span data-ttu-id="58e91-130">700</span><span class="sxs-lookup"><span data-stu-id="58e91-130">700</span></span> |

3. <span data-ttu-id="58e91-131">Abra la pestaña **Automatizar**. Si no ve la pestaña **Automatizar**, presione la flecha desplegable para comprobar el desbordamiento de la cinta de opciones.</span><span class="sxs-lookup"><span data-stu-id="58e91-131">Open the **Automate** tab. If you don't see the **Automate** tab, check the ribbon overflow by pressing the drop-down arrow.</span></span> <span data-ttu-id="58e91-132">Si aún no aparece, siga los consejos del artículo [Solución de problemas de Scripts de Office](../testing/troubleshooting.md#automate-tab-not-appearing-or-office-scripts-unavailable).</span><span class="sxs-lookup"><span data-stu-id="58e91-132">If it's still not there, follow the advice in the article [Troubleshoot Office Scripts](../testing/troubleshooting.md#automate-tab-not-appearing-or-office-scripts-unavailable).</span></span>
4. <span data-ttu-id="58e91-133">Haga clic en el botón **Guardar acciones**.</span><span class="sxs-lookup"><span data-stu-id="58e91-133">Press the **Record Actions** button.</span></span>
5. <span data-ttu-id="58e91-134">Seleccione las celdas **A2:C2** (la fila "Naranjas") y configure el color de relleno como naranja.</span><span class="sxs-lookup"><span data-stu-id="58e91-134">Select cells **A2:C2** (the "Oranges" row) and set the fill color to orange.</span></span>
6. <span data-ttu-id="58e91-135">Detenga la grabación pulsando el botón **Detener**.</span><span class="sxs-lookup"><span data-stu-id="58e91-135">Stop the recording by pressing the **Stop** button.</span></span>

    <span data-ttu-id="58e91-136">La hoja de cálculo debe tener este aspecto (no se preocupe si el color es diferente):</span><span class="sxs-lookup"><span data-stu-id="58e91-136">Your worksheet should look like this (don't worry if the color is different):</span></span>

    :::image type="content" source="../images/tutorial-1.png" alt-text="Una hoja de cálculo donde se muestra la fila de datos de ventas de fruta con la fila que contiene «Naranjas» resaltada en color naranja.":::

## <a name="edit-an-existing-script"></a><span data-ttu-id="58e91-138">Editar un script existente</span><span class="sxs-lookup"><span data-stu-id="58e91-138">Edit an existing script</span></span>

<span data-ttu-id="58e91-139">El script anterior pinta la fila "Naranja" de color naranja.</span><span class="sxs-lookup"><span data-stu-id="58e91-139">The previous script colored the "Oranges" row to be orange.</span></span> <span data-ttu-id="58e91-140">Ahora, agreguemos una fila amarilla a "Limones".</span><span class="sxs-lookup"><span data-stu-id="58e91-140">Let's add a yellow row for the "Lemons".</span></span>

1. <span data-ttu-id="58e91-141">En el panel **Detalles** ya abierto, presione el botón **Editar**.</span><span class="sxs-lookup"><span data-stu-id="58e91-141">From the now-open **Details** pane, press the **Edit** button.</span></span>
2. <span data-ttu-id="58e91-142">Debería ver algo parecido a este código:</span><span class="sxs-lookup"><span data-stu-id="58e91-142">You should see something similar to this code:</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
      // Set fill color to FFC000 for range Sheet1!A2:C2
      let selectedSheet = workbook.getActiveWorksheet();
      selectedSheet.getRange("A2:C2").getFormat().getFill().setColor("FFC000");
    }
    ```

    <span data-ttu-id="58e91-143">Este código obtiene la hoja de cálculo actual del libro.</span><span class="sxs-lookup"><span data-stu-id="58e91-143">This code gets the current worksheet from the workbook.</span></span> <span data-ttu-id="58e91-144">Después, establece el color de relleno del rango **A2:C2**.</span><span class="sxs-lookup"><span data-stu-id="58e91-144">Then, it sets the fill color of the range **A2:C2**.</span></span>

    <span data-ttu-id="58e91-145">Los rangos son una parte fundamental de las secuencias de comandos de Office en Excel en la Web.</span><span class="sxs-lookup"><span data-stu-id="58e91-145">Ranges are a fundamental part of Office Scripts in Excel on the web.</span></span> <span data-ttu-id="58e91-146">Un rango es un bloque de celdas contiguo y rectangular que contiene valores, fórmulas y formatos.</span><span class="sxs-lookup"><span data-stu-id="58e91-146">A range is a contiguous, rectangular block of cells that contains values, formula, and formatting.</span></span> <span data-ttu-id="58e91-147">Constituyen la estructura básica de las celdas y se usan para realizar la mayoría de las tareas de scripts.</span><span class="sxs-lookup"><span data-stu-id="58e91-147">They are the basic structure of cells through which you'll perform most of your scripting tasks.</span></span>

3. <span data-ttu-id="58e91-148">Agregue la línea siguiente al final del script (entre el lugar en el que se establece el `color` y aparece el `}` de cierre):</span><span class="sxs-lookup"><span data-stu-id="58e91-148">Add the following line to the end of the script (between where the `color` is set and the closing `}`):</span></span>

    ```TypeScript
    selectedSheet.getRange("A3:C3").getFormat().getFill().setColor("yellow");
    ```

4. <span data-ttu-id="58e91-149">Para probar el script, presione **Ejecutar**.</span><span class="sxs-lookup"><span data-stu-id="58e91-149">Test the script by pressing **Run**.</span></span> <span data-ttu-id="58e91-150">El libro tendrá ahora el siguiente aspecto:</span><span class="sxs-lookup"><span data-stu-id="58e91-150">Your workbook should now look like this:</span></span>

    :::image type="content" source="../images/tutorial-2.png" alt-text=" Una hoja de cálculo donde se muestra una fila de datos de ventas con la fila «Naranjas» resaltada en naranja y la fila «Limones» resaltada en amarillo.":::

## <a name="create-a-table"></a><span data-ttu-id="58e91-152">Crear una tabla</span><span class="sxs-lookup"><span data-stu-id="58e91-152">Create a table</span></span>

<span data-ttu-id="58e91-153">Vamos a convertir estos datos de ventas de frutas en una tabla.</span><span class="sxs-lookup"><span data-stu-id="58e91-153">Let's convert this fruit sales data into a table.</span></span> <span data-ttu-id="58e91-154">Usaremos nuestro script para todo este proceso.</span><span class="sxs-lookup"><span data-stu-id="58e91-154">We'll use our script for the entire process.</span></span>

1. <span data-ttu-id="58e91-155">Agregue la línea siguiente al final del script (antes del `}` de cierre):</span><span class="sxs-lookup"><span data-stu-id="58e91-155">Add the following line to the end of the script (before the closing `}`):</span></span>

    ```TypeScript
    let table = selectedSheet.addTable("A1:C5", true);
    ```

2. <span data-ttu-id="58e91-156">Esa llamada devuelve un objeto de `Table`.</span><span class="sxs-lookup"><span data-stu-id="58e91-156">That call returns a `Table` object.</span></span> <span data-ttu-id="58e91-157">Vamos a usar la tabla para ordenar los datos.</span><span class="sxs-lookup"><span data-stu-id="58e91-157">Let's use that table to sort the data.</span></span> <span data-ttu-id="58e91-158">Ordenaremos los datos de menor a mayor en función de los valores de la columna "Frutas".</span><span class="sxs-lookup"><span data-stu-id="58e91-158">We'll sort the data in ascending order based on the values in the "Fruit" column.</span></span> <span data-ttu-id="58e91-159">Agregue la siguiente línea después de la creación de tabla:</span><span class="sxs-lookup"><span data-stu-id="58e91-159">Add the following line after the table creation:</span></span>

    ```TypeScript
    table.getSort().apply([{ key: 0, ascending: true }]);
    ```

    <span data-ttu-id="58e91-160">Su script debe tener este aspecto:</span><span class="sxs-lookup"><span data-stu-id="58e91-160">Your script should look like this:</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
        // Set fill color to FFC000 for range Sheet1!A2:C2
        let selectedSheet = workbook.getActiveWorksheet();
        selectedSheet.getRange("A2:C2").getFormat().getFill().setColor("FFC000");
        selectedSheet.getRange("A3:C3").getFormat().getFill().setColor("yellow");
        let table = selectedSheet.addTable("A1:C5", true);
        table.getSort().apply([{ key: 0, ascending: true }]);
    }
    ```

    <span data-ttu-id="58e91-161">Las tablas tienen un objeto `TableSort` al que se accede mediante el método `Table.getSort`.</span><span class="sxs-lookup"><span data-stu-id="58e91-161">Tables have a `TableSort` object, accessed through the `Table.getSort` method.</span></span> <span data-ttu-id="58e91-162">Puede aplicar un criterio de ordenación a ese objeto.</span><span class="sxs-lookup"><span data-stu-id="58e91-162">You can apply sorting criteria to that object.</span></span> <span data-ttu-id="58e91-163">El método `apply` acepta una matriz de objetos `SortField`.</span><span class="sxs-lookup"><span data-stu-id="58e91-163">The `apply` method takes in an array of `SortField` objects.</span></span> <span data-ttu-id="58e91-164">En este caso, solo tenemos un criterio de ordenación, por lo que solo usamos un `SortField`.</span><span class="sxs-lookup"><span data-stu-id="58e91-164">In this case, we only have one sorting criteria, so we only use one `SortField`.</span></span> <span data-ttu-id="58e91-165">`key: 0` establece los valores que definen la ordenación de la columna como "0" (que es la primera columna de la tabla **A** en este caso).</span><span class="sxs-lookup"><span data-stu-id="58e91-165">`key: 0` sets the column with the sort-defining values to "0" (which is the first column on the table, **A** in this case).</span></span> <span data-ttu-id="58e91-166">`ascending: true` ordena los datos de menor a mayor (en lugar de mayor a menor).</span><span class="sxs-lookup"><span data-stu-id="58e91-166">`ascending: true` sorts the data in ascending order (instead of descending order).</span></span>

3. <span data-ttu-id="58e91-p111">Ejecute el script. Debería ver una tabla como esta:</span><span class="sxs-lookup"><span data-stu-id="58e91-p111">Run the script. You should see a table like this:</span></span>

    :::image type="content" source="../images/tutorial-3.png" alt-text="Una hoja de cálculo donde se muestra la tabla de ventas de frutas ordenadas.":::

    > [!NOTE]
    > <span data-ttu-id="58e91-170">Si vuelve a ejecutar el script, se producirá un error.</span><span class="sxs-lookup"><span data-stu-id="58e91-170">If you re-run the script, you'll get an error.</span></span> <span data-ttu-id="58e91-171">Esto se debe a que no se puede crear una tabla encima de otra.</span><span class="sxs-lookup"><span data-stu-id="58e91-171">This is because you cannot create a table on top of another table.</span></span> <span data-ttu-id="58e91-172">Sin embargo, puede ejecutar el script en otra hoja de cálculo o en un libro.</span><span class="sxs-lookup"><span data-stu-id="58e91-172">However, you can run the script on a different worksheet or workbook.</span></span>

### <a name="re-run-the-script"></a><span data-ttu-id="58e91-173">Ejecute el script de nuevo</span><span class="sxs-lookup"><span data-stu-id="58e91-173">Re-run the script</span></span>

1. <span data-ttu-id="58e91-174">Crear una nueva hoja de cálculo en el libro actual.</span><span class="sxs-lookup"><span data-stu-id="58e91-174">Create a new worksheet in the current workbook.</span></span>
2. <span data-ttu-id="58e91-175">Copie los datos de frutas del principio del tutorial y péguelos en la nueva hoja de cálculo, comenzando en la celda **A1**.</span><span class="sxs-lookup"><span data-stu-id="58e91-175">Copy the fruit data from the beginning of the tutorial and paste it into the new worksheet, starting at cell **A1**.</span></span>
3. <span data-ttu-id="58e91-176">Ejecute el script.</span><span class="sxs-lookup"><span data-stu-id="58e91-176">Run the script.</span></span>

## <a name="next-steps"></a><span data-ttu-id="58e91-177">Pasos siguientes</span><span class="sxs-lookup"><span data-stu-id="58e91-177">Next steps</span></span>

<span data-ttu-id="58e91-178">Complete el tutorial [Leer datos de libros con scripts de Office en Excel en la Web](excel-read-tutorial.md).</span><span class="sxs-lookup"><span data-stu-id="58e91-178">Complete the [Read workbook data with Office Scripts in Excel on the web](excel-read-tutorial.md) tutorial.</span></span> <span data-ttu-id="58e91-179">En él aprenderá a leer datos de un libro con un script de Office.</span><span class="sxs-lookup"><span data-stu-id="58e91-179">It teaches you how to read data from a workbook with an Office Script.</span></span>
