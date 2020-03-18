---
title: Grabar, editar y crear scripts de Office en Excel en la Web
description: Un tutorial sobre los conceptos básicos de scripts de Office que incluye la grabación de scripts en la Grabadora de acciones y la escritura de datos en un libro.
ms.date: 01/27/2020
localization_priority: Priority
ms.openlocfilehash: 1971ff2ffd80554beb6ac561677ee3384f87ca81
ms.sourcegitcommit: b075eed5a6f275274fbbf6d62633219eac416f26
ms.translationtype: HT
ms.contentlocale: es-ES
ms.lasthandoff: 03/10/2020
ms.locfileid: "42700313"
---
# <a name="record-edit-and-create-office-scripts-in-excel-on-the-web"></a><span data-ttu-id="1a856-103">Grabar, editar y crear scripts de Office en Excel en la Web</span><span class="sxs-lookup"><span data-stu-id="1a856-103">Record, edit, and create Office Scripts in Excel on the web</span></span>

<span data-ttu-id="1a856-104">Este tutorial le enseñará los conceptos básicos de la grabación, la edición y la escritura de un script de Office para Excel en la Web.</span><span class="sxs-lookup"><span data-stu-id="1a856-104">This tutorial will teach you the basics of recording, editing, and writing an Office Script for Excel on the web.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="1a856-105">Requisitos previos</span><span class="sxs-lookup"><span data-stu-id="1a856-105">Prerequisites</span></span>

[!INCLUDE [Preview note](../includes/preview-note.md)]

<span data-ttu-id="1a856-106">Antes de iniciar este tutorial, necesitará acceder a los scripts de Office. Esto requiere lo siguiente:</span><span class="sxs-lookup"><span data-stu-id="1a856-106">Before starting this tutorial, you'll need access to Office Scripts, which requires the following:</span></span>

- <span data-ttu-id="1a856-107">[Excel en la Web](https://www.office.com/launch/excel)</span><span class="sxs-lookup"><span data-stu-id="1a856-107">[Excel on the web](https://www.office.com/launch/excel).</span></span>
- <span data-ttu-id="1a856-108">Pida a su administrador que [habilite los scripts de Office para su organización](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf), lo que agrega la barra de herramientas **Automatizar** a la cinta de opciones.</span><span class="sxs-lookup"><span data-stu-id="1a856-108">Ask your administrator to [enable Office Scripts for your organization](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf), which adds the **Automate** tab to the ribbon.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="1a856-109">Este tutorial está diseñado para las personas con conocimientos a nivel intermedio de JavaScript o TypeScript.</span><span class="sxs-lookup"><span data-stu-id="1a856-109">This tutorial is intended for people with beginner to intermediate-level knowledge of JavaScript or TypeScript.</span></span> <span data-ttu-id="1a856-110">Si no está familiarizado con JavaScript, le recomendamos que revise el [Tutorial de JavaScript de Mozilla](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span><span class="sxs-lookup"><span data-stu-id="1a856-110">If you're new to JavaScript, we recommend reviewing the [Mozilla JavaScript tutorial](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span></span> <span data-ttu-id="1a856-111">Para obtener más información sobre el entorno de los scripts, visite [Scripts de Office en Excel en la Web](../overview/excel.md).</span><span class="sxs-lookup"><span data-stu-id="1a856-111">Visit [Office Scripts in Excel on the web](../overview/excel.md) to learn more about the script environment.</span></span>

## <a name="add-data-and-record-a-basic-script"></a><span data-ttu-id="1a856-112">Agregar datos y grabar un script básico</span><span class="sxs-lookup"><span data-stu-id="1a856-112">Add data and record a basic script</span></span>

<span data-ttu-id="1a856-113">En primer lugar, necesitaremos algunos datos y un pequeño script inicial.</span><span class="sxs-lookup"><span data-stu-id="1a856-113">First, we'll need some data and a small starting script.</span></span>

1. <span data-ttu-id="1a856-114">Cree un libro nuevo en Excel para la Web.</span><span class="sxs-lookup"><span data-stu-id="1a856-114">Create a new workbook in Excel for the Web.</span></span>
2. <span data-ttu-id="1a856-115">Copie los siguientes datos de ventas de frutas y péguelos en la hoja de cálculo, comenzando por la celda **A1**.</span><span class="sxs-lookup"><span data-stu-id="1a856-115">Copy the following fruit sales data and paste it into the worksheet, starting at cell **A1**.</span></span>

    |<span data-ttu-id="1a856-116">Fruta</span><span class="sxs-lookup"><span data-stu-id="1a856-116">Fruit</span></span> |<span data-ttu-id="1a856-117">2018</span><span class="sxs-lookup"><span data-stu-id="1a856-117">2018</span></span> |<span data-ttu-id="1a856-118">2019</span><span class="sxs-lookup"><span data-stu-id="1a856-118">2019</span></span> |
    |:---|:---|:---|
    |<span data-ttu-id="1a856-119">Naranjas</span><span class="sxs-lookup"><span data-stu-id="1a856-119">Oranges</span></span> |<span data-ttu-id="1a856-120">1000</span><span class="sxs-lookup"><span data-stu-id="1a856-120">1000</span></span> |<span data-ttu-id="1a856-121">1200</span><span class="sxs-lookup"><span data-stu-id="1a856-121">1200</span></span> |
    |<span data-ttu-id="1a856-122">Limones</span><span class="sxs-lookup"><span data-stu-id="1a856-122">Lemons</span></span> |<span data-ttu-id="1a856-123">800</span><span class="sxs-lookup"><span data-stu-id="1a856-123">800</span></span> |<span data-ttu-id="1a856-124">900</span><span class="sxs-lookup"><span data-stu-id="1a856-124">900</span></span> |
    |<span data-ttu-id="1a856-125">Limas</span><span class="sxs-lookup"><span data-stu-id="1a856-125">Limes</span></span> |<span data-ttu-id="1a856-126">600</span><span class="sxs-lookup"><span data-stu-id="1a856-126">600</span></span> |<span data-ttu-id="1a856-127">500</span><span class="sxs-lookup"><span data-stu-id="1a856-127">500</span></span> |
    |<span data-ttu-id="1a856-128">Pomelos</span><span class="sxs-lookup"><span data-stu-id="1a856-128">Grapefruits</span></span> |<span data-ttu-id="1a856-129">900</span><span class="sxs-lookup"><span data-stu-id="1a856-129">900</span></span> |<span data-ttu-id="1a856-130">700</span><span class="sxs-lookup"><span data-stu-id="1a856-130">700</span></span> |

3. <span data-ttu-id="1a856-131">Abra la pestaña **Automatizar**. Si no ve la pestaña **Automatizar**, presione la flecha desplegable para comprobar el desbordamiento de la cinta de opciones.</span><span class="sxs-lookup"><span data-stu-id="1a856-131">Open the **Automate** tab. If you do not see the **Automate** tab, check the ribbon overflow by pressing the drop-down arrow.</span></span>
4. <span data-ttu-id="1a856-132">Haga clic en el botón **Guardar acciones**.</span><span class="sxs-lookup"><span data-stu-id="1a856-132">Press the **Record Actions** button.</span></span>
5. <span data-ttu-id="1a856-133">Seleccione las celdas **A2:C2** (la fila "Naranjas") y configure el color de relleno como naranja.</span><span class="sxs-lookup"><span data-stu-id="1a856-133">Select cells **A2:C2** (the "Oranges" row) and set the fill color to orange.</span></span>
6. <span data-ttu-id="1a856-134">Detenga la grabación pulsando el botón **Detener**.</span><span class="sxs-lookup"><span data-stu-id="1a856-134">Stop the recording by pressing the **Stop** button.</span></span>
7. <span data-ttu-id="1a856-135">Rellene el campo **Nombre del script** con un nombre que luego vaya a recordar.</span><span class="sxs-lookup"><span data-stu-id="1a856-135">Fill in the **Script Name** field with a memorable name.</span></span>
8. <span data-ttu-id="1a856-136">*Opcional:* rellenar el campo **Descripción** con una descripción relevante.</span><span class="sxs-lookup"><span data-stu-id="1a856-136">*Optional:* Fill in the **Description** field with a meaningful description.</span></span> <span data-ttu-id="1a856-137">Esto ofrece un contexto sobre lo que hace el script.</span><span class="sxs-lookup"><span data-stu-id="1a856-137">This is used to provide context as to what the script does.</span></span> <span data-ttu-id="1a856-138">Para el tutorial, puede usar "Distinguiendo las filas de una tabla con colores".</span><span class="sxs-lookup"><span data-stu-id="1a856-138">For the tutorial, you can use "Color-codes rows of a table".</span></span>

   > [!TIP]
   > <span data-ttu-id="1a856-139">Puede editar una descripción del script más adelante desde el panel de **Detalles del script**, que se encuentra en el menú **...** del Editor de código.</span><span class="sxs-lookup"><span data-stu-id="1a856-139">You can edit a script's description later from the **Script Details** pane, which is located under the Code Editor's **...** menu.</span></span>

9. <span data-ttu-id="1a856-140">Para guardar el script, presione el botón **Guardar**.</span><span class="sxs-lookup"><span data-stu-id="1a856-140">Save the script by pressing the **Save** button.</span></span>

    <span data-ttu-id="1a856-141">La hoja de cálculo debe tener este aspecto (no se preocupe si el color es diferente):</span><span class="sxs-lookup"><span data-stu-id="1a856-141">Your worksheet should look like this (don't worry if the color is different):</span></span>

    ![Una fila de datos de ventas de frutas con la fila "Naranjas" resaltada en naranja.](../images/tutorial-1.png)

## <a name="edit-an-existing-script"></a><span data-ttu-id="1a856-143">Editar un script existente</span><span class="sxs-lookup"><span data-stu-id="1a856-143">Edit an existing script</span></span>

<span data-ttu-id="1a856-144">El script anterior pinta la fila "Naranja" de color naranja.</span><span class="sxs-lookup"><span data-stu-id="1a856-144">The previous script colored the "Oranges" row to be orange.</span></span> <span data-ttu-id="1a856-145">Ahora, agreguemos una fila amarilla a "Limones".</span><span class="sxs-lookup"><span data-stu-id="1a856-145">Let's add a yellow row for the "Lemons".</span></span>

1. <span data-ttu-id="1a856-146">Abra la pestaña **Automatizar**.</span><span class="sxs-lookup"><span data-stu-id="1a856-146">Open the **Automate** tab.</span></span>
2. <span data-ttu-id="1a856-147">Presione el botón del **Editor de código**.</span><span class="sxs-lookup"><span data-stu-id="1a856-147">Press the **Code Editor** button.</span></span>
3. <span data-ttu-id="1a856-148">Abra el script que anotó en la sección anterior.</span><span class="sxs-lookup"><span data-stu-id="1a856-148">Open the script you recorded in the previous section.</span></span> <span data-ttu-id="1a856-149">Debería ver algo parecido a esto en el código:</span><span class="sxs-lookup"><span data-stu-id="1a856-149">You should see something similar to this code:</span></span>

    ```TypeScript
    async function main(context: Excel.RequestContext) {
      // Set fill color to FFC000 for range Sheet1!A2:C2
      let workbook = context.workbook;
      let worksheets = workbook.worksheets;
      let selectedSheet = worksheets.getActiveWorksheet();
      selectedSheet.getRange("A2:C2").format.fill.color = "FFC000";
    }
    ```

    <span data-ttu-id="1a856-150">Este código accede a la colección de hojas de cálculo del libro para obtener la hoja de cálculo actual.</span><span class="sxs-lookup"><span data-stu-id="1a856-150">This code gets the current worksheet by first accessing the workbook's worksheet collection.</span></span> <span data-ttu-id="1a856-151">Después, establece el color de relleno del rango **A2:C2**.</span><span class="sxs-lookup"><span data-stu-id="1a856-151">Then, it sets the fill color of the range **A2:C2**.</span></span>

    <span data-ttu-id="1a856-152">Los rangos son una parte fundamental de las secuencias de comandos de Office en Excel en la Web.</span><span class="sxs-lookup"><span data-stu-id="1a856-152">Ranges are a fundamental part of Office Scripts in Excel on the web.</span></span> <span data-ttu-id="1a856-153">Un rango es un bloque de celdas contiguo y rectangular que contiene valores, fórmulas y formatos.</span><span class="sxs-lookup"><span data-stu-id="1a856-153">A range is a contiguous, rectangular block of cells that contains values, formula, and formatting.</span></span> <span data-ttu-id="1a856-154">Constituyen la estructura básica de las celdas y se usan para realizar la mayoría de las tareas de scripts.</span><span class="sxs-lookup"><span data-stu-id="1a856-154">They are the basic structure of cells through which you'll perform most of your scripting tasks.</span></span>

4. <span data-ttu-id="1a856-155">Agregue la línea siguiente al final del script (entre el lugar en el que se establece el  y aparece el `}` de cierre):</span><span class="sxs-lookup"><span data-stu-id="1a856-155">Add the following line to the end of the script (between where the `color` is set and the closing `}`):</span></span>

    ```TypeScript
    selectedSheet.getRange("A3:C3").format.fill.color = "yellow";
    ```

5. <span data-ttu-id="1a856-156">Para probar el script, presione **Ejecutar**.</span><span class="sxs-lookup"><span data-stu-id="1a856-156">Test the script by pressing **Run**.</span></span> <span data-ttu-id="1a856-157">El libro tendrá ahora el siguiente aspecto:</span><span class="sxs-lookup"><span data-stu-id="1a856-157">Your workbook should now look like this:</span></span>

    ![Una fila de datos de ventas con la fila "Naranjas" resaltada en naranja y la fila "Limones" ahora resaltada en amarillo.](../images/tutorial-2.png)

## <a name="create-a-table"></a><span data-ttu-id="1a856-159">Crear una tabla</span><span class="sxs-lookup"><span data-stu-id="1a856-159">Create a table</span></span>

<span data-ttu-id="1a856-160">Vamos a convertir estos datos de ventas de frutas en una tabla.</span><span class="sxs-lookup"><span data-stu-id="1a856-160">Let's convert this fruit sales data into a table.</span></span> <span data-ttu-id="1a856-161">Usaremos nuestro script para todo este proceso.</span><span class="sxs-lookup"><span data-stu-id="1a856-161">We'll use our script for the entire process.</span></span>

1. <span data-ttu-id="1a856-162">Agregue la línea siguiente al final del script (antes del `}` de cierre):</span><span class="sxs-lookup"><span data-stu-id="1a856-162">Add the following line to the end of the script (before the closing `}`):</span></span>

    ```TypeScript
    let table = selectedSheet.tables.add("A1:C5", true);
    ```

2. <span data-ttu-id="1a856-163">Esa llamada devuelve un objeto de `Table`.</span><span class="sxs-lookup"><span data-stu-id="1a856-163">That call returns a `Table` object.</span></span> <span data-ttu-id="1a856-164">Vamos a usar la tabla para ordenar los datos.</span><span class="sxs-lookup"><span data-stu-id="1a856-164">Let's use that table to sort the data.</span></span> <span data-ttu-id="1a856-165">Ordenaremos los datos de menor a mayor en función de los valores de la columna "Frutas".</span><span class="sxs-lookup"><span data-stu-id="1a856-165">We'll sort the data in ascending order based on the values in the "Fruit" column.</span></span> <span data-ttu-id="1a856-166">Agregue la siguiente línea después de la creación de tabla:</span><span class="sxs-lookup"><span data-stu-id="1a856-166">Add the following line after the table creation:</span></span>

    ```TypeScript
    table.sort.apply([{ key: 0, ascending: true }]);
    ```

    <span data-ttu-id="1a856-167">Su script debe tener este aspecto:</span><span class="sxs-lookup"><span data-stu-id="1a856-167">Your script should look like this:</span></span>

    ```TypeScript
    async function main(context: Excel.RequestContext) {
      // Set fill color to FFC000 for range Sheet1!A2:C2
      let workbook = context.workbook;
      let worksheets = workbook.worksheets;
      let selectedSheet = worksheets.getActiveWorksheet();
      selectedSheet.getRange("A2:C2").format.fill.color = "FFC000";
      selectedSheet.getRange("A3:C3").format.fill.color = "yellow";
      let table = selectedSheet.tables.add("A1:C5", true);
      table.sort.apply([{ key: 0, ascending: true }]);
    }
    ```

    <span data-ttu-id="1a856-168">Las tablas tienen un objeto `TableSort` al que se accede mediante la propiedad `Table.sort`.</span><span class="sxs-lookup"><span data-stu-id="1a856-168">Tables have a `TableSort` object, accessed through the `Table.sort` property.</span></span> <span data-ttu-id="1a856-169">Puede aplicar un criterio de ordenación a ese objeto.</span><span class="sxs-lookup"><span data-stu-id="1a856-169">You can apply sorting criteria to that object.</span></span> <span data-ttu-id="1a856-170">El método `apply` acepta una matriz de objetos `SortField`.</span><span class="sxs-lookup"><span data-stu-id="1a856-170">The `apply` method takes in an array of `SortField` objects.</span></span> <span data-ttu-id="1a856-171">En este caso, solo tenemos un criterio de ordenación, por lo que solo usamos un `SortField`.</span><span class="sxs-lookup"><span data-stu-id="1a856-171">In this case, we only have one sorting criteria, so we only use one `SortField`.</span></span> <span data-ttu-id="1a856-172">`key: 0` establece los valores que definen la ordenación de la columna como "0" (que es la primera columna de la tabla **A** en este caso).</span><span class="sxs-lookup"><span data-stu-id="1a856-172">`key: 0` sets the column with the sort-defining values to "0" (which is the first column on the table, **A** in this case).</span></span> <span data-ttu-id="1a856-173">`ascending: true` ordena los datos de menor a mayor (en lugar de mayor a menor).</span><span class="sxs-lookup"><span data-stu-id="1a856-173">`ascending: true` sorts the data in ascending order (instead of descending order).</span></span>

3. <span data-ttu-id="1a856-174">Ejecute el script.</span><span class="sxs-lookup"><span data-stu-id="1a856-174">Run the script.</span></span> <span data-ttu-id="1a856-175">Debería ver una tabla como esta:</span><span class="sxs-lookup"><span data-stu-id="1a856-175">You should see a table like this:</span></span>

    ![Una tabla de ventas de frutas ordenada.](../images/tutorial-3.png)

    > [!NOTE]
    > <span data-ttu-id="1a856-177">Si vuelve a ejecutar el script, se producirá un error.</span><span class="sxs-lookup"><span data-stu-id="1a856-177">If you re-run the script, you'll get an error.</span></span> <span data-ttu-id="1a856-178">Esto se debe a que no se puede crear una tabla encima de otra.</span><span class="sxs-lookup"><span data-stu-id="1a856-178">This is because you cannot create a table on top of another table.</span></span> <span data-ttu-id="1a856-179">Sin embargo, puede ejecutar el script en otra hoja de cálculo o en un libro.</span><span class="sxs-lookup"><span data-stu-id="1a856-179">However, you can run the script on a different worksheet or workbook.</span></span>

### <a name="re-run-the-script"></a><span data-ttu-id="1a856-180">Ejecute el script de nuevo</span><span class="sxs-lookup"><span data-stu-id="1a856-180">Re-run the script</span></span>

1. <span data-ttu-id="1a856-181">Crear una nueva hoja de cálculo en el libro actual.</span><span class="sxs-lookup"><span data-stu-id="1a856-181">Create a new worksheet in the current workbook.</span></span>
2. <span data-ttu-id="1a856-182">Copie los datos de frutas del principio del tutorial y péguelos en la nueva hoja de cálculo, comenzando en la celda **A1**.</span><span class="sxs-lookup"><span data-stu-id="1a856-182">Copy the fruit data from the beginning of the tutorial and paste it into the new worksheet, starting at cell **A1**.</span></span>
3. <span data-ttu-id="1a856-183">Ejecute el script.</span><span class="sxs-lookup"><span data-stu-id="1a856-183">Run the script.</span></span>

## <a name="next-steps"></a><span data-ttu-id="1a856-184">Pasos siguientes</span><span class="sxs-lookup"><span data-stu-id="1a856-184">Next steps</span></span>

<span data-ttu-id="1a856-185">Complete el tutorial [Leer datos de libros con scripts de Office en Excel en la Web](excel-read-tutorial.md).</span><span class="sxs-lookup"><span data-stu-id="1a856-185">Complete the [Read workbook data with Office Scripts in Excel on the web](excel-read-tutorial.md) tutorial.</span></span> <span data-ttu-id="1a856-186">En él aprenderá a leer datos de un libro con un script de Office.</span><span class="sxs-lookup"><span data-stu-id="1a856-186">It teaches you how to read data from a workbook with an Office Script.</span></span>
