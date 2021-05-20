---
title: Ejecutar scripts mediante un flujo manual de Power Automate
description: Tutorial sobre el uso de scripts de Office en Power Automate mediante un desencadenador manual.
ms.date: 05/17/2021
localization_priority: Priority
ms.openlocfilehash: f4feb14f70c43497f40dae3a521353dfee63c082
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: HT
ms.contentlocale: es-ES
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545832"
---
# <a name="call-scripts-from-a-manual-power-automate-flow"></a><span data-ttu-id="72a45-103">Ejecutar scripts mediante un flujo manual de Power Automate</span><span class="sxs-lookup"><span data-stu-id="72a45-103">Call scripts from a manual Power Automate flow</span></span>

<span data-ttu-id="72a45-104">Este tutorial le enseña a ejecutar un script de Office para Excel en la Web mediante [Power Automate](https://flow.microsoft.com).</span><span class="sxs-lookup"><span data-stu-id="72a45-104">This tutorial teaches you how to run an Office Script for Excel on the web through [Power Automate](https://flow.microsoft.com).</span></span> <span data-ttu-id="72a45-105">Va a escribir un script que actualiza los valores de dos celdas con la hora actual.</span><span class="sxs-lookup"><span data-stu-id="72a45-105">You'll make a script that updates the values of two cells with the current time.</span></span> <span data-ttu-id="72a45-106">Luego, conectaremos ese script a un flujo de Power Automate activado manualmente, de modo que el script se ejecute siempre que se presione un botón en Power Automate.</span><span class="sxs-lookup"><span data-stu-id="72a45-106">You'll then connect that script to a manually triggered Power Automate flow, so that the script is run whenever a button in Power Automate is pressed.</span></span> <span data-ttu-id="72a45-107">Cuando entienda el funcionamiento básico, podrá expandir el flujo para incluir otras aplicaciones y automatizar una mayor parte de su flujo de trabajo diario.</span><span class="sxs-lookup"><span data-stu-id="72a45-107">Once you understand the basic pattern, you can expand the flow to include other applications and automate more of your daily workflow.</span></span>

> [!TIP]
> <span data-ttu-id="72a45-108">Si no está familiarizado con los scripts de Office, le recomendamos que empiece con el tutorial de [Guardar, editar y crear scripts de Office en Excel en la Web](excel-tutorial.md).</span><span class="sxs-lookup"><span data-stu-id="72a45-108">If you are new to Office Scripts, we recommend starting with the [Record, edit, and create Office Scripts in Excel on the web](excel-tutorial.md) tutorial.</span></span> <span data-ttu-id="72a45-109">[Scripts de Office usa TypeScript](../overview/code-editor-environment.md) y este tutorial está diseñado para las personas con conocimientos de nivel intermedio de JavaScript o TypeScript.</span><span class="sxs-lookup"><span data-stu-id="72a45-109">[Office Scripts use TypeScript](../overview/code-editor-environment.md) and this tutorial is intended for people with beginner to intermediate-level knowledge of JavaScript or TypeScript.</span></span> <span data-ttu-id="72a45-110">Si no está familiarizado con JavaScript, le recomendamos que comience con el [Tutorial de JavaScript de Mozilla](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span><span class="sxs-lookup"><span data-stu-id="72a45-110">If you're new to JavaScript, we recommend starting with the [Mozilla JavaScript tutorial](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="72a45-111">Requisitos previos</span><span class="sxs-lookup"><span data-stu-id="72a45-111">Prerequisites</span></span>

[!INCLUDE [Tutorial prerequisites](../includes/power-automate-tutorial-prerequisites.md)]

## <a name="prepare-the-workbook"></a><span data-ttu-id="72a45-112">Preparar el libro</span><span class="sxs-lookup"><span data-stu-id="72a45-112">Prepare the workbook</span></span>

<span data-ttu-id="72a45-113">Power Automate no debe usar [referencias relativas](../testing/power-automate-troubleshooting.md#avoid-relative-references) como `Workbook.getActiveWorksheet` para acceder a componentes de libros de trabajo.</span><span class="sxs-lookup"><span data-stu-id="72a45-113">Power Automate shouldn't use [relative references](../testing/power-automate-troubleshooting.md#avoid-relative-references) like `Workbook.getActiveWorksheet` to access workbook components.</span></span> <span data-ttu-id="72a45-114">Por lo tanto, es necesario contar con un libro de trabajo y una hoja de cálculo con nombres coherentes a los que pueda hacer referencia Power Automate.</span><span class="sxs-lookup"><span data-stu-id="72a45-114">So, we need a workbook and worksheet with consistent names that Power Automate can reference.</span></span>

1. <span data-ttu-id="72a45-115">Cree un nuevo libro de trabajo y llámelo **Mi libro de trabajo**.</span><span class="sxs-lookup"><span data-stu-id="72a45-115">Create a new workbook named **MyWorkbook**.</span></span>

2. <span data-ttu-id="72a45-116">En el libro **Mi libro de trabajo**, cree una hoja de cálculo y llámela **Hoja de cálculo del tutorial**.</span><span class="sxs-lookup"><span data-stu-id="72a45-116">In the **MyWorkbook** workbook, create a worksheet called **TutorialWorksheet**.</span></span>

## <a name="create-an-office-script"></a><span data-ttu-id="72a45-117">Cree un script de Office</span><span class="sxs-lookup"><span data-stu-id="72a45-117">Create an Office Script</span></span>

1. <span data-ttu-id="72a45-118">Vaya a la pestaña **Automatizar** y seleccione **Todos los scripts**.</span><span class="sxs-lookup"><span data-stu-id="72a45-118">Go to the **Automate** tab and select **All Scripts**.</span></span>

2. <span data-ttu-id="72a45-119">Seleccione **Nuevo script**.</span><span class="sxs-lookup"><span data-stu-id="72a45-119">Select **New Script**.</span></span>

3. <span data-ttu-id="72a45-120">Reemplace el script predeterminado con el siguiente script.</span><span class="sxs-lookup"><span data-stu-id="72a45-120">Replace the default script with the following script.</span></span> <span data-ttu-id="72a45-121">Este script agrega la fecha y la hora actuales a las dos primeras celdas de la hoja de cálculo **Hoja de cálculo del tutorial**.</span><span class="sxs-lookup"><span data-stu-id="72a45-121">This script adds the current date and time to the first two cells of the **TutorialWorksheet** worksheet.</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
      // Get the "TutorialWorksheet" worksheet from the workbook.
      let worksheet = workbook.getWorksheet("TutorialWorksheet");

      // Get the cells at A1 and B1.
      let dateRange = worksheet.getRange("A1");
      let timeRange = worksheet.getRange("B1");

      // Get the current date and time using the JavaScript Date object.
      let date = new Date(Date.now());

      // Add the date string to A1.
      dateRange.setValue(date.toLocaleDateString());

      // Add the time string to B1.
      timeRange.setValue(date.toLocaleTimeString());
    }
    ```

4. <span data-ttu-id="72a45-122">Cambie el nombre del script a **Establecer la fecha y la hora**.</span><span class="sxs-lookup"><span data-stu-id="72a45-122">Rename the script to **Set date and time**.</span></span> <span data-ttu-id="72a45-123">Presione el nombre del script para cambiarlo.</span><span class="sxs-lookup"><span data-stu-id="72a45-123">Press the script name to change it.</span></span>

5. <span data-ttu-id="72a45-124">Para guardar el script, presione el botón **Guardar script**.</span><span class="sxs-lookup"><span data-stu-id="72a45-124">Save the script by pressing **Save Script**.</span></span>

## <a name="create-an-automated-workflow-with-power-automate"></a><span data-ttu-id="72a45-125">Crear un flujo de trabajo automatizado con Power Automate</span><span class="sxs-lookup"><span data-stu-id="72a45-125">Create an automated workflow with Power Automate</span></span>

1. <span data-ttu-id="72a45-126">Inicie sesión en el [sitio de Power Automate](https://flow.microsoft.com).</span><span class="sxs-lookup"><span data-stu-id="72a45-126">Sign in to the [Power Automate site](https://flow.microsoft.com).</span></span>

2. <span data-ttu-id="72a45-127">En el menú que se muestra en la parte izquierda de la pantalla, presione **Crear**.</span><span class="sxs-lookup"><span data-stu-id="72a45-127">In the menu that's displayed on the left side of the screen, press **Create**.</span></span> <span data-ttu-id="72a45-128">Se mostrará una lista de maneras de crear flujos de trabajo nuevos.</span><span class="sxs-lookup"><span data-stu-id="72a45-128">This brings you to list of ways to create new workflows.</span></span>

    :::image type="content" source="../images/power-automate-tutorial-1.png" alt-text="El botón &quot;Crear&quot; de Power Automate.":::

3. <span data-ttu-id="72a45-130">En la sección **Inicio desde cero**, seleccione **Flujo instantáneo**.</span><span class="sxs-lookup"><span data-stu-id="72a45-130">In the **Start from blank** section, select **Instant flow**.</span></span> <span data-ttu-id="72a45-131">Se creará un flujo de trabajo activado manualmente.</span><span class="sxs-lookup"><span data-stu-id="72a45-131">This creates a manually activated workflow.</span></span>

    :::image type="content" source="../images/power-automate-tutorial-2.png" alt-text="La opción de flujo instantáneo de Power Automate para crear un nuevo flujo de trabajo.":::

4. <span data-ttu-id="72a45-133">En la ventana de diálogo que aparece, escriba un nombre para el flujo en el cuadro de texto **Nombre de flujo**, seleccione **Activar manualmente un flujo** de la lista de opciones en **Elija cómo desencadenar el flujo** y presione **Crear**.</span><span class="sxs-lookup"><span data-stu-id="72a45-133">In the dialog window that appears, enter a name for your flow in the **Flow name** text box, select **Manually trigger a flow** from the list of options under **Choose how to trigger the flow**, and press **Create**.</span></span>

    :::image type="content" source="../images/power-automate-tutorial-3.png" alt-text="La opción &quot;Activar manualmente un flujo&quot; de Power Automate.":::

    <span data-ttu-id="72a45-135">Tenga en cuenta que un flujo activado manualmente es solo uno de los distintos tipos de flujos.</span><span class="sxs-lookup"><span data-stu-id="72a45-135">Note that a manually triggered flow is just one of many types of flows.</span></span> <span data-ttu-id="72a45-136">En el siguiente tutorial, podrá crear un flujo que se ejecuta automáticamente al recibir un correo electrónico.</span><span class="sxs-lookup"><span data-stu-id="72a45-136">In the next tutorial, you'll make a flow that automatically runs when you receive an email.</span></span>

5. <span data-ttu-id="72a45-137">Presione **Nuevo paso**.</span><span class="sxs-lookup"><span data-stu-id="72a45-137">Press **New step**.</span></span>

6. <span data-ttu-id="72a45-138">Seleccione la pestaña **Estándar** y, a continuación, seleccione **Excel Online (empresa)**.</span><span class="sxs-lookup"><span data-stu-id="72a45-138">Select the **Standard** tab, then select **Excel Online (Business)**.</span></span>

    :::image type="content" source="../images/power-automate-tutorial-4.png" alt-text="Opción de Excel Online (empresa) en Power Automate":::

7. <span data-ttu-id="72a45-140">En **Acciones**, seleccione **Ejecutar script**.</span><span class="sxs-lookup"><span data-stu-id="72a45-140">Under **Actions**, select **Run script**.</span></span>

    :::image type="content" source="../images/power-automate-tutorial-5.png" alt-text="Opción de acción Ejecutar script en Power Automate":::

8. <span data-ttu-id="72a45-142">A continuación, seleccione el libro y el script que va a usar en el paso de flujo.</span><span class="sxs-lookup"><span data-stu-id="72a45-142">Next, you'll select the workbook and script to use in the flow step.</span></span> <span data-ttu-id="72a45-143">En el tutorial, usará el libro que creó en OneDrive, pero puede usar cualquier libro en un sitio de OneDrive o SharePoint.</span><span class="sxs-lookup"><span data-stu-id="72a45-143">For the tutorial, you'll use the workbook you created in your OneDrive, but you could use any workbook in a OneDrive or SharePoint site.</span></span> <span data-ttu-id="72a45-144">Especifique la siguiente configuración para el conector **Ejecutar script**:</span><span class="sxs-lookup"><span data-stu-id="72a45-144">Specify the following settings for the **Run script** connector:</span></span>

    - <span data-ttu-id="72a45-145">**Ubicación**: OneDrive para la Empresa</span><span class="sxs-lookup"><span data-stu-id="72a45-145">**Location**: OneDrive for Business</span></span>
    - <span data-ttu-id="72a45-146">**Biblioteca de documentos**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="72a45-146">**Document Library**: OneDrive</span></span>
    - <span data-ttu-id="72a45-147">**Archivo**: MyWorkbook.xlsx *(seleccionado por el explorador de archivos)*</span><span class="sxs-lookup"><span data-stu-id="72a45-147">**File**: MyWorkbook.xlsx *(Chosen through the file browser)*</span></span>
    - <span data-ttu-id="72a45-148">**Script**: Establecer fecha y hora</span><span class="sxs-lookup"><span data-stu-id="72a45-148">**Script**: Set date and time</span></span>

    :::image type="content" source="../images/power-automate-tutorial-6.png" alt-text="La configuración del conector de Power Automate para ejecutar un script.":::

9. <span data-ttu-id="72a45-150">Presione **Guardar**.</span><span class="sxs-lookup"><span data-stu-id="72a45-150">Press **Save**.</span></span>

<span data-ttu-id="72a45-151">El flujo ya está listo para ejecutarse mediante Power Automate.</span><span class="sxs-lookup"><span data-stu-id="72a45-151">Your flow is now ready to be run through Power Automate.</span></span> <span data-ttu-id="72a45-152">Para probarlo, pulse el botón **Probar** en el editor de flujos o siga los pasos restantes del tutorial para ejecutar el flujo de la colección de flujos.</span><span class="sxs-lookup"><span data-stu-id="72a45-152">You can test it using the **Test** button in the flow editor or follow the remaining tutorial steps to run the flow from your flow collection.</span></span>

## <a name="run-the-script-through-power-automate"></a><span data-ttu-id="72a45-153">Ejecutar el script mediante Power Automate</span><span class="sxs-lookup"><span data-stu-id="72a45-153">Run the script through Power Automate</span></span>

1. <span data-ttu-id="72a45-154">En la página principal de Power Automate, seleccione **Mis flujos**.</span><span class="sxs-lookup"><span data-stu-id="72a45-154">From the main Power Automate page, select **My flows**.</span></span>

    :::image type="content" source="../images/power-automate-tutorial-7.png" alt-text="El botón Mis flujos en Power Automate":::

2. <span data-ttu-id="72a45-156">Seleccione **Mi flujo de tutoriales** en la lista de flujos que se muestra en la pestaña **Mis flujos**. Se mostrarán los detalles del flujo que creó anteriormente.</span><span class="sxs-lookup"><span data-stu-id="72a45-156">Select **My tutorial flow** from the list of flows displayed in the **My flows** tab. This shows the details of the flow we previously created.</span></span>

3. <span data-ttu-id="72a45-157">Pulse **Ejecutar**.</span><span class="sxs-lookup"><span data-stu-id="72a45-157">Press **Run**.</span></span>

    :::image type="content" source="../images/power-automate-tutorial-8.png" alt-text="El botón Ejecutar en Power Automate":::

4. <span data-ttu-id="72a45-159">Se mostrará un panel de tareas para ejecutar el flujo.</span><span class="sxs-lookup"><span data-stu-id="72a45-159">A task pane will appear for running the flow.</span></span> <span data-ttu-id="72a45-160">Si se le solicita **Iniciar sesión** en Excel Online, presione **Continuar**.</span><span class="sxs-lookup"><span data-stu-id="72a45-160">If you are asked to **Sign in** to Excel Online, do so by pressing **Continue**.</span></span>

5. <span data-ttu-id="72a45-161">Pulse **Ejecutar flujo**.</span><span class="sxs-lookup"><span data-stu-id="72a45-161">Press **Run flow**.</span></span> <span data-ttu-id="72a45-162">Se ejecutará el flujo, que ejecuta a su vez el script de Office relacionado.</span><span class="sxs-lookup"><span data-stu-id="72a45-162">This runs the flow, which runs the related Office Script.</span></span>

6. <span data-ttu-id="72a45-163">Presione **Listo**.</span><span class="sxs-lookup"><span data-stu-id="72a45-163">Press **Done**.</span></span> <span data-ttu-id="72a45-164">En consecuencia, debería actualizarse la sección **Ejecuciones**.</span><span class="sxs-lookup"><span data-stu-id="72a45-164">You should see the **Runs** section update accordingly.</span></span>

7. <span data-ttu-id="72a45-165">Actualice la página para ver los resultados de Power Automate.</span><span class="sxs-lookup"><span data-stu-id="72a45-165">Refresh the page to see the results of the Power Automate.</span></span> <span data-ttu-id="72a45-166">Si se ha realizado correctamente, podrá ver las celdas actualizadas en el libro de trabajo.</span><span class="sxs-lookup"><span data-stu-id="72a45-166">If it succeeded, go to the workbook to see the updated cells.</span></span> <span data-ttu-id="72a45-167">Si se ha producido un error, compruebe la configuración del flujo y ejecútelo de nuevo.</span><span class="sxs-lookup"><span data-stu-id="72a45-167">If it failed, verify the flow's settings and run it a second time.</span></span>

    :::image type="content" source="../images/power-automate-tutorial-9.png" alt-text="Resultado de Power Automate que muestra una ejecución de flujo satisfactoria":::

## <a name="next-steps"></a><span data-ttu-id="72a45-169">Pasos siguientes</span><span class="sxs-lookup"><span data-stu-id="72a45-169">Next steps</span></span>

<span data-ttu-id="72a45-170">Complete el tutorial [Pasar datos a scripts en un flujo de Power Automate ejecutado automáticamente](excel-power-automate-trigger.md).</span><span class="sxs-lookup"><span data-stu-id="72a45-170">Complete the [Pass data to scripts in an automatically-run Power Automate flow](excel-power-automate-trigger.md) tutorial.</span></span> <span data-ttu-id="72a45-171">Aprenderá a pasar datos desde un servicio de flujo de trabajo al script de Office y a ejecutar el flujo de Power Automate cuando se producen determinados eventos.</span><span class="sxs-lookup"><span data-stu-id="72a45-171">It teaches you how to pass data from a workflow service to your Office Script and run the Power Automate flow when certain events occur.</span></span>
