---
title: Ejecutar scripts mediante un flujo manual de Power Automate
description: Tutorial sobre el uso de scripts de Office en Power Automate mediante un desencadenador manual.
ms.date: 07/14/2020
localization_priority: Priority
ms.openlocfilehash: 70fca2620973ecefe9eda40f02e28f064b713677
ms.sourcegitcommit: ebd1079c7e2695ac0e7e4c616f2439975e196875
ms.translationtype: HT
ms.contentlocale: es-ES
ms.lasthandoff: 07/17/2020
ms.locfileid: "45160440"
---
# <a name="call-scripts-from-a-manual-power-automate-flow-preview"></a><span data-ttu-id="40dc6-103">Ejecutar scripts mediante un flujo manual de Power Automate (versión preliminar)</span><span class="sxs-lookup"><span data-stu-id="40dc6-103">Call scripts from a manual Power Automate flow (preview)</span></span>

<span data-ttu-id="40dc6-104">Este tutorial le enseña a ejecutar un script de Office para Excel en la Web mediante [Power Automate](https://flow.microsoft.com).</span><span class="sxs-lookup"><span data-stu-id="40dc6-104">This tutorial teaches you how to run an Office Script for Excel on the web through [Power Automate](https://flow.microsoft.com).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="40dc6-105">Requisitos previos</span><span class="sxs-lookup"><span data-stu-id="40dc6-105">Prerequisites</span></span>

[!INCLUDE [Tutorial prerequisites](../includes/power-automate-tutorial-prerequisites.md)]

> [!IMPORTANT]
> <span data-ttu-id="40dc6-106">Este tutorial presupone que usted ha completado el tutorial [Registro, edición y creación de scripts de Office en Excel en la web](excel-tutorial.md).</span><span class="sxs-lookup"><span data-stu-id="40dc6-106">This tutorial assumes you have completed the [Record, edit, and create Office Scripts in Excel on the web](excel-tutorial.md) tutorial.</span></span>

## <a name="prepare-the-workbook"></a><span data-ttu-id="40dc6-107">Preparar el libro de trabajo</span><span class="sxs-lookup"><span data-stu-id="40dc6-107">Prepare the workbook</span></span>

<span data-ttu-id="40dc6-108">Power Automate no puede usar referencias relativas como `Workbook.getActiveWorksheet` para acceder a componentes de libros de trabajo.</span><span class="sxs-lookup"><span data-stu-id="40dc6-108">Power Automate can't use relative references like `Workbook.getActiveWorksheet` to access workbook components.</span></span> <span data-ttu-id="40dc6-109">Por lo tanto, es necesario contar con un libro de trabajo y una hoja de cálculo con nombres coherentes a los que pueda hacer referencia Power Automate.</span><span class="sxs-lookup"><span data-stu-id="40dc6-109">So, we need a workbook and worksheet with consistent names that Power Automate can reference.</span></span>

1. <span data-ttu-id="40dc6-110">Cree un nuevo libro de trabajo y llámelo **Mi libro de trabajo**.</span><span class="sxs-lookup"><span data-stu-id="40dc6-110">Create a new workbook named **MyWorkbook**.</span></span>

2. <span data-ttu-id="40dc6-111">En el libro **Mi libro de trabajo**, cree una hoja de cálculo y llámela **Hoja de cálculo del tutorial**.</span><span class="sxs-lookup"><span data-stu-id="40dc6-111">In the **MyWorkbook** workbook, create a worksheet called **TutorialWorksheet**.</span></span>

## <a name="create-an-office-script"></a><span data-ttu-id="40dc6-112">Cree un script de Office</span><span class="sxs-lookup"><span data-stu-id="40dc6-112">Create an Office Script</span></span>

1. <span data-ttu-id="40dc6-113">Vaya a la pestaña **Automatizar** y seleccione **Editor de código**.</span><span class="sxs-lookup"><span data-stu-id="40dc6-113">Go to the **Automate** tab and select **Code Editor**.</span></span>

2. <span data-ttu-id="40dc6-114">Seleccione **Nuevo script**.</span><span class="sxs-lookup"><span data-stu-id="40dc6-114">Select **New Script**.</span></span>

3. <span data-ttu-id="40dc6-115">Reemplace el script predeterminado con el siguiente script.</span><span class="sxs-lookup"><span data-stu-id="40dc6-115">Replace the default script with the following script.</span></span> <span data-ttu-id="40dc6-116">Este script agrega la fecha y la hora actuales a las dos primeras celdas de la hoja de cálculo **Hoja de cálculo del tutorial**.</span><span class="sxs-lookup"><span data-stu-id="40dc6-116">This script adds the current date and time to the first two cells of the **TutorialWorksheet** worksheet.</span></span>

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

4. <span data-ttu-id="40dc6-117">Cambie el nombre del script a **Establecer la fecha y la hora**.</span><span class="sxs-lookup"><span data-stu-id="40dc6-117">Rename the script to **Set date and time**.</span></span> <span data-ttu-id="40dc6-118">Presione el nombre del script para cambiarlo.</span><span class="sxs-lookup"><span data-stu-id="40dc6-118">Press the script name to change it.</span></span>

5. <span data-ttu-id="40dc6-119">Para guardar el script, presione el botón **Guardar script**.</span><span class="sxs-lookup"><span data-stu-id="40dc6-119">Save the script by pressing **Save Script**.</span></span>

## <a name="create-an-automated-workflow-with-power-automate"></a><span data-ttu-id="40dc6-120">Crear un flujo de trabajo automatizado con Power Automate</span><span class="sxs-lookup"><span data-stu-id="40dc6-120">Create an automated workflow with Power Automate</span></span>

1. <span data-ttu-id="40dc6-121">Inicie sesión en el [sitio de Power Automate](https://flow.microsoft.com).</span><span class="sxs-lookup"><span data-stu-id="40dc6-121">Sign in to the [Power Automate site](https://flow.microsoft.com).</span></span>

2. <span data-ttu-id="40dc6-122">En el menú que se muestra en la parte izquierda de la pantalla, presione **Crear**.</span><span class="sxs-lookup"><span data-stu-id="40dc6-122">In the menu that's displayed on the left side of the screen, press **Create**.</span></span> <span data-ttu-id="40dc6-123">Se mostrará una lista de maneras de crear flujos de trabajo nuevos.</span><span class="sxs-lookup"><span data-stu-id="40dc6-123">This brings you to list of ways to create new workflows.</span></span>

    ![El botón Crear en Power Automate.](../images/power-automate-tutorial-1.png)

3. <span data-ttu-id="40dc6-125">En la sección **Inicio desde cero**, seleccione **Flujo instantáneo**.</span><span class="sxs-lookup"><span data-stu-id="40dc6-125">In the **Start from blank** section, select **Instant flow**.</span></span> <span data-ttu-id="40dc6-126">Se creará un flujo de trabajo activado manualmente.</span><span class="sxs-lookup"><span data-stu-id="40dc6-126">This creates a manually activated workflow.</span></span>

    ![La opción Flujo instantáneo para crear un nuevo flujo de trabajo.](../images/power-automate-tutorial-2.png)

4. <span data-ttu-id="40dc6-128">En la ventana de diálogo que aparece, escriba un nombre para el flujo en el cuadro de texto **Nombre de flujo**, seleccione **Activar manualmente un flujo** de la lista de opciones en **Elija cómo desencadenar el flujo** y presione **Crear**.</span><span class="sxs-lookup"><span data-stu-id="40dc6-128">In the dialog window that appears, enter a name for your flow in the **Flow name** text box, select **Manually trigger a flow** from the list of options under **Choose how to trigger the flow**, and press **Create**.</span></span>

    ![La opción de desencadenador manual para crear un nuevo flujo instantáneo.](../images/power-automate-tutorial-3.png)

    <span data-ttu-id="40dc6-130">Tenga en cuenta que un flujo activado manualmente es solo uno de los distintos tipos de flujos.</span><span class="sxs-lookup"><span data-stu-id="40dc6-130">Note that a manually triggered flow is just one of many types of flows.</span></span> <span data-ttu-id="40dc6-131">En el siguiente tutorial, podrá crear un flujo que se ejecuta automáticamente al recibir un correo electrónico.</span><span class="sxs-lookup"><span data-stu-id="40dc6-131">In the next tutorial, you'll make a flow that automatically runs when you receive an email.</span></span>

5. <span data-ttu-id="40dc6-132">Presione **Nuevo paso**.</span><span class="sxs-lookup"><span data-stu-id="40dc6-132">Press **New step**.</span></span>

6. <span data-ttu-id="40dc6-133">Seleccione la pestaña **Estándar** y, a continuación, seleccione **Excel Online (empresa)**.</span><span class="sxs-lookup"><span data-stu-id="40dc6-133">Select the **Standard** tab, then select **Excel Online (Business)**.</span></span>

    ![La opción Power Automate para Excel Online (empresa).](../images/power-automate-tutorial-4.png)

7. <span data-ttu-id="40dc6-135">En **Acciones**, seleccione **Ejecutar script (versión preliminar)**.</span><span class="sxs-lookup"><span data-stu-id="40dc6-135">Under **Actions**, select **Run script (preview)**.</span></span>

    ![La opción de acción de Power Automate para Ejecutar script (versión preliminar).](../images/power-automate-tutorial-5.png)

8. <span data-ttu-id="40dc6-137">Especifique la siguiente configuración para el conector **Ejecutar script**:</span><span class="sxs-lookup"><span data-stu-id="40dc6-137">Specify the following settings for the **Run script** connector:</span></span>

    - <span data-ttu-id="40dc6-138">**Ubicación**: OneDrive para la Empresa</span><span class="sxs-lookup"><span data-stu-id="40dc6-138">**Location**: OneDrive for Business</span></span>
    - <span data-ttu-id="40dc6-139">**Biblioteca de documentos**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="40dc6-139">**Document Library**: OneDrive</span></span>
    - <span data-ttu-id="40dc6-140">**Archivo**: MyWorkbook.xlsx</span><span class="sxs-lookup"><span data-stu-id="40dc6-140">**File**: MyWorkbook.xlsx</span></span>
    - <span data-ttu-id="40dc6-141">**Script**: Establecer fecha y hora</span><span class="sxs-lookup"><span data-stu-id="40dc6-141">**Script**: Set date and time</span></span>

    ![La configuración del conector para ejecutar un script en Power Automate.](../images/power-automate-tutorial-6.png)

9. <span data-ttu-id="40dc6-143">Presione **Guardar**.</span><span class="sxs-lookup"><span data-stu-id="40dc6-143">Press **Save**.</span></span>

<span data-ttu-id="40dc6-144">El flujo ya está listo para ejecutarse mediante Power Automate.</span><span class="sxs-lookup"><span data-stu-id="40dc6-144">Your flow is now ready to be run through Power Automate.</span></span> <span data-ttu-id="40dc6-145">Para probarlo, pulse el botón **Probar** en el editor de flujos o siga los pasos restantes del tutorial para ejecutar el flujo de la colección de flujos.</span><span class="sxs-lookup"><span data-stu-id="40dc6-145">You can test it using the **Test** button in the flow editor or follow the remaining tutorial steps to run the flow from your flow collection.</span></span>

## <a name="run-the-script-through-power-automate"></a><span data-ttu-id="40dc6-146">Ejecutar el script mediante Power Automate</span><span class="sxs-lookup"><span data-stu-id="40dc6-146">Run the script through Power Automate</span></span>

1. <span data-ttu-id="40dc6-147">En la página principal de Power Automate, seleccione **Mis flujos**.</span><span class="sxs-lookup"><span data-stu-id="40dc6-147">From the main Power Automate page, select **My flows**.</span></span>

    ![El botón Mis flujos en Power Automate.](../images/power-automate-tutorial-7.png)

2. <span data-ttu-id="40dc6-149">Seleccione **Mi flujo de tutoriales** en la lista de flujos que se muestra en la pestaña **Mis flujos**. Se mostrarán los detalles del flujo que creó anteriormente.</span><span class="sxs-lookup"><span data-stu-id="40dc6-149">Select **My tutorial flow** from the list of flows displayed in the **My flows** tab. This shows the details of the flow we previously created.</span></span>

3. <span data-ttu-id="40dc6-150">Pulse **Ejecutar**.</span><span class="sxs-lookup"><span data-stu-id="40dc6-150">Press **Run**.</span></span>

    ![El botón Ejecutar en Power Automate.](../images/power-automate-tutorial-8.png)

4. <span data-ttu-id="40dc6-152">Se mostrará un panel de tareas para ejecutar el flujo.</span><span class="sxs-lookup"><span data-stu-id="40dc6-152">A task pane will appear for running the flow.</span></span> <span data-ttu-id="40dc6-153">Si se le solicita **Iniciar sesión** en Excel Online, presione **Continuar**.</span><span class="sxs-lookup"><span data-stu-id="40dc6-153">If you are asked to **Sign in** to Excel Online, do so by pressing **Continue**.</span></span>

5. <span data-ttu-id="40dc6-154">Pulse **Ejecutar flujo**.</span><span class="sxs-lookup"><span data-stu-id="40dc6-154">Press **Run flow**.</span></span> <span data-ttu-id="40dc6-155">Se ejecutará el flujo, que ejecuta a su vez el script de Office relacionado.</span><span class="sxs-lookup"><span data-stu-id="40dc6-155">This runs the flow, which runs the related Office Script.</span></span>

6. <span data-ttu-id="40dc6-156">Presione **Listo**.</span><span class="sxs-lookup"><span data-stu-id="40dc6-156">Press **Done**.</span></span> <span data-ttu-id="40dc6-157">En consecuencia, debería actualizarse la sección **Ejecuciones**.</span><span class="sxs-lookup"><span data-stu-id="40dc6-157">You should see the **Runs** section update accordingly.</span></span>

7. <span data-ttu-id="40dc6-158">Actualice la página para ver los resultados de Power Automate.</span><span class="sxs-lookup"><span data-stu-id="40dc6-158">Refresh the page to see the results of the Power Automate.</span></span> <span data-ttu-id="40dc6-159">Si se ha realizado correctamente, podrá ver las celdas actualizadas en el libro de trabajo.</span><span class="sxs-lookup"><span data-stu-id="40dc6-159">If it succeeded, go to the workbook to see the updated cells.</span></span> <span data-ttu-id="40dc6-160">Si se ha producido un error, compruebe la configuración del flujo y ejecútelo de nuevo.</span><span class="sxs-lookup"><span data-stu-id="40dc6-160">If it failed, verify the flow's settings and run it a second time.</span></span>

    ![Salida de Power Automate que muestra una ejecución de flujo satisfactoria.](../images/power-automate-tutorial-9.png)

## <a name="next-steps"></a><span data-ttu-id="40dc6-162">Pasos siguientes</span><span class="sxs-lookup"><span data-stu-id="40dc6-162">Next steps</span></span>

<span data-ttu-id="40dc6-163">Complete el tutorial [Pasar datos a scripts en un flujo de Power Automate ejecutado automáticamente](excel-power-automate-trigger.md).</span><span class="sxs-lookup"><span data-stu-id="40dc6-163">Complete the [Pass data to scripts in an automatically-run Power Automate flow](excel-power-automate-trigger.md) tutorial.</span></span> <span data-ttu-id="40dc6-164">Aprenderá a pasar datos desde un servicio de flujo de trabajo al script de Office y a ejecutar el flujo de Power Automate cuando se producen determinados eventos.</span><span class="sxs-lookup"><span data-stu-id="40dc6-164">It teaches you how to pass data from a workflow service to your Office Script and run the Power Automate flow when certain events occur.</span></span>
