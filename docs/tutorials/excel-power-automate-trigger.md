---
title: Pasar datos a scripts en un flujo de Power Automate ejecutado automáticamente
description: Un tutorial sobre la ejecución de Scripts de Office para Excel en la Web mediante Power Automate cuando se reciba el correo y el paso de datos de flujo al script.
ms.date: 07/24/2020
localization_priority: Priority
ms.openlocfilehash: aed34f4b93bbe22768aab73d7a7264cc7d3c3ee6
ms.sourcegitcommit: ff7fde04ce5a66d8df06ed505951c8111e2e9833
ms.translationtype: HT
ms.contentlocale: es-ES
ms.lasthandoff: 08/11/2020
ms.locfileid: "46616769"
---
# <a name="pass-data-to-scripts-in-an-automatically-run-power-automate-flow-preview"></a><span data-ttu-id="0dda4-103">Pasar datos a scripts en un flujo de Power Automate ejecutado automáticamente (versión preliminar)</span><span class="sxs-lookup"><span data-stu-id="0dda4-103">Pass data to scripts in an automatically-run Power Automate flow (preview)</span></span>

<span data-ttu-id="0dda4-104">Este tutorial le enseña cómo usar un script de Office para Excel en la web con un flujo de trabajo automatizado de [Power Automate](https://flow.microsoft.com).</span><span class="sxs-lookup"><span data-stu-id="0dda4-104">This tutorial teaches you how to use an Office Script for Excel on the web with an automated [Power Automate](https://flow.microsoft.com) workflow.</span></span> <span data-ttu-id="0dda4-105">El script se ejecutará automáticamente cada vez que reciba un correo electrónico, grabando información del correo en un libro de Excel.</span><span class="sxs-lookup"><span data-stu-id="0dda4-105">Your script will automatically run each time you receive an email, recording information from the email in an Excel workbook.</span></span> <span data-ttu-id="0dda4-106">Pasar datos de otras aplicaciones a un script de Office le ofrece una gran flexibilidad y libertad para sus procesos automatizados.</span><span class="sxs-lookup"><span data-stu-id="0dda4-106">Being able to pass data from other applications into an Office Script gives you a great deal of flexibility and freedom in your automated processes.</span></span>

> [!TIP]
> <span data-ttu-id="0dda4-107">Si no está familiarizado con los scripts de Office, le recomendamos que empiece con el tutorial de [Guardar, editar y crear scripts de Office en Excel en la Web](excel-tutorial.md).</span><span class="sxs-lookup"><span data-stu-id="0dda4-107">If you are new to Office Scripts, we recommend starting with the [Record, edit, and create Office Scripts in Excel on the web](excel-tutorial.md) tutorial.</span></span> <span data-ttu-id="0dda4-108">Si es la primera vez que usa Power Automate, le recomendamos que comience con el tutorial [Llamar scripts desde un flujo manual de Power Automate](excel-power-automate-manual.md).</span><span class="sxs-lookup"><span data-stu-id="0dda4-108">If you are new to Power Automate, we recommend starting with the [Call scripts from a manual Power Automate flow](excel-power-automate-manual.md) tutorial.</span></span> <span data-ttu-id="0dda4-109">[Scripts de Office usan TypeScript](../overview/code-editor-environment.md) y este tutorial está diseñado para las personas con conocimientos a nivel intermedio de JavaScript o TypeScript.</span><span class="sxs-lookup"><span data-stu-id="0dda4-109">[Office Scripts use TypeScript](../overview/code-editor-environment.md) and this tutorial is intended for people with beginner to intermediate-level knowledge of JavaScript or TypeScript.</span></span> <span data-ttu-id="0dda4-110">Si no está familiarizado con JavaScript, le recomendamos que comience con el [Tutorial de JavaScript de Mozilla](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span><span class="sxs-lookup"><span data-stu-id="0dda4-110">If you're new to JavaScript, we recommend starting with the [Mozilla JavaScript tutorial](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="0dda4-111">Requisitos previos</span><span class="sxs-lookup"><span data-stu-id="0dda4-111">Prerequisites</span></span>

[!INCLUDE [Tutorial prerequisites](../includes/power-automate-tutorial-prerequisites.md)]

## <a name="prepare-the-workbook"></a><span data-ttu-id="0dda4-112">Preparar el libro</span><span class="sxs-lookup"><span data-stu-id="0dda4-112">Prepare the workbook</span></span>

<span data-ttu-id="0dda4-113">Power Automate no puede usar [referencias relativas](../develop/power-automate-integration.md#avoid-using-relative-references) como `Workbook.getActiveWorksheet` para acceder a componentes de libros de trabajo.</span><span class="sxs-lookup"><span data-stu-id="0dda4-113">Power Automate can't use [relative references](../develop/power-automate-integration.md#avoid-using-relative-references) like `Workbook.getActiveWorksheet` to access workbook components.</span></span> <span data-ttu-id="0dda4-114">Por lo tanto, es necesario un libro de trabajo y una hoja de cálculo con nombres coherentes para que Power Automate haga referencia.</span><span class="sxs-lookup"><span data-stu-id="0dda4-114">So, we need a workbook and worksheet with consistent names for Power Automate to reference.</span></span>

1. <span data-ttu-id="0dda4-115">Cree un nuevo libro de trabajo y llámelo **Mi libro de trabajo**.</span><span class="sxs-lookup"><span data-stu-id="0dda4-115">Create a new workbook named **MyWorkbook**.</span></span>

2. <span data-ttu-id="0dda4-116">Vaya a la pestaña **Automatizar** y seleccione **Editor de código**.</span><span class="sxs-lookup"><span data-stu-id="0dda4-116">Go to the **Automate** tab and select **Code Editor**.</span></span>

3. <span data-ttu-id="0dda4-117">Seleccione **Nuevo script**.</span><span class="sxs-lookup"><span data-stu-id="0dda4-117">Select **New Script**.</span></span>

4. <span data-ttu-id="0dda4-118">Reemplace el código existente con el siguiente script y presione **Ejecutar**.</span><span class="sxs-lookup"><span data-stu-id="0dda4-118">Replace the existing code with the following script and press **Run**.</span></span> <span data-ttu-id="0dda4-119">Esto configurará el libro con nombres de tabla dinámica, hoja de cálculo y tabla coherentes.</span><span class="sxs-lookup"><span data-stu-id="0dda4-119">This will setup the workbook with consistent worksheet, table, and PivotTable names.</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
      // Add a new worksheet to store our email table
      let emailsSheet = workbook.addWorksheet("Emails");

      // Add data and create a table
      emailsSheet.getRange("A1:D1").setValues([
        ["Date", "Day of the week", "Email address", "Subject"]
      ]);
      let newTable = workbook.addTable(emailsSheet.getRange("A1:D2"), true);
      newTable.setName("EmailTable");

      // Add a new PivotTable to a new worksheet
      let pivotWorksheet = workbook.addWorksheet("Subjects");
      let newPivotTable = workbook.addPivotTable("Pivot", "EmailTable", pivotWorksheet.getRange("A3:C20"));

      // Setup the pivot hierarchies
      newPivotTable.addRowHierarchy(newPivotTable.getHierarchy("Day of the week"));
      newPivotTable.addRowHierarchy(newPivotTable.getHierarchy("Email address"));
      newPivotTable.addDataHierarchy(newPivotTable.getHierarchy("Subject"));
    }
    ```

## <a name="create-an-office-script"></a><span data-ttu-id="0dda4-120">Crear un script de Office</span><span class="sxs-lookup"><span data-stu-id="0dda4-120">Create an Office Script</span></span>

<span data-ttu-id="0dda4-121">Comencemos a crear un script que registre información de un correo electrónico.</span><span class="sxs-lookup"><span data-stu-id="0dda4-121">Let's create a script that logs information from an email.</span></span> <span data-ttu-id="0dda4-122">Queremos saber cuál es el número de días de la semana que recibimos más correo electrónico y cuántos remitentes únicos envían ese correo.</span><span class="sxs-lookup"><span data-stu-id="0dda4-122">We want to know how which days of the week we receive the most mail and how many unique senders are sending that mail.</span></span> <span data-ttu-id="0dda4-123">Nuestro libro tiene una tabla con columnas de **Fecha**, **Día de la semana**, **Dirección de correo electrónico** y **Asunto**.</span><span class="sxs-lookup"><span data-stu-id="0dda4-123">Our workbook has a table with **Date**, **Day of the week**, **Email address**, and **Subject** columns.</span></span> <span data-ttu-id="0dda4-124">Nuestra hoja de cálculo también tiene una tabla dinámica que se dinamiza en el **Día de la semana** y **Dirección de correo electrónico** (que son las jerarquías de fila).</span><span class="sxs-lookup"><span data-stu-id="0dda4-124">Our worksheet also has a PivotTable that is pivoting on the **Day of the week** and **Email address** (those are the row hierarchies).</span></span> <span data-ttu-id="0dda4-125">El recuento de **Asuntos** únicos es la información agregada que se muestra (la jerarquía de datos).</span><span class="sxs-lookup"><span data-stu-id="0dda4-125">The count of unique **Subjects** is the aggregated information being displayed (the data hierarchy).</span></span> <span data-ttu-id="0dda4-126">Haremos que nuestro script actualice esa tabla dinámica después de actualizar la tabla de correo electrónico.</span><span class="sxs-lookup"><span data-stu-id="0dda4-126">We'll have our script refresh that PivotTable after updating the email table.</span></span>

1. <span data-ttu-id="0dda4-127">Desde el **Editor de código**, seleccione **Nuevo script**.</span><span class="sxs-lookup"><span data-stu-id="0dda4-127">From within the **Code Editor**, select **New Script**.</span></span>

2. <span data-ttu-id="0dda4-128">El flujo que crearemos más adelante en el tutorial enviará la información de script de cada mensaje de correo electrónico que se reciba.</span><span class="sxs-lookup"><span data-stu-id="0dda4-128">The flow that we'll create later in the tutorial will send our script information about each email that's received.</span></span> <span data-ttu-id="0dda4-129">El script necesita aceptar esa entrada mediante parámetros en la función `main`.</span><span class="sxs-lookup"><span data-stu-id="0dda4-129">The script needs to accept that input through parameters in the `main` function.</span></span> <span data-ttu-id="0dda4-130">Reemplace el script predeterminado con el siguiente script:</span><span class="sxs-lookup"><span data-stu-id="0dda4-130">Replace the default script with the following script:</span></span>

    ```TypeScript
    function main(
      workbook: ExcelScript.Workbook,
      from: string,
      dateReceived: string,
      subject: string) {

    }
    ```

3. <span data-ttu-id="0dda4-131">El script necesita acceso a la tabla del libro de trabajo y a la tabla dinámica.</span><span class="sxs-lookup"><span data-stu-id="0dda4-131">The script needs access to the workbook's table and PivotTable.</span></span> <span data-ttu-id="0dda4-132">Agregue el siguiente código al cuerpo del script, después de la apertura `{`:</span><span class="sxs-lookup"><span data-stu-id="0dda4-132">Add the following code to the body of the script, after the opening `{`:</span></span>

    ```TypeScript
    // Get the email table.
    let emailWorksheet = workbook.getWorksheet("Emails");
    let table = emailWorksheet.getTable("EmailTable");
  
    // Get the PivotTable.
    let pivotTableWorksheet = workbook.getWorksheet("Subjects");
    let pivotTable = pivotTableWorksheet.getPivotTable("Pivot");
    ```

4. <span data-ttu-id="0dda4-133">El parámetro `dateReceived` es de tipo `string`.</span><span class="sxs-lookup"><span data-stu-id="0dda4-133">The `dateReceived` parameter is of type `string`.</span></span> <span data-ttu-id="0dda4-134">Vamos a convertir esto en un [objeto `Date`](../develop/javascript-objects.md#date) para que podamos obtener fácilmente el día de la semana.</span><span class="sxs-lookup"><span data-stu-id="0dda4-134">Let's convert that to a [`Date` object](../develop/javascript-objects.md#date) so we can easily get the day of the week.</span></span> <span data-ttu-id="0dda4-135">Después de hacerlo, deberemos asignar el valor numérico del día a una versión más legible.</span><span class="sxs-lookup"><span data-stu-id="0dda4-135">After doing that, we'll need to map the day's number value to a more readable version.</span></span> <span data-ttu-id="0dda4-136">Agregue el código siguiente al final del script, antes del cierre `}`:</span><span class="sxs-lookup"><span data-stu-id="0dda4-136">Add the following code to the end of your script, before the closing `}`:</span></span>

    ```TypeScript
      // Parse the received date string to determine the day of the week.
      let emailDate = new Date(dateReceived);
      let dayName = emailDate.toLocaleDateString("en-US", { weekday: 'long' });
    ```

5. <span data-ttu-id="0dda4-137">La cadena `subject` puede incluir la etiqueta de respuesta "RE:".</span><span class="sxs-lookup"><span data-stu-id="0dda4-137">The `subject` string may include the "RE:" reply tag.</span></span> <span data-ttu-id="0dda4-138">Eliminemos eso de la cadena para que los correos electrónicos en el mismo hilo tengan el mismo asunto para la tabla.</span><span class="sxs-lookup"><span data-stu-id="0dda4-138">Let's remove that from the string so that emails in the same thread have the same subject for the table.</span></span> <span data-ttu-id="0dda4-139">Agregue el código siguiente al final del script, antes del cierre `}`:</span><span class="sxs-lookup"><span data-stu-id="0dda4-139">Add the following code to the end of your script, before the closing `}`:</span></span>

    ```TypeScript
    // Remove the reply tag from the email subject to group emails on the same thread.
    let subjectText = subject.replace("Re: ", "");
    subjectText = subjectText.replace("RE: ", "");
    ```

6. <span data-ttu-id="0dda4-140">Ahora que se ha dado formato a los datos de correo electrónico a nuestro gusto, agreguemos una fila a la tabla de correo electrónico.</span><span class="sxs-lookup"><span data-stu-id="0dda4-140">Now that the email data has been formatted to our liking, let's add a row to the email table.</span></span> <span data-ttu-id="0dda4-141">Agregue el código siguiente al final del script, antes del cierre `}`:</span><span class="sxs-lookup"><span data-stu-id="0dda4-141">Add the following code to the end of your script, before the closing `}`:</span></span>

    ```TypeScript
    // Add the parsed text to the table.
    table.addRow(-1, [dateReceived, dayName, from, subjectText]);
    ```

7. <span data-ttu-id="0dda4-142">Por último, vamos a asegurarnos de que se actualiza la tabla dinámica.</span><span class="sxs-lookup"><span data-stu-id="0dda4-142">Finally, let's make sure the PivotTable is refreshed.</span></span> <span data-ttu-id="0dda4-143">Agregue el código siguiente al final del script, antes del cierre `}`:</span><span class="sxs-lookup"><span data-stu-id="0dda4-143">Add the following code to the end of your script, before the closing `}`:</span></span>

    ```TypeScript
    // Refresh the PivotTable to include the new row.
    pivotTable.refresh();
    ```

8. <span data-ttu-id="0dda4-144">Cambie el nombre del script **Registrar correo electrónico** y presione **Guardar script**.</span><span class="sxs-lookup"><span data-stu-id="0dda4-144">Rename your script **Record Email** and press **Save script**.</span></span>

<span data-ttu-id="0dda4-145">El script ya está preparado para un flujo de trabajo de Power Automate.</span><span class="sxs-lookup"><span data-stu-id="0dda4-145">Your script is now ready for a Power Automate workflow.</span></span> <span data-ttu-id="0dda4-146">Debería ser similar al siguiente script:</span><span class="sxs-lookup"><span data-stu-id="0dda4-146">It should look like the following script:</span></span>

```TypeScript
function main(
  workbook: ExcelScript.Workbook,
  from: string,
  dateReceived: string,
  subject: string) {
  // Get the email table.
  let emailWorksheet = workbook.getWorksheet("Emails");
  let table = emailWorksheet.getTable("EmailTable");

  // Get the PivotTable.
  let pivotTableWorksheet = workbook.getWorksheet("Subjects");
  let pivotTable = pivotTableWorksheet.getPivotTable("Pivot");

  // Parse the received date string to determine the day of the week.
  let emailDate = new Date(dateReceived);
  let dayName = emailDate.toLocaleDateString("en-US", { weekday: 'long' });

  // Remove the reply tag from the email subject to group emails on the same thread.
  let subjectText = subject.replace("Re: ", "");
  subjectText = subjectText.replace("RE: ", "");

  // Add the parsed text to the table.
  table.addRow(-1, [dateReceived, dayName, from, subjectText]);

  // Refresh the PivotTable to include the new row.
  pivotTable.refresh();
}
```

## <a name="create-an-automated-workflow-with-power-automate"></a><span data-ttu-id="0dda4-147">Crear un flujo de trabajo automatizado con Power Automate</span><span class="sxs-lookup"><span data-stu-id="0dda4-147">Create an automated workflow with Power Automate</span></span>

1. <span data-ttu-id="0dda4-148">Inicie sesión en el [sitio de Power Automate](https://flow.microsoft.com).</span><span class="sxs-lookup"><span data-stu-id="0dda4-148">Sign in to the [Power Automate site](https://flow.microsoft.com).</span></span>

2. <span data-ttu-id="0dda4-149">En el menú que se muestra en la parte izquierda de la pantalla, presione **Crear**.</span><span class="sxs-lookup"><span data-stu-id="0dda4-149">In the menu that's displayed on the left side of the screen, press **Create**.</span></span> <span data-ttu-id="0dda4-150">Se mostrará una lista de maneras de crear flujos de trabajo nuevos.</span><span class="sxs-lookup"><span data-stu-id="0dda4-150">This brings you to list of ways to create new workflows.</span></span>

    ![El botón Crear en Power Automate.](../images/power-automate-tutorial-1.png)

3. <span data-ttu-id="0dda4-152">En la sección **Inicio desde cero**, seleccione **Flujo automatizado**.</span><span class="sxs-lookup"><span data-stu-id="0dda4-152">In the **Start from blank** section, select **Automated flow**.</span></span> <span data-ttu-id="0dda4-153">Esto creará un flujo de trabajo desencadenado por un evento, como la recepción de un correo electrónico.</span><span class="sxs-lookup"><span data-stu-id="0dda4-153">This creates a workflow triggered by an event, such as receiving an email.</span></span>

    ![La opción de flujo automatizada en Power Automate.](../images/power-automate-params-tutorial-1.png)

4. <span data-ttu-id="0dda4-155">En la ventana de diálogo que aparece, escriba un nombre para su flujo en el cuadro de texto **Nombre de flujo**.</span><span class="sxs-lookup"><span data-stu-id="0dda4-155">In the dialog window that appears, enter a name for your flow in the **Flow name** text box.</span></span> <span data-ttu-id="0dda4-156">A continuación, seleccione **Cuando llegue un nuevo correo electrónico** de la lista de opciones de **Elegir el desencadenador de flujo**.</span><span class="sxs-lookup"><span data-stu-id="0dda4-156">Then select **When a new email arrives** from the list of options under **Choose your flow's trigger**.</span></span> <span data-ttu-id="0dda4-157">Es posible que tenga que buscar la opción con el cuadro de búsqueda.</span><span class="sxs-lookup"><span data-stu-id="0dda4-157">You may need to search for the option using the search box.</span></span> <span data-ttu-id="0dda4-158">Por último, pulse **Crear**.</span><span class="sxs-lookup"><span data-stu-id="0dda4-158">Finally, press **Create**.</span></span>

    ![Parte de la ventana Generar un flujo automatizado en Power Automate que muestra la opción "llega un nuevo correo electrónico".](../images/power-automate-params-tutorial-2.png)

    > [!NOTE]
    > <span data-ttu-id="0dda4-160">Este tutorial usa Outlook.</span><span class="sxs-lookup"><span data-stu-id="0dda4-160">This tutorial uses Outlook.</span></span> <span data-ttu-id="0dda4-161">Puede usar el servicio de correo electrónico que prefiera, aunque algunas opciones pueden ser diferentes.</span><span class="sxs-lookup"><span data-stu-id="0dda4-161">Feel free to use your preferred email service instead, though some options may be different.</span></span>

5. <span data-ttu-id="0dda4-162">Presione **Nuevo paso**.</span><span class="sxs-lookup"><span data-stu-id="0dda4-162">Press **New step**.</span></span>

6. <span data-ttu-id="0dda4-163">Seleccione la pestaña **Estándar** y, a continuación, seleccione **Excel Online (empresa)**.</span><span class="sxs-lookup"><span data-stu-id="0dda4-163">Select the **Standard** tab, then select **Excel Online (Business)**.</span></span>

    ![La opción Power Automate para Excel Online (empresa).](../images/power-automate-tutorial-4.png)

7. <span data-ttu-id="0dda4-165">En **Acciones**, seleccione **Ejecutar script (versión preliminar)**.</span><span class="sxs-lookup"><span data-stu-id="0dda4-165">Under **Actions**, select **Run script (preview)**.</span></span>

    ![La opción de acción de Power Automate para Ejecutar script (versión preliminar).](../images/power-automate-tutorial-5.png)

8. <span data-ttu-id="0dda4-167">A continuación, seleccione el libro, el script y los argumentos de entrada del script a usar en el paso de flujo.</span><span class="sxs-lookup"><span data-stu-id="0dda4-167">Next, you'll select the workbook, script, and script input arguments to use in the flow step.</span></span> <span data-ttu-id="0dda4-168">En el tutorial, usará el libro que creó en OneDrive, pero puede usar cualquier libro en un sitio de OneDrive o SharePoint.</span><span class="sxs-lookup"><span data-stu-id="0dda4-168">For the tutorial, you'll use the workbook you created in your OneDrive, but you could use any workbook in a OneDrive or SharePoint site.</span></span> <span data-ttu-id="0dda4-169">Especifique la siguiente configuración para el conector **Ejecutar script**:</span><span class="sxs-lookup"><span data-stu-id="0dda4-169">Specify the following settings for the **Run script** connector:</span></span>

    - <span data-ttu-id="0dda4-170">**Ubicación**: OneDrive para la Empresa</span><span class="sxs-lookup"><span data-stu-id="0dda4-170">**Location**: OneDrive for Business</span></span>
    - <span data-ttu-id="0dda4-171">**Biblioteca de documentos**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="0dda4-171">**Document Library**: OneDrive</span></span>
    - <span data-ttu-id="0dda4-172">**Archivo**: MyWorkbook.xlsx</span><span class="sxs-lookup"><span data-stu-id="0dda4-172">**File**: MyWorkbook.xlsx</span></span>
    - <span data-ttu-id="0dda4-173">**Script**: Registrar correo electrónico</span><span class="sxs-lookup"><span data-stu-id="0dda4-173">**Script**: Record Email</span></span>
    - <span data-ttu-id="0dda4-174">**de**: De *(contenido dinámico de Outlook)*</span><span class="sxs-lookup"><span data-stu-id="0dda4-174">**from**: From *(dynamic content from Outlook)*</span></span>
    - <span data-ttu-id="0dda4-175">**dateReceived**: Hora de recepción *(contenido dinámico de Outlook)*</span><span class="sxs-lookup"><span data-stu-id="0dda4-175">**dateReceived**: Received Time *(dynamic content from Outlook)*</span></span>
    - <span data-ttu-id="0dda4-176">**asunto**: Asunto *(contenido dinámico de Outlook)*</span><span class="sxs-lookup"><span data-stu-id="0dda4-176">**subject**: Subject *(dynamic content from Outlook)*</span></span>

    <span data-ttu-id="0dda4-177">*Tenga en cuenta que los parámetros del script solo aparecen cuando se selecciona el script.*</span><span class="sxs-lookup"><span data-stu-id="0dda4-177">*Note that the parameters for the script will only appear once the script is selected.*</span></span>

    ![La opción de acción de Power Automate para Ejecutar script (versión preliminar).](../images/power-automate-params-tutorial-3.png)

9. <span data-ttu-id="0dda4-179">Presione **Guardar**.</span><span class="sxs-lookup"><span data-stu-id="0dda4-179">Press **Save**.</span></span>

<span data-ttu-id="0dda4-180">El flujo está ahora habilitado.</span><span class="sxs-lookup"><span data-stu-id="0dda4-180">Your flow is now enabled.</span></span> <span data-ttu-id="0dda4-181">El script se ejecutará automáticamente cada vez que reciba un correo electrónico a través de Outlook.</span><span class="sxs-lookup"><span data-stu-id="0dda4-181">It will automatically run your script each time you receive an email through Outlook.</span></span>

## <a name="manage-the-script-in-power-automate"></a><span data-ttu-id="0dda4-182">Administrar el script en Power Automate</span><span class="sxs-lookup"><span data-stu-id="0dda4-182">Manage the script in Power Automate</span></span>

1. <span data-ttu-id="0dda4-183">En la página principal de Power Automate, seleccione **Mis flujos**.</span><span class="sxs-lookup"><span data-stu-id="0dda4-183">From the main Power Automate page, select **My flows**.</span></span>

    ![El botón Mis flujos en Power Automate.](../images/power-automate-tutorial-7.png)

2. <span data-ttu-id="0dda4-185">Seleccione el flujo.</span><span class="sxs-lookup"><span data-stu-id="0dda4-185">Select your flow.</span></span> <span data-ttu-id="0dda4-186">Aquí puede ver el historial de ejecución.</span><span class="sxs-lookup"><span data-stu-id="0dda4-186">Here you can see the run history.</span></span> <span data-ttu-id="0dda4-187">Puede actualizar la página o presionar el botón actualizar **Todas las ejecuciones** para actualizar el historial.</span><span class="sxs-lookup"><span data-stu-id="0dda4-187">You can refresh the page or press the refresh **All runs** button to update the history.</span></span> <span data-ttu-id="0dda4-188">El flujo se desencadenará poco después de que se reciba un correo electrónico.</span><span class="sxs-lookup"><span data-stu-id="0dda4-188">The flow will trigger shortly after an email is received.</span></span> <span data-ttu-id="0dda4-189">Pruebe el flujo enviándose un correo electrónico a sí mismo.</span><span class="sxs-lookup"><span data-stu-id="0dda4-189">Test the flow by sending yourself mail.</span></span>

<span data-ttu-id="0dda4-190">Cuando se desencadene el flujo y se ejecute correctamente el script, debería ver que se actualizan la tabla dinámica y la tabla del libro.</span><span class="sxs-lookup"><span data-stu-id="0dda4-190">When the flow is triggered and successfully runs your script, you should see the workbook's table and PivotTable update.</span></span>

![La tabla de correo electrónico después de ejecutar el flujo un par de veces.](../images/power-automate-params-tutorial-4.png)

![La tabla dinámica después de ejecutar el flujo un par de veces.](../images/power-automate-params-tutorial-5.png)

## <a name="next-steps"></a><span data-ttu-id="0dda4-193">Pasos siguientes</span><span class="sxs-lookup"><span data-stu-id="0dda4-193">Next steps</span></span>

<span data-ttu-id="0dda4-194">Visite [Ejecutar scripts de Office con Power Automate](../develop/power-automate-integration.md) para más información sobre la conexión de Scripts de Office con Power Automate.</span><span class="sxs-lookup"><span data-stu-id="0dda4-194">Visit [Run Office Scripts with Power Automate](../develop/power-automate-integration.md) to learn more about connecting Office Scripts with Power Automate.</span></span>

<span data-ttu-id="0dda4-195">También puede consultar el [Escenario de muestra de recordatorios de tareas automatizados](../resources/scenarios/task-reminders.md) para aprender a combinar los Scripts de Office y Power Automate con las Tarjetas adaptables de Teams.</span><span class="sxs-lookup"><span data-stu-id="0dda4-195">You can also check out the [Automated task reminders sample scenario](../resources/scenarios/task-reminders.md) to learn how to combine Office Scripts and Power Automate with Teams Adaptive Cards.</span></span>
