---
title: Pasar datos a scripts en un flujo de Power Automate ejecutado automáticamente
description: Un tutorial sobre la ejecución de Scripts de Office para Excel en la Web mediante Power Automate cuando se reciba el correo y el paso de datos de flujo al script.
ms.date: 07/14/2020
localization_priority: Priority
ms.openlocfilehash: c024891e187f22b7d10f6e9d52d262dc2ec4057f
ms.sourcegitcommit: ebd1079c7e2695ac0e7e4c616f2439975e196875
ms.translationtype: HT
ms.contentlocale: es-ES
ms.lasthandoff: 07/17/2020
ms.locfileid: "45160484"
---
# <a name="pass-data-to-scripts-in-an-automatically-run-power-automate-flow-preview"></a><span data-ttu-id="92c6c-103">Pasar datos a scripts en un flujo de Power Automate ejecutado automáticamente (versión preliminar)</span><span class="sxs-lookup"><span data-stu-id="92c6c-103">Pass data to scripts in an automatically-run Power Automate flow (preview)</span></span>

<span data-ttu-id="92c6c-104">Este tutorial le enseña cómo usar un script de Office para Excel en la web con un flujo de trabajo automatizado de [Power Automate](https://flow.microsoft.com).</span><span class="sxs-lookup"><span data-stu-id="92c6c-104">This tutorial teaches you how to use an Office Script for Excel on the web with an automated [Power Automate](https://flow.microsoft.com) workflow.</span></span> <span data-ttu-id="92c6c-105">El script se ejecutará automáticamente cada vez que reciba un correo electrónico, grabando información del correo en un libro de Excel.</span><span class="sxs-lookup"><span data-stu-id="92c6c-105">Your script will automatically run each time you receive an email, recording information from the email in an Excel workbook.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="92c6c-106">Requisitos previos</span><span class="sxs-lookup"><span data-stu-id="92c6c-106">Prerequisites</span></span>

[!INCLUDE [Tutorial prerequisites](../includes/power-automate-tutorial-prerequisites.md)]

> [!IMPORTANT]
> <span data-ttu-id="92c6c-107">Este tutorial presupone que ha completado el tutorial [Ejecutar Scripts de Office en Excel en la Web con Power Automate](excel-power-automate-manual.md).</span><span class="sxs-lookup"><span data-stu-id="92c6c-107">This tutorial assumes you have completed the [Run Office Scripts in Excel on the web with Power Automate](excel-power-automate-manual.md) tutorial.</span></span>

## <a name="prepare-the-workbook"></a><span data-ttu-id="92c6c-108">Preparar el libro</span><span class="sxs-lookup"><span data-stu-id="92c6c-108">Prepare the workbook</span></span>

<span data-ttu-id="92c6c-109">Power Automate no puede usar [referencias relativas](../develop/power-automate-integration.md#avoid-using-relative-references) como `Workbook.getActiveWorksheet` para acceder a componentes de libros de trabajo.</span><span class="sxs-lookup"><span data-stu-id="92c6c-109">Power Automate can't use [relative references](../develop/power-automate-integration.md#avoid-using-relative-references) like `Workbook.getActiveWorksheet` to access workbook components.</span></span> <span data-ttu-id="92c6c-110">Por lo tanto, es necesario un libro de trabajo y una hoja de cálculo con nombres coherentes para que Power Automate haga referencia.</span><span class="sxs-lookup"><span data-stu-id="92c6c-110">So, we need a workbook and worksheet with consistent names for Power Automate to reference.</span></span>

1. <span data-ttu-id="92c6c-111">Cree un nuevo libro de trabajo y llámelo **Mi libro de trabajo**.</span><span class="sxs-lookup"><span data-stu-id="92c6c-111">Create a new workbook named **MyWorkbook**.</span></span>

2. <span data-ttu-id="92c6c-112">Vaya a la pestaña **Automatizar** y seleccione **Editor de código**.</span><span class="sxs-lookup"><span data-stu-id="92c6c-112">Go to the **Automate** tab and select **Code Editor**.</span></span>

3. <span data-ttu-id="92c6c-113">Seleccione **Nuevo script**.</span><span class="sxs-lookup"><span data-stu-id="92c6c-113">Select **New Script**.</span></span>

4. <span data-ttu-id="92c6c-114">Reemplace el código existente con el siguiente script y presione **Ejecutar**.</span><span class="sxs-lookup"><span data-stu-id="92c6c-114">Replace the existing code with the following script and press **Run**.</span></span> <span data-ttu-id="92c6c-115">Esto configurará el libro con nombres de tabla dinámica, hoja de cálculo y tabla coherentes.</span><span class="sxs-lookup"><span data-stu-id="92c6c-115">This will setup the workbook with consistent worksheet, table, and PivotTable names.</span></span>

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
      let pivotWorksheet = workbook.addWorksheet("SubjectPivot");
      let newPivotTable = workbook.addPivotTable("Pivot", "EmailTable", pivotWorksheet.getRange("A3:C20"));

      // Setup the pivot hierarchies
      newPivotTable.addRowHierarchy(newPivotTable.getHierarchy("Day of the week"));
      newPivotTable.addRowHierarchy(newPivotTable.getHierarchy("Email address"));
      newPivotTable.addDataHierarchy(newPivotTable.getHierarchy("Subject"));
    }
    ```

## <a name="create-an-office-script-for-your-automated-workflow"></a><span data-ttu-id="92c6c-116">Crear un script de Office para el flujo de trabajo automatizado</span><span class="sxs-lookup"><span data-stu-id="92c6c-116">Create an Office Script for your automated workflow</span></span>

<span data-ttu-id="92c6c-117">Comencemos a crear un script que registre información de un correo electrónico.</span><span class="sxs-lookup"><span data-stu-id="92c6c-117">Let's create a script that logs information from an email.</span></span> <span data-ttu-id="92c6c-118">Queremos saber cuál es el número de días de la semana que recibimos más correo electrónico y cuántos remitentes únicos envían ese correo.</span><span class="sxs-lookup"><span data-stu-id="92c6c-118">We want to know how which days of the week we receive the most mail and how many unique senders are sending that mail.</span></span> <span data-ttu-id="92c6c-119">Nuestro libro tiene una tabla con columnas de **Fecha**, **Día de la semana**, **Dirección de correo electrónico** y **Asunto**.</span><span class="sxs-lookup"><span data-stu-id="92c6c-119">Our workbook has a table with **Date**, **Day of the week**, **Email address**, and **Subject** columns.</span></span> <span data-ttu-id="92c6c-120">Nuestra hoja de cálculo también tiene una tabla dinámica que se dinamiza en el **Día de la semana** y **Dirección de correo electrónico** (que son las jerarquías de fila).</span><span class="sxs-lookup"><span data-stu-id="92c6c-120">Our worksheet also has a PivotTable that is pivoting on the **Day of the week** and **Email address** (those are the row hierarchies).</span></span> <span data-ttu-id="92c6c-121">El recuento de **Asuntos** únicos es la información agregada que se muestra (la jerarquía de datos).</span><span class="sxs-lookup"><span data-stu-id="92c6c-121">The count of unique **Subjects** is the aggregated information being displayed (the data hierarchy).</span></span> <span data-ttu-id="92c6c-122">Haremos que nuestro script actualice esa tabla dinámica después de actualizar la tabla de correo electrónico.</span><span class="sxs-lookup"><span data-stu-id="92c6c-122">We'll have our script refresh that PivotTable after updating the email table.</span></span>

1. <span data-ttu-id="92c6c-123">Desde el **Editor de código**, seleccione **Nuevo script**.</span><span class="sxs-lookup"><span data-stu-id="92c6c-123">From within the **Code Editor**, select **New Script**.</span></span>

2. <span data-ttu-id="92c6c-124">El flujo que crearemos más adelante en el tutorial enviará la información de script de cada mensaje de correo electrónico que se reciba.</span><span class="sxs-lookup"><span data-stu-id="92c6c-124">The flow that we'll create later in the tutorial will send our script information about each email that's received.</span></span> <span data-ttu-id="92c6c-125">El script necesita aceptar esa entrada mediante parámetros en la función `main`.</span><span class="sxs-lookup"><span data-stu-id="92c6c-125">The script needs to accept that input through parameters in the `main` function.</span></span> <span data-ttu-id="92c6c-126">Reemplace el script predeterminado con el siguiente script:</span><span class="sxs-lookup"><span data-stu-id="92c6c-126">Replace the default script with the following script:</span></span>

    ```TypeScript
    function main(
      workbook: ExcelScript.Workbook,
      from: string,
      dateReceived: string,
      subject: string) {

    }
    ```

3. <span data-ttu-id="92c6c-127">El script necesita acceso a la tabla del libro de trabajo y a la tabla dinámica.</span><span class="sxs-lookup"><span data-stu-id="92c6c-127">The script needs access to the workbook's table and PivotTable.</span></span> <span data-ttu-id="92c6c-128">Agregue el siguiente código al cuerpo del script, después de la apertura `{`:</span><span class="sxs-lookup"><span data-stu-id="92c6c-128">Add the following code to the body of the script, after the opening `{`:</span></span>

    ```TypeScript
    // Get the email table.
    let emailWorksheet = workbook.getWorksheet("Emails");
    let table = emailWorksheet.getTable("EmailTable");
  
    // Get the PivotTable.
    let pivotTableWorksheet = workbook.getWorksheet("SubjectPivot");
    let pivotTable = pivotTableWorksheet.getPivotTable("Pivot");
    ```

4. <span data-ttu-id="92c6c-129">El parámetro `dateReceived` es de tipo `string`.</span><span class="sxs-lookup"><span data-stu-id="92c6c-129">The `dateReceived` parameter is of type `string`.</span></span> <span data-ttu-id="92c6c-130">Vamos a convertir esto en un [objeto `Date`](../develop/javascript-objects.md#date) para que podamos obtener fácilmente el día de la semana.</span><span class="sxs-lookup"><span data-stu-id="92c6c-130">Let's convert that to a [`Date` object](../develop/javascript-objects.md#date) so we can easily get the day of the week.</span></span> <span data-ttu-id="92c6c-131">Después de hacerlo, deberemos asignar el valor numérico del día a una versión más legible.</span><span class="sxs-lookup"><span data-stu-id="92c6c-131">After doing that, we'll need to map the day's number value to a more readable version.</span></span> <span data-ttu-id="92c6c-132">Agregue el código siguiente al final del script, antes del cierre `}`:</span><span class="sxs-lookup"><span data-stu-id="92c6c-132">Add the following code to the end of your script, before the closing `}`:</span></span>

    ```TypeScript
    // Parse the received date string.
    let date = new Date(dateReceived);

    // Convert number representing the day of the week into the name of the day.
    let dayText : string;
    switch (date.getDay()) {
      case 0:
        dayText = "Sunday";
        break;
      case 1:
        dayText = "Monday";
        break;
      case 2:
        dayText = "Tuesday";
        break;
      case 3:
        dayText = "Wednesday";
        break;
      case 4:
        dayText = "Thursday";
        break;
      case 5:
        dayText = "Friday";
        break;
      default:
        dayText = "Saturday";
        break;
    }
    ```

5. <span data-ttu-id="92c6c-133">La cadena `subject` puede incluir la etiqueta de respuesta "RE:".</span><span class="sxs-lookup"><span data-stu-id="92c6c-133">The `subject` string may include the "RE:" reply tag.</span></span> <span data-ttu-id="92c6c-134">Eliminemos eso de la cadena para que los correos electrónicos en el mismo hilo tengan el mismo asunto para la tabla.</span><span class="sxs-lookup"><span data-stu-id="92c6c-134">Let's remove that from the string so that emails in the same thread have the same subject for the table.</span></span> <span data-ttu-id="92c6c-135">Agregue el código siguiente al final del script, antes del cierre `}`:</span><span class="sxs-lookup"><span data-stu-id="92c6c-135">Add the following code to the end of your script, before the closing `}`:</span></span>

    ```TypeScript
    // Remove the reply tag from the email subject to group emails on the same thread.
    let subjectText = subject.replace("Re: ", "");
    subjectText = subjectText.replace("RE: ", "");
    ```

6. <span data-ttu-id="92c6c-136">Ahora que se ha dado formato a los datos de correo electrónico a nuestro gusto, agreguemos una fila a la tabla de correo electrónico.</span><span class="sxs-lookup"><span data-stu-id="92c6c-136">Now that the email data has been formatted to our liking, let's add a row to the email table.</span></span> <span data-ttu-id="92c6c-137">Agregue el código siguiente al final del script, antes del cierre `}`:</span><span class="sxs-lookup"><span data-stu-id="92c6c-137">Add the following code to the end of your script, before the closing `}`:</span></span>

    ```TypeScript
    // Add the parsed text to the table.
    table.addRow(-1, [dateReceived, dayText, from, subjectText]);
    ```

7. <span data-ttu-id="92c6c-138">Por último, vamos a asegurarnos de que se actualiza la tabla dinámica.</span><span class="sxs-lookup"><span data-stu-id="92c6c-138">Finally, let's make sure the PivotTable is refreshed.</span></span> <span data-ttu-id="92c6c-139">Agregue el código siguiente al final del script, antes del cierre `}`:</span><span class="sxs-lookup"><span data-stu-id="92c6c-139">Add the following code to the end of your script, before the closing `}`:</span></span>

    ```TypeScript
    // Refresh the PivotTable to include the new row.
    pivotTable.refresh();
    ```

8. <span data-ttu-id="92c6c-140">Cambie el nombre del script **Registrar correo electrónico** y presione **Guardar script**.</span><span class="sxs-lookup"><span data-stu-id="92c6c-140">Rename your script **Record Email** and press **Save script**.</span></span>

<span data-ttu-id="92c6c-141">El script ya está preparado para un flujo de trabajo de Power Automate.</span><span class="sxs-lookup"><span data-stu-id="92c6c-141">Your script is now ready for a Power Automate workflow.</span></span> <span data-ttu-id="92c6c-142">Debería ser similar al siguiente script:</span><span class="sxs-lookup"><span data-stu-id="92c6c-142">It should look like the following script:</span></span>

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
  let pivotTableWorksheet = workbook.getWorksheet("Pivot");
  let pivotTable = pivotTableWorksheet.getPivotTable("SubjectPivot");

  // Parse the received date string.
  let date = new Date(dateReceived);

  // Convert number representing the day of the week into the name of the day.
  let dayText: string;
  switch (date.getDay()) {
    case 0:
      dayText = "Sunday";
      break;
    case 1:
      dayText = "Monday";
      break;
    case 2:
      dayText = "Tuesday";
      break;
    case 3:
      dayText = "Wednesday";
      break;
    case 4:
      dayText = "Thursday";
      break;
    case 5:
      dayText = "Friday";
      break;
    default:
      dayText = "Saturday";
      break;
  }

  // Remove the reply tag from the email subject to group emails on the same thread.
  let subjectText = subject.replace("Re: ", "");
  subjectText = subjectText.replace("RE: ", "");

  // Add the parsed text to the table.
  table.addRow(-1, [dateReceived, dayText, from, subjectText]);

  // Refresh the PivotTable to include the new row.
  pivotTable.refresh();
}
```

## <a name="create-an-automated-workflow-with-power-automate"></a><span data-ttu-id="92c6c-143">Crear un flujo de trabajo automatizado con Power Automate</span><span class="sxs-lookup"><span data-stu-id="92c6c-143">Create an automated workflow with Power Automate</span></span>

1. <span data-ttu-id="92c6c-144">Inicie sesión en el [sitio de Power Automate](https://flow.microsoft.com).</span><span class="sxs-lookup"><span data-stu-id="92c6c-144">Sign in to the [Power Automate site](https://flow.microsoft.com).</span></span>

2. <span data-ttu-id="92c6c-145">En el menú que se muestra en la parte izquierda de la pantalla, presione **Crear**.</span><span class="sxs-lookup"><span data-stu-id="92c6c-145">In the menu that's displayed on the left side of the screen, press **Create**.</span></span> <span data-ttu-id="92c6c-146">Se mostrará una lista de maneras de crear flujos de trabajo nuevos.</span><span class="sxs-lookup"><span data-stu-id="92c6c-146">This brings you to list of ways to create new workflows.</span></span>

    ![El botón Crear en Power Automate.](../images/power-automate-tutorial-1.png)

3. <span data-ttu-id="92c6c-148">En la sección **Inicio desde cero**, seleccione **Flujo automatizado**.</span><span class="sxs-lookup"><span data-stu-id="92c6c-148">In the **Start from blank** section, select **Automated flow**.</span></span> <span data-ttu-id="92c6c-149">Esto creará un flujo de trabajo desencadenado por un evento, como la recepción de un correo electrónico.</span><span class="sxs-lookup"><span data-stu-id="92c6c-149">This creates a workflow triggered by an event, such as receiving an email.</span></span>

    ![La opción de flujo automatizada en Power Automate.](../images/power-automate-params-tutorial-1.png)

4. <span data-ttu-id="92c6c-151">En la ventana de diálogo que aparece, escriba un nombre para su flujo en el cuadro de texto **Nombre de flujo**.</span><span class="sxs-lookup"><span data-stu-id="92c6c-151">In the dialog window that appears, enter a name for your flow in the **Flow name** text box.</span></span> <span data-ttu-id="92c6c-152">A continuación, seleccione **Cuando llegue un nuevo correo electrónico** de la lista de opciones de **Elegir el desencadenador de flujo**.</span><span class="sxs-lookup"><span data-stu-id="92c6c-152">Then select **When a new email arrives** from the list of options under **Choose your flow's trigger**.</span></span> <span data-ttu-id="92c6c-153">Es posible que tenga que buscar la opción con el cuadro de búsqueda.</span><span class="sxs-lookup"><span data-stu-id="92c6c-153">You may need to search for the option using the search box.</span></span> <span data-ttu-id="92c6c-154">Por último, pulse **Crear**.</span><span class="sxs-lookup"><span data-stu-id="92c6c-154">Finally, press **Create**.</span></span>

    ![Parte de la ventana Generar un flujo automatizado en Power Automate que muestra la opción "llega un nuevo correo electrónico".](../images/power-automate-params-tutorial-2.png)

    > [!NOTE]
    > <span data-ttu-id="92c6c-156">Este tutorial usa Outlook.</span><span class="sxs-lookup"><span data-stu-id="92c6c-156">This tutorial uses Outlook.</span></span> <span data-ttu-id="92c6c-157">Puede usar el servicio de correo electrónico que prefiera, aunque algunas opciones pueden ser diferentes.</span><span class="sxs-lookup"><span data-stu-id="92c6c-157">Feel free to use your preferred email service instead, though some options may be different.</span></span>

5. <span data-ttu-id="92c6c-158">Presione **Nuevo paso**.</span><span class="sxs-lookup"><span data-stu-id="92c6c-158">Press **New step**.</span></span>

6. <span data-ttu-id="92c6c-159">Seleccione la pestaña **Estándar** y, a continuación, seleccione **Excel Online (empresa)**.</span><span class="sxs-lookup"><span data-stu-id="92c6c-159">Select the **Standard** tab, then select **Excel Online (Business)**.</span></span>

    ![La opción Power Automate para Excel Online (empresa).](../images/power-automate-tutorial-4.png)

7. <span data-ttu-id="92c6c-161">En **Acciones**, seleccione **Ejecutar script (versión preliminar)**.</span><span class="sxs-lookup"><span data-stu-id="92c6c-161">Under **Actions**, select **Run script (preview)**.</span></span>

    ![La opción de acción de Power Automate para Ejecutar script (versión preliminar).](../images/power-automate-tutorial-5.png)

8. <span data-ttu-id="92c6c-163">Especifique la siguiente configuración para el conector **Ejecutar script**:</span><span class="sxs-lookup"><span data-stu-id="92c6c-163">Specify the following settings for the **Run script** connector:</span></span>

    - <span data-ttu-id="92c6c-164">**Ubicación**: OneDrive para la Empresa</span><span class="sxs-lookup"><span data-stu-id="92c6c-164">**Location**: OneDrive for Business</span></span>
    - <span data-ttu-id="92c6c-165">**Biblioteca de documentos**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="92c6c-165">**Document Library**: OneDrive</span></span>
    - <span data-ttu-id="92c6c-166">**Archivo**: MyWorkbook.xlsx</span><span class="sxs-lookup"><span data-stu-id="92c6c-166">**File**: MyWorkbook.xlsx</span></span>
    - <span data-ttu-id="92c6c-167">**Script**: Registrar correo electrónico</span><span class="sxs-lookup"><span data-stu-id="92c6c-167">**Script**: Record Email</span></span>
    - <span data-ttu-id="92c6c-168">**de**: De *(contenido dinámico de Outlook)*</span><span class="sxs-lookup"><span data-stu-id="92c6c-168">**from**: From *(dynamic content from Outlook)*</span></span>
    - <span data-ttu-id="92c6c-169">**dateReceived**: Hora de recepción *(contenido dinámico de Outlook)*</span><span class="sxs-lookup"><span data-stu-id="92c6c-169">**dateReceived**: Received Time *(dynamic content from Outlook)*</span></span>
    - <span data-ttu-id="92c6c-170">**asunto**: Asunto *(contenido dinámico de Outlook)*</span><span class="sxs-lookup"><span data-stu-id="92c6c-170">**subject**: Subject *(dynamic content from Outlook)*</span></span>

    <span data-ttu-id="92c6c-171">*Tenga en cuenta que los parámetros del script solo aparecen cuando se selecciona el script.*</span><span class="sxs-lookup"><span data-stu-id="92c6c-171">*Note that the parameters for the script will only appear once the script is selected.*</span></span>

    ![La opción de acción de Power Automate para Ejecutar script (versión preliminar).](../images/power-automate-params-tutorial-3.png)

9. <span data-ttu-id="92c6c-173">Presione **Guardar**.</span><span class="sxs-lookup"><span data-stu-id="92c6c-173">Press **Save**.</span></span>

<span data-ttu-id="92c6c-174">El flujo está ahora habilitado.</span><span class="sxs-lookup"><span data-stu-id="92c6c-174">Your flow is now enabled.</span></span> <span data-ttu-id="92c6c-175">El script se ejecutará automáticamente cada vez que reciba un correo electrónico a través de Outlook.</span><span class="sxs-lookup"><span data-stu-id="92c6c-175">It will automatically run your script each time you receive an email through Outlook.</span></span>

## <a name="manage-the-script-in-power-automate"></a><span data-ttu-id="92c6c-176">Administrar el script en Power Automate</span><span class="sxs-lookup"><span data-stu-id="92c6c-176">Manage the script in Power Automate</span></span>

1. <span data-ttu-id="92c6c-177">En la página principal de Power Automate, seleccione **Mis flujos**.</span><span class="sxs-lookup"><span data-stu-id="92c6c-177">From the main Power Automate page, select **My flows**.</span></span>

    ![El botón Mis flujos en Power Automate.](../images/power-automate-tutorial-7.png)

2. <span data-ttu-id="92c6c-179">Seleccione el flujo.</span><span class="sxs-lookup"><span data-stu-id="92c6c-179">Select your flow.</span></span> <span data-ttu-id="92c6c-180">Aquí puede ver el historial de ejecución.</span><span class="sxs-lookup"><span data-stu-id="92c6c-180">Here you can see the run history.</span></span> <span data-ttu-id="92c6c-181">Puede actualizar la página o presionar el botón actualizar **Todas las ejecuciones** para actualizar el historial.</span><span class="sxs-lookup"><span data-stu-id="92c6c-181">You can refresh the page or press the refresh **All runs** button to update the history.</span></span> <span data-ttu-id="92c6c-182">El flujo se desencadenará poco después de que se reciba un correo electrónico.</span><span class="sxs-lookup"><span data-stu-id="92c6c-182">The flow will trigger shortly after an email is received.</span></span> <span data-ttu-id="92c6c-183">Pruebe el flujo enviándose un correo electrónico a sí mismo.</span><span class="sxs-lookup"><span data-stu-id="92c6c-183">Test the flow by sending yourself mail.</span></span>

<span data-ttu-id="92c6c-184">Cuando se desencadene el flujo y se ejecute correctamente el script, debería ver que se actualizan la tabla dinámica y la tabla del libro.</span><span class="sxs-lookup"><span data-stu-id="92c6c-184">When the flow is triggered and successfully runs your script, you should see the workbook's table and PivotTable update.</span></span>

![La tabla de correo electrónico después de ejecutar el flujo un par de veces.](../images/power-automate-params-tutorial-4.png)

![La tabla dinámica después de ejecutar el flujo un par de veces.](../images/power-automate-params-tutorial-5.png)

## <a name="next-steps"></a><span data-ttu-id="92c6c-187">Pasos siguientes</span><span class="sxs-lookup"><span data-stu-id="92c6c-187">Next steps</span></span>

<span data-ttu-id="92c6c-188">Visite [Ejecutar scripts de Office con Power Automate](../develop/power-automate-integration.md) para más información sobre la conexión de Scripts de Office con Power Automate.</span><span class="sxs-lookup"><span data-stu-id="92c6c-188">Visit [Run Office Scripts with Power Automate](../develop/power-automate-integration.md) to learn more about connecting Office Scripts with Power Automate.</span></span>

<span data-ttu-id="92c6c-189">También puede consultar el [Escenario de muestra de recordatorios de tareas automatizados](../resources/scenarios/task-reminders.md) para aprender a combinar los Scripts de Office y Power Automate con las Tarjetas adaptables de Teams.</span><span class="sxs-lookup"><span data-stu-id="92c6c-189">You can also check out the [Automated task reminders sample scenario](../resources/scenarios/task-reminders.md) to learn how to combine Office Scripts and Power Automate with Teams Adaptive Cards.</span></span>
