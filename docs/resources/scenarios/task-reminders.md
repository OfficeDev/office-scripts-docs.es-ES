---
title: 'Escenario de ejemplo de scripts de Office: avisos de tareas automatizadas'
description: Un ejemplo que usa la automatización de energía y las tarjetas adaptables automatiza los avisos de tareas en una hoja de cálculo de administración de proyectos.
ms.date: 06/09/2020
localization_priority: Normal
ms.openlocfilehash: f764c37dafdd964e9435d504770d10b1608428b8
ms.sourcegitcommit: aec3c971c6640429f89b6bb99d2c95ea06725599
ms.translationtype: Auto
ms.contentlocale: es-ES
ms.lasthandoff: 06/25/2020
ms.locfileid: "44878909"
---
# <a name="office-scripts-sample-scenario-automated-task-reminders"></a><span data-ttu-id="2e77e-103">Escenario de ejemplo de scripts de Office: avisos de tareas automatizadas</span><span class="sxs-lookup"><span data-stu-id="2e77e-103">Office Scripts sample scenario: Automated task reminders</span></span>

<span data-ttu-id="2e77e-104">En este escenario, está administrando un proyecto.</span><span class="sxs-lookup"><span data-stu-id="2e77e-104">In this scenario you're managing a project.</span></span> <span data-ttu-id="2e77e-105">Use una hoja de cálculo de Excel para realizar un seguimiento del estado de sus empleados cada mes.</span><span class="sxs-lookup"><span data-stu-id="2e77e-105">You use an Excel worksheet to track your employees' status every month.</span></span> <span data-ttu-id="2e77e-106">Con frecuencia, debe recordar a los usuarios que deben rellenar su estado, por lo que decidió automatizar el proceso de recordatorio.</span><span class="sxs-lookup"><span data-stu-id="2e77e-106">You often need to remind people to fill out their status, so you've decided to automate that reminder process.</span></span>

<span data-ttu-id="2e77e-107">Creará un flujo de automatización de potencia para avisar a los usuarios de los campos de estado que faltan y aplicar sus respuestas a la hoja de cálculo.</span><span class="sxs-lookup"><span data-stu-id="2e77e-107">You'll create a Power Automate flow to message people with missing status fields and apply their responses to the spreadsheet.</span></span> <span data-ttu-id="2e77e-108">Para ello, desarrollará un par de scripts para controlar el trabajo con el libro.</span><span class="sxs-lookup"><span data-stu-id="2e77e-108">To do this, you'll develop a pair of scripts to handle the working with the workbook.</span></span> <span data-ttu-id="2e77e-109">El primer script obtiene una lista de personas con Estados en blanco y el segundo script agrega una cadena de estado a la fila de la derecha.</span><span class="sxs-lookup"><span data-stu-id="2e77e-109">The first script gets a list of people with blank statuses and the second script adds a status string to the right row.</span></span> <span data-ttu-id="2e77e-110">También usará las [tarjetas adaptables de Microsoft Teams](/microsoftteams/platform/task-modules-and-cards/what-are-cards) para que los empleados especifiquen su estado directamente desde la notificación.</span><span class="sxs-lookup"><span data-stu-id="2e77e-110">You'll also make use of [Teams Adaptive Cards](/microsoftteams/platform/task-modules-and-cards/what-are-cards) to have employees enter their status directly from the notification.</span></span>

## <a name="scripting-skills-covered"></a><span data-ttu-id="2e77e-111">Habilidades de scripting cubiertas</span><span class="sxs-lookup"><span data-stu-id="2e77e-111">Scripting skills covered</span></span>

- <span data-ttu-id="2e77e-112">Crear flujos con Power automatization</span><span class="sxs-lookup"><span data-stu-id="2e77e-112">Create flows in Power Automate</span></span>
- <span data-ttu-id="2e77e-113">Pasar datos a scripts</span><span class="sxs-lookup"><span data-stu-id="2e77e-113">Pass data to scripts</span></span>
- <span data-ttu-id="2e77e-114">Devolver datos de scripts</span><span class="sxs-lookup"><span data-stu-id="2e77e-114">Return data from scripts</span></span>
- <span data-ttu-id="2e77e-115">Tarjetas adaptables de Teams</span><span class="sxs-lookup"><span data-stu-id="2e77e-115">Teams Adaptive Cards</span></span>
- <span data-ttu-id="2e77e-116">Tablas</span><span class="sxs-lookup"><span data-stu-id="2e77e-116">Tables</span></span>

## <a name="prerequisites"></a><span data-ttu-id="2e77e-117">Requisitos previos</span><span class="sxs-lookup"><span data-stu-id="2e77e-117">Prerequisites</span></span>

<span data-ttu-id="2e77e-118">Este escenario usa la [funcionalidad de automatización](https://flow.microsoft.com) y [Microsoft Teams](https://www.microsoft.com/microsoft-365/microsoft-teams/group-chat-software).</span><span class="sxs-lookup"><span data-stu-id="2e77e-118">This scenario uses [Power Automate](https://flow.microsoft.com) and [Microsoft Teams](https://www.microsoft.com/microsoft-365/microsoft-teams/group-chat-software).</span></span> <span data-ttu-id="2e77e-119">Necesitará ambos asociados con la cuenta que usa para el desarrollo de scripts de Office.</span><span class="sxs-lookup"><span data-stu-id="2e77e-119">You will need both associated with the account that you use for developing Office Scripts.</span></span> <span data-ttu-id="2e77e-120">Para obtener acceso gratuito a una suscripción de Microsoft Developer para conocer y trabajar con estas aplicaciones, considere la posibilidad de unirse al [programa de desarrolladores de microsoft 365](https://developer.microsoft.com/microsoft-365/dev-program).</span><span class="sxs-lookup"><span data-stu-id="2e77e-120">For free access to a Microsoft Developer subscription to learn about and work with these applications, consider joining the [Microsoft 365 Developer Program](https://developer.microsoft.com/microsoft-365/dev-program).</span></span>

## <a name="setup-instructions"></a><span data-ttu-id="2e77e-121">Instrucciones de instalación</span><span class="sxs-lookup"><span data-stu-id="2e77e-121">Setup instructions</span></span>

1. <span data-ttu-id="2e77e-122">Descargue <a href="task-reminders.xlsx">task-reminders.xlsx</a> a su OneDrive.</span><span class="sxs-lookup"><span data-stu-id="2e77e-122">Download <a href="task-reminders.xlsx">task-reminders.xlsx</a> to your OneDrive.</span></span>

2. <span data-ttu-id="2e77e-123">Abra el libro en Excel en la Web.</span><span class="sxs-lookup"><span data-stu-id="2e77e-123">Open the workbook in Excel on the web.</span></span>

3. <span data-ttu-id="2e77e-124">En la ficha **automatizar** , abra el **Editor de código**.</span><span class="sxs-lookup"><span data-stu-id="2e77e-124">Under the **Automate** tab, open the **Code Editor**.</span></span>

4. <span data-ttu-id="2e77e-125">En primer lugar, necesitamos un script para obtener todos los empleados con informes de estado que faltan en la hoja de cálculo.</span><span class="sxs-lookup"><span data-stu-id="2e77e-125">First, we need a script to get all the employees with status reports that are missing from the spreadsheet.</span></span> <span data-ttu-id="2e77e-126">En el panel de tareas **Editor de código** , presione **nueva secuencia** de comandos y pegue el siguiente script en el editor.</span><span class="sxs-lookup"><span data-stu-id="2e77e-126">In the **Code Editor** task pane, press **New Script** and paste the following script into the editor.</span></span>

    ```typescript
    /**
     * This script looks for missing status reports in a project management table.
     *
     * @returns An array of Employee objects (containing their names and emails).
     */
    function main(workbook: ExcelScript.Workbook): Employee[] {
      // Get the first worksheet and the first table on that worksheet.
      let sheet = workbook.getFirstWorksheet()
      let table = sheet.getTables()[0];

      // Give the column indices names matching their expected content.
      const NAME_INDEX = 0;
      const EMAIL_INDEX = 1;
      const STATUS_REPORT_INDEX = 2;

      // Get the data for the whole table.
      let bodyRangeValues = table.getRangeBetweenHeaderAndTotal().getValues();

      // Create the array of Employee objects to return.
      let people: Employee[] = [];

      // Loop through the table and check each row for completion.
      for (let i = 0; i < bodyRangeValues.length; i++) {
        let row = bodyRangeValues[i];
        if (row[STATUS_REPORT_INDEX] === "") {
          // Save the email to return.
          people.push({ name: row[NAME_INDEX], email: row[EMAIL_INDEX] });
        }
      }

      // Log the array to verify we're getting the right rows.
      console.log(people);

      // Return the array of Employees.
      return people;
    }

    /**
     * An interface representing an employee.
     * An array of Employees will be returned from the script
     * for the Power Automate flow.
     */
    interface Employee {
      name: string;
      email: string;
    }
    ```

5. <span data-ttu-id="2e77e-127">Guarde el script con el nombre **Get People**.</span><span class="sxs-lookup"><span data-stu-id="2e77e-127">Save the script with the name **Get People**.</span></span>

6. <span data-ttu-id="2e77e-128">A continuación, necesitamos un segundo script para procesar las tarjetas del informe de estado y colocar la nueva información en la hoja de cálculo.</span><span class="sxs-lookup"><span data-stu-id="2e77e-128">Next, we need a second script to process the status report cards and put the new information in the spreadsheet.</span></span> <span data-ttu-id="2e77e-129">En el panel de tareas **Editor de código** , presione **nueva secuencia** de comandos y pegue el siguiente script en el editor.</span><span class="sxs-lookup"><span data-stu-id="2e77e-129">In the **Code Editor** task pane, press **New Script** and paste the following script into the editor.</span></span>

    ```typescript
    /**
     * This script applies the results of a Teams Adaptive Card about
     * a status update to a project management table.
     *
     * @param senderEmail - The email address of the employee updating their status.
     * @param statusReportResponse - The employee's status report.
     */
    function main(workbook: ExcelScript.Workbook,
      senderEmail: string,
      statusReportResponse: string) {

      // Get the first worksheet and the first table in that worksheet.
      let sheet = workbook.getFirstWorksheet();
      let table = sheet.getTables()[0];

      // Give the column indices names matching their expected content.
      const NAME_INDEX = 0;
      const EMAIL_INDEX = 1;
      const STATUS_REPORT_INDEX = 2;

      // Get the range and data for the whole table.
      let bodyRange = table.getRangeBetweenHeaderAndTotal();
      let tableRowCount = bodyRange.getRowCount();
      let bodyRangeValues = bodyRange.getValues();

      // Create a flag to denote success.
      let statusAdded = false;

      // Loop through the table and check each row for a matching email address.
      for (let i = 0; i < tableRowCount && !statusAdded; i++) {
        let row = bodyRangeValues[i];

        // Check if the row's email address matches.
        if (row[EMAIL_INDEX] === senderEmail) {
          // Add the Teams Adaptive Card response to the table.
          bodyRange.getCell(i, STATUS_REPORT_INDEX).setValues([
            [statusReportResponse]
          ]);
          statusAdded = true;
        }
      }

      // If successful, log the status update.
      if (statusAdded) {
        console.log(
          `Successfully added status report for ${senderEmail} containing: ${statusReportResponse}`
        );
      }
    }
    ```

7. <span data-ttu-id="2e77e-130">Guarde el script con el nombre **Guardar estado**.</span><span class="sxs-lookup"><span data-stu-id="2e77e-130">Save the script with the name **Save Status**.</span></span>

8. <span data-ttu-id="2e77e-131">Ahora, es necesario crear el flujo.</span><span class="sxs-lookup"><span data-stu-id="2e77e-131">Now, we need to create the flow.</span></span> <span data-ttu-id="2e77e-132">Abra [energía automatizada](https://flow.microsoft.com/).</span><span class="sxs-lookup"><span data-stu-id="2e77e-132">Open [Power Automate](https://flow.microsoft.com/).</span></span>

    > [!TIP]
    > <span data-ttu-id="2e77e-133">Si no ha creado un flujo antes, consulte nuestro tutorial comenzar a [usar scripts con Power automatization](../../tutorials/excel-power-automate-manual.md) para conocer los conceptos básicos.</span><span class="sxs-lookup"><span data-stu-id="2e77e-133">If you haven't created a flow before, please check out our tutorial [Start using scripts with Power Automate](../../tutorials/excel-power-automate-manual.md) to learn the basics.</span></span>

9. <span data-ttu-id="2e77e-134">Cree un nuevo **flujo instantáneo**.</span><span class="sxs-lookup"><span data-stu-id="2e77e-134">Create a new **Instant flow**.</span></span>

10. <span data-ttu-id="2e77e-135">Elija **desencadenar manualmente un flujo** de las opciones y pulse **crear**.</span><span class="sxs-lookup"><span data-stu-id="2e77e-135">Choose **Manually trigger a flow** from the options and press **Create**.</span></span>

11. <span data-ttu-id="2e77e-136">El flujo tiene que llamar al script **Get People** para obtener todos los empleados con campos de estado vacíos.</span><span class="sxs-lookup"><span data-stu-id="2e77e-136">The flow needs to call the **Get People** script to get all the employees with empty status fields.</span></span> <span data-ttu-id="2e77e-137">Presione **nuevo paso** y seleccione **Excel online (empresa)**.</span><span class="sxs-lookup"><span data-stu-id="2e77e-137">Press **New step** and select **Excel Online (Business)**.</span></span> <span data-ttu-id="2e77e-138">En **acciones**, seleccione **Ejecutar script (vista previa)**.</span><span class="sxs-lookup"><span data-stu-id="2e77e-138">Under **Actions**, select **Run script (preview)**.</span></span> <span data-ttu-id="2e77e-139">Proporcione las siguientes entradas para el paso flujo:</span><span class="sxs-lookup"><span data-stu-id="2e77e-139">Provide the following entries for the flow step:</span></span>

    - <span data-ttu-id="2e77e-140">**Ubicación**: OneDrive para la empresa</span><span class="sxs-lookup"><span data-stu-id="2e77e-140">**Location**: OneDrive for Business</span></span>
    - <span data-ttu-id="2e77e-141">**Biblioteca de documentos**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="2e77e-141">**Document Library**: OneDrive</span></span>
    - <span data-ttu-id="2e77e-142">**Archivo**: task-reminders.xlsx</span><span class="sxs-lookup"><span data-stu-id="2e77e-142">**File**: task-reminders.xlsx</span></span>
    - <span data-ttu-id="2e77e-143">**Script**: obtener personas</span><span class="sxs-lookup"><span data-stu-id="2e77e-143">**Script**: Get People</span></span>

    ![Paso del flujo de script de primera ejecución.](../../images/scenario-task-reminders-first-flow-step.png)

12. <span data-ttu-id="2e77e-145">A continuación, el flujo debe procesar cada empleado en la matriz que devuelve el script.</span><span class="sxs-lookup"><span data-stu-id="2e77e-145">Next, the flow needs to process each Employee in the array returned by the script.</span></span> <span data-ttu-id="2e77e-146">Presione **nuevo paso** y seleccione **publicar una tarjeta adaptable en un usuario de Teams y espere una respuesta**.</span><span class="sxs-lookup"><span data-stu-id="2e77e-146">Press **New step** and select **Post an Adaptive Card to a Teams user and wait for a response**.</span></span>

13. <span data-ttu-id="2e77e-147">Para el campo de **destinatario** , agregue **correo electrónico** del contenido dinámico (la selección tendrá el logotipo de Excel).</span><span class="sxs-lookup"><span data-stu-id="2e77e-147">For the **Recipient** field, add **email** from the dynamic content (the selection will have the Excel logo by it).</span></span> <span data-ttu-id="2e77e-148">Agregar **correo electrónico** hace que el paso de flujo esté rodeado por una **aplicación a cada** bloque.</span><span class="sxs-lookup"><span data-stu-id="2e77e-148">Adding **email** causes the flow step to be surrounded by an **Apply to each** block.</span></span> <span data-ttu-id="2e77e-149">Esto significa que la matriz se recorrerá en iteración mediante la automatización de energía.</span><span class="sxs-lookup"><span data-stu-id="2e77e-149">That means the array will be iterated over by Power Automate.</span></span>

14. <span data-ttu-id="2e77e-150">El envío de una tarjeta adaptable requiere que se proporcione el JSON de la tarjeta como **mensaje**.</span><span class="sxs-lookup"><span data-stu-id="2e77e-150">Sending an Adaptive Card requires the card's JSON to be provided as the **Message**.</span></span> <span data-ttu-id="2e77e-151">Puede usar el [Diseñador de tarjetas adaptable](https://adaptivecards.io/designer/) para crear tarjetas personalizadas.</span><span class="sxs-lookup"><span data-stu-id="2e77e-151">You can use the [Adaptive Card Designer](https://adaptivecards.io/designer/) to create custom cards.</span></span> <span data-ttu-id="2e77e-152">Para este ejemplo, use el siguiente JSON.</span><span class="sxs-lookup"><span data-stu-id="2e77e-152">For this sample, use the following JSON.</span></span>  

    ```json
    {
      "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
      "type": "AdaptiveCard",
      "version": "1.0",
      "body": [
        {
          "type": "TextBlock",
          "size": "Medium",
          "weight": "Bolder",
          "text": "Update your Status Report"
        },
        {
          "type": "Image",
          "altText": "",
          "url": "https://i.imgur.com/f5RcuF3.png"
        },
        {
          "type": "TextBlock",
          "text": "This is a reminder to update your status report for this month's review. You can do so right here in this card, or by adding it directly to the spreadsheet.",
          "wrap": true
        },
        {
          "type": "Input.Text",
          "placeholder": "My status report for this month is...",
          "id": "response",
          "isMultiline": true
        }
      ],
      "actions": [
        {
          "type": "Action.Submit",
          "title": "Submit",
          "id": "submit"
        }
      ]
    }
    ```

15. <span data-ttu-id="2e77e-153">Rellene los campos restantes de la siguiente manera:</span><span class="sxs-lookup"><span data-stu-id="2e77e-153">Fill out the remaining fields as follows:</span></span>

    - <span data-ttu-id="2e77e-154">**Mensaje de actualización**: Gracias por enviar el informe de estado.</span><span class="sxs-lookup"><span data-stu-id="2e77e-154">**Update message**: Thank you for submitting your status report.</span></span> <span data-ttu-id="2e77e-155">La respuesta se ha agregado correctamente a la hoja de cálculo.</span><span class="sxs-lookup"><span data-stu-id="2e77e-155">Your response has been successfully added to the spreadsheet.</span></span>
    - <span data-ttu-id="2e77e-156">**Debe actualizar la tarjeta**: sí</span><span class="sxs-lookup"><span data-stu-id="2e77e-156">**Should update card**: Yes</span></span>

16. <span data-ttu-id="2e77e-157">En el bloque **aplicar a cada** bloque, después de **publicar una tarjeta adaptable a un usuario de Teams y esperar una respuesta**, presione **Agregar una acción**.</span><span class="sxs-lookup"><span data-stu-id="2e77e-157">In the **Apply to each** block, following the **Post an Adaptive Card to a Teams user and wait for a response**, press **Add an action**.</span></span> <span data-ttu-id="2e77e-158">Seleccione **Excel online (empresa)**.</span><span class="sxs-lookup"><span data-stu-id="2e77e-158">Select **Excel Online (Business)**.</span></span> <span data-ttu-id="2e77e-159">En **acciones**, seleccione **Ejecutar script (vista previa)**.</span><span class="sxs-lookup"><span data-stu-id="2e77e-159">Under **Actions**, select **Run script (preview)**.</span></span> <span data-ttu-id="2e77e-160">Proporcione las siguientes entradas para el paso flujo:</span><span class="sxs-lookup"><span data-stu-id="2e77e-160">Provide the following entries for the flow step:</span></span>

    - <span data-ttu-id="2e77e-161">**Ubicación**: OneDrive para la empresa</span><span class="sxs-lookup"><span data-stu-id="2e77e-161">**Location**: OneDrive for Business</span></span>
    - <span data-ttu-id="2e77e-162">**Biblioteca de documentos**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="2e77e-162">**Document Library**: OneDrive</span></span>
    - <span data-ttu-id="2e77e-163">**Archivo**: task-reminders.xlsx</span><span class="sxs-lookup"><span data-stu-id="2e77e-163">**File**: task-reminders.xlsx</span></span>
    - <span data-ttu-id="2e77e-164">**Script**: estado de guardado</span><span class="sxs-lookup"><span data-stu-id="2e77e-164">**Script**: Save Status</span></span>
    - <span data-ttu-id="2e77e-165">**senderEmail**: correo electrónico *(contenido dinámico de Excel)*</span><span class="sxs-lookup"><span data-stu-id="2e77e-165">**senderEmail**: email *(dynamic content from Excel)*</span></span>
    - <span data-ttu-id="2e77e-166">**statusReportResponse**: respuesta *(contenido dinámico de Teams)*</span><span class="sxs-lookup"><span data-stu-id="2e77e-166">**statusReportResponse**: response *(dynamic content from Teams)*</span></span>

    ![El paso aplicar a cada flujo.](../../images/scenario-task-reminders-last-flow-step.png)

17. <span data-ttu-id="2e77e-168">Guarde el flujo.</span><span class="sxs-lookup"><span data-stu-id="2e77e-168">Save the flow.</span></span>

## <a name="running-the-flow"></a><span data-ttu-id="2e77e-169">Ejecución del flujo</span><span class="sxs-lookup"><span data-stu-id="2e77e-169">Running the flow</span></span>

<span data-ttu-id="2e77e-170">Para probar el flujo, asegúrese de que todas las filas de tabla con el estado en blanco usan una dirección de correo electrónico ligada a una cuenta de Teams (probablemente debería usar su propia dirección de correo electrónico durante las pruebas).</span><span class="sxs-lookup"><span data-stu-id="2e77e-170">To test the flow, make sure any table rows with blank status use an email address tied to a Teams account (you should probably use your own email address while testing).</span></span>

<span data-ttu-id="2e77e-171">Puede seleccionar **Test** desde el diseñador de flujo o ejecutar el flujo desde la página **Mis flujos** .</span><span class="sxs-lookup"><span data-stu-id="2e77e-171">You can either select **Test** from the flow designer, or run the flow from the **My flows** page.</span></span> <span data-ttu-id="2e77e-172">Después de iniciar el flujo y aceptar el uso de las conexiones necesarias, debe recibir una tarjeta adaptable de la alimentación automatizada a través de Teams.</span><span class="sxs-lookup"><span data-stu-id="2e77e-172">After starting the flow and accepting the use of the required connections, you should receive an Adaptive Card from Power Automate through Teams.</span></span> <span data-ttu-id="2e77e-173">Una vez que rellene el campo Estado en la tarjeta, el flujo continuará y actualizará la hoja de cálculo con el estado proporcionado.</span><span class="sxs-lookup"><span data-stu-id="2e77e-173">Once you fill out the status field in the card, the flow will continue and update the spreadsheet with the status you provide.</span></span>

### <a name="before-running-the-flow"></a><span data-ttu-id="2e77e-174">Antes de ejecutar el flujo</span><span class="sxs-lookup"><span data-stu-id="2e77e-174">Before running the flow</span></span>

![Una hoja de cálculo con un informe de estado que contiene una entrada de estado que falta.](../../images/scenario-task-reminders-spreadsheet-before.png)

### <a name="receiving-the-adaptive-card"></a><span data-ttu-id="2e77e-176">Recibir la tarjeta adaptable</span><span class="sxs-lookup"><span data-stu-id="2e77e-176">Receiving the Adaptive Card</span></span>

![Una tarjeta adaptable de teams que solicita al empleado una actualización de estado.](../../images/scenario-task-reminders-adaptive-card.png)

### <a name="after-running-the-flow"></a><span data-ttu-id="2e77e-178">Después de ejecutar el flujo</span><span class="sxs-lookup"><span data-stu-id="2e77e-178">After running the flow</span></span>

![Una hoja de cálculo con un informe de estado con una entrada de estado ahora rellenada.](../../images/scenario-task-reminders-spreadsheet-after.png)
