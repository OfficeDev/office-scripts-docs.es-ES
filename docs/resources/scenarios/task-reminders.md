---
title: 'Office Escenario de ejemplo scripts: avisos de tareas automatizados'
description: Un ejemplo que usa Power Automate tarjetas adaptables automatizan los avisos de tareas en una hoja de cálculo de administración de proyectos.
ms.date: 11/30/2020
localization_priority: Normal
ms.openlocfilehash: c254a627da8442c0974263908a41275182740b6e
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545609"
---
# <a name="office-scripts-sample-scenario-automated-task-reminders"></a><span data-ttu-id="bce6a-103">Office Escenario de ejemplo scripts: avisos de tareas automatizados</span><span class="sxs-lookup"><span data-stu-id="bce6a-103">Office Scripts sample scenario: Automated task reminders</span></span>

<span data-ttu-id="bce6a-104">En este escenario, está administrando un proyecto.</span><span class="sxs-lookup"><span data-stu-id="bce6a-104">In this scenario you're managing a project.</span></span> <span data-ttu-id="bce6a-105">Use una hoja de Excel para realizar un seguimiento del estado de sus empleados cada mes.</span><span class="sxs-lookup"><span data-stu-id="bce6a-105">You use an Excel worksheet to track your employees' status every month.</span></span> <span data-ttu-id="bce6a-106">A menudo debes recordar a los usuarios que rellenen su estado, por lo que has decidido automatizar ese proceso de aviso.</span><span class="sxs-lookup"><span data-stu-id="bce6a-106">You often need to remind people to fill out their status, so you've decided to automate that reminder process.</span></span>

<span data-ttu-id="bce6a-107">Crearás un flujo de Power Automate para enviar mensajes a personas con campos de estado que faltan y aplicar sus respuestas a la hoja de cálculo.</span><span class="sxs-lookup"><span data-stu-id="bce6a-107">You'll create a Power Automate flow to message people with missing status fields and apply their responses to the spreadsheet.</span></span> <span data-ttu-id="bce6a-108">Para ello, desarrollará un par de scripts para controlar el trabajo con el libro.</span><span class="sxs-lookup"><span data-stu-id="bce6a-108">To do this, you'll develop a pair of scripts to handle the working with the workbook.</span></span> <span data-ttu-id="bce6a-109">El primer script obtiene una lista de personas con estados en blanco y el segundo script agrega una cadena de estado a la fila derecha.</span><span class="sxs-lookup"><span data-stu-id="bce6a-109">The first script gets a list of people with blank statuses and the second script adds a status string to the right row.</span></span> <span data-ttu-id="bce6a-110">También usarás las tarjetas adaptables Teams [para](/microsoftteams/platform/task-modules-and-cards/what-are-cards) que los empleados escriban su estado directamente desde la notificación.</span><span class="sxs-lookup"><span data-stu-id="bce6a-110">You'll also make use of [Teams Adaptive Cards](/microsoftteams/platform/task-modules-and-cards/what-are-cards) to have employees enter their status directly from the notification.</span></span>

## <a name="scripting-skills-covered"></a><span data-ttu-id="bce6a-111">Habilidades de scripting cubiertas</span><span class="sxs-lookup"><span data-stu-id="bce6a-111">Scripting skills covered</span></span>

- <span data-ttu-id="bce6a-112">Crear flujos en Power Automate</span><span class="sxs-lookup"><span data-stu-id="bce6a-112">Create flows in Power Automate</span></span>
- <span data-ttu-id="bce6a-113">Pasar datos a scripts</span><span class="sxs-lookup"><span data-stu-id="bce6a-113">Pass data to scripts</span></span>
- <span data-ttu-id="bce6a-114">Devolver datos de scripts</span><span class="sxs-lookup"><span data-stu-id="bce6a-114">Return data from scripts</span></span>
- <span data-ttu-id="bce6a-115">Teams Tarjetas adaptables</span><span class="sxs-lookup"><span data-stu-id="bce6a-115">Teams Adaptive Cards</span></span>
- <span data-ttu-id="bce6a-116">Tablas</span><span class="sxs-lookup"><span data-stu-id="bce6a-116">Tables</span></span>

## <a name="prerequisites"></a><span data-ttu-id="bce6a-117">Requisitos previos</span><span class="sxs-lookup"><span data-stu-id="bce6a-117">Prerequisites</span></span>

<span data-ttu-id="bce6a-118">Este escenario usa [Power Automate](https://flow.microsoft.com) y [Microsoft Teams](https://www.microsoft.com/microsoft-365/microsoft-teams/group-chat-software).</span><span class="sxs-lookup"><span data-stu-id="bce6a-118">This scenario uses [Power Automate](https://flow.microsoft.com) and [Microsoft Teams](https://www.microsoft.com/microsoft-365/microsoft-teams/group-chat-software).</span></span> <span data-ttu-id="bce6a-119">Necesitarás ambos asociados con la cuenta que usas para desarrollar Office scripts.</span><span class="sxs-lookup"><span data-stu-id="bce6a-119">You will need both associated with the account that you use for developing Office Scripts.</span></span> <span data-ttu-id="bce6a-120">Para obtener acceso gratuito a una suscripción de Microsoft Developer para obtener información sobre estas aplicaciones y trabajar con ellas, considere la posibilidad de unirse al [Microsoft 365 Developer Program](https://developer.microsoft.com/microsoft-365/dev-program).</span><span class="sxs-lookup"><span data-stu-id="bce6a-120">For free access to a Microsoft Developer subscription to learn about and work with these applications, consider joining the [Microsoft 365 Developer Program](https://developer.microsoft.com/microsoft-365/dev-program).</span></span>

## <a name="setup-instructions"></a><span data-ttu-id="bce6a-121">Instrucciones de configuración</span><span class="sxs-lookup"><span data-stu-id="bce6a-121">Setup instructions</span></span>

1. <span data-ttu-id="bce6a-122">Descargue <a href="task-reminders.xlsx">task-reminders.xlsx</a> a su OneDrive.</span><span class="sxs-lookup"><span data-stu-id="bce6a-122">Download <a href="task-reminders.xlsx">task-reminders.xlsx</a> to your OneDrive.</span></span>

2. <span data-ttu-id="bce6a-123">Abra el libro en Excel en la Web.</span><span class="sxs-lookup"><span data-stu-id="bce6a-123">Open the workbook in Excel on the web.</span></span>

3. <span data-ttu-id="bce6a-124">En la **pestaña Automatizar,** abra **Todos los scripts**.</span><span class="sxs-lookup"><span data-stu-id="bce6a-124">Under the **Automate** tab, open **All Scripts**.</span></span>

4. <span data-ttu-id="bce6a-125">En primer lugar, necesitamos un script para obtener todos los empleados con informes de estado que faltan en la hoja de cálculo.</span><span class="sxs-lookup"><span data-stu-id="bce6a-125">First, we need a script to get all the employees with status reports that are missing from the spreadsheet.</span></span> <span data-ttu-id="bce6a-126">En el **panel de tareas Editor** de código, presione Nuevo **script** y pegue el siguiente script en el editor.</span><span class="sxs-lookup"><span data-stu-id="bce6a-126">In the **Code Editor** task pane, press **New Script** and paste the following script into the editor.</span></span>

    ```TypeScript
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
          people.push({ name: row[NAME_INDEX].toString(), email: row[EMAIL_INDEX].toString() });
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

5. <span data-ttu-id="bce6a-127">Guarde el script con el nombre **Get People**.</span><span class="sxs-lookup"><span data-stu-id="bce6a-127">Save the script with the name **Get People**.</span></span>

6. <span data-ttu-id="bce6a-128">A continuación, necesitamos un segundo script para procesar las tarjetas de informe de estado y colocar la nueva información en la hoja de cálculo.</span><span class="sxs-lookup"><span data-stu-id="bce6a-128">Next, we need a second script to process the status report cards and put the new information in the spreadsheet.</span></span> <span data-ttu-id="bce6a-129">En el **panel de tareas Editor** de código, presione Nuevo **script** y pegue el siguiente script en el editor.</span><span class="sxs-lookup"><span data-stu-id="bce6a-129">In the **Code Editor** task pane, press **New Script** and paste the following script into the editor.</span></span>

    ```TypeScript
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

7. <span data-ttu-id="bce6a-130">Guarde el script con el nombre **Guardar estado**.</span><span class="sxs-lookup"><span data-stu-id="bce6a-130">Save the script with the name **Save Status**.</span></span>

8. <span data-ttu-id="bce6a-131">Ahora, debemos crear el flujo.</span><span class="sxs-lookup"><span data-stu-id="bce6a-131">Now, we need to create the flow.</span></span> <span data-ttu-id="bce6a-132">Abra [Power Automate](https://flow.microsoft.com/).</span><span class="sxs-lookup"><span data-stu-id="bce6a-132">Open [Power Automate](https://flow.microsoft.com/).</span></span>

    > [!TIP]
    > <span data-ttu-id="bce6a-133">Si no has creado un flujo antes, consulta nuestro tutorial Empezar a usar [scripts con](../../tutorials/excel-power-automate-manual.md) Power Automate para aprender los conceptos básicos.</span><span class="sxs-lookup"><span data-stu-id="bce6a-133">If you haven't created a flow before, please check out our tutorial [Start using scripts with Power Automate](../../tutorials/excel-power-automate-manual.md) to learn the basics.</span></span>

9. <span data-ttu-id="bce6a-134">Crear un nuevo **flujo instantáneo**.</span><span class="sxs-lookup"><span data-stu-id="bce6a-134">Create a new **Instant flow**.</span></span>

10. <span data-ttu-id="bce6a-135">Elija **Desencadenar manualmente un flujo de** las opciones y presione **Crear**.</span><span class="sxs-lookup"><span data-stu-id="bce6a-135">Choose **Manually trigger a flow** from the options and press **Create**.</span></span>

11. <span data-ttu-id="bce6a-136">El flujo debe llamar al script **Obtener personas** para obtener todos los empleados con campos de estado vacíos.</span><span class="sxs-lookup"><span data-stu-id="bce6a-136">The flow needs to call the **Get People** script to get all the employees with empty status fields.</span></span> <span data-ttu-id="bce6a-137">Presione **Nuevo paso** y seleccione Excel Online **(Empresa).**</span><span class="sxs-lookup"><span data-stu-id="bce6a-137">Press **New step** and select **Excel Online (Business)**.</span></span> <span data-ttu-id="bce6a-138">En **Acciones**, seleccione **Ejecutar script**.</span><span class="sxs-lookup"><span data-stu-id="bce6a-138">Under **Actions**, select **Run script**.</span></span> <span data-ttu-id="bce6a-139">Proporcione las siguientes entradas para el paso de flujo:</span><span class="sxs-lookup"><span data-stu-id="bce6a-139">Provide the following entries for the flow step:</span></span>

    - <span data-ttu-id="bce6a-140">**Ubicación**: OneDrive para la Empresa</span><span class="sxs-lookup"><span data-stu-id="bce6a-140">**Location**: OneDrive for Business</span></span>
    - <span data-ttu-id="bce6a-141">**Biblioteca de documentos**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="bce6a-141">**Document Library**: OneDrive</span></span>
    - <span data-ttu-id="bce6a-142">**Archivo**: task-reminders.xlsx *(elegido a través del explorador de archivos)*</span><span class="sxs-lookup"><span data-stu-id="bce6a-142">**File**: task-reminders.xlsx *(Chosen through the file browser)*</span></span>
    - <span data-ttu-id="bce6a-143">**Script**: Obtener personas</span><span class="sxs-lookup"><span data-stu-id="bce6a-143">**Script**: Get People</span></span>

    :::image type="content" source="../../images/scenario-task-reminders-first-flow-step.png" alt-text="El Power Automate que muestra el primer paso ejecutar flujo de script":::

12. <span data-ttu-id="bce6a-145">A continuación, el flujo debe procesar cada empleado de la matriz devuelta por el script.</span><span class="sxs-lookup"><span data-stu-id="bce6a-145">Next, the flow needs to process each Employee in the array returned by the script.</span></span> <span data-ttu-id="bce6a-146">Presione **Nuevo paso** y seleccione Publicar una tarjeta adaptable en un Teams usuario y espere una **respuesta**.</span><span class="sxs-lookup"><span data-stu-id="bce6a-146">Press **New step** and select **Post an Adaptive Card to a Teams user and wait for a response**.</span></span>

13. <span data-ttu-id="bce6a-147">Para el **campo Destinatario,** agregue **correo** electrónico desde el contenido dinámico (la selección tendrá Excel logotipo).</span><span class="sxs-lookup"><span data-stu-id="bce6a-147">For the **Recipient** field, add **email** from the dynamic content (the selection will have the Excel logo by it).</span></span> <span data-ttu-id="bce6a-148">Agregar **correo** electrónico hace que el paso de flujo esté rodeado por un **aplicar a cada** bloque.</span><span class="sxs-lookup"><span data-stu-id="bce6a-148">Adding **email** causes the flow step to be surrounded by an **Apply to each** block.</span></span> <span data-ttu-id="bce6a-149">Esto significa que la matriz se iterará por Power Automate.</span><span class="sxs-lookup"><span data-stu-id="bce6a-149">That means the array will be iterated over by Power Automate.</span></span>

14. <span data-ttu-id="bce6a-150">El envío de una tarjeta adaptable requiere que el JSON de la tarjeta se proporciona como **message**.</span><span class="sxs-lookup"><span data-stu-id="bce6a-150">Sending an Adaptive Card requires the card's JSON to be provided as the **Message**.</span></span> <span data-ttu-id="bce6a-151">Puede usar el Diseñador de [tarjetas adaptables](https://adaptivecards.io/designer/) para crear tarjetas personalizadas.</span><span class="sxs-lookup"><span data-stu-id="bce6a-151">You can use the [Adaptive Card Designer](https://adaptivecards.io/designer/) to create custom cards.</span></span> <span data-ttu-id="bce6a-152">Para este ejemplo, use el siguiente JSON.</span><span class="sxs-lookup"><span data-stu-id="bce6a-152">For this sample, use the following JSON.</span></span>  

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

15. <span data-ttu-id="bce6a-153">Rellene los campos restantes de la siguiente manera:</span><span class="sxs-lookup"><span data-stu-id="bce6a-153">Fill out the remaining fields as follows:</span></span>

    - <span data-ttu-id="bce6a-154">**Mensaje de actualización:** gracias por enviar el informe de estado.</span><span class="sxs-lookup"><span data-stu-id="bce6a-154">**Update message**: Thank you for submitting your status report.</span></span> <span data-ttu-id="bce6a-155">La respuesta se ha agregado correctamente a la hoja de cálculo.</span><span class="sxs-lookup"><span data-stu-id="bce6a-155">Your response has been successfully added to the spreadsheet.</span></span>
    - <span data-ttu-id="bce6a-156">**Debe actualizar la tarjeta**: Sí</span><span class="sxs-lookup"><span data-stu-id="bce6a-156">**Should update card**: Yes</span></span>

16. <span data-ttu-id="bce6a-157">En el **bloque Aplicar** a cada bloque, después de publicar una tarjeta adaptable a un usuario de Teams y esperar **una** respuesta, presione Agregar **una acción**.</span><span class="sxs-lookup"><span data-stu-id="bce6a-157">In the **Apply to each** block, following the **Post an Adaptive Card to a Teams user and wait for a response**, press **Add an action**.</span></span> <span data-ttu-id="bce6a-158">Seleccione **Excel online (empresa).**</span><span class="sxs-lookup"><span data-stu-id="bce6a-158">Select **Excel Online (Business)**.</span></span> <span data-ttu-id="bce6a-159">En **Acciones**, seleccione **Ejecutar script**.</span><span class="sxs-lookup"><span data-stu-id="bce6a-159">Under **Actions**, select **Run script**.</span></span> <span data-ttu-id="bce6a-160">Proporcione las siguientes entradas para el paso de flujo:</span><span class="sxs-lookup"><span data-stu-id="bce6a-160">Provide the following entries for the flow step:</span></span>

    - <span data-ttu-id="bce6a-161">**Ubicación**: OneDrive para la Empresa</span><span class="sxs-lookup"><span data-stu-id="bce6a-161">**Location**: OneDrive for Business</span></span>
    - <span data-ttu-id="bce6a-162">**Biblioteca de documentos**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="bce6a-162">**Document Library**: OneDrive</span></span>
    - <span data-ttu-id="bce6a-163">**Archivo**: task-reminders.xlsx *(elegido a través del explorador de archivos)*</span><span class="sxs-lookup"><span data-stu-id="bce6a-163">**File**: task-reminders.xlsx *(Chosen through the file browser)*</span></span>
    - <span data-ttu-id="bce6a-164">**Script**: Guardar estado</span><span class="sxs-lookup"><span data-stu-id="bce6a-164">**Script**: Save Status</span></span>
    - <span data-ttu-id="bce6a-165">**senderEmail:** correo *electrónico (contenido dinámico de Excel)*</span><span class="sxs-lookup"><span data-stu-id="bce6a-165">**senderEmail**: email *(dynamic content from Excel)*</span></span>
    - <span data-ttu-id="bce6a-166">**statusReportResponse**: response *(contenido dinámico de Teams)*</span><span class="sxs-lookup"><span data-stu-id="bce6a-166">**statusReportResponse**: response *(dynamic content from Teams)*</span></span>

    :::image type="content" source="../../images/scenario-task-reminders-last-flow-step.png" alt-text="El Power Automate que muestra el paso aplicar a cada paso":::

17. <span data-ttu-id="bce6a-168">Guarde el flujo.</span><span class="sxs-lookup"><span data-stu-id="bce6a-168">Save the flow.</span></span>

## <a name="running-the-flow"></a><span data-ttu-id="bce6a-169">Ejecución del flujo</span><span class="sxs-lookup"><span data-stu-id="bce6a-169">Running the flow</span></span>

<span data-ttu-id="bce6a-170">Para probar el flujo, asegúrese de que cualquier fila de tabla con estado en blanco use una dirección de correo electrónico vinculada a una cuenta de Teams (probablemente debería usar su propia dirección de correo electrónico durante las pruebas).</span><span class="sxs-lookup"><span data-stu-id="bce6a-170">To test the flow, make sure any table rows with blank status use an email address tied to a Teams account (you should probably use your own email address while testing).</span></span>

<span data-ttu-id="bce6a-171">Puede seleccionar Probar **en** el diseñador de flujos o ejecutar el flujo desde la **página Mis flujos.**</span><span class="sxs-lookup"><span data-stu-id="bce6a-171">You can either select **Test** from the flow designer, or run the flow from the **My flows** page.</span></span> <span data-ttu-id="bce6a-172">Después de iniciar el flujo y aceptar el uso de las conexiones necesarias, debe recibir una tarjeta adaptable de Power Automate a Teams.</span><span class="sxs-lookup"><span data-stu-id="bce6a-172">After starting the flow and accepting the use of the required connections, you should receive an Adaptive Card from Power Automate through Teams.</span></span> <span data-ttu-id="bce6a-173">Una vez rellenado el campo de estado en la tarjeta, el flujo continuará y actualizará la hoja de cálculo con el estado que proporcione.</span><span class="sxs-lookup"><span data-stu-id="bce6a-173">Once you fill out the status field in the card, the flow will continue and update the spreadsheet with the status you provide.</span></span>

### <a name="before-running-the-flow"></a><span data-ttu-id="bce6a-174">Antes de ejecutar el flujo</span><span class="sxs-lookup"><span data-stu-id="bce6a-174">Before running the flow</span></span>

:::image type="content" source="../../images/scenario-task-reminders-spreadsheet-before.png" alt-text="Hoja de cálculo con un informe de estado que contiene una entrada de estado que falta":::

### <a name="receiving-the-adaptive-card"></a><span data-ttu-id="bce6a-176">Recepción de la tarjeta adaptable</span><span class="sxs-lookup"><span data-stu-id="bce6a-176">Receiving the Adaptive Card</span></span>

:::image type="content" source="../../images/scenario-task-reminders-adaptive-card.png" alt-text="Una tarjeta adaptable en Teams solicitar al empleado una actualización de estado":::

### <a name="after-running-the-flow"></a><span data-ttu-id="bce6a-178">Después de ejecutar el flujo</span><span class="sxs-lookup"><span data-stu-id="bce6a-178">After running the flow</span></span>

:::image type="content" source="../../images/scenario-task-reminders-spreadsheet-after.png" alt-text="Hoja de cálculo con un informe de estado con una entrada de estado ahora rellenada":::
