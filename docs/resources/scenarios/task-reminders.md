---
title: 'Office ejemplo scripts: avisos de tareas automatizadas'
description: Un ejemplo que usa Power Automate y tarjetas adaptables automatizan los avisos de tareas en una hoja de cálculo de administración de proyectos.
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: f157601d067bda0d5501ae865d7f63f99926d347
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585950"
---
# <a name="office-scripts-sample-scenario-automated-task-reminders"></a>Office ejemplo scripts: avisos de tareas automatizadas

En este escenario, está administrando un proyecto. Use una hoja de Excel para realizar un seguimiento del estado de sus empleados cada mes. A menudo debes recordar a los usuarios que rellenen su estado, por lo que has decidido automatizar ese proceso de aviso.

Crearás un flujo de Power Automate para enviar mensajes a personas con campos de estado ausentes y aplicar sus respuestas a la hoja de cálculo. Para ello, desarrollará un par de scripts para controlar el trabajo con el libro. El primer script obtiene una lista de personas con estados en blanco y el segundo script agrega una cadena de estado a la fila derecha. También usarás las tarjetas adaptables [Teams para](/microsoftteams/platform/task-modules-and-cards/what-are-cards) que los empleados escriban su estado directamente desde la notificación.

## <a name="scripting-skills-covered"></a>Habilidades de scripting cubiertas

- Crear flujos en Power Automate
- Pasar datos a scripts
- Devolver datos de scripts
- Teams adaptables
- Tablas

## <a name="prerequisites"></a>Requisitos previos

Este escenario usa [Power Automate](https://flow.microsoft.com) y [Microsoft Teams](https://www.microsoft.com/microsoft-365/microsoft-teams/group-chat-software). Necesitarás ambos asociados con la cuenta que usas para desarrollar scripts Office usuario. Para obtener acceso gratuito a una suscripción de Microsoft Developer para obtener información sobre estas aplicaciones y trabajar con ellas, considere la posibilidad de unirse al [Microsoft 365 Developer Program](https://developer.microsoft.com/microsoft-365/dev-program).

## <a name="setup-instructions"></a>Instrucciones de configuración

1. Descargue <a href="task-reminders.xlsx">task-reminders.xlsx</a> a su OneDrive.

1. Abra el libro en Excel en la Web.

1. En primer lugar, necesitamos un script para obtener todos los empleados con informes de estado que faltan en la hoja de cálculo. En la **pestaña Automatizar** , seleccione **Nuevo script** y pegue el siguiente script en el editor.

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

1. Guarde el script con el nombre **Get People**.

1. A continuación, necesitamos un segundo script para procesar las tarjetas de informe de estado y colocar la nueva información en la hoja de cálculo. En el panel de tareas Editor de código, seleccione **Nuevo script** y pegue el siguiente script en el editor.

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

1. Guarde el script con el nombre **Guardar estado**.

1. Ahora, debemos crear el flujo. Abra [Power Automate](https://flow.microsoft.com/).

    > [!TIP]
    > Si no has creado un flujo antes, consulta nuestro tutorial Empezar a usar [scripts con](../../tutorials/excel-power-automate-manual.md) Power Automate para aprender los conceptos básicos.

1. Cree un nuevo **flujo instantáneo**.

1. Elija **Desencadenar manualmente un flujo de** las opciones y seleccione **Crear**.

1. El flujo debe llamar al script **Obtener personas** para obtener todos los empleados con campos de estado vacíos. Seleccione **Nuevo paso** y, **a continuación, Excel Online (Empresa).**. En **Acciones**, seleccione **Ejecutar script**. Proporcione las siguientes entradas para el paso de flujo:

    - **Ubicación**: OneDrive para la Empresa
    - **Biblioteca de documentos**: OneDrive
    - **Archivo**: task-reminders.xlsx *(elegido a través del explorador de archivos)*
    - **Script**: Obtener personas

    :::image type="content" source="../../images/scenario-task-reminders-first-flow-step.png" alt-text="Flujo Power Automate muestra el primer paso de flujo ejecutar script.":::

1. A continuación, el flujo debe procesar cada empleado de la matriz devuelta por el script. Seleccione **Nuevo paso** y, a continuación, elija Publicar una tarjeta adaptable **en un Teams usuario y esperar una respuesta**.

1. Para el **campo Destinatario**, agregue **correo** electrónico desde el contenido dinámico (la selección tendrá Excel logotipo). Agregar **correo** electrónico hace que el paso de flujo esté rodeado por un **aplicar a cada** bloque. Esto significa que la matriz se iterará por Power Automate.

1. El envío de una tarjeta adaptable requiere que el JSON de la tarjeta se proporciona como **mensaje**. Puede usar el Diseñador de [tarjetas adaptables](https://adaptivecards.io/designer/) para crear tarjetas personalizadas. Para este ejemplo, use el siguiente JSON.  

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

1. Rellene los campos restantes de la siguiente manera:

    - **Mensaje de actualización**: gracias por enviar el informe de estado. La respuesta se ha agregado correctamente a la hoja de cálculo.
    - **Debe actualizar la tarjeta**: Sí

1. En el **bloque Aplicar a cada** bloque, después de publicar una tarjeta adaptable a un usuario de Teams y esperar **una** respuesta, seleccione **Agregar una acción**. Seleccione **Excel online (empresa).**. En **Acciones**, seleccione **Ejecutar script**. Proporcione las siguientes entradas para el paso de flujo:

    - **Ubicación**: OneDrive para la Empresa
    - **Biblioteca de documentos**: OneDrive
    - **Archivo**: task-reminders.xlsx *(elegido a través del explorador de archivos)*
    - **Script**: Guardar estado
    - **senderEmail**: correo *electrónico (contenido dinámico de Excel)*
    - **statusReportResponse**: respuesta *(contenido dinámico de Teams)*

    :::image type="content" source="../../images/scenario-task-reminders-last-flow-step.png" alt-text="El Power Automate que muestra el paso aplicar a cada paso.":::

1. Guarde el flujo.

## <a name="running-the-flow"></a>Ejecución del flujo

Para probar el flujo, asegúrese de que cualquier fila de tabla con estado en blanco use una dirección de correo electrónico vinculada a una cuenta de Teams (probablemente debería usar su propia dirección de correo electrónico durante las pruebas). Use el **botón Probar** en la página del editor de flujo o ejecute el flujo a través de la **pestaña Mis flujos** . Asegúrese de permitir el acceso cuando se le pida.

Debe recibir una tarjeta adaptable de Power Automate a Teams. Una vez rellenado el campo de estado en la tarjeta, el flujo continuará y actualizará la hoja de cálculo con el estado que proporcione.

### <a name="before-running-the-flow"></a>Antes de ejecutar el flujo

:::image type="content" source="../../images/scenario-task-reminders-spreadsheet-before.png" alt-text="Hoja de cálculo con un informe de estado que contiene una entrada de estado que falta.":::

### <a name="receiving-the-adaptive-card"></a>Recepción de la tarjeta adaptable

:::image type="content" source="../../images/scenario-task-reminders-adaptive-card.png" alt-text="Una tarjeta adaptable en Teams solicita al empleado una actualización de estado.":::

### <a name="after-running-the-flow"></a>Después de ejecutar el flujo

:::image type="content" source="../../images/scenario-task-reminders-spreadsheet-after.png" alt-text="Hoja de cálculo con un informe de estado con una entrada de estado ahora rellenada.":::
