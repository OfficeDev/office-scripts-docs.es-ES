---
title: 'escenario de ejemplo de scripts de Office: recordatorios de tareas automatizados'
description: Un ejemplo que usa Power Automate y tarjetas adaptables automatiza los recordatorios de tareas en una hoja de cálculo de administración de proyectos.
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: 08f3713210e83162f86d38bc8eb33d76bf8a7288
ms.sourcegitcommit: dd01979d34b3499360d2f79a56f8a8f24f480eed
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 06/15/2022
ms.locfileid: "66088116"
---
# <a name="office-scripts-sample-scenario-automated-task-reminders"></a>escenario de ejemplo de scripts de Office: recordatorios de tareas automatizados

En este escenario, va a administrar un proyecto. Use una hoja de cálculo de Excel para realizar un seguimiento del estado de los empleados cada mes. A menudo es necesario recordar a los usuarios que rellenen su estado, por lo que ha decidido automatizar ese proceso de recordatorio.

Creará un flujo de Power Automate a los usuarios de mensajes con campos de estado que faltan y aplicará sus respuestas a la hoja de cálculo. Para ello, desarrollará un par de scripts para controlar el trabajo con el libro. El primer script obtiene una lista de personas con estados en blanco y el segundo script agrega una cadena de estado a la fila derecha. También usará [Teams tarjetas adaptables](/microsoftteams/platform/task-modules-and-cards/what-are-cards) para que los empleados escriban su estado directamente desde la notificación.

## <a name="scripting-skills-covered"></a>Aptitudes de scripting cubiertas

- Creación de flujos en Power Automate
- Pasar datos a scripts
- Devolver datos de scripts
- tarjetas adaptables Teams
- Tablas

## <a name="prerequisites"></a>Requisitos previos

En este escenario se usan [Power Automate](https://flow.microsoft.com) y [Microsoft Teams](https://www.microsoft.com/microsoft-365/microsoft-teams/group-chat-software). Necesitará ambos asociados a la cuenta que use para desarrollar scripts de Office. Para obtener acceso gratuito a una suscripción de Microsoft Developer para obtener información sobre estas aplicaciones y trabajar con ellas, considere la posibilidad de unirse al [Programa para desarrolladores de Microsoft 365](https://developer.microsoft.com/microsoft-365/dev-program).

## <a name="setup-instructions"></a>Instrucciones de instalación

1. Descargue <a href="task-reminders.xlsx">task-reminders.xlsx</a> en su OneDrive.

1. Abra el libro en Excel en la Web.

1. En primer lugar, necesitamos un script para obtener todos los empleados con informes de estado que faltan en la hoja de cálculo. En la pestaña **Automatizar** , seleccione **Nuevo script** y pegue el siguiente script en el editor.

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

1. Guarde el script con el nombre **Obtener personas**.

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

1. Ahora, tenemos que crear el flujo. Abra [Power Automate](https://flow.microsoft.com/).

    > [!TIP]
    > Si no ha creado un flujo antes, consulte nuestro tutorial [Empezar a usar scripts con Power Automate](../../tutorials/excel-power-automate-manual.md) para aprender los conceptos básicos.

1. Cree un nuevo **flujo instantáneo**.

1. Elija **Desencadenar manualmente un flujo** en las opciones y seleccione **Crear**.

1. El flujo debe llamar al script **Obtener personas** para obtener todos los empleados con campos de estado vacíos. Seleccione **Nuevo paso** y, a continuación, **Excel Online (Empresa).** En **Acciones**, seleccione **Ejecutar script**. Proporcione las siguientes entradas para el paso de flujo:

    - **Ubicación**: OneDrive para la Empresa
    - **Biblioteca de documentos**: OneDrive
    - **Archivo**: task-reminders.xlsx *(elegido mediante el explorador de archivos)*
    - **Script**: Obtener personas

    :::image type="content" source="../../images/scenario-task-reminders-first-flow-step.png" alt-text="Flujo de Power Automate que muestra el primer paso Ejecutar flujo de script.":::

1. A continuación, el flujo debe procesar cada empleado de la matriz devuelta por el script. Seleccione **Nuevo paso** y, **después, Post an Adaptive Card to a Teams user (Publicar una tarjeta adaptable a un usuario Teams) y espere una respuesta**.

1. En el campo **Destinatario**, agregue **correo electrónico** desde el contenido dinámico (la selección tendrá el logotipo de Excel). La adición de **correo electrónico** hace que el paso de flujo esté rodeado por un **apply a cada** bloque. Esto significa que la matriz se iterará por Power Automate.

1. El envío de una tarjeta adaptable requiere que el [JSON](https://www.w3schools.com/whatis/whatis_json.asp) de la tarjeta se proporcione como **mensaje**. Puede usar el [Diseñador de tarjetas adaptables](https://adaptivecards.io/designer/) para crear tarjetas personalizadas. Para este ejemplo, use el siguiente CÓDIGO JSON.  

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

1. En el bloque **Aplicar a cada** bloque, después de **publicar una tarjeta adaptable en un usuario Teams y esperar una respuesta**, seleccione **Agregar una acción**. Seleccione **Excel Online (Empresa).** En **Acciones**, seleccione **Ejecutar script**. Proporcione las siguientes entradas para el paso de flujo:

    - **Ubicación**: OneDrive para la Empresa
    - **Biblioteca de documentos**: OneDrive
    - **Archivo**: task-reminders.xlsx *(elegido mediante el explorador de archivos)*
    - **Script**: Guardar estado
    - **senderEmail**: correo electrónico *(contenido dinámico de Excel)*
    - **statusReportResponse**: respuesta *(contenido dinámico de Teams)*

    :::image type="content" source="../../images/scenario-task-reminders-last-flow-step.png" alt-text="Flujo de Power Automate que muestra el paso aplicar a cada paso.":::

1. Guarde el flujo.

## <a name="running-the-flow"></a>Ejecución del flujo

Para probar el flujo, asegúrese de que las filas de tabla con estado en blanco usen una dirección de correo electrónico asociada a una cuenta de Teams (probablemente debería usar su propia dirección de correo electrónico durante las pruebas). Use el botón **Probar** de la página del editor de flujo o ejecute el flujo a través de la pestaña **Mis flujos** . Asegúrese de permitir el acceso cuando se le solicite.

Debería recibir una tarjeta adaptable de Power Automate a través de Teams. Una vez rellenado el campo de estado en la tarjeta, el flujo continuará y actualizará la hoja de cálculo con el estado que proporcione.

### <a name="before-running-the-flow"></a>Antes de ejecutar el flujo

:::image type="content" source="../../images/scenario-task-reminders-spreadsheet-before.png" alt-text="Hoja de cálculo con un informe de estado que contiene una entrada de estado que falta.":::

### <a name="receiving-the-adaptive-card"></a>Recepción de la tarjeta adaptable

:::image type="content" source="../../images/scenario-task-reminders-adaptive-card.png" alt-text="Una tarjeta adaptable en Teams pedir al empleado una actualización de estado.":::

### <a name="after-running-the-flow"></a>Después de ejecutar el flujo

:::image type="content" source="../../images/scenario-task-reminders-spreadsheet-after.png" alt-text="Hoja de cálculo con un informe de estado con una entrada de estado ahora rellenada.":::
