---
title: 'Office Escenario de ejemplo de scripts: recordatorios de tareas automatizados'
description: Ejemplo que usa tarjetas Power Automate y adaptables automatiza los recordatorios de tareas en una hoja de cálculo de administración de proyectos.
ms.date: 11/30/2020
localization_priority: Normal
ms.openlocfilehash: c254a627da8442c0974263908a41275182740b6e
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545609"
---
# <a name="office-scripts-sample-scenario-automated-task-reminders"></a>Office Escenario de ejemplo de scripts: recordatorios de tareas automatizados

En este escenario está administrando un proyecto. Utilice una hoja de trabajo Excel para realizar un seguimiento del estado de sus empleados cada mes. A menudo necesitas recordar a la gente que llene su estado, así que has decidido automatizar ese proceso de recordatorio.

Creará un flujo de Power Automate para enviar mensajes a personas con campos de estado que faltan y aplicará sus respuestas a la hoja de cálculo. Para ello, desarrollará un par de scripts para controlar el trabajo con el libro. El primer script obtiene una lista de personas con estados en blanco y el segundo script agrega una cadena de estado a la fila derecha. También hará uso de [Teams tarjetas adaptables](/microsoftteams/platform/task-modules-and-cards/what-are-cards) para que los empleados introduzcan su estado directamente desde la notificación.

## <a name="scripting-skills-covered"></a>Habilidades de scripting cubiertas

- Crear flujos en Power Automate
- Pasar datos a scripts
- Devolver datos de scripts
- Teams Tarjetas adaptables
- Tablas

## <a name="prerequisites"></a>Requisitos previos

Este escenario utiliza [Power Automate](https://flow.microsoft.com) y [Microsoft Teams](https://www.microsoft.com/microsoft-365/microsoft-teams/group-chat-software). Necesitará ambos asociados a la cuenta que usa para desarrollar scripts de Office. Para obtener acceso gratuito a una suscripción a Microsoft Developer para obtener información y trabajar con estas aplicaciones, considere la posibilidad de unirse al [programa para desarrolladores de Microsoft 365.](https://developer.microsoft.com/microsoft-365/dev-program)

## <a name="setup-instructions"></a>Instrucciones de configuración

1. Descarga <a href="task-reminders.xlsx">task-reminders.xlsx</a> en tu OneDrive.

2. Abra el libro en Excel en la Web.

3. En la pestaña **Automatizar,** abra **Todos los scripts**.

4. En primer lugar, necesitamos un script para obtener todos los empleados con informes de estado que faltan en la hoja de cálculo. En el panel de tareas **Editor de código,** presione **Nuevo script** y pegue el siguiente script en el editor.

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

5. Guarde el script con el nombre **Get People**.

6. A continuación, necesitamos un segundo script para procesar las tarjetas de informe de estado y poner la nueva información en la hoja de cálculo. En el panel de tareas **Editor de código,** presione **Nuevo script** y pegue el siguiente script en el editor.

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

7. Guarde el script con el nombre **Guardar estado**.

8. Ahora, necesitamos crear el flujo. Abra [Power Automate](https://flow.microsoft.com/).

    > [!TIP]
    > Si no ha creado un flujo antes, consulte nuestro tutorial Empezar a [usar scripts con Power Automate](../../tutorials/excel-power-automate-manual.md) para aprender lo básico.

9. Cree un nuevo **flujo instantáneo.**

10. Elija **Activar manualmente un flujo** desde las opciones y pulse **Crear**.

11. El flujo debe llamar al script **Get People** para obtener todos los empleados con campos de estado vacíos. Pulse **Nuevo paso** y seleccione Excel En línea **(Negocio).** En **Acciones**, seleccione **Ejecutar script**. Proporcione las siguientes entradas para el paso de flujo:

    - **Ubicación**: OneDrive para la Empresa
    - **Biblioteca de documentos**: OneDrive
    - **Archivo**: task-reminders.xlsx *(Elegido a través del explorador de archivos)*
    - **Guión**: Obtener personas

    :::image type="content" source="../../images/scenario-task-reminders-first-flow-step.png" alt-text="El flujo Power Automate que muestra el primer paso de flujo de script run":::

12. A continuación, el flujo debe procesar cada empleado de la matriz devuelta por el script. Pulse **Nuevo paso** y seleccione Publicar una tarjeta adaptable en un usuario Teams y espere una **respuesta.**

13. Para el campo **Destinatario,** agregue **correo electrónico** desde el contenido dinámico (la selección tendrá el logotipo de Excel por él). Agregar **correo electrónico** hace que el paso de flujo esté rodeado por un Aplicar a **cada** bloque. Esto significa que la matriz se iterará en iterado por Power Automate.

14. El envío de una tarjeta adaptable requiere que el JSON de la tarjeta se proporcione como **mensaje.** Puede usar el [Diseñador de tarjetas adaptables](https://adaptivecards.io/designer/) para crear tarjetas personalizadas. Para este ejemplo, utilice el siguiente JSON.  

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

15. Rellene los campos restantes de la siguiente manera:

    - **Mensaje de actualización**: Gracias por enviar su informe de estado. Su respuesta se ha agregado correctamente a la hoja de cálculo.
    - **Debe actualizar la tarjeta**: Sí

16. En el bloque **Aplicar a cada** bloque, siguiendo la tecla Registrar una tarjeta adaptable a un usuario Teams y esperar una **respuesta**, pulse Agregar **una acción**. Seleccione **Excel en línea (negocio).** En **Acciones**, seleccione **Ejecutar script**. Proporcione las siguientes entradas para el paso de flujo:

    - **Ubicación**: OneDrive para la Empresa
    - **Biblioteca de documentos**: OneDrive
    - **Archivo**: task-reminders.xlsx *(Elegido a través del explorador de archivos)*
    - **Script**: Guardar estado
    - **senderEmail**: correo electrónico *(contenido dinámico de Excel)*
    - **statusReportResponse**: respuesta *(contenido dinámico de Teams)*

    :::image type="content" source="../../images/scenario-task-reminders-last-flow-step.png" alt-text="El flujo Power Automate que muestra el paso de aplicar a cada":::

17. Guarde el flujo.

## <a name="running-the-flow"></a>Ejecutar el flujo

Para probar el flujo, asegúrese de que las filas de la tabla con estado en blanco usen una dirección de correo electrónico vinculada a una cuenta Teams (probablemente debería usar su propia dirección de correo electrónico durante las pruebas).

Puede seleccionar **Probar** en el diseñador de flujo o ejecutar el flujo desde la página **Mis flujos.** Después de iniciar el flujo y aceptar el uso de las conexiones necesarias, debe recibir una tarjeta adaptable de Power Automate a través de Teams. Una vez que rellene el campo de estado en la tarjeta, el flujo continuará y actualizará la hoja de cálculo con el estado que proporcione.

### <a name="before-running-the-flow"></a>Antes de ejecutar el flujo

:::image type="content" source="../../images/scenario-task-reminders-spreadsheet-before.png" alt-text="Una hoja de cálculo con un informe de estado que contiene una entrada de estado que falta":::

### <a name="receiving-the-adaptive-card"></a>Recepción de la tarjeta adaptativa

:::image type="content" source="../../images/scenario-task-reminders-adaptive-card.png" alt-text="Una tarjeta adaptable en Teams pidiendo al empleado una actualización de estado":::

### <a name="after-running-the-flow"></a>Después de ejecutar el flujo

:::image type="content" source="../../images/scenario-task-reminders-spreadsheet-after.png" alt-text="Una hoja de cálculo con un informe de estado con una entrada de estado ahora rellenada":::
