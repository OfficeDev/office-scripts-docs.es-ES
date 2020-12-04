---
title: 'Escenario de ejemplo de scripts de Office: avisos de tareas automatizadas'
description: Un ejemplo que usa la automatización de energía y las tarjetas adaptables automatiza los avisos de tareas en una hoja de cálculo de administración de proyectos.
ms.date: 11/30/2020
localization_priority: Normal
ms.openlocfilehash: 90769eca0201e450e25778db0eb5c62284b9feb0
ms.sourcegitcommit: af487756dffea0f8f0cd62710c586842cb08073c
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 12/04/2020
ms.locfileid: "49571454"
---
# <a name="office-scripts-sample-scenario-automated-task-reminders"></a>Escenario de ejemplo de scripts de Office: avisos de tareas automatizadas

En este escenario, está administrando un proyecto. Use una hoja de cálculo de Excel para realizar un seguimiento del estado de sus empleados cada mes. Con frecuencia, debe recordar a los usuarios que deben rellenar su estado, por lo que decidió automatizar el proceso de recordatorio.

Creará un flujo de automatización de potencia para avisar a los usuarios de los campos de estado que faltan y aplicar sus respuestas a la hoja de cálculo. Para ello, desarrollará un par de scripts para controlar el trabajo con el libro. El primer script obtiene una lista de personas con Estados en blanco y el segundo script agrega una cadena de estado a la fila de la derecha. También usará las [tarjetas adaptables de Microsoft Teams](/microsoftteams/platform/task-modules-and-cards/what-are-cards) para que los empleados especifiquen su estado directamente desde la notificación.

## <a name="scripting-skills-covered"></a>Habilidades de scripting cubiertas

- Crear flujos con Power automatization
- Pasar datos a scripts
- Devolver datos de scripts
- Tarjetas adaptables de Teams
- Tablas

## <a name="prerequisites"></a>Requisitos previos

Este escenario usa la [funcionalidad de automatización](https://flow.microsoft.com) y [Microsoft Teams](https://www.microsoft.com/microsoft-365/microsoft-teams/group-chat-software). Necesitará ambos asociados con la cuenta que usa para el desarrollo de scripts de Office. Para obtener acceso gratuito a una suscripción de Microsoft Developer para conocer y trabajar con estas aplicaciones, considere la posibilidad de unirse al [programa de desarrolladores de microsoft 365](https://developer.microsoft.com/microsoft-365/dev-program).

## <a name="setup-instructions"></a>Instrucciones de instalación

1. Descargue <a href="task-reminders.xlsx">task-reminders.xlsx</a> a su OneDrive.

2. Abra el libro en Excel en la Web.

3. En la ficha **automatizar** , abra el **Editor de código**.

4. En primer lugar, necesitamos un script para obtener todos los empleados con informes de estado que faltan en la hoja de cálculo. En el panel de tareas **Editor de código** , presione **nueva secuencia** de comandos y pegue el siguiente script en el editor.

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

6. A continuación, necesitamos un segundo script para procesar las tarjetas del informe de estado y colocar la nueva información en la hoja de cálculo. En el panel de tareas **Editor de código** , presione **nueva secuencia** de comandos y pegue el siguiente script en el editor.

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

7. Guarde el script con el nombre **Guardar estado**.

8. Ahora, es necesario crear el flujo. Abra [energía automatizada](https://flow.microsoft.com/).

    > [!TIP]
    > Si no ha creado un flujo antes, consulte nuestro tutorial comenzar a [usar scripts con Power automatization](../../tutorials/excel-power-automate-manual.md) para conocer los conceptos básicos.

9. Cree un nuevo **flujo instantáneo**.

10. Elija **desencadenar manualmente un flujo** de las opciones y pulse **crear**.

11. El flujo tiene que llamar al script **Get People** para obtener todos los empleados con campos de estado vacíos. Presione **nuevo paso** y seleccione **Excel online (empresa)**. En **Acciones**, seleccione **Ejecutar script (versión preliminar)**. Proporcione las siguientes entradas para el paso flujo:

    - **Ubicación**: OneDrive para la Empresa
    - **Biblioteca de documentos**: OneDrive
    - **File**: task-reminders.xlsx *(elegido mediante el explorador de archivos)*
    - **Script**: obtener personas

    ![Paso del flujo de script de primera ejecución.](../../images/scenario-task-reminders-first-flow-step.png)

12. A continuación, el flujo debe procesar cada empleado en la matriz que devuelve el script. Presione **nuevo paso** y seleccione **publicar una tarjeta adaptable en un usuario de Teams y espere una respuesta**.

13. Para el campo de **destinatario** , agregue **correo electrónico** del contenido dinámico (la selección tendrá el logotipo de Excel). Agregar **correo electrónico** hace que el paso de flujo esté rodeado por una **aplicación a cada** bloque. Esto significa que la matriz se recorrerá en iteración mediante la automatización de energía.

14. El envío de una tarjeta adaptable requiere que se proporcione el JSON de la tarjeta como **mensaje**. Puede usar el [Diseñador de tarjetas adaptable](https://adaptivecards.io/designer/) para crear tarjetas personalizadas. Para este ejemplo, use el siguiente JSON.  

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

    - **Mensaje de actualización**: Gracias por enviar el informe de estado. La respuesta se ha agregado correctamente a la hoja de cálculo.
    - **Debe actualizar la tarjeta**: sí

16. En el bloque **aplicar a cada** bloque, después de **publicar una tarjeta adaptable a un usuario de Teams y esperar una respuesta**, presione **Agregar una acción**. Seleccione **Excel online (empresa)**. En **Acciones**, seleccione **Ejecutar script (versión preliminar)**. Proporcione las siguientes entradas para el paso flujo:

    - **Ubicación**: OneDrive para la Empresa
    - **Biblioteca de documentos**: OneDrive
    - **File**: task-reminders.xlsx *(elegido mediante el explorador de archivos)*
    - **Script**: estado de guardado
    - **senderEmail**: correo electrónico *(contenido dinámico de Excel)*
    - **statusReportResponse**: respuesta *(contenido dinámico de Teams)*

    ![El paso aplicar a cada flujo.](../../images/scenario-task-reminders-last-flow-step.png)

17. Guarde el flujo.

## <a name="running-the-flow"></a>Ejecución del flujo

Para probar el flujo, asegúrese de que todas las filas de tabla con el estado en blanco usan una dirección de correo electrónico ligada a una cuenta de Teams (probablemente debería usar su propia dirección de correo electrónico durante las pruebas).

Puede seleccionar **Test** desde el diseñador de flujo o ejecutar el flujo desde la página **Mis flujos** . Después de iniciar el flujo y aceptar el uso de las conexiones necesarias, debe recibir una tarjeta adaptable de la alimentación automatizada a través de Teams. Una vez que rellene el campo Estado en la tarjeta, el flujo continuará y actualizará la hoja de cálculo con el estado proporcionado.

### <a name="before-running-the-flow"></a>Antes de ejecutar el flujo

![Una hoja de cálculo con un informe de estado que contiene una entrada de estado que falta.](../../images/scenario-task-reminders-spreadsheet-before.png)

### <a name="receiving-the-adaptive-card"></a>Recibir la tarjeta adaptable

![Una tarjeta adaptable de teams que solicita al empleado una actualización de estado.](../../images/scenario-task-reminders-adaptive-card.png)

### <a name="after-running-the-flow"></a>Después de ejecutar el flujo

![Una hoja de cálculo con un informe de estado con una entrada de estado ahora rellenada.](../../images/scenario-task-reminders-spreadsheet-after.png)
