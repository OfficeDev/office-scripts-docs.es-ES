---
title: Pasar datos a scripts en un flujo de Power Automate ejecutado automáticamente
description: Un tutorial sobre la ejecución de Scripts de Office para Excel en la Web mediante Power Automate cuando se reciba el correo y el paso de datos de flujo al script.
ms.date: 06/10/2022
ms.localizationpriority: high
ms.openlocfilehash: 73a551df09eadba1f6e75de35e17e1c5a93498e9
ms.sourcegitcommit: dd01979d34b3499360d2f79a56f8a8f24f480eed
ms.translationtype: HT
ms.contentlocale: es-ES
ms.lasthandoff: 06/15/2022
ms.locfileid: "66088144"
---
# <a name="pass-data-to-scripts-in-an-automatically-run-power-automate-flow"></a>Pasar datos a scripts en un flujo de Power Automate ejecutado automáticamente

Este tutorial le enseña cómo usar un script de Office para Excel en la web con un flujo de trabajo automatizado de [Power Automate](https://flow.microsoft.com). El script se ejecutará automáticamente cada vez que reciba un correo electrónico, grabando información del correo en un libro de Excel. Pasar datos de otras aplicaciones a un script de Office le ofrece una gran flexibilidad y libertad para sus procesos automatizados.

> [!TIP]
> Si no está familiarizado con los scripts de Office, le recomendamos que empiece con el tutorial de [Guardar, editar y crear scripts de Office en Excel en la Web](excel-tutorial.md). Si es la primera vez que usa Power Automate, le recomendamos que comience con el tutorial [Llamar a scripts desde un flujo manual de Power Automate](excel-power-automate-manual.md). [Scripts de Office usa TypeScript](../overview/code-editor-environment.md) y este tutorial está diseñado para las personas con conocimientos de nivel intermedio de JavaScript o TypeScript. Si no está familiarizado con JavaScript, le recomendamos que comience con el [Tutorial de JavaScript de Mozilla](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).

## <a name="prerequisites"></a>Requisitos previos

[!INCLUDE [Tutorial prerequisites](../includes/power-automate-tutorial-prerequisites.md)]

## <a name="prepare-the-workbook"></a>Preparar el libro

Power Automate no debe usar [referencias relativas](../testing/power-automate-troubleshooting.md#avoid-relative-references) como `Workbook.getActiveWorksheet` para acceder a componentes de libros de trabajo. Por lo tanto, es necesario un libro de trabajo y una hoja de cálculo con nombres coherentes para que Power Automate haga referencia.

1. Cree un nuevo libro de trabajo y llámelo **Mi libro de trabajo**.

2. Vaya a la pestaña **Automatizar** y seleccione **Todos los scripts**.

3. Seleccione **Nuevo script**.

4. Reemplace el código existente con el siguiente script y seleccione **Ejecutar**. Esto configurará el libro con nombres de tabla dinámica, hoja de cálculo y tabla coherentes.

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

## <a name="create-an-office-script"></a>Crear un script de Office

Comencemos a crear un script que registre información de un correo electrónico. Queremos saber en qué días de la semana recibimos más correos electrónicos y cuántos remitentes únicos nos los envían. Nuestro libro tiene una tabla con columnas de **Fecha**, **Día de la semana**, **Dirección de correo electrónico** y **Asunto**. Nuestra hoja de cálculo también tiene una tabla dinámica que se dinamiza en el **Día de la semana** y **Dirección de correo electrónico** (que son las jerarquías de fila). El recuento de **Asuntos** únicos es la información agregada que se muestra (la jerarquía de datos). Haremos que nuestro script actualice esa tabla dinámica después de actualizar la tabla de correo electrónico.

1. Desde el panel de tareas del Editor de código, seleccione **Nuevo script**.

2. El flujo que crearemos más adelante en el tutorial enviará la información de script de cada mensaje de correo electrónico que se reciba. El script necesita aceptar esa entrada mediante parámetros en la función `main`. Reemplace el script predeterminado con el siguiente script:

    ```TypeScript
    function main(
      workbook: ExcelScript.Workbook,
      from: string,
      dateReceived: string,
      subject: string) {

    }
    ```

3. El script necesita acceso a la tabla del libro de trabajo y a la tabla dinámica. Agregue el siguiente código al cuerpo del script, después de la apertura `{`:

    ```TypeScript
    // Get the email table.
    let emailWorksheet = workbook.getWorksheet("Emails");
    let table = emailWorksheet.getTable("EmailTable");
  
    // Get the PivotTable.
    let pivotTableWorksheet = workbook.getWorksheet("Subjects");
    let pivotTable = pivotTableWorksheet.getPivotTable("Pivot");
    ```

4. El parámetro `dateReceived` es de tipo `string`. Vamos a convertir esto en un [objeto `Date`](../develop/javascript-objects.md#date) para que podamos obtener fácilmente el día de la semana. Después de hacerlo, deberemos asignar el valor numérico del día a una versión más legible. Agregue el código siguiente al final del script, antes del cierre `}`:

    ```TypeScript
      // Parse the received date string to determine the day of the week.
      let emailDate = new Date(dateReceived);
      let dayName = emailDate.toLocaleDateString("en-US", { weekday: 'long' });
    ```

5. La cadena `subject` puede incluir la etiqueta de respuesta "RE:". Eliminemos eso de la cadena para que los correos electrónicos en el mismo hilo tengan el mismo asunto para la tabla. Agregue el código siguiente al final del script, antes del cierre `}`:

    ```TypeScript
    // Remove the reply tag from the email subject to group emails on the same thread.
    let subjectText = subject.replace("Re: ", "");
    subjectText = subjectText.replace("RE: ", "");
    ```

6. Ahora que se ha dado formato a los datos de correo electrónico a nuestro gusto, agreguemos una fila a la tabla de correo electrónico. Agregue el código siguiente al final del script, antes del cierre `}`:

    ```TypeScript
    // Add the parsed text to the table.
    table.addRow(-1, [dateReceived, dayName, from, subjectText]);
    ```

7. Por último, vamos a asegurarnos de que se actualiza la tabla dinámica. Agregue el código siguiente al final del script, antes del cierre `}`:

    ```TypeScript
    // Refresh the PivotTable to include the new row.
    pivotTable.refresh();
    ```

8. Cambie el nombre del script **Registrar correo electrónico** y seleccione **Guardar script**.

El script ya está preparado para un flujo de trabajo de Power Automate. Debería ser similar al siguiente script:

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

## <a name="create-an-automated-workflow-with-power-automate"></a>Crear un flujo de trabajo automatizado con Power Automate

1. Inicie sesión en el [sitio de Power Automate](https://flow.microsoft.com).

2. En el menú que se muestra en la parte izquierda de la pantalla, seleccione **Crear**. Se mostrará una lista de maneras de crear flujos de trabajo nuevos.

    :::image type="content" source="../images/power-automate-tutorial-1.png" alt-text="El botón Crear en Power Automate.":::

3. En la sección **Inicio desde cero**, seleccione **Flujo automatizado**. Esto creará un flujo de trabajo desencadenado por un evento, como la recepción de un correo electrónico.

    :::image type="content" source="../images/power-automate-params-tutorial-1.png" alt-text="La opción de Flujo automatizado en Power Automate.":::

4. En la ventana de diálogo que aparece, escriba un nombre para su flujo en el cuadro de texto **Nombre de flujo**. A continuación, seleccione **Cuando llegue un nuevo correo electrónico** de la lista de opciones de **Elegir el desencadenador de flujo**. Es posible que tenga que buscar la opción con el cuadro de búsqueda. Por último, seleccione **Crear**.

    :::image type="content" source="../images/power-automate-params-tutorial-2.png" alt-text="Forma parte del flujo de Power Automate que muestra el «nombre del flujo» y las opciones de «elegir el desencadenador del flujo». El nombre del flujo es «Registrar flujo de correo electrónico» y el desencadenador es la opción «Cuando llega un correo electrónico nuevo a Outlook».":::

    > [!NOTE]
    > Este tutorial usa Outlook. Usted puede usar el servicio de correo electrónico que prefiera, aunque algunas opciones pueden ser diferentes.

5. Seleccione **Nuevo paso**.

6. Seleccione la pestaña **Estándar** y, a continuación, seleccione **Excel Online (empresa)**.

    :::image type="content" source="../images/power-automate-tutorial-4.png" alt-text="Opción de Excel Online (empresa) en Power Automate.":::

7. En **Acciones**, seleccione **Ejecutar script**.

    :::image type="content" source="../images/power-automate-tutorial-5.png" alt-text="Opción de acción ejecutar script en Power Automate":::

8. A continuación, seleccione el libro, el script y los argumentos de entrada del script que se usará en el paso de flujo. En el tutorial, usará el libro que creó en OneDrive, pero puede usar cualquier libro en un sitio de OneDrive o SharePoint. Especifique la siguiente configuración para el conector **Ejecutar script**:

    - **Ubicación**: OneDrive para la Empresa
    - **Biblioteca de documentos**: OneDrive
    - **Archivo**: MyWorkbook.xlsx *(seleccionado por el explorador de archivos)*
    - **Script**: Registrar correo electrónico
    - **de**: De *(contenido dinámico de Outlook)*
    - **dateReceived**: Hora de recepción *(contenido dinámico de Outlook)*
    - **asunto**: Asunto *(contenido dinámico de Outlook)*

    *Tenga en cuenta que los parámetros del script solo aparecen cuando se selecciona el script.*

    :::image type="content" source="../images/power-automate-params-tutorial-3.png" alt-text="La acción de ejecutar script de Power Automate muestra las opciones que aparecen una vez seleccionado el script.":::

9. Seleccione **Guardar**.

El flujo ya está habilitado. El script se ejecutará automáticamente cada vez que reciba un correo electrónico a través de Outlook.

## <a name="manage-the-script-in-power-automate"></a>Administrar el script en Power Automate

1. En la página principal de Power Automate, seleccione **Mis flujos**.

    :::image type="content" source="../images/power-automate-tutorial-7.png" alt-text="El botón Mis flujos en Power Automate.":::

2. Seleccione el flujo. Aquí puede ver el historial de ejecución. Puede actualizar la página o seleccionar el botón actualizar **Todas las ejecuciones** para actualizar el historial. El flujo se desencadenará poco después de que se reciba un correo electrónico. Pruebe el flujo enviándose un correo electrónico a sí mismo.

Cuando se desencadene el flujo y se ejecute correctamente el script, debería ver que se actualizan la tabla dinámica y la tabla del libro.

:::image type="content" source="../images/power-automate-params-tutorial-4.png" alt-text="Una hoja de cálculo que muestra la tabla de correo electrónico después de que el flujo se haya ejecutado tres veces.":::

:::image type="content" source="../images/power-automate-params-tutorial-5.png" alt-text="Una hoja de cálculo que muestra la tabla dinámica después de que el flujo se haya ejecutado tres veces.":::

## <a name="troubleshooting"></a>Solución de problemas

Recibir varios correos electrónicos al mismo tiempo puede causar conflictos de combinación en Excel. Este riesgo se mitiga configurando el conector de correo electrónico para que solo actúe en un correo electrónico a la vez. Para hacerlo:

1. Seleccione el botón **Menu (…)** en el conector de correo electrónico y, a continuación, seleccione **Configuración**.

    :::image type="content" source="../images/outlook-connector-settings-1.png" alt-text="La opción de configuración resaltada en el menú del conector.":::

1. En la **Configuración** en las opciones emergentes, establezca **Control de simultaneidad** en **Activado**. A continuación, establezca el **grado de paralelismo** en **1**.

    :::image type="content" source="../images/outlook-connector-settings-2.png" alt-text="Las opciones de simultaneidad en el menú de configuración.":::

## <a name="next-steps"></a>Pasos siguientes

Complete el tutorial [Devolver datos de un script a un flujo de Power Automate ejecutado automáticamente](excel-power-automate-returns.md). Muestra cómo devolver datos de un script al flujo.

También puede consultar el [Escenario de muestra de recordatorios de tareas automatizados](../resources/scenarios/task-reminders.md) para aprender a combinar los Scripts de Office y Power Automate con las Tarjetas adaptables de Teams.
