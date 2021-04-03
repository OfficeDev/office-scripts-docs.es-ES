---
title: Devolver datos de un script a un flujo de Power Automate ejecutado automáticamente
description: Un tutorial que muestra cómo enviar recordatorios por correo electrónico mediante la ejecución de Scripts de Office para Excel en la Web con Power Automate.
ms.date: 12/15/2020
localization_priority: Priority
ms.openlocfilehash: 31ba31ddbfb36f20087be6aa7d83b1b896a698d1
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: HT
ms.contentlocale: es-ES
ms.lasthandoff: 04/02/2021
ms.locfileid: "51570533"
---
# <a name="return-data-from-a-script-to-an-automatically-run-power-automate-flow-preview"></a>Devolver datos de un script a un flujo de Power Automate (vista previa)

Este tutorial le enseña cómo devolver información de un script de Office para Excel en la Web con un flujo de trabajo automatizado de [Power Automate](https://flow.microsoft.com). Creará un script que busca en una programación y funciona con un flujo para enviar correos electrónicos de recordatorio. Este flujo se ejecutará de forma periódica y proporcionará los recordatorios en su nombre.

> [!TIP]
> Si no está familiarizado con los scripts de Office, le recomendamos que empiece con el tutorial de [Guardar, editar y crear scripts de Office en Excel en la Web](excel-tutorial.md).
>
> Si no está familiarizado con Power Automate, le recomendamos que empiece con los tutoriales de [Llamar a scripts desde un flujo de Power Automate manual](excel-power-automate-manual.md) y [Pasar datos a scripts en un flujo ejecutado automáticamente de Power Automate](excel-power-automate-trigger.md).
>
> [Scripts de Office usa TypeScript](../overview/code-editor-environment.md) y este tutorial está diseñado para las personas con conocimientos de nivel intermedio de JavaScript o TypeScript. Si no está familiarizado con JavaScript, le recomendamos que comience con el [Tutorial de JavaScript de Mozilla](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).

## <a name="prerequisites"></a>Requisitos previos

[!INCLUDE [Tutorial prerequisites](../includes/power-automate-tutorial-prerequisites.md)]

## <a name="prepare-the-workbook"></a>Preparar el libro

1. Descargue el libro <a href="on-call-rotation.xlsx">on-call-rotation.xlsx</a> en su OneDrive.

1. Abra **on-call-rotation.xlsx** en Excel en la Web.

1. Agregue una fila a la tabla con su nombre, dirección de correo electrónico y fechas de inicio y finalización que se solapen con la fecha actual.

    > [!IMPORTANT]
    > El script que escribirá usa la primera entrada coincidente de la tabla, así que asegúrese de que su nombre se encuentre encima de cualquier fila con la semana actual.

    ![Captura de pantalla de la tabla de rotación de guardias en la hoja de cálculo de Excel](../images/power-automate-return-tutorial-1.png)

## <a name="create-an-office-script"></a>Crear un script de Office

1. Vaya a la pestaña **Automatizar** y seleccione **Todos los scripts**.

1. Seleccione **Nuevo script**.

1. Asigne al script el nombre **Obtener persona de guardia**.

1. Ahora debería tener un script vacío. Queremos usar el script para obtener una dirección de correo electrónico de la hoja de cálculo. Cambie `main` para devolver una cadena, de la siguiente manera:

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) : string {
    }
    ```

1. Después, debemos obtener todos los datos de la tabla. Eso nos permite buscar en cada fila con el script. Agregue el código siguiente dentro de la función `main`.

    ```TypeScript
    // Get the H1 worksheet.
    let worksheet = workbook.getWorksheet("H1");

    // Get the first (and only) table in the worksheet.
    let table = worksheet.getTables()[0];

    // Get the data from the table.
    let tableValues = table.getRangeBetweenHeaderAndTotal().getValues();
    ```

1. Las fechas de la tabla se almacenan con [el número de serie de la fecha de Excel](https://support.microsoft.com/office/date-systems-in-excel-e7fe7167-48a9-4b96-bb53-5612a800b487). Es necesario convertir esas fechas en fechas de JavaScript para poder compararlas. Agregaremos una función auxiliar al script. Agregue el código siguiente fuera de la función `main`:

    ```TypeScript
    // Convert the Excel date to a JavaScript Date object.
    function convertDate(excelDateValue: number) {
        let javaScriptDate = new Date(Math.round((excelDateValue - 25569) * 86400 * 1000));
        return javaScriptDate;
    }
    ```

1. Ahora, debemos saber cuál es el usuario de guardia en este momento. La fila tendrá una fecha de inicio y una de finalización en torno a la fecha actual. Escribiremos el script para asumir que solo un usuario está de guardia cada vez. Los scripts pueden devolver matrices para manejar varios valores, pero por ahora devolveremos la primera dirección de correo electrónico coincidente. Agregue el siguiente código al final de la función `main`.

    ```TypeScript
    // Look for the first row where today's date is between the row's start and end dates.
    let currentDate = new Date();
    for (let row = 0; row < tableValues.length; row++) {
        let startDate = convertDate(tableValues[row][2] as number);
        let endDate = convertDate(tableValues[row][3] as number);
        if (startDate <= currentDate && endDate >= currentDate) {
            // Return the first matching email address.
            return tableValues[row][1].toString();
        }
    }
    ```

1. El script final debería tener un aspecto similar al siguiente:

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) : string {
        // Get the H1 worksheet.
        let worksheet = workbook.getWorksheet("H1");

        // Get the first (and only) table in the worksheet.
        let table = worksheet.getTables()[0];
    
        // Get the data from the table.
        let tableValues = table.getRangeBetweenHeaderAndTotal().getValues();
    
        // Look for the first row where today's date is between the row's start and end dates.
        let currentDate = new Date();
        for (let row = 0; row < tableValues.length; row++) {
            let startDate = convertDate(tableValues[row][2] as number);
            let endDate = convertDate(tableValues[row][3] as number);
            if (startDate <= currentDate && endDate >= currentDate) {
                // Return the first matching email address.
                return tableValues[row][1].toString();
            }
        }
    }

    // Convert the Excel date to a JavaScript Date object.
    function convertDate(excelDateValue: number) {
        let javaScriptDate = new Date(Math.round((excelDateValue - 25569) * 86400 * 1000));
        return javaScriptDate;
    }
    ```

## <a name="create-an-automated-workflow-with-power-automate"></a>Crear un flujo de trabajo automatizado con Power Automate

1. Inicie sesión en el [sitio de Power Automate](https://flow.microsoft.com).

1. En el menú que se muestra en la parte izquierda de la pantalla, presione **Crear**. Se mostrará una lista de maneras de crear flujos de trabajo nuevos.

    ![El botón Crear en Power Automate](../images/power-automate-tutorial-1.png)

1. En la sección **Empezar desde cero**, seleccione **Flujo de nube programado**.

    ![El botón de Flujo de nube programado en Power Automate](../images/power-automate-return-tutorial-2.png)

1. Ahora debemos establecer la programación para este flujo. Nuestra hoja de cálculo tiene una nueva asignación de guardia que empieza todos los lunes en la primera mitad de 2021. Vamos a configurar el flujo para que se ejecute a primera hora los lunes por la mañana. Use las opciones siguientes para configurar el flujo de ejecución el lunes de cada semana.

    - **Nombre de flujo**: notificar a la persona de guardia
    - **Inicio**: 4/1/21 a la 01:00
    - **Repetir cada**: 1 semana
    - **En estos días**: L

    ![Ventana que muestra las opciones especificadas para el flujo programado](../images/power-automate-return-tutorial-3.png)

1. Presione **Crear**.

1. Presione **Nuevo paso**.

1. Seleccione la pestaña **Estándar** y, a continuación, seleccione **Excel Online (empresa)**.

    ![Opción de Excel Online (empresa) en Power Automate](../images/power-automate-tutorial-4.png)

1. En **Acciones**, seleccione **Ejecutar script (versión preliminar)**.

    ![Opción de acción Ejecutar script (vista previa) en Power Automate](../images/power-automate-tutorial-5.png)

1. Después, seleccione el libro y el script que va a usar en el paso del flujo. Use el libro **on-call-rotation.xlsx** que creó en OneDrive. Especifique la siguiente configuración para el conector **Ejecutar script**:

    - **Ubicación**: OneDrive para la Empresa
    - **Biblioteca de documentos**: OneDrive
    - **Archivo**: on-call-rotation.xlsx *(seleccionado en el explorador de archivos)*
    - **Script**: obtener persona de guardia

    ![La configuración del conector para ejecutar un script en Power Automate](../images/power-automate-return-tutorial-4.png)

1. Presione **Nuevo paso**.

1. Finalizaremos el flujo enviando el correo electrónico de recordatorio. Seleccione **Enviar un correo electrónico (V2)** mediante la barra de búsqueda del conector. Use el control **Agregar contenido dinámico** para agregar la dirección de correo electrónico que devuelve el script. Se etiquetará como **resultado** con el icono de Excel situado al lado. Puede proporcionar el asunto y el texto de cuerpo que prefiera.

    ![La configuración del conector para enviar un correo electrónico en Power Automate](../images/power-automate-return-tutorial-5.png)

    > [!NOTE]
    > Este tutorial usa Outlook. Puede usar el servicio de correo electrónico que prefiera, aunque algunas opciones pueden ser diferentes.

1. Presione **Guardar**.

## <a name="test-the-script-in-power-automate"></a>Probar el script en Power Automate

El flujo se ejecutará cada lunes por la mañana. Para probar el script ahora, presione el botón **Probar** en la esquina superior derecha de la pantalla. Seleccione **Manualmente** y presione **Ejecutar prueba** para ejecutar el flujo ahora y probar el comportamiento. Es posible que deba conceder permisos a Excel y Outlook para continuar.

![El botón de prueba de Power Automate](../images/power-automate-return-tutorial-6.png)

> [!TIP]
> Si el flujo no puede enviar un correo electrónico, vuelva a comprobar en la hoja de cálculo que se muestra un correo electrónico válido para el intervalo de fechas actual en la parte superior de la tabla.

## <a name="next-steps"></a>Pasos siguientes

Visite [Ejecutar scripts de Office con Power Automate](../develop/power-automate-integration.md) para más información sobre la conexión de Scripts de Office con Power Automate.

También puede consultar el [Escenario de muestra de recordatorios de tareas automatizados](../resources/scenarios/task-reminders.md) para aprender a combinar los Scripts de Office y Power Automate con las Tarjetas adaptables de Teams.
