---
title: Ejecutar scripts mediante un flujo manual de Power Automate
description: Tutorial sobre el uso de scripts de Office en Power Automate mediante un desencadenador manual.
ms.date: 06/29/2021
ms.localizationpriority: high
ms.openlocfilehash: e926540976dc066b3f07620c1e710dfa3abc7660
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: HT
ms.contentlocale: es-ES
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585943"
---
# <a name="call-scripts-from-a-manual-power-automate-flow"></a>Ejecutar scripts mediante un flujo manual de Power Automate

Este tutorial le enseña a ejecutar un script de Office para Excel en la Web mediante [Power Automate](https://flow.microsoft.com). Va a escribir un script que actualiza los valores de dos celdas con la hora actual. Luego, conectaremos ese script a un flujo de Power Automate activado manualmente, de modo que el script se ejecute siempre que se seleccione un botón en Power Automate. Cuando entienda el funcionamiento básico, podrá expandir el flujo para incluir otras aplicaciones y automatizar una mayor parte de su flujo de trabajo diario.

> [!TIP]
> Si no está familiarizado con los scripts de Office, le recomendamos que empiece con el tutorial de [Guardar, editar y crear scripts de Office en Excel en la Web](excel-tutorial.md). [Scripts de Office usa TypeScript](../overview/code-editor-environment.md) y este tutorial está diseñado para las personas con conocimientos de nivel intermedio de JavaScript o TypeScript. Si no está familiarizado con JavaScript, le recomendamos que comience con el [Tutorial de JavaScript de Mozilla](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).

## <a name="prerequisites"></a>Requisitos previos

[!INCLUDE [Tutorial prerequisites](../includes/power-automate-tutorial-prerequisites.md)]

## <a name="prepare-the-workbook"></a>Preparar el libro

Power Automate no debe usar [referencias relativas](../testing/power-automate-troubleshooting.md#avoid-relative-references) como `Workbook.getActiveWorksheet` para acceder a componentes de libros de trabajo. Por lo tanto, es necesario contar con un libro de trabajo y una hoja de cálculo con nombres coherentes a los que pueda hacer referencia Power Automate.

1. Cree un nuevo libro de trabajo y llámelo **Mi libro de trabajo**.

2. En el libro **Mi libro de trabajo**, cree una hoja de cálculo y llámela **Hoja de cálculo del tutorial**.

## <a name="create-an-office-script"></a>Cree un script de Office

1. Vaya a la pestaña **Automatizar** y seleccione **Todos los scripts**.

2. Seleccione **Nuevo script**.

3. Reemplace el script predeterminado con el siguiente script. Este script agrega la fecha y la hora actuales a las dos primeras celdas de la hoja de cálculo **Hoja de cálculo del tutorial**.

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
      // Get the "TutorialWorksheet" worksheet from the workbook.
      let worksheet = workbook.getWorksheet("TutorialWorksheet");

      // Get the cells at A1 and B1.
      let dateRange = worksheet.getRange("A1");
      let timeRange = worksheet.getRange("B1");

      // Get the current date and time using the JavaScript Date object.
      let date = new Date(Date.now());

      // Add the date string to A1.
      dateRange.setValue(date.toLocaleDateString());

      // Add the time string to B1.
      timeRange.setValue(date.toLocaleTimeString());
    }
    ```

4. Cambie el nombre del script a **Establecer la fecha y la hora**. Seleccione el nombre del script para cambiarlo.

5. Para guardar el script, seleccione **Guardar script**.

## <a name="create-an-automated-workflow-with-power-automate"></a>Crear un flujo de trabajo automatizado con Power Automate

1. Inicie sesión en el [sitio de Power Automate](https://flow.microsoft.com).

2. En el menú que se muestra en la parte izquierda de la pantalla, seleccione **Crear**. Se mostrará una lista de maneras de crear flujos de trabajo nuevos.

    :::image type="content" source="../images/power-automate-tutorial-1.png" alt-text="El botón 'Crear' de Power Automate.":::

3. En la sección **Inicio desde cero**, seleccione **Flujo instantáneo**. Se creará un flujo de trabajo activado manualmente.

    :::image type="content" source="../images/power-automate-tutorial-2.png" alt-text="La opción de flujo instantáneo de Power Automate para crear un nuevo flujo de trabajo.":::

4. En la ventana de diálogo que aparece, escriba un nombre para el flujo en el cuadro de texto **Nombre de flujo**, seleccione **Activar manualmente un flujo** de la lista de opciones en **Elija cómo desencadenar el flujo** y después seleccione **Crear**.

    :::image type="content" source="../images/power-automate-tutorial-3.png" alt-text="La opción 'Activar manualmente un flujo' de Power Automate.":::

    Tenga en cuenta que un flujo activado manualmente es solo uno de los distintos tipos de flujos. En el siguiente tutorial, podrá crear un flujo que se ejecuta automáticamente al recibir un correo electrónico.

5. Seleccione **Nuevo paso**.

6. Seleccione la pestaña **Estándar** y, a continuación, seleccione **Excel Online (empresa)**.

    :::image type="content" source="../images/power-automate-tutorial-4.png" alt-text="Opción de Excel Online (empresa) en Power Automate.":::

7. En **Acciones**, seleccione **Ejecutar script**.

    :::image type="content" source="../images/power-automate-tutorial-5.png" alt-text="Opción de acción ejecutar script en Power Automate":::

8. A continuación, seleccione el libro y el script que va a usar en el paso de flujo. En el tutorial, usará el libro que creó en OneDrive, pero puede usar cualquier libro en un sitio de OneDrive o SharePoint. Especifique la siguiente configuración para el conector **Ejecutar script**:

    - **Ubicación**: OneDrive para la Empresa
    - **Biblioteca de documentos**: OneDrive
    - **Archivo**: MyWorkbook.xlsx *(seleccionado por el explorador de archivos)*
    - **Script**: Establecer fecha y hora

    :::image type="content" source="../images/power-automate-tutorial-6.png" alt-text="La configuración del conector de Power Automate para ejecutar un script.":::

9. Seleccione **Guardar**.

El flujo ya está listo para ejecutarse mediante Power Automate. Puede probarlo con el botón **Probar** en el editor de flujos o siga los pasos restantes del tutorial para ejecutar el flujo desde la colección de flujos.

## <a name="run-the-script-through-power-automate"></a>Ejecutar el script mediante Power Automate

1. En la página principal de Power Automate, seleccione **Mis flujos**.

    :::image type="content" source="../images/power-automate-tutorial-7.png" alt-text="El botón Mis flujos en Power Automate":::

2. Seleccione **Mi flujo de tutoriales** en la lista de flujos que se muestra en la pestaña **Mis flujos**. Se mostrarán los detalles del flujo que creó anteriormente.

3. Seleccione **Ejecutar**.

    :::image type="content" source="../images/power-automate-tutorial-8.png" alt-text="El botón Ejecutar en Power Automate":::

4. Se mostrará un panel de tareas para ejecutar el flujo. Si se le solicita **Iniciar sesión** en Excel Online, seleccione **Continuar**.

5. Seleccione **Ejecutar flujo**. Esto ejecuta el flujo, que ejecuta el script de Office relacionado.

6. Seleccione **Listo**. En consecuencia, debería actualizarse la sección **Ejecuciones**.

7. Actualice la página para ver los resultados de Power Automate. Si se ha realizado correctamente, podrá ver las celdas actualizadas en el libro de trabajo. Si se ha producido un error, compruebe la configuración del flujo y ejecútelo de nuevo.

    :::image type="content" source="../images/power-automate-tutorial-9.png" alt-text="Resultado de Power Automate que muestra una ejecución de flujo satisfactoria":::

## <a name="next-steps"></a>Pasos siguientes

Complete el tutorial [Pasar datos a scripts en un flujo de Power Automate ejecutado automáticamente](excel-power-automate-trigger.md). Aprenderá a pasar datos desde un servicio de flujo de trabajo al script de Office y a ejecutar el flujo de Power Automate cuando se producen determinados eventos.
