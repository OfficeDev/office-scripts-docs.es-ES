---
title: Ejecutar scripts mediante un flujo manual de Power Automate
description: Tutorial sobre el uso de scripts de Office en Power Automate mediante un desencadenador manual.
ms.date: 07/14/2020
localization_priority: Priority
ms.openlocfilehash: 70fca2620973ecefe9eda40f02e28f064b713677
ms.sourcegitcommit: ebd1079c7e2695ac0e7e4c616f2439975e196875
ms.translationtype: HT
ms.contentlocale: es-ES
ms.lasthandoff: 07/17/2020
ms.locfileid: "45160440"
---
# <a name="call-scripts-from-a-manual-power-automate-flow-preview"></a>Ejecutar scripts mediante un flujo manual de Power Automate (versión preliminar)

Este tutorial le enseña a ejecutar un script de Office para Excel en la Web mediante [Power Automate](https://flow.microsoft.com).

## <a name="prerequisites"></a>Requisitos previos

[!INCLUDE [Tutorial prerequisites](../includes/power-automate-tutorial-prerequisites.md)]

> [!IMPORTANT]
> Este tutorial presupone que usted ha completado el tutorial [Registro, edición y creación de scripts de Office en Excel en la web](excel-tutorial.md).

## <a name="prepare-the-workbook"></a>Preparar el libro de trabajo

Power Automate no puede usar referencias relativas como `Workbook.getActiveWorksheet` para acceder a componentes de libros de trabajo. Por lo tanto, es necesario contar con un libro de trabajo y una hoja de cálculo con nombres coherentes a los que pueda hacer referencia Power Automate.

1. Cree un nuevo libro de trabajo y llámelo **Mi libro de trabajo**.

2. En el libro **Mi libro de trabajo**, cree una hoja de cálculo y llámela **Hoja de cálculo del tutorial**.

## <a name="create-an-office-script"></a>Cree un script de Office

1. Vaya a la pestaña **Automatizar** y seleccione **Editor de código**.

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

4. Cambie el nombre del script a **Establecer la fecha y la hora**. Presione el nombre del script para cambiarlo.

5. Para guardar el script, presione el botón **Guardar script**.

## <a name="create-an-automated-workflow-with-power-automate"></a>Crear un flujo de trabajo automatizado con Power Automate

1. Inicie sesión en el [sitio de Power Automate](https://flow.microsoft.com).

2. En el menú que se muestra en la parte izquierda de la pantalla, presione **Crear**. Se mostrará una lista de maneras de crear flujos de trabajo nuevos.

    ![El botón Crear en Power Automate.](../images/power-automate-tutorial-1.png)

3. En la sección **Inicio desde cero**, seleccione **Flujo instantáneo**. Se creará un flujo de trabajo activado manualmente.

    ![La opción Flujo instantáneo para crear un nuevo flujo de trabajo.](../images/power-automate-tutorial-2.png)

4. En la ventana de diálogo que aparece, escriba un nombre para el flujo en el cuadro de texto **Nombre de flujo**, seleccione **Activar manualmente un flujo** de la lista de opciones en **Elija cómo desencadenar el flujo** y presione **Crear**.

    ![La opción de desencadenador manual para crear un nuevo flujo instantáneo.](../images/power-automate-tutorial-3.png)

    Tenga en cuenta que un flujo activado manualmente es solo uno de los distintos tipos de flujos. En el siguiente tutorial, podrá crear un flujo que se ejecuta automáticamente al recibir un correo electrónico.

5. Presione **Nuevo paso**.

6. Seleccione la pestaña **Estándar** y, a continuación, seleccione **Excel Online (empresa)**.

    ![La opción Power Automate para Excel Online (empresa).](../images/power-automate-tutorial-4.png)

7. En **Acciones**, seleccione **Ejecutar script (versión preliminar)**.

    ![La opción de acción de Power Automate para Ejecutar script (versión preliminar).](../images/power-automate-tutorial-5.png)

8. Especifique la siguiente configuración para el conector **Ejecutar script**:

    - **Ubicación**: OneDrive para la Empresa
    - **Biblioteca de documentos**: OneDrive
    - **Archivo**: MyWorkbook.xlsx
    - **Script**: Establecer fecha y hora

    ![La configuración del conector para ejecutar un script en Power Automate.](../images/power-automate-tutorial-6.png)

9. Presione **Guardar**.

El flujo ya está listo para ejecutarse mediante Power Automate. Para probarlo, pulse el botón **Probar** en el editor de flujos o siga los pasos restantes del tutorial para ejecutar el flujo de la colección de flujos.

## <a name="run-the-script-through-power-automate"></a>Ejecutar el script mediante Power Automate

1. En la página principal de Power Automate, seleccione **Mis flujos**.

    ![El botón Mis flujos en Power Automate.](../images/power-automate-tutorial-7.png)

2. Seleccione **Mi flujo de tutoriales** en la lista de flujos que se muestra en la pestaña **Mis flujos**. Se mostrarán los detalles del flujo que creó anteriormente.

3. Pulse **Ejecutar**.

    ![El botón Ejecutar en Power Automate.](../images/power-automate-tutorial-8.png)

4. Se mostrará un panel de tareas para ejecutar el flujo. Si se le solicita **Iniciar sesión** en Excel Online, presione **Continuar**.

5. Pulse **Ejecutar flujo**. Se ejecutará el flujo, que ejecuta a su vez el script de Office relacionado.

6. Presione **Listo**. En consecuencia, debería actualizarse la sección **Ejecuciones**.

7. Actualice la página para ver los resultados de Power Automate. Si se ha realizado correctamente, podrá ver las celdas actualizadas en el libro de trabajo. Si se ha producido un error, compruebe la configuración del flujo y ejecútelo de nuevo.

    ![Salida de Power Automate que muestra una ejecución de flujo satisfactoria.](../images/power-automate-tutorial-9.png)

## <a name="next-steps"></a>Pasos siguientes

Complete el tutorial [Pasar datos a scripts en un flujo de Power Automate ejecutado automáticamente](excel-power-automate-trigger.md). Aprenderá a pasar datos desde un servicio de flujo de trabajo al script de Office y a ejecutar el flujo de Power Automate cuando se producen determinados eventos.
