---
title: Ejecución de scripts de Office en Excel con botones
description: Agregue botones a los libros que controlan Office scripts en Excel.
ms.topic: overview
ms.date: 05/09/2022
ms.localizationpriority: medium
ms.openlocfilehash: cc19a13a97d4d11f73cb91bc46b70afff3eadf03
ms.sourcegitcommit: aecbd5baf1e2122d836c3eef3b15649e132bc68e
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 06/16/2022
ms.locfileid: "66128219"
---
# <a name="run-office-scripts-in-excel-with-buttons"></a>Ejecución de scripts de Office en Excel con botones

Ayude a sus compañeros a buscar y ejecutar los scripts agregando botones de script a un libro.

:::image type="content" source="../images/run-from-button.png" alt-text="Un botón de la hoja de cálculo que ejecuta un script al hacer clic en él.":::

## <a name="create-script-buttons"></a>Crear botones de script

Con cualquier script, vaya al menú **Más opciones (...)** en la página de detalles del script o en el panel de tareas del Editor de código y seleccione el **botón Agregar**. Esto creará un botón en el libro que ejecutará el script asociado cuando se seleccione. También compartirá el script con el libro, por lo que todos los usuarios con permisos de escritura en el libro pueden usar esta automatización útil.

En la captura de pantalla siguiente se muestra la página de detalles del script de un script titulado **Crear tabla dinámica** y tiene la opción **de botón Agregar** en el menú **Más opciones (...)** resaltado.

:::image type="content" source="../images/add-button.png" alt-text="La opción &quot;Agregar botón&quot; en el menú de la página de detalles del script.":::

## <a name="remove-script-buttons"></a>Quitar botones de script

Para dejar de compartir un script mediante un botón, vaya al menú **Más opciones (...)** de la página de detalles del script y seleccione **Detener el uso compartido**. Esto quitará todos los botones que ejecutan el script. Al eliminar un solo botón, se quita el script de ese botón, incluso si la operación se deshace o si el botón se corta y pega.

## <a name="script-buttons-with-excel-on-windows"></a>Botones de script con Excel en Windows

Estos botones de script también funcionan en Windows. Cree el botón en Excel en la Web y los usuarios de Windows pueden ejecutar el script con el clic de un botón. Tenga en cuenta que no puede editar scripts en Excel en Windows. Solo puede editar scripts en Excel en la Web.

Es posible que algunas API de scripts de Office no sean compatibles con Excel en Windows, especialmente las compilaciones más antiguas. Estas incluyen API y API más recientes para las características de solo web. Si un script contiene API no admitidas, el script no se ejecuta y, en su lugar, el panel de tareas **Estado de ejecución** de script muestra un mensaje de advertencia que indica: "Este script debe ejecutarse actualmente en Excel para la Web. Abra el libro en el explorador e inténtelo de nuevo o póngase en contacto con el propietario del script para obtener ayuda."  

> [!IMPORTANT]
> Los botones de script requieren [que WebView2](/deployoffice/webview2-install) funcione con Excel en Windows. Esto se instala de forma predeterminada con las versiones más recientes de Excel en el escritorio, pero si no puede hacer clic en los botones de scripts, visite [Descargar el motor en tiempo de ejecución de WebView2](https://developer.microsoft.com/microsoft-edge/webview2/#download-section) y descargue el motor del explorador.
