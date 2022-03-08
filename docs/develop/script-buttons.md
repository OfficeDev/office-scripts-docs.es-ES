---
title: Ejecutar Office scripts en Excel con botones
description: Agregue botones a los libros que controlan Office scripts en Excel.
ms.topic: overview
ms.date: 02/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: 0d88a6bcd928e6b4931b2374313cc17f4161ebf7
ms.sourcegitcommit: 49f527a7f54aba00e843ad4a92385af59c1d7bfa
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 03/08/2022
ms.locfileid: "63352202"
---
# <a name="run-office-scripts-in-excel-with-buttons"></a>Ejecutar Office scripts en Excel con botones

Ayude a sus compañeros a buscar y ejecutar los scripts agregando botones de script a un libro.

:::image type="content" source="../images/run-from-button.png" alt-text="Un botón de la hoja de cálculo que ejecuta un script al hacer clic en él.":::

## <a name="create-script-buttons"></a>Crear botones de script

Con cualquier script, vaya al menú **Más opciones (...)** en la página de detalles del script o en el panel de tareas del Editor de código y seleccione **Agregar botón**. Esto creará un botón en el libro que ejecutará el script asociado cuando se seleccione. También compartirá el script con el libro, por lo que todos los usuarios con permisos de escritura en el libro pueden usar esta automatización útil.

La siguiente captura de pantalla muestra la página Detalles del script de un script titulado **Crear** tabla dinámica y  tiene resaltada la opción Agregar botón dentro del menú Más opciones **(...)** .

:::image type="content" source="../images/add-button.png" alt-text="La opción &quot;Agregar botón&quot; en el menú de la página Detalles del script.":::

## <a name="remove-script-buttons"></a>Quitar botones de script

Para dejar de compartir un script a través de un botón, vaya al menú **Más opciones (...)** en la página Detalles del script y seleccione **Detener uso compartido**. Esto quitará todos los botones que ejecutan el script. Al eliminar un solo botón, se quita el script de ese botón, incluso si la operación se deshace o si el botón se corta y pega.

## <a name="script-buttons-on-excel-for-windows"></a>Botones de script Excel para Windows

Estos botones de script también funcionan en Windows. Cree el botón en Excel en la Web y los usuarios de Windows pueden ejecutar el script con el clic de un botón. Tenga en cuenta que no puede editar scripts en Excel en Windows. Solo puede editar scripts en Excel en la Web.

> [!NOTE]
> Esta característica se está implementando para usuarios con una suscripción de Microsoft 365, pero no está disponible para todos los usuarios. Poco a poco vamos lanzando esta característica a un número mayor de usuarios para garantizar que funciona como se esperaba. Esta característica está sujeta a cambios en función de sus comentarios. Las plataformas no admitidas o las versiones de Office sin esta característica, mostrarán la forma del botón de script, pero no se podrá hacer clic en el mismo.
