---
title: Almacenamiento y propiedad de archivos de Scripts de Office
description: Información sobre cómo se almacenan los scripts de Office en Microsoft OneDrive y se transfieren entre propietarios.
ms.date: 08/31/2022
ms.localizationpriority: medium
ms.openlocfilehash: 573f65f299c29b4f481c9a2e23ebe7e36181706b
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572510"
---
# <a name="office-scripts-file-storage-and-ownership"></a>Almacenamiento y propiedad de archivos de Scripts de Office

Los scripts de Office se almacenan como archivos **.osts** en Microsoft OneDrive o en una carpeta de SharePoint. Se almacenan por separado de un libro. Para proporcionar a los usuarios que están fuera del sitio de SharePoint acceso al script, [comparta el script con un libro de Excel](excel.md#share-office-scripts). Esto significa que va a vincular el script con el archivo, no adjuntarlo. Quien tenga acceso al archivo de Excel también podrá ver, ejecutar o hacer una copia del script.

Excel solo reconoce y ejecuta un script si está en la carpeta de OneDrive, en una carpeta de SharePoint o en el libro.

## <a name="onedrive"></a>OneDrive

El comportamiento predeterminado es que los scripts de Office se almacenan en OneDrive. Los archivos **.osts** se encuentran en la carpeta **/Documents/Office Scripts/** . Las modificaciones realizadas en estos archivos **.osts** , como cambiar el nombre o eliminar archivos, se reflejarán en el Editor de código y la Galería de scripts.

Los scripts que se comparten con uno de los libros permanecen en OneDrive del creador del script. No se copian en ninguna de las carpetas locales o de OneDrive al ejecutar el script compartido en Excel. El botón **Crear una copia** del Editor de código guarda una copia independiente del script en OneDrive. Los cambios en la copia no afectan al script original.

A menos que comparta sus scripts personales, nadie más puede acceder a ellos. La configuración de OneDrive controla el acceso compartido y los permisos de todos los archivos **.osts** de script, independientemente de la configuración de Excel. Los scripts no se pueden vincular desde un disco local ni desde ubicaciones de nube personalizadas.

## <a name="sharepoint"></a>SharePoint

Los scripts de Office que se guardan en un sitio de SharePoint pertenecen a su equipo. Usted y los miembros de su organización con el acceso adecuado pueden ejecutar y editar scripts desde SharePoint. También verá que estos scripts aparecen en la galería de scripts de la pestaña **Automatizar** .

Para cargar un script desde SharePoint, vaya a **Todos los scripts** y seleccione **Ver más scripts** en la parte inferior de la lista. Esto abre un selector de archivos donde puede elegir archivos **.osts** desde cualquier sitio de SharePoint al que tenga acceso. Tenga en cuenta que los scripts de SharePoint que ya ha abierto se mostrarán en la lista de scripts recientes.

Para guardar un script en SharePoint, vaya al menú **Más opciones (...)** y seleccione **Guardar como**. Se abre un selector de archivos donde puede seleccionar carpetas en el sitio de SharePoint. Guardar en una nueva ubicación crea una copia del script en esa ubicación. La versión original sigue estando en su ubicación de OneDrive u otra ubicación de SharePoint.

> [!IMPORTANT]
> Los scripts con [llamadas externas](../develop/external-calls.md) no se pueden ejecutar desde SharePoint. Recibirá un error que indica que "Las llamadas de acceso a la red no se admiten en este momento para los scripts guardados en un sitio de SharePoint".

> [!IMPORTANT]
> Power Automate **no** admite scripts almacenados en SharePoint en este momento.

## <a name="restore-deleted-scripts"></a>Restauración de scripts eliminados

Al eliminar un script en Excel, se dirige a su papelera de reciclaje de OneDrive o SharePoint. Para restaurar un script eliminado, siga los pasos enumerados en [Recuperación de elementos que faltan, eliminados o dañados en SharePoint y OneDrive para el trabajo o la escuela](https://support.microsoft.com/office/how-to-recover-missing-deleted-or-corrupted-items-in-sharepoint-and-onedrive-for-work-or-school-3d748edf-c072-46c9-81a4-4989056ebc87). La restauración de un archivo **.osts** lo devuelve a la lista **Todos los scripts** .

Un script eliminado no se comparte con el libro. Al restaurar un script, **no** conserva su acceso al script. Tendrá que volver a compartir el script.

Los scripts restaurados siguen funcionando según lo esperado con los flujos de Power Automate. No es necesario volver a crear el conector de flujo.

## <a name="file-ownership-and-retention"></a>Propiedad y retención de archivos

Los scripts de Office siguen las directivas de retención y eliminación especificadas por Microsoft OneDrive y Microsoft SharePoint. Para obtener información sobre cómo controlar los scripts creados y compartidos por un usuario que se quita de su organización, consulte [Información sobre la retención para SharePoint y OneDrive](/microsoft-365/compliance/retention-policies-sharepoint?view=o365-worldwide&preserve-view=true).

Durante la edición, los archivos se almacenan temporalmente en el explorador. Debe guardar el script antes de cerrar la ventana de Excel para guardarlo en la ubicación de OneDrive. No olvide guardar el archivo después de las modificaciones o, de lo contrario, esas modificaciones solo estarán en la versión del explorador del archivo.

## <a name="audit-office-scripts-usage-at-the-admin-level"></a>Auditar el uso de scripts de Office en el nivel de administrador

Descubra quién usa scripts de Office en su organización con el registro de auditoría del Centro de cumplimiento. Los detalles sobre el registro de auditoría se encuentran en [Buscar el registro de auditoría en el Centro de cumplimiento de seguridad &](/microsoft-365/compliance/search-the-audit-log-in-security-and-compliance?view=o365-worldwide&preserve-view=true#search-the-audit-log).

Para auditar específicamente la actividad relacionada con los scripts de Office como administrador, siga estos pasos.

1. En una ventana del explorador InPrivate (o Incognito u otro modo de seguimiento limitado específico del explorador), abra e inicie sesión en el [Centro de cumplimiento](https://compliance.microsoft.com/).
1. Vaya a la página **Auditoría** .
1. *(Solo una vez)* En la pestaña **Buscar** , seleccione **Iniciar la grabación de la actividad de usuario y administrador**.

    > [!IMPORTANT]
    > Puede tardar una o dos horas después de activar la grabación antes de que se graben todas las actividades en todo el inquilino.

1. Establezca las opciones de búsqueda deseadas y presione **Buscar**. Filtre **el script Activities** to **Ran (Actividades a ejecución) en el libro** para ver cada vez que se ejecutó un script. También puede filtrar el campo **Archivo, carpeta o sitio** por `.osts`. Esto revela quién en su organización está creando o modificando scripts.

    :::image type="content" source="../images/audit-log-example.png" alt-text="Algunas filas de resultados de búsqueda de registros de auditoría, incluida la acción &quot;Ejecutar script en el libro&quot; y la carga y modificación de un archivo .osts.":::

## <a name="see-also"></a>Vea también

- [Compartir Scripts de Office en Excel para la Web](https://support.microsoft.com/office/226eddbc-3a44-4540-acfe-fccda3d1122b)
- [Solución de problemas de scripts de Office](../testing/troubleshooting.md)
- [Configuración de scripts de Office en M365](/microsoft-365/admin/manage/manage-office-scripts-settings)
- [Deshacer los efectos de Scripts de Office](../testing/undo.md)
