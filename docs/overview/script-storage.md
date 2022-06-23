---
title: almacenamiento y propiedad de archivos de scripts de Office
description: Información sobre cómo se almacenan Office scripts en Microsoft OneDrive y se transfieren entre propietarios.
ms.date: 06/21/2022
ms.localizationpriority: medium
ms.openlocfilehash: 9dbf53292cb16b0be32afe3cdb93409f3dbb2612
ms.sourcegitcommit: 4f2164ac4dd61d123ea5442a4c446be2d139e8ff
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 06/23/2022
ms.locfileid: "66211303"
---
# <a name="office-scripts-file-storage-and-ownership"></a>almacenamiento y propiedad de archivos de scripts de Office

> [!IMPORTANT]
> SharePoint compatibilidad con scripts de Office se está implementando y no está disponible para todos. Se distribuye lentamente a un mayor número de usuarios para asegurarse de que funciona según lo previsto. Esta característica está sujeta a cambios en función de sus comentarios.

Office scripts se almacenan como archivos **.osts** en la Microsoft OneDrive o una carpeta SharePoint. Se almacenan por separado de un libro. Para proporcionar a los usuarios que están fuera de la SharePoint sitio acceso al script, [comparta el script con un libro de Excel](excel.md#share-office-scripts). Esto significa que va a vincular el script con el archivo, no adjuntarlo. Quien tenga acceso al archivo Excel también podrá ver, ejecutar o hacer una copia del script.

Excel solo reconoce y ejecuta un script si está en la carpeta OneDrive, una carpeta de SharePoint o compartida con el libro.

## <a name="onedrive"></a>OneDrive

El comportamiento predeterminado es que Office scripts se almacenan en el OneDrive. Los archivos **.osts** se encuentran en la carpeta **/Documents/Office Scripts/**. Las modificaciones realizadas en estos archivos **.osts** , como cambiar el nombre o eliminar archivos, se reflejarán en el Editor de código y la Galería de scripts.

Los scripts que se comparten con uno de los libros permanecen en el OneDrive del creador del script. No se copian en ninguna de las carpetas locales o OneDrive al ejecutar el script compartido en Excel. El botón **Crear una copia** del Editor de código guarda una copia independiente del script en el OneDrive. Los cambios en la copia no afectan al script original.

A menos que comparta sus scripts personales, nadie más puede acceder a ellos. La configuración de OneDrive controla el acceso compartido y los permisos de todos los archivos **.osts** de script, independientemente de la configuración de Excel. Los scripts no se pueden vincular desde un disco local ni desde ubicaciones de nube personalizadas.

## <a name="sharepoint"></a>SharePoint

Office los scripts que se guardan en un sitio SharePoint son propiedad del equipo. Usted y los miembros de la organización con el acceso adecuado pueden ejecutar y editar scripts desde SharePoint. También verá que estos scripts aparecen en la galería de scripts de la pestaña **Automatizar** .

Para cargar un script desde SharePoint, vaya a **Todos los scripts** y seleccione **Ver más scripts** en la parte inferior de la lista. Esto abre un selector de archivos en el que puede elegir archivos **.osts** de cualquier sitio SharePoint al que tenga acceso. Tenga en cuenta que los scripts de SharePoint que ya ha abierto se mostrarán en la lista de scripts recientes.

Para guardar un script en SharePoint, vaya al menú **Más opciones (...)** y seleccione **Guardar como**. Se abrirá un selector de archivos en el que puede seleccionar carpetas en el sitio de SharePoint. Guardar en una nueva ubicación crea una copia del script en esa ubicación. La versión original sigue estando en la OneDrive u otra ubicación SharePoint.

> [!IMPORTANT]
> Los scripts con [llamadas externas](../develop/external-calls.md) no se pueden ejecutar desde SharePoint. Recibirá un error que indica que "Las llamadas de acceso a la red no se admiten en este momento para los scripts guardados en un sitio de SharePoint".

> [!IMPORTANT]
> Power Automate **no** admite scripts almacenados en SharePoint en este momento.

## <a name="restore-deleted-scripts"></a>Restauración de scripts eliminados

Al eliminar un script en Excel, se dirige a la papelera de reciclaje de OneDrive o SharePoint. Para restaurar un script eliminado, siga los pasos indicados en [Recuperación de elementos que faltan, eliminados o dañados en SharePoint y OneDrive para el trabajo o la escuela](https://support.microsoft.com/office/how-to-recover-missing-deleted-or-corrupted-items-in-sharepoint-and-onedrive-for-work-or-school-3d748edf-c072-46c9-81a4-4989056ebc87). La restauración de un archivo **.osts** lo devuelve a la lista **Todos los scripts** .

Un script eliminado no se comparte con el libro. Al restaurar un script, **no** conserva su acceso al script. Tendrá que volver a compartir el script.

Los scripts restaurados siguen funcionando según lo esperado con los flujos de Power Automate. No es necesario volver a crear el conector de flujo.

## <a name="file-ownership-and-retention"></a>Propiedad y retención de archivos

Office scripts siguen las directivas de retención y eliminación especificadas por Microsoft OneDrive y Microsoft SharePoint. Para obtener información sobre cómo controlar los scripts creados y compartidos por un usuario que se quita de su organización, consulte [Información sobre la retención de SharePoint y OneDrive](/microsoft-365/compliance/retention-policies-sharepoint?view=o365-worldwide&preserve-view=true).

Durante la edición, los archivos se almacenan temporalmente en el explorador. Debe guardar el script antes de cerrar la ventana de Excel para guardarlo en la ubicación de OneDrive. No olvide guardar el archivo después de las modificaciones o, de lo contrario, esas modificaciones solo estarán en la versión del explorador del archivo.

## <a name="audit-office-scripts-usage-at-the-admin-level"></a>Auditar el uso de scripts de Office en el nivel de administrador

Descubra quién usa scripts de Office en su organización con el registro de auditoría del Centro de cumplimiento. Los detalles sobre el registro de auditoría se encuentran en [Buscar el registro de auditoría en el Centro de cumplimiento de seguridad &](/microsoft-365/compliance/search-the-audit-log-in-security-and-compliance?view=o365-worldwide&preserve-view=true#search-the-audit-log).

Para auditar específicamente Office actividad relacionada con scripts como administrador, siga estos pasos.

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
