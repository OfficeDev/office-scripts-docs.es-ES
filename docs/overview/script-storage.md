---
title: almacenamiento y propiedad de archivos de scripts de Office
description: Información sobre cómo se almacenan Office scripts en Microsoft OneDrive y se transfieren entre propietarios.
ms.date: 05/10/2022
ms.localizationpriority: medium
ms.openlocfilehash: 5e2bc89db54ee5520c3b911ebd0f182777a78e2b
ms.sourcegitcommit: 8ae932e8b4e521fec8576ab16126eb9fe22a8dd7
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 05/11/2022
ms.locfileid: "65310760"
---
# <a name="office-scripts-file-storage-and-ownership"></a>almacenamiento y propiedad de archivos de scripts de Office

Office scripts se almacenan como archivos **.osts** en el Microsoft OneDrive. Se almacenan por separado de un libro. Para conceder acceso a otros usuarios, [comparta el script con un libro de Excel](excel.md#share-office-scripts). Esto significa que va a vincular el script con el archivo, no adjuntarlo. Quien tenga acceso al archivo Excel también podrá ver, ejecutar o hacer una copia del script.

A menos que comparta los scripts, nadie más puede acceder a ellos. La configuración de OneDrive controla el acceso compartido y los permisos de todos los archivos **.osts** de script, independientemente de la configuración de Excel. Los scripts no se pueden vincular desde un disco local ni desde ubicaciones de nube personalizadas. Office Scripts solo reconoce y ejecuta un script si está en la carpeta OneDrive o se comparte con el libro.

## <a name="file-storage"></a>Almacenamiento de archivos.

Los scripts Office se almacenan en el OneDrive. Los archivos **.osts** se encuentran en la carpeta **/Documents/Office Scripts/**. Las modificaciones realizadas en estos archivos **.osts** , como cambiar el nombre o eliminar archivos, se reflejarán en el Editor de código y la Galería de scripts.

Los scripts que se comparten con uno de los libros permanecen en el OneDrive del creador del script. No se copian en ninguna de las carpetas locales o OneDrive al ejecutar el script compartido en Excel. El botón **Crear una copia** del Editor de código guarda una copia independiente del script en el OneDrive. Los cambios en la copia no afectan al script original.

### <a name="restore-deleted-scripts"></a>Restauración de scripts eliminados

Al eliminar un script en Excel, se dirige a la papelera de reciclaje de OneDrive. Para restaurar un script eliminado, siga los pasos indicados en [Restauración de archivos o carpetas eliminados en OneDrive](https://support.microsoft.com/office/949ada80-0026-4db3-a953-c99083e6a84f). La restauración de un archivo **.osts** lo devuelve a la lista **Todos los scripts** .

Un script eliminado no se comparte con el libro. Al restaurar un script, **no** conserva su acceso al script. Tendrá que volver a compartir el script.

Los scripts restaurados siguen funcionando según lo esperado con los flujos de Power Automate. No es necesario volver a crear el conector de flujo.

## <a name="file-ownership-and-retention"></a>Propiedad y retención de archivos

Office scripts se almacenan en la OneDrive de un usuario. Siguen las directivas de retención y eliminación especificadas por Microsoft OneDrive. Para obtener información sobre cómo administrar scripts creados y compartidos por un usuario que fue quitado de la organización, consulte [Retención y eliminación de OneDrive](/onedrive/retention-and-deletion).

Durante la edición, los archivos se almacenan temporalmente en el explorador. Debe guardar el script antes de cerrar la ventana de Excel para guardarlo en la ubicación de OneDrive. No olvide guardar el archivo después de las modificaciones o, de lo contrario, esas modificaciones solo estarán en la versión del explorador del archivo.

## <a name="audit-office-scripts-usage-at-the-admin-level"></a>Auditar el uso de scripts de Office en el nivel de administrador

Descubra qué inquilinos usan scripts de Office con el registro de auditoría en el centro de cumplimiento. Para obtener información sobre cómo usar esta herramienta, visite [Buscar en el registro de auditoría en el Centro de cumplimiento de seguridad &](/microsoft-365/compliance/search-the-audit-log-in-security-and-compliance?view=o365-worldwide&preserve-view=true#search-the-audit-log).

Para buscar quién usa Office scripts con la herramienta de búsqueda, agregue `.osts` en el campo **Archivo, carpeta o sitio**. Esto busca todos los archivos con la extensión de archivo scripts de Office. Si alguien de su organización ha usado la característica scripts de Office, la actividad de usuario se muestra en los resultados de la búsqueda del registro de auditoría.

## <a name="see-also"></a>Vea también

- [Compartir Scripts de Office en Excel para la Web](https://support.microsoft.com/office/226eddbc-3a44-4540-acfe-fccda3d1122b)
- [Solución de problemas de scripts de Office](../testing/troubleshooting.md)
- [Configuración de scripts de Office en M365](/microsoft-365/admin/manage/manage-office-scripts-settings)
- [Deshacer los efectos de Scripts de Office](../testing/undo.md)
