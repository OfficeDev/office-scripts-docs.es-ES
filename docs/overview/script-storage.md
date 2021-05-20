---
title: Office Scripts almacenamiento de archivos y propiedad
description: Información sobre cómo se almacenan Office scripts en Microsoft OneDrive y se transfieren entre propietarios.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 556d784dc1fe64873866c49ab2726a4c68abc1a7
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545804"
---
# <a name="office-scripts-file-storage-and-ownership"></a>Office Scripts almacenamiento de archivos y propiedad

Office Los scripts se almacenan como archivos **.osts** en el Microsoft OneDrive. Se almacenan por separado de un libro de trabajo. Para dar acceso a otros, [comparta el script con un libro de trabajo Excel.](excel.md#sharing-scripts) Esto significa que está vinculando el script con el archivo, no adjuntándolo. Quien tenga acceso al archivo Excel también podrá ver, ejecutar o hacer una copia del script.

A menos que compartas tus scripts, nadie más puede acceder a ellos. La configuración de OneDrive controla el acceso compartido y los permisos para todos los archivos **.osts** de script, independientemente de cualquier configuración Excel. Los scripts no se pueden vincular desde un disco local ni desde ubicaciones de nube personalizadas. Office Scripts solo reconoce y ejecuta un script si está en la carpeta OneDrive o se comparte con el libro.

## <a name="file-storage"></a>Almacenamiento de archivos.

Office Scripts se almacenan en el OneDrive. Los archivos **.osts** se encuentran en la carpeta **/Documents/Office Scripts/.** Las ediciones realizadas en estos archivos **.osts,** como cambiar el nombre o eliminar archivos, se reflejarán en el Editor de código y la Galería de scripts.

Los scripts que se comparten con uno de sus libros permanecen en la OneDrive del creador de scripts. No se copian en ninguna de las carpetas locales o OneDrive al ejecutar el script compartido en Excel. El botón **Crear una copia** del Editor de código guarda una copia independiente del script en el OneDrive. Los cambios en la copia no afectan al script original.

## <a name="file-ownership-and-retention"></a>Propiedad y retención de archivos

Office Los scripts se almacenan en la OneDrive de un usuario. Siguen las directivas de retención y eliminación especificadas por Microsoft OneDrive. Para obtener información sobre cómo administrar scripts creados y compartidos por un usuario que fue quitado de la organización, consulte [Retención y eliminación de OneDrive](/onedrive/retention-and-deletion).

Durante la edición, los archivos se almacenan temporalmente en el navegador. Debe guardar el script antes de cerrar la ventana de Excel para guardarla en la ubicación OneDrive. No olvide guardar el archivo después de las ediciones, o de lo contrario esas ediciones sólo estarán en la versión del navegador del archivo.

## <a name="see-also"></a>Vea también

- [Compartir Scripts de Office en Excel para la web](https://support.microsoft.com/office/sharing-office-scripts-in-excel-for-the-web-226eddbc-3a44-4540-acfe-fccda3d1122b)
- [Solución de problemas de scripts de Office](../testing/troubleshooting.md)
- [Configuración de scripts de Office en M365](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf)
- [Deshacer los efectos de Scripts de Office](../testing/undo.md)
