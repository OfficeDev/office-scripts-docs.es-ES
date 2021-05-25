---
title: Office Propiedad y almacenamiento de archivos de scripts
description: Información sobre cómo Office scripts se almacenan en Microsoft OneDrive y se transfieren entre propietarios.
ms.date: 05/21/2021
localization_priority: Normal
ms.openlocfilehash: 25683d2b6ac2e8ac47b465b24fa087af83175806
ms.sourcegitcommit: 09d8859d5269ada8f1d0e141f6b5a4f96d95a739
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 05/24/2021
ms.locfileid: "52631660"
---
# <a name="office-scripts-file-storage-and-ownership"></a>Office Propiedad y almacenamiento de archivos de scripts

Office Los scripts se almacenan **como archivos .osts** en el Microsoft OneDrive. Se almacenan por separado de un libro. Para dar acceso a otros usuarios, [comparta el script con un libro Excel .](excel.md#sharing-scripts) Esto significa que vinculas el script con el archivo, no lo adjuntas. Quien tenga acceso al archivo Excel también podrá ver, ejecutar o hacer una copia del script.

A menos que comparta los scripts, nadie más podrá acceder a ellos. La OneDrive controle el acceso compartido y los permisos de todos los archivos **.osts** de script, independientemente de Excel configuración. Los scripts no se pueden vincular desde un disco local ni desde ubicaciones de nube personalizadas. Office Los scripts solo reconocen y ejecutan un script si está en la carpeta OneDrive o si se comparte con el libro.

## <a name="file-storage"></a>Almacenamiento de archivos.

Puede Office scripts se almacenan en su OneDrive. Los **archivos .osts** se encuentran en la **carpeta /Documents/Office Scripts/.** Las modificaciones realizadas en estos archivos **.osts,** como cambiar el nombre o eliminar archivos, se reflejarán en el Editor de código y la Galería de scripts.

Los scripts que se comparten con uno de los libros permanecen en el sitio del creador de scripts OneDrive. No se copian en ninguna de las carpetas locales o OneDrive cuando se ejecuta el script compartido en Excel. El **botón Hacer una copia** del Editor de código guarda una copia independiente del script en el OneDrive. Los cambios realizados en la copia no afectan al script original.

## <a name="file-ownership-and-retention"></a>Retención y propiedad de archivos

Office Los scripts se almacenan en el OneDrive. Siguen las directivas de retención y eliminación especificadas por Microsoft OneDrive. Para obtener información sobre cómo administrar scripts creados y compartidos por un usuario que fue quitado de la organización, consulte [Retención y eliminación de OneDrive](/onedrive/retention-and-deletion).

Durante la edición, los archivos se almacenan temporalmente en el explorador. Debe guardar el script antes de cerrar la Excel para guardarlo en la OneDrive ubicación. No olvide guardar el archivo después de las ediciones, o de lo contrario, dichas ediciones solo estarán en la versión del explorador del archivo.

## <a name="audit-office-scripts-usage-at-the-admin-level"></a>Auditar Office de scripts en el nivel de administración

Descubra qué inquilinos usan Office scripts con el registro de auditoría en el centro de cumplimiento. Para obtener información sobre cómo usar esta herramienta, visite Buscar en el registro de auditoría en el Centro de [seguridad & cumplimiento](/microsoft-365/compliance/search-the-audit-log-in-security-and-compliance?view=o365-worldwide#search-the-audit-log).

Para buscar quién usa Office scripts con la herramienta de búsqueda, agregue `.osts` el **campo Archivo, carpeta o sitio.** Esto busca todos los archivos con la Office de archivos scripts. Si alguien de la organización ha usado la característica Office scripts, la actividad del usuario aparece en los resultados de la búsqueda del registro de auditoría.

> [!NOTE]
> Actualmente no se registra la ejecución de un script. Solo se registran las acciones crear, ver y modificar.

## <a name="see-also"></a>Vea también

- [Compartir Scripts de Office en Excel para la web](https://support.microsoft.com/office/sharing-office-scripts-in-excel-for-the-web-226eddbc-3a44-4540-acfe-fccda3d1122b)
- [Solución de problemas de scripts de Office](../testing/troubleshooting.md)
- [Configuración de scripts de Office en M365](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf)
- [Deshacer los efectos de Scripts de Office](../testing/undo.md)
