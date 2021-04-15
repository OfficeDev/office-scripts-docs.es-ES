---
title: Propiedad y almacenamiento de archivos de Scripts de Office
description: Información sobre cómo los scripts de Office se almacenan en Microsoft OneDrive y se transfieren entre propietarios.
ms.date: 11/13/2020
localization_priority: Normal
ms.openlocfilehash: bd868c1dbfd0b33d3cd9fc4ee774c654d86f9b07
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 04/14/2021
ms.locfileid: "51755108"
---
# <a name="office-scripts-file-storage-and-ownership"></a>Propiedad y almacenamiento de archivos de Scripts de Office

Los scripts de Office se almacenan **como archivos .osts** en Microsoft OneDrive. Esto permite que los scripts existan fuera de cualquier libro en particular. La configuración de OneDrive controla el acceso compartido y los permisos de todos los archivos **.osts** de script; independiente de cualquier configuración de Excel.

## <a name="file-storage"></a>Almacenamiento de archivos.

Los scripts de Office se almacenan en OneDrive. Los **archivos .osts** se encuentran en la **carpeta /Documents/Office Scripts/.** Las modificaciones realizadas en estos archivos **.osts,** como cambiar el nombre o eliminar archivos, se reflejarán en el Editor de código y la Galería de scripts.

Los scripts que se comparten con uno de los libros permanecen en OneDrive del creador de scripts. No se copian en ninguna de las carpetas locales o de OneDrive cuando se ejecuta el script compartido en Excel. El **botón Hacer una copia** del Editor de código guarda una copia independiente del script en OneDrive. Los cambios realizados en la copia no afectan al script original.

### <a name="script-folders"></a>Carpetas de script

Agregar carpetas a OneDrive ayuda a mantener los scripts organizados. Las carpetas **de /Documents/Office Scripts/** se muestran en la sección **Mis scripts** del Editor de código. Tenga en cuenta que estas carpetas no se pueden crear ni eliminar mediante el Editor de código. Del mismo modo, los scripts no se pueden colocar en carpetas ni moverse entre carpetas mediante el Editor de código.

:::image type="content" source="../images/script-folders.png" alt-text="Cuadro de diálogo Nuevo script en el Editor de código que muestra los scripts contenidos en carpetas, tal como se muestra en el panel de tareas.":::

## <a name="file-ownership-and-retention"></a>Retención y propiedad de archivos

Los scripts de Office se almacenan en OneDrive de un usuario. Siguen las directivas de retención y eliminación especificadas por Microsoft OneDrive. Para obtener información sobre cómo administrar scripts creados y compartidos por un usuario que fue quitado de la organización, consulte [Retención y eliminación de OneDrive](/onedrive/retention-and-deletion).

## <a name="see-also"></a>Consulte también

- [Compartir Scripts de Office en Excel para la web](https://support.microsoft.com/office/sharing-office-scripts-in-excel-for-the-web-226eddbc-3a44-4540-acfe-fccda3d1122b)
- [Solución de problemas de scripts de Office](../testing/troubleshooting.md)
- [Configuración de scripts de Office en M365](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf)
- [Deshacer los efectos de un script de Office](../testing/undo.md)
