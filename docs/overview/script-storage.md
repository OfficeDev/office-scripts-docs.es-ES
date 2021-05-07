---
title: Office Propiedad y almacenamiento de archivos de scripts
description: Información sobre cómo Office scripts se almacenan en Microsoft OneDrive y se transfieren entre propietarios.
ms.date: 11/13/2020
localization_priority: Normal
ms.openlocfilehash: 47b732399c3068bea78b027e01324bbd73a83bc7
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232532"
---
# <a name="office-scripts-file-storage-and-ownership"></a>Office Propiedad y almacenamiento de archivos de scripts

Office Los scripts se almacenan **como archivos .osts** en el Microsoft OneDrive. Esto permite que los scripts existan fuera de cualquier libro en particular. La OneDrive controle el acceso compartido y los permisos para todos los archivos **.osts** de script; independiente de cualquier Excel configuración.

## <a name="file-storage"></a>Almacenamiento de archivos.

Puede Office scripts se almacenan en su OneDrive. Los **archivos .osts** se encuentran en la **carpeta /Documents/Office Scripts/.** Las modificaciones realizadas en estos archivos **.osts,** como cambiar el nombre o eliminar archivos, se reflejarán en el Editor de código y la Galería de scripts.

Los scripts que se comparten con uno de los libros permanecen en el sitio del creador de scripts OneDrive. No se copian en ninguna de las carpetas locales o OneDrive cuando se ejecuta el script compartido en Excel. El **botón Hacer una copia** del Editor de código guarda una copia independiente del script en el OneDrive. Los cambios realizados en la copia no afectan al script original.

### <a name="script-folders"></a>Carpetas de script

Agregar carpetas a su OneDrive ayuda a mantener los scripts organizados. Las carpetas **en /Documents/Office Scripts/** se muestran en la sección **Mis scripts** del Editor de código. Tenga en cuenta que estas carpetas no se pueden crear ni eliminar mediante el Editor de código. Del mismo modo, los scripts no se pueden colocar en carpetas ni moverse entre carpetas mediante el Editor de código.

:::image type="content" source="../images/script-folders.png" alt-text="Cuadro de diálogo Nuevo script en el Editor de código que muestra scripts contenidos en carpetas, como se muestra en el panel de tareas":::

## <a name="file-ownership-and-retention"></a>Retención y propiedad de archivos

Office Los scripts se almacenan en el OneDrive. Siguen las directivas de retención y eliminación especificadas por Microsoft OneDrive. Para obtener información sobre cómo administrar scripts creados y compartidos por un usuario que fue quitado de la organización, consulte [Retención y eliminación de OneDrive](/onedrive/retention-and-deletion).

## <a name="see-also"></a>Vea también

- [Compartir Scripts de Office en Excel para la web](https://support.microsoft.com/office/sharing-office-scripts-in-excel-for-the-web-226eddbc-3a44-4540-acfe-fccda3d1122b)
- [Solución de problemas de scripts de Office](../testing/troubleshooting.md)
- [Configuración de scripts de Office en M365](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf)
- [Deshacer los efectos de un script de Office](../testing/undo.md)
