---
title: Almacenamiento y propiedad de archivos de scripts de Office
description: Información sobre cómo se almacenan los scripts de Office en Microsoft OneDrive y cómo se transfieren entre propietarios.
ms.date: 11/13/2020
localization_priority: Normal
ms.openlocfilehash: 648f3b2cf7e7d8d3bab2cf07a090e116e267a99a
ms.sourcegitcommit: 82d3c0ef1e187bcdeceb2b5fc3411186674fe150
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 11/18/2020
ms.locfileid: "49346873"
---
# <a name="office-scripts-file-storage-and-ownership"></a>Almacenamiento y propiedad de archivos de scripts de Office

Los scripts de Office se almacenan como archivos **. osts** en su Microsoft OneDrive. Esto permite que los scripts existan fuera de un libro determinado. Su configuración de OneDrive controla el acceso compartido y los permisos para todos los archivos script **. osts** ; independiente de cualquier configuración de Excel.

## <a name="file-storage"></a>Almacenamiento de archivos.

Los scripts de Office se almacenan en su OneDrive. Los archivos **. osts** se encuentran en la carpeta **scripts//Documents/Office** . Los cambios realizados en estos archivos **. osts** , como cambiar el nombre o eliminar archivos, se reflejarán en el editor de código y en la galería de scripts.

Los scripts compartidos con uno de los libros permanecen en el OneDrive del creador de scripts. No se copian en ninguna de las carpetas locales o de OneDrive al ejecutar el script compartido en Excel. El botón **hacer una copia** del editor de código guarda una copia independiente del script en su OneDrive. Los cambios realizados en la copia no afectan al script original.

### <a name="script-folders"></a>Carpetas de script

Agregar carpetas a OneDrive ayuda a mantener los scripts organizados. Las carpetas en **/Documents/Office scripts/** se muestran en la sección **mis scripts** del editor de código. Tenga en cuenta que estas carpetas no se pueden crear ni eliminar con el editor de código. Del mismo modo, los scripts no se pueden colocar en carpetas ni se pueden mover entre carpetas mediante el editor de código.

![Algunos scripts contenidos en carpetas, tal como se muestra en el panel de tareas del editor de código](../images/script-folders.png)

## <a name="file-ownership-and-retention"></a>Posesión y retención de archivos

Los scripts de Office se almacenan en el OneDrive de un usuario. Siguen las directivas de retención y eliminación especificadas por Microsoft OneDrive. Para obtener información sobre cómo administrar scripts creados y compartidos por un usuario que fue quitado de la organización, consulte [Retención y eliminación de OneDrive](/onedrive/retention-and-deletion).

## <a name="see-also"></a>Consulte también

- [Compartir Scripts de Office en Excel para la web](https://support.microsoft.com/office/sharing-office-scripts-in-excel-for-the-web-226eddbc-3a44-4540-acfe-fccda3d1122b)
- [Solución de problemas de scripts de Office](../testing/troubleshooting.md)
- [Configuración de scripts de Office en M365](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf)
- [Deshacer los efectos de un script de Office](../testing/undo.md)
