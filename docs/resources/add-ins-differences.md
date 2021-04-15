---
title: Diferencias entre los scripts de Office y los complementos de Office
description: El comportamiento y las diferencias de API entre scripts de Office y complementos de Office.
ms.date: 06/01/2020
localization_priority: Normal
ms.openlocfilehash: 96af98ca9f247406c5cc916f38892c318d33c560
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 04/14/2021
ms.locfileid: "51755101"
---
# <a name="differences-between-office-scripts-and-office-add-ins"></a>Diferencias entre los scripts de Office y los complementos de Office

Los complementos de Office y los scripts de Office tienen mucho en común. Ambos ofrecen control automatizado de un libro de Excel una API de JavaScript. Sin embargo, las API de Scripts de Office son una versión sincrónica y especializada de la API de JavaScript de Office.

:::image type="content" source="../images/office-programmability-diagram.png" alt-text="Diagrama de cuatro cuadrantes que muestra las áreas de enfoque para diferentes soluciones de extensibilidad de Office. Tanto los scripts de Office como los complementos web de Office se centran en la web y la colaboración, pero los scripts de Office se adaptan a los usuarios finales (mientras que los complementos web de Office están dirigidos a desarrolladores profesionales).":::

Los scripts de Office se ejecutan hasta completarse con un botón manual o como un paso en [Power Automate,](https://flow.microsoft.com/)mientras que los complementos de Office persisten mientras sus paneles de tareas están abiertos. Esto significa que los complementos pueden mantener el estado durante una sesión, mientras que los scripts de Office no mantienen un estado interno entre ejecuciones. Si encuentra que la extensión de Excel debe superar las capacidades de la plataforma de scripting, visite la documentación de complementos de [Office](/office/dev/add-ins) para obtener más información sobre los complementos de Office.

El resto de este artículo describe las principales diferencias entre los complementos de Office y los scripts de Office.

## <a name="platform-support"></a>Compatibilidad con plataformas

Los complementos de Office son multiplataforma. Funcionan en plataformas de escritorio, Mac, iOS y web de Windows y proporcionan la misma experiencia en cada una de ellas. Cualquier excepción a esto se indica en la documentación de la API individual.

Actualmente, los scripts de Office solo son compatibles con Excel en la web. Toda la grabación, edición y ejecución se realiza en la plataforma web.

## <a name="apis"></a>API

No hay ninguna versión sincrónica de las API de JavaScript de Office para complementos de Office. Las API estándar de Scripts de Office son únicas para la plataforma y tienen numerosas optimizaciones y alteraciones para evitar el uso del `load` / `sync` paradigma.

Algunas de las [API de JavaScript de Excel](/javascript/api/excel?view=excel-js-preview&preserve-view=true) son compatibles con las API asincrónicas de Scripts de [Office.](../develop/excel-async-model.md) Algunos ejemplos y bloques de código de complemento se pueden porte a `Excel.run` bloques con traducción mínima. Aunque las dos plataformas comparten funcionalidad, hay diferencias. Los dos conjuntos de API principales que tienen los complementos de Office, pero los scripts de Office no son eventos y las API comunes.

### <a name="events"></a>Eventos

Los scripts de Office no [admiten eventos](/office/dev/add-ins/excel/excel-add-ins-events). Cada script ejecuta el código en un único método y, a `main` continuación, finaliza. No se reactiva cuando se desencadenan los eventos y, por lo tanto, no se pueden registrar eventos.

### <a name="common-apis"></a>API comunes

Los scripts de Office no [pueden usar API comunes.](/javascript/api/office) Si necesita autenticación, ventanas de diálogo u otras características que solo son compatibles con las API comunes, es probable que necesite crear un complemento de Office en lugar de un script de Office.

## <a name="see-also"></a>Consulte también

- [Scripts de Office en Excel en la Web](../overview/excel.md)
- [Diferencias entre scripts de Office y macros de VBA](vba-differences.md)
- [Solución de problemas de scripts de Office](../testing/troubleshooting.md)
- [Crear un complemento de panel de tareas de Excel](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)
