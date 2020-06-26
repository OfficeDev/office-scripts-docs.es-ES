---
title: Diferencias entre los scripts de Office y los complementos de Office
description: El comportamiento y las diferencias de API entre scripts de Office y complementos de Office.
ms.date: 06/01/2020
localization_priority: Normal
ms.openlocfilehash: fc2029780190672c633e00e26f44273e4311c754
ms.sourcegitcommit: aec3c971c6640429f89b6bb99d2c95ea06725599
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 06/25/2020
ms.locfileid: "44878664"
---
# <a name="differences-between-office-scripts-and-office-add-ins"></a>Diferencias entre los scripts de Office y los complementos de Office

Los complementos de Office y los scripts de Office tienen mucho en común. Ambos ofrecen un control automatizado de un libro de Excel con una API de JavaScript. Sin embargo, las API de scripts de Office son una versión de la API de JavaScript de Office especializada.

![Un diagrama de cuatro fases que muestra las áreas de enfoque para diferentes soluciones de extensibilidad de Office. Los scripts de Office y los complementos Web de Office se centran en la web y la colaboración, pero los scripts de Office se adaptan a los usuarios finales (mientras que los complementos Web de Office tienen como objetivo desarrolladores profesionales)).](../images/office-programmability-diagram.png)

Los scripts de Office se ejecutan hasta el final con una pulsación de botón manual o como un paso de la [automatización de energía](https://flow.microsoft.com/), mientras que los complementos de Office se conservan mientras los paneles de tareas están abiertos. Esto significa que los complementos pueden mantener el estado durante una sesión, mientras que los scripts de Office no mantienen un estado interno entre ejecuciones. Si observa que su extensión de Excel debe superar las capacidades de la plataforma de scripting, visite la documentación de los complementos de [Office](/office/dev/add-ins) para obtener más información sobre los complementos de Office.

En el resto de este artículo se describen las principales diferencias entre los complementos de Office y los scripts de Office.

## <a name="platform-support"></a>Compatibilidad con plataformas

Los complementos de Office son para varias plataformas. Funcionan en plataformas de escritorio de Windows, Mac, iOS y Web y proporcionan la misma experiencia en cada uno de ellos. Cualquier excepción a esto se indica en la documentación de la API individual.

Los scripts de Office solo están actualmente admitidos por para Excel en la Web. Todas las operaciones de grabación, edición y ejecución se realizan en la plataforma Web.

## <a name="apis"></a>API

No hay ninguna versión sincrónica de las API de JavaScript para Office para complementos de Office. Las API de scripts estándar de Office son únicas para la plataforma y tienen numerosas optimizaciones y alteraciones para evitar el uso del `load` / `sync` paradigma.

Algunas de las [API de JavaScript de Excel](/javascript/api/excel?view=excel-js-preview) son compatibles con las [API asincrónicas de scripts de Office](../develop/excel-async-model.md). Algunos ejemplos y bloques de código de complementos se pueden trasladar a `Excel.run` bloques con una traducción mínima. Mientras que las dos plataformas comparten la funcionalidad, hay brechas. Los dos conjuntos de API principales que tienen los complementos de Office, pero los scripts de Office no son eventos y las API comunes.

### <a name="events"></a>Eventos

Los scripts de Office no admiten [eventos](/office/dev/add-ins/excel/excel-add-ins-events). Cada secuencia de comandos ejecuta el código en un solo `main` método y, a continuación, finaliza. No se reactiva cuando se desencadenan eventos y, por lo tanto, no pueden registrar los eventos.

### <a name="common-apis"></a>API comunes

Los scripts de Office no pueden usar [API comunes](/javascript/api/office). Si necesita la autenticación, ventanas de cuadro de diálogo u otras características que solo se admiten en las API comunes, es probable que deba crear un complemento de Office en lugar de un script de Office.

## <a name="see-also"></a>Vea también

- [Scripts de Office en Excel en la Web](../overview/excel.md)
- [Diferencias entre scripts de Office y macros de VBA](vba-differences.md)
- [Solución de problemas de scripts de Office](../testing/troubleshooting.md)
- [Crear un complemento de panel de tareas de Excel](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)
