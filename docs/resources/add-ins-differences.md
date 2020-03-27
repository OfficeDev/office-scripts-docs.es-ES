---
title: Diferencias entre scripts de Office y complementos de Office
description: El comportamiento y las diferencias de API entre scripts de Office y complementos de Office.
ms.date: 03/23/2020
localization_priority: Normal
ms.openlocfilehash: 2290d4e34b7a7286d67443de9e9c64bad4fcd4b7
ms.sourcegitcommit: d556aaefac80e55f53ac56b7f6ecbc657ebd426f
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 03/26/2020
ms.locfileid: "42978730"
---
# <a name="differences-between-office-scripts-and-office-add-ins"></a>Diferencias entre scripts de Office y complementos de Office

Los complementos de Office y los scripts de Office tienen mucho en común. Ambos ofrecen el control automatizado de un libro de Excel `Excel` a través del espacio de nombres de la API de JavaScript de Office. Sin embargo, las secuencias de comandos de Office están más limitadas en su ámbito.

![Un diagrama de cuatro fases que muestra las áreas de enfoque para diferentes soluciones de extensibilidad de Office. Los scripts de Office y los complementos Web de Office se centran en la web y la colaboración, pero los scripts de Office se adaptan a los usuarios finales (mientras que los complementos Web de Office tienen como objetivo desarrolladores profesionales)).](../images/office-programmability-diagram.png)

Los scripts de Office se ejecutan hasta el final con una pulsación de botón manual o como un paso de la [automatización de energía](https://flow.microsoft.com/), mientras que los complementos de Office se conservan mientras los paneles de tareas están abiertos. Esto significa que los complementos pueden mantener el estado durante una sesión, mientras que los scripts de Office no mantienen un estado interno entre ejecuciones. Si observa que su extensión de Excel debe superar las capacidades de la plataforma de scripting, visite la documentación de los complementos de [Office](/office/dev/add-ins) para obtener más información sobre los complementos de Office.

En el resto de este artículo se describen las principales diferencias entre los complementos de Office y los scripts de Office.

## <a name="platform-support"></a>Compatibilidad con plataformas

Los complementos de Office son para varias plataformas. Funcionan en plataformas de escritorio de Windows, Mac, iOS y Web y proporcionan la misma experiencia en cada uno de ellos. Cualquier excepción a esto se indica en la documentación de la API individual.

Los scripts de Office solo están actualmente admitidos por para Excel en la Web. Todas las operaciones de grabación, edición y ejecución se realizan en la plataforma Web.

## <a name="apis"></a>API

Los scripts de Office admiten la mayoría de las API de JavaScript de Excel, lo que significa que hay mucha funcionalidad superpuesta entre las dos plataformas. Hay dos excepciones: Events y Common API.

### <a name="events"></a>Eventos

Los scripts de Office no admiten [eventos](/office/dev/add-ins/excel/excel-add-ins-events). Cada secuencia de comandos ejecuta el código en `main` un solo método y, a continuación, finaliza. No se reactiva cuando se desencadenan eventos y, por lo tanto, no pueden registrar los eventos.

### <a name="common-apis"></a>API comunes

Los scripts de Office no pueden usar [API comunes](/javascript/api/office). Si necesita la autenticación, ventanas de cuadro de diálogo u otras características que solo se admiten en las API comunes, es probable que deba crear un complemento de Office en lugar de un script de Office.

## <a name="see-also"></a>Consulte también

- [Scripts de Office en Excel en la Web](../overview/excel.md)
- [Diferencias entre scripts de Office y macros de VBA](vba-differences.md)
- [Solución de problemas de scripts de Office](../testing/troubleshooting.md)
- [Crear un complemento de panel de tareas de Excel](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)
