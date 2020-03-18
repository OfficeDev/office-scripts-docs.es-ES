---
title: Diferencias entre scripts de Office y complementos de Office
description: El comportamiento y las diferencias de API entre scripts de Office y complementos de Office.
ms.date: 12/12/2019
localization_priority: Normal
ms.openlocfilehash: 4626afb66b54c94a72f29b039c601435c089d64d
ms.sourcegitcommit: b075eed5a6f275274fbbf6d62633219eac416f26
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 03/10/2020
ms.locfileid: "42700397"
---
# <a name="differences-between-office-scripts-and-office-add-ins"></a>Diferencias entre scripts de Office y complementos de Office

Los complementos de Office y los scripts de Office tienen mucho en común. Ambos ofrecen el control automatizado de un libro de Excel `Excel` a través del espacio de nombres de la API de JavaScript de Office. Sin embargo, las secuencias de comandos de Office están más limitadas en su ámbito.

Los scripts de Office se ejecutan hasta el final con una pulsación de botón manual, mientras que los complementos de Office se basan en la interacción del usuario y se conservan mientras el libro está en uso. Si observa que su extensión de Excel debe superar las capacidades de la plataforma de scripting, visite la documentación de los complementos de [Office](/office/dev/add-ins) para obtener más información sobre los complementos de Office.

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

## <a name="see-also"></a>Vea también

- [Scripts de Office en Excel en la web](../overview/excel.md)
- [Solución de problemas de scripts de Office](../testing/troubleshooting.md)
- [Crear un complemento de panel de tareas de Excel](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)