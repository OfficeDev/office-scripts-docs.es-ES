---
title: Diferencias entre los scripts de Office y los complementos de Office
description: El comportamiento y las diferencias de API entre Office scripts y Office complementos.
ms.date: 06/01/2020
localization_priority: Normal
ms.openlocfilehash: 45993d08d85cfceb299216dddbe2e7da9fd2e404
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232637"
---
# <a name="differences-between-office-scripts-and-office-add-ins"></a>Diferencias entre los scripts de Office y los complementos de Office

Office Los complementos y Office scripts tienen mucho en común. Ambos ofrecen control automatizado de un libro Excel una API de JavaScript. Sin embargo, las API Office Scripts son una versión sincrónica y especializada de la API Office JavaScript.

:::image type="content" source="../images/office-programmability-diagram.png" alt-text="Diagrama de cuatro cuadrantes que muestra las áreas de enfoque para diferentes Office extensibilidad. Tanto Office scripts como Office complementos web se centran en la web y la colaboración, pero los scripts de Office se adaptan a los usuarios finales (mientras que Office complementos web están dirigidos a desarrolladores profesionales)":::

Office Los scripts se ejecutan hasta completarse con un botón manual o como un paso en [Power Automate](https://flow.microsoft.com/), mientras que Office complementos persisten mientras sus paneles de tareas están abiertos. Esto significa que los complementos pueden mantener el estado durante una sesión, mientras que Office scripts no mantienen un estado interno entre ejecuciones. Si encuentra que la extensión Excel debe superar las capacidades de la plataforma de scripting, visite la documentación de complementos de [Office](/office/dev/add-ins) para obtener más información sobre Office complementos.

El resto de este artículo describe las principales diferencias entre Office complementos y Office scripts.

## <a name="platform-support"></a>Compatibilidad con plataformas

Office Los complementos son multiplataforma. Funcionan en Windows escritorio, Mac, iOS y plataformas web y proporcionan la misma experiencia en cada una de ellas. Cualquier excepción a esto se indica en la documentación de la API individual.

Office Actualmente, los scripts solo son compatibles con Excel en la Web. Toda la grabación, edición y ejecución se realiza en la plataforma web.

## <a name="apis"></a>API

No hay ninguna versión sincrónica de las API Office JavaScript para Office complementos. Las API Office scripts estándar son únicas en la plataforma y tienen numerosas optimizaciones y alteraciones para evitar el uso del `load` / `sync` paradigma.

Algunas de las [API Excel JavaScript](/javascript/api/excel?view=excel-js-preview&preserve-view=true) son compatibles con las API Office [Scripts Async](../develop/excel-async-model.md). Algunos ejemplos y bloques de código de complemento se pueden porte a `Excel.run` bloques con traducción mínima. Aunque las dos plataformas comparten funcionalidad, hay diferencias. Los dos conjuntos de API principales que Office complementos tienen pero Office scripts no son eventos y las API comunes.

### <a name="events"></a>Eventos

Office Los scripts no admiten [eventos](/office/dev/add-ins/excel/excel-add-ins-events). Cada script ejecuta el código en un único método y, a `main` continuación, finaliza. No se reactiva cuando se desencadenan los eventos y, por lo tanto, no se pueden registrar eventos.

### <a name="common-apis"></a>API comunes

Office Los scripts no pueden [usar API comunes](/javascript/api/office). Si necesitas autenticación, ventanas de diálogo u otras características que solo sean compatibles con las API comunes, es probable que necesites crear un complemento de Office en lugar de un script Office.

## <a name="see-also"></a>Vea también

- [Scripts de Office en Excel en la Web](../overview/excel.md)
- [Diferencias entre Office scripts y macros de VBA](vba-differences.md)
- [Solución de problemas de scripts de Office](../testing/troubleshooting.md)
- [Crear un complemento de panel de tareas de Excel](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)
