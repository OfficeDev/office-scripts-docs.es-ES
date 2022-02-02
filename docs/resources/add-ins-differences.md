---
title: Diferencias entre los scripts de Office y los complementos de Office
description: El comportamiento y las diferencias de API entre Office scripts y Office complementos.
ms.date: 01/24/2022
ms.localizationpriority: medium
ms.openlocfilehash: f4422203911aeb1b2667856991bc7a006070ee97
ms.sourcegitcommit: 9e7111b183c7117e05f38b1b13050b5397476d74
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 02/02/2022
ms.locfileid: "62319170"
---
# <a name="differences-between-office-scripts-and-office-add-ins"></a>Diferencias entre los scripts de Office y los complementos de Office

Comprenda las diferencias entre Office scripts y Office complementos para saber cuándo usar cada uno. Office scripts están diseñados para ser creados rápidamente por cualquier persona que quiera mejorar su flujo de trabajo. Office complementos se integran con la interfaz de usuario Office para una experiencia más interactiva a través de botones de cinta y paneles de tareas. Office complementos también pueden expandir las funciones integradas Excel funciones personalizadas.

:::image type="content" source="../images/office-programmability-diagram.png" alt-text="Diagrama de cuatro cuadrantes que muestra las áreas de enfoque para diferentes Office extensibilidad. Tanto los scripts Office como los Office web de Office se centran en la web y la colaboración, pero los scripts de Office se adaptan a los usuarios finales (mientras que Office complementos web están dirigidos a desarrolladores profesionales).":::

Office Scripts se ejecutan hasta su finalización con un botón manual o como un paso en [Power Automate](https://flow.microsoft.com/), mientras que Office complementos siguen ejecutándose en función de cómo estén configurados. Por ejemplo, puede configurar un complemento Office para continuar ejecutándose incluso cuando se cierre su panel de tareas. Esto significa que Office complementos mantienen el estado durante una sesión, mientras que los scripts Office no mantienen un estado interno entre ejecuciones. Si la solución que está creando requiere un estado de mantenimiento, debe visitar la documentación de Office [Complementos](/office/dev/add-ins) para obtener más información sobre Office complementos.

El resto de este artículo describe las principales diferencias entre Office complementos y Office scripts.

## <a name="platform-support"></a>Compatibilidad con plataformas

Office complementos son multiplataforma. Funcionan en Windows escritorio, Mac, iOS y plataformas web y proporcionan la misma experiencia en cada una de ellas. Cualquier excepción a esto se indica en la documentación de la API individual.

Office scripts solo son compatibles actualmente con Excel en la Web. Toda la administración de grabaciones, edición y scripts se realiza en la plataforma web.

### <a name="script-support-for-excel-on-windows-preview"></a>Compatibilidad con scripts para Excel en Windows (versión preliminar)

[!INCLUDE [Run-from-button support](../includes/run-from-button-desktop-support.md)]

## <a name="apis"></a>API

Aunque las API Office JavaScript para Office Complementos y las API de scripts de Office comparten cierta funcionalidad, son plataformas diferentes. Las API Office scripts son un subconjunto optimizado y sincrónico del Excel api de JavaScript. La diferencia principal es el uso del `load`/`sync` paradigma con complementos. Además, los complementos ofrecen API para eventos y un conjunto más amplio de funciones fuera de Excel, conocidas como API comunes.

### <a name="events"></a>Eventos

Office scripts no admiten eventos de nivel de [libro](/office/dev/add-ins/excel/excel-add-ins-events). Los scripts los desencadenan los usuarios que seleccionan el **botón Ejecutar** para un script o a través de Power Automate. Cada script ejecuta el código en un único `main` método y, a continuación, finaliza.

### <a name="common-apis"></a>API comunes

Office scripts no pueden usar [API comunes](/javascript/api/office). Si necesitas autenticación, ventanas de diálogo u otras características que solo sean compatibles con las API comunes, es probable que necesites crear un complemento de Office en lugar de un script Office.

## <a name="see-also"></a>Recursos adicionales

- [Scripts de Office en Excel en la Web](../overview/excel.md)
- [Diferencias entre Office scripts y macros de VBA](vba-differences.md)
- [Solución de problemas de scripts de Office](../testing/troubleshooting.md)
- [Crear un complemento de panel de tareas de Excel](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)
