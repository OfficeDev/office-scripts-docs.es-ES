---
title: Diferencias entre los scripts de Office y los complementos de Office
description: El comportamiento y las diferencias de API entre los scripts de Office y los complementos de Office.
ms.date: 02/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: bd483f928e3e153b8a08537f6b333c3ea8d724dd
ms.sourcegitcommit: 34c7740c9bff0e4c7426e01029f967724bfee566
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 05/13/2022
ms.locfileid: "65393624"
---
# <a name="differences-between-office-scripts-and-office-add-ins"></a>Diferencias entre los scripts de Office y los complementos de Office

Comprenda las diferencias entre Office Scripts y Office Complementos para saber cuándo usar cada uno de ellos. Office scripts están diseñados para que los realice rápidamente cualquier persona que quiera mejorar su flujo de trabajo. Office complementos se integran con la interfaz de usuario de Office para una experiencia más interactiva a través de botones de cinta de opciones y paneles de tareas. Office complementos también pueden expandir funciones integradas de Excel proporcionando funciones personalizadas.

:::image type="content" source="../images/office-programmability-diagram.png" alt-text="Diagrama de cuatro cuadrantes que muestra las áreas de enfoque para diferentes soluciones de extensibilidad de Office. Tanto los scripts de Office como los complementos web de Office se centran en la web y la colaboración, pero Office scripts atienden a los usuarios finales (mientras que los complementos web de Office se dirigen a desarrolladores profesionales).":::

Office scripts se ejecutan hasta la finalización con un botón manual o como un paso en [Power Automate](https://flow.microsoft.com/), mientras que Office complementos continúan ejecutándose en función de cómo se configuren. Por ejemplo, puede configurar un complemento de Office para que siga ejecutándose incluso cuando se cierre su panel de tareas. Esto significa que Office complementos mantienen el estado durante una sesión, mientras que Office scripts no mantienen un estado interno entre ejecuciones. Si la solución que está creando requiere un estado mantenido, debe visitar la [documentación de complementos de Office](/office/dev/add-ins) para obtener más información sobre los complementos de Office.

En el resto de este artículo se describen las principales diferencias entre los complementos de Office y los scripts de Office.

## <a name="platform-support"></a>Compatibilidad con la plataforma

Office complementos son multiplataforma. Funcionan en Windows plataformas de escritorio, Mac, iOS y web y proporcionan la misma experiencia en cada una de ellas. Cualquier excepción a esto se anota en la documentación de la API individual.

Actualmente, los scripts de Office solo son compatibles con Excel en la Web. Toda la administración de scripts, edición y grabación se realiza en la plataforma web.

### <a name="script-support-for-excel-on-windows"></a>Compatibilidad con scripts para Excel en Windows

[!INCLUDE [Run-from-button support](../includes/run-from-button-desktop-support.md)]

## <a name="apis"></a>API

Aunque las API de JavaScript de Office para complementos de Office y las API de scripts de Office comparten cierta funcionalidad, son plataformas diferentes. Las API de scripts de Office son un subconjunto sincrónico optimizado del modelo de API de JavaScript Excel. La principal diferencia es el uso del `load`/`sync` paradigma con complementos. Además, los complementos ofrecen API para eventos y un conjunto más amplio de funciones fuera de Excel, conocidas como API comunes.

### <a name="events"></a>Eventos

Office scripts no admiten [eventos](/office/dev/add-ins/excel/excel-add-ins-events) de nivel de libro. Los scripts los desencadenan los usuarios que seleccionan el botón **Ejecutar** para un script o a través de Power Automate. Cada script ejecuta el código en un solo `main` método y, a continuación, finaliza.

### <a name="common-apis"></a>API comunes

Office scripts no pueden usar [las API comunes](/javascript/api/office). Si necesita autenticación, ventanas de diálogo u otras características que solo son compatibles con las API comunes, es probable que tenga que crear un complemento de Office en lugar de un script de Office.

## <a name="see-also"></a>Consulte también

- [scripts de Office en Excel](../overview/excel.md)
- [Diferencias entre scripts de Office y macros de VBA](vba-differences.md)
- [Solución de problemas de scripts de Office](../testing/troubleshooting.md)
- [Crear un complemento de panel de tareas de Excel](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)
