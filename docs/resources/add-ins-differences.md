---
title: Diferencias entre los scripts de Office y los complementos de Office
description: El comportamiento y las diferencias de API entre los scripts de Office y los complementos de Office.
ms.date: 02/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: a3df4daf04f963598d2cb31f82dd2c1c9923fdc8
ms.sourcegitcommit: 33fe0f6807daefb16b148fd73c863de101f47cea
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 08/08/2022
ms.locfileid: "67281913"
---
# <a name="differences-between-office-scripts-and-office-add-ins"></a>Diferencias entre los scripts de Office y los complementos de Office

Comprenda las diferencias entre los scripts de Office y los complementos de Office para saber cuándo usar cada uno de ellos. Los scripts de Office están diseñados para ser hechos rápidamente por cualquier persona que quiera mejorar su flujo de trabajo. Los complementos de Office se integran con la interfaz de usuario de Office para una experiencia más interactiva a través de botones de cinta de opciones y paneles de tareas. Los complementos de Office también pueden expandir funciones integradas de Excel proporcionando funciones personalizadas.

:::image type="content" source="../images/office-programmability-diagram.png" alt-text="Diagrama de cuatro cuadrantes que muestra las áreas de enfoque de las diferentes soluciones de extensibilidad de Office. Tanto los scripts de Office como los complementos web de Office se centran en la web y la colaboración, pero los scripts de Office atienden a los usuarios finales (mientras que los complementos web de Office se dirigen a desarrolladores profesionales).":::

Los scripts de Office se ejecutan hasta la finalización con un botón manual o como un paso en [Power Automate](https://flow.microsoft.com/), mientras que los complementos de Office siguen ejecutándose en función de cómo se configuren. Por ejemplo, puede configurar un complemento de Office para que siga ejecutándose incluso cuando se cierre su panel de tareas. Esto significa que los complementos de Office mantienen el estado durante una sesión, mientras que los scripts de Office no mantienen un estado interno entre ejecuciones. Si la solución que está creando requiere un estado mantenido, debe visitar la [documentación de complementos de Office](/office/dev/add-ins) para obtener más información sobre los complementos de Office.

En el resto de este artículo se describen las principales diferencias entre los complementos de Office y los scripts de Office.

## <a name="platform-support"></a>Compatibilidad con la plataforma

Los complementos de Office son multiplataforma. Funcionan en plataformas de escritorio, Mac, iOS y web de Windows y proporcionan la misma experiencia en cada una de ellas. Cualquier excepción a esto se anota en la documentación de la API individual.

Actualmente, los scripts de Office solo son compatibles con Excel en la Web. Toda la administración de scripts, edición y grabación se realiza en la plataforma web.

### <a name="script-support-for-excel-on-windows"></a>Compatibilidad con scripts para Excel en Windows

[!INCLUDE [Run-from-button support](../includes/run-from-button-desktop-support.md)]

## <a name="apis"></a>API

Aunque las API de JavaScript de Office para complementos de Office y las API de scripts de Office comparten algunas funciones, son plataformas diferentes. Las API de scripts de Office son un subconjunto sincrónico optimizado del modelo de API de JavaScript de Excel. La principal diferencia es el uso del `load`/`sync` paradigma con complementos. Además, los complementos ofrecen API para eventos y un conjunto más amplio de funciones fuera de Excel, conocidas como API comunes.

### <a name="events"></a>Events

Los scripts de Office no admiten [eventos](/office/dev/add-ins/excel/excel-add-ins-events) de nivel de libro. Los scripts los desencadenan los usuarios que seleccionan el botón **Ejecutar** para un script o a través de Power Automate. Cada script ejecuta el código en una sola `main` función y, a continuación, finaliza.

### <a name="common-apis"></a>API comunes

Los scripts de Office no pueden usar [las API comunes](/javascript/api/office). Si necesita autenticación, ventanas de diálogo u otras características que solo son compatibles con las API comunes, es probable que tenga que crear un complemento de Office en lugar de un script de Office.

## <a name="see-also"></a>Vea también

- [Scripts de Office en Excel](../overview/excel.md)
- [Diferencias entre scripts de Office y macros de VBA](vba-differences.md)
- [Solución de problemas de scripts de Office](../testing/troubleshooting.md)
- [Crear un complemento de panel de tareas de Excel](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)
