---
title: Diferencias entre scripts de Office y macros de VBA
description: El comportamiento y las diferencias de API entre scripts de Office y macros de VBA de Excel.
ms.date: 12/14/2020
localization_priority: Normal
ms.openlocfilehash: a56409a5de3eb07876faa88bfbfe78eeca59f70f
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 04/14/2021
ms.locfileid: "51755024"
---
# <a name="differences-between-office-scripts-and-vba-macros"></a>Diferencias entre scripts de Office y macros de VBA

Los scripts de Office y las macros de VBA tienen mucho en común. Ambos permiten a los usuarios automatizar soluciones a través de una grabadora de acciones fácil de usar y permitir ediciones de esas grabaciones. Ambos marcos están diseñados para habilitar a las personas que pueden no considerarse programadores para crear pequeños programas en Excel.
La diferencia fundamental es que las macros de VBA se desarrollan para soluciones de escritorio y los scripts de Office están diseñados con compatibilidad y seguridad multiplataforma como principios de guía. Actualmente, los scripts de Office solo se admiten en Excel en la web.

:::image type="content" source="../images/office-programmability-diagram.png" alt-text="Diagrama de cuatro cuadrantes que muestra las áreas de enfoque para diferentes soluciones de extensibilidad de Office. Tanto los scripts de Office como las macros de VBA están diseñados para ayudar a los usuarios finales a crear soluciones, pero los scripts de Office se crean para la web y la colaboración (mientras que VBA es para el escritorio).":::

En este artículo se describen las principales diferencias entre macros de VBA (así como VBA en general) y scripts de Office. Dado que los scripts de Office solo están disponibles para Excel, este es el único host que se trata aquí.

## <a name="platform-and-ecosystem"></a>Plataforma y ecosistema

VBA está diseñado para el escritorio y los scripts de Office están diseñados para la web. VBA puede interactuar con el escritorio de un usuario para conectarse con tecnologías similares, como COM y OLE. Sin embargo, VBA no tiene una forma cómoda de llamar a Internet.

Los scripts de Office usan un tiempo de ejecución universal para JavaScript. Esto proporciona un comportamiento y accesibilidad coherentes, independientemente de la máquina que se esté utilizando para ejecutar el script. También pueden realizar llamadas a otros servicios web.

## <a name="security"></a>Seguridad

Las macros de VBA tienen el mismo espacio de seguridad que Excel. Esto les da acceso completo al escritorio. Los scripts de Office solo tienen acceso al libro, no al equipo que hospeda el libro. Además, no se pueden compartir tokens de autenticación de JavaScript con scripts. Esto significa que el script no tiene los tokens del usuario que ha iniciado sesión ni hay capacidades de API para iniciar sesión en un servicio externo, por lo que no pueden usar tokens existentes para realizar llamadas externas en nombre del usuario.

Los administradores tienen tres opciones para macros de VBA: permitir todas las macros en el inquilino, no permitir macros en el inquilino o permitir solo macros con certificados firmados. Esta falta de granularidad hace que sea difícil aislar un solo actor malo. Actualmente, los scripts de Office están desactivados o desactivados para un inquilino. Sin embargo, estamos trabajando para dar a los administradores más control sobre scripts individuales y creadores de scripts.

## <a name="coverage"></a>Cobertura

Actualmente, VBA ofrece una cobertura más completa de las características de Excel, especialmente las disponibles en el cliente de escritorio. Los scripts de Office cubren casi todos los escenarios para Excel en la web. Además, a medida que las nuevas características debutan en la web, los scripts de Office las admitirán tanto para la Grabadora de acciones como para las API de JavaScript.

Los scripts de Office no admiten eventos de nivel de [Excel.](/office/vba/excel/concepts/events-worksheetfunctions-shapes/using-events-with-excel-objects) Los scripts solo se ejecutan cuando un usuario los inicia manualmente o cuando un flujo de Power Automate llama al script.

## <a name="power-automate"></a>Power Automate

Los scripts de Office se pueden ejecutar a través de Power Automate. El libro se puede actualizar a través de flujos programados o controlados por eventos, lo que le permite automatizar flujos de trabajo sin siquiera abrir Excel. Esto significa que, mientras el libro esté almacenado en OneDrive (y accesible para Power Automate), un flujo puede ejecutar los scripts independientemente de si usted y su organización usan el escritorio, Mac o cliente web de Excel.

VBA no tiene un conector de Power Automate. Todos los escenarios de VBA admitidos implicaban a un usuario que asistía a la ejecución de la macro.

## <a name="see-also"></a>Consulte también

- [Scripts de Office en Excel en la Web](../overview/excel.md)
- [Diferencias entre los scripts de Office y los complementos de Office](add-ins-differences.md)
- [Solución de problemas de scripts de Office](../testing/troubleshooting.md)
- [Referencia de VBA para Excel](/office/vba/api/overview/excel)
