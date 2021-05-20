---
title: Diferencias entre Office scripts y macros VBA
description: El comportamiento y las diferencias de API entre scripts de Office y macros de VBA Excel.
ms.date: 05/18/2021
localization_priority: Normal
ms.openlocfilehash: 612a5f21d935fd262a6e9fd12a3431956105636a
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545591"
---
# <a name="differences-between-office-scripts-and-vba-macros"></a>Diferencias entre Office scripts y macros VBA

Office Los scripts y las macros VBA tienen mucho en común. Ambos permiten a los usuarios automatizar soluciones a través de una grabadora de acciones fácil de usar y permiten la edición de esas grabaciones. Ambos marcos están diseñados para empoderar a las personas que tal vez no se consideren programadoras para crear pequeños programas en Excel.
La diferencia fundamental es que las macros VBA se desarrollan para soluciones de escritorio y Office scripts están diseñados para soluciones seguras basadas en la nube. Actualmente, los scripts de Office solo se admiten en Excel en la Web.

:::image type="content" source="../images/office-programmability-diagram.png" alt-text="Un diagrama de cuatro cuadrantes que muestra las áreas de enfoque para diferentes soluciones de extensibilidad Office. Tanto Office scripts como las macros VBA están diseñadas para ayudar a los usuarios finales a crear soluciones, pero Office scripts están diseñados para la web y la colaboración (mientras que VBA es para el escritorio)":::

En este artículo se describen las principales diferencias entre las macros VBA (así como VBA en general) y los scripts de Office. Dado que los scripts de Office solo están disponibles para Excel, ese es el único host que se discute aquí.

## <a name="platform-and-ecosystem"></a>Plataforma y ecosistema

VBA está diseñado para el escritorio y Office scripts están diseñados para la web. VBA puede interactuar con el escritorio de un usuario para conectarse con tecnologías similares, como COM y OLE. Sin embargo, VBA no tiene una manera conveniente de llamar a Internet.

Office Los scripts utilizan un tiempo de ejecución universal para JavaScript. Esto proporciona un comportamiento y una accesibilidad coherentes, independientemente de que el equipo se utilice para ejecutar el script. También pueden realizar llamadas a otros servicios web.

## <a name="security"></a>Seguridad

Las macros VBA tienen la misma autorización de seguridad que Excel. Esto les da acceso completo a su escritorio. Office Los scripts solo tienen acceso al libro, no a la máquina que hospeda el libro. Además, no se pueden compartir tokens de autenticación de JavaScript con scripts. Esto significa que el script no tiene ni los tokens del usuario que ha iniciado sesión ni hay ninguna funcionalidad de API para iniciar sesión en un servicio externo, por lo que no pueden usar tokens existentes para realizar llamadas externas en nombre del usuario.

Los administradores tienen tres opciones para macros de VBA: permitir todas las macros en el inquilino, no permitir macros en el inquilino o permitir solo macros con certificados firmados. Esta falta de granularidad hace que sea difícil aislar a un solo mal actor. Actualmente, Office scripts están activados o desactivando para un inquilino. Sin embargo, estamos trabajando para dar a los administradores más control sobre los scripts individuales y los creadores de scripts.

## <a name="coverage"></a>cobertura

Actualmente, VBA ofrece una cobertura más completa de las características de Excel, particularmente las disponibles en el cliente de escritorio. Office Los scripts cubren casi todos los escenarios para Excel en la Web. Además, a medida que las nuevas características debutan en la web, Office Scripts las admitirá tanto para el Grabador de acciones como para las API de JavaScript.

Office Los scripts no admiten [eventos](/office/vba/excel/concepts/events-worksheetfunctions-shapes/using-events-with-excel-objects)de nivel Excel. Los scripts solo se ejecutan cuando un usuario los inicia manualmente o cuando un flujo de Power Automate llama al script.

## <a name="power-automate"></a>Power Automate

Office Los scripts se pueden ejecutar a través de Power Automate. El libro de trabajo se puede actualizar a través de flujos programados o controlados por eventos, lo que le permite automatizar flujos de trabajo sin siquiera abrir Excel. Esto significa que mientras el libro de trabajo se almacene en OneDrive (y sea accesible para Power Automate), un flujo puede ejecutar los scripts independientemente de si usted y su organización usan el escritorio, Mac o cliente web de Excel.

VBA no tiene un conector Power Automate. Todos los escenarios de VBA admitidos implican un usuario que atiende a la ejecución de la macro.

Pruebe los [scripts de llamada de un](../tutorials/excel-power-automate-manual.md) tutorial manual de flujo Power Automate para empezar a aprender sobre Power Automate. También puede consultar el ejemplo [recordatorios de tareas automatizadas](scenarios/task-reminders.md) para ver Office scripts conectados a Teams a través de Power Automate en un escenario del mundo real.

## <a name="see-also"></a>Vea también

- [Scripts de Office en Excel en la Web](../overview/excel.md)
- [Ejecute scripts de Office con Power Automate](../develop/power-automate-integration.md)
- [Diferencias entre los scripts de Office y los complementos de Office](add-ins-differences.md)
- [Solución de problemas de scripts de Office](../testing/troubleshooting.md)
- [Referencia de VBA para Excel](/office/vba/api/overview/excel)
