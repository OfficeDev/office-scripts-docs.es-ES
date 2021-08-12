---
title: Diferencias entre Office scripts y macros de VBA
description: El comportamiento y las diferencias de API entre Office scripts y Excel macros de VBA.
ms.date: 05/21/2021
localization_priority: Normal
ms.openlocfilehash: 0d94607902fa62e07ce378b94ec3b9c328937e16535b1882b6cad5bd76212b33
ms.sourcegitcommit: 75f7ed8c2d23a104acc293f8ce29ea580b4fcdc5
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 08/11/2021
ms.locfileid: "57847271"
---
# <a name="differences-between-office-scripts-and-vba-macros"></a>Diferencias entre Office scripts y macros de VBA

Office Los scripts y las macros de VBA tienen mucho en común. Ambos permiten a los usuarios automatizar soluciones a través de una grabadora de acciones fácil de usar y permitir ediciones de esas grabaciones. Ambos marcos están diseñados para habilitar a las personas que pueden no considerarse programadores para crear pequeños programas en Excel.
La diferencia fundamental es que las macros de VBA se desarrollan para soluciones de escritorio y Office scripts están diseñados para soluciones seguras basadas en la nube. Actualmente, Office scripts solo se admiten en Excel en la Web.

:::image type="content" source="../images/office-programmability-diagram.png" alt-text="Diagrama de cuatro cuadrantes que muestra las áreas de enfoque para diferentes Office extensibilidad. Tanto Office scripts como macros de VBA están diseñadas para ayudar a los usuarios finales a crear soluciones, pero Office Scripts se crean para la web y la colaboración (mientras que VBA es para el escritorio).":::

En este artículo se describen las principales diferencias entre macros de VBA (así como VBA en general) y Office scripts. Dado que Office scripts solo están disponibles para Excel, este es el único host que se trata aquí.

## <a name="platform-and-ecosystem"></a>Plataforma y ecosistema

VBA está diseñado para el escritorio y Office scripts están diseñados para la web. VBA puede interactuar con el escritorio de un usuario para conectarse con tecnologías similares, como COM y OLE. Sin embargo, VBA no tiene una forma cómoda de llamar a Internet.

Office Los scripts usan un tiempo de ejecución universal para JavaScript. Esto proporciona un comportamiento y accesibilidad coherentes, independientemente de la máquina que se esté utilizando para ejecutar el script. También pueden realizar llamadas a otros servicios web.

## <a name="security"></a>Seguridad

Las macros de VBA tienen el mismo espacio de seguridad que Excel. Esto les da acceso completo al escritorio. Office Los scripts solo tienen acceso al libro, no al equipo que hospeda el libro. Además, no se pueden compartir tokens de autenticación de JavaScript con scripts. Esto significa que el script no tiene los tokens del usuario que ha iniciado sesión ni hay capacidades de API para iniciar sesión en un servicio externo, por lo que no pueden usar tokens existentes para realizar llamadas externas en nombre del usuario.

Los administradores tienen tres opciones para macros de VBA: permitir todas las macros en el inquilino, no permitir macros en el inquilino o permitir solo macros con certificados firmados. Esta falta de granularidad hace que sea difícil aislar un solo actor malo. Actualmente, Office scripts pueden estar desactivados para todo un inquilino, para todo un inquilino o para un grupo de usuarios de un inquilino. Los administradores también tienen control sobre quién puede compartir scripts con otros usuarios y quién puede usar scripts en Power Automate.

## <a name="coverage"></a>Cobertura

Actualmente, VBA ofrece una cobertura más completa de Excel características, especialmente las disponibles en el cliente de escritorio. Office Los scripts abarcan casi todos los escenarios para Excel en la Web. Además, a medida que las nuevas características debutan en la web, Office scripts las admitirán tanto para la Grabadora de acciones como para las API de JavaScript.

Office Los scripts no admiten Excel de [nivel de archivo](/office/vba/excel/concepts/events-worksheetfunctions-shapes/using-events-with-excel-objects). Los scripts solo se ejecutan cuando un usuario los inicia manualmente o cuando un flujo Power Automate llama al script.

## <a name="power-automate"></a>Power Automate

Office Los scripts se pueden ejecutar Power Automate. El libro se puede actualizar a través de flujos programados o controlados por eventos, lo que le permite automatizar flujos de trabajo sin siquiera abrir Excel. Esto significa que, mientras el libro esté almacenado en OneDrive (y accesible para Power Automate), un flujo puede ejecutar los scripts independientemente de si usted y su organización usan el escritorio, Mac o cliente web de Excel.

VBA no tiene un conector Power Automate. Todos los escenarios de VBA admitidos implican que un usuario asista a la ejecución de la macro.

Pruebe los [scripts de llamada desde un tutorial de flujo Power Automate manual](../tutorials/excel-power-automate-manual.md) para empezar a aprender sobre Power Automate. También puede consultar el ejemplo [De avisos](scenarios/task-reminders.md) de tareas automatizadas para ver Office scripts conectados a Teams a través de Power Automate en un escenario real.

## <a name="see-also"></a>Vea también

- [Scripts de Office en Excel en la Web](../overview/excel.md)
- [Ejecute Office scripts con Power Automate](../develop/power-automate-integration.md)
- [Diferencias entre los scripts de Office y los complementos de Office](add-ins-differences.md)
- [Solución de problemas de scripts de Office](../testing/troubleshooting.md)
- [Referencia de VBA para Excel](/office/vba/api/overview/excel)
