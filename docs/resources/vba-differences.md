---
title: Diferencias entre Office scripts y macros de VBA
description: El comportamiento y las diferencias de API entre Office scripts y Excel macros de VBA.
ms.date: 02/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: 53cd2d9b163a3d3c3f9ac9196b5f5126b539611a
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 03/31/2022
ms.locfileid: "64586020"
---
# <a name="differences-between-office-scripts-and-vba-macros"></a>Diferencias entre Office scripts y macros de VBA

Office scripts y macros de VBA tienen mucho en común. Ambos permiten a los usuarios automatizar soluciones a través de una grabadora de acciones fácil de usar y permitir ediciones de esas grabaciones. Ambos marcos están diseñados para habilitar a las personas que pueden no considerarse programadores para crear pequeños programas en Excel.

La diferencia fundamental es que las macros de VBA se desarrollan para soluciones de escritorio y los scripts Office están diseñados para soluciones seguras basadas en la nube. Actualmente, Office scripts solo se admiten en Excel en la Web.

:::image type="content" source="../images/office-programmability-diagram.png" alt-text="Diagrama de cuatro cuadrantes que muestra las áreas de enfoque para diferentes Office extensibilidad. Tanto Office scripts como macros vba están diseñadas para ayudar a los usuarios finales a crear soluciones, pero Office scripts se crean para la web y la colaboración (mientras que VBA es para el escritorio).":::

En este artículo se describen las principales diferencias entre las macros de VBA (así como VBA en general) y Office scripts. Dado que Office scripts solo están disponibles para Excel, este es el único host que se trata aquí.

## <a name="platform-and-ecosystem"></a>Plataforma y ecosistema

VBA es compatible con Excel en Windows y Mac. Office scripts es compatible con Excel en la Web.

Las dos soluciones se diseñaron para sus respectivas plataformas. VBA puede interactuar con el escritorio de un usuario para conectarse con tecnologías similares, como COM y OLE. Sin embargo, VBA no tiene una forma cómoda de llamar a Internet. Office scripts usan un tiempo de ejecución universal para JavaScript. Esto proporciona un comportamiento y accesibilidad coherentes, independientemente de la máquina que se esté utilizando para ejecutar el script. También pueden realizar llamadas a otros servicios web.

### <a name="script-support-for-excel-on-windows"></a>Compatibilidad con scripts para Excel en Windows

[!INCLUDE [Run-from-button support](../includes/run-from-button-desktop-support.md)]

## <a name="security"></a>Seguridad

Las macros de VBA tienen el mismo espacio de seguridad que Excel. Esto les da acceso completo al escritorio. Office los scripts solo tienen acceso al libro, no al equipo que hospeda el libro. Además, no se pueden compartir tokens de autenticación de JavaScript con scripts. Esto significa que el script no tiene los tokens del usuario que ha iniciado sesión ni hay capacidades de API para iniciar sesión en un servicio externo, por lo que no pueden usar tokens existentes para realizar llamadas externas en nombre del usuario.

Los administradores tienen tres opciones para macros de VBA: permitir todas las macros en el inquilino, no permitir macros en el inquilino o permitir solo macros con certificados firmados. Esta falta de granularidad hace que sea difícil aislar un solo actor malo. Actualmente, Office scripts pueden estar desactivados para todo un inquilino, para todo un inquilino o para un grupo de usuarios de un espacio empresarial. Los administradores también tienen control sobre quién puede compartir scripts con otros usuarios y quién puede usar scripts en Power Automate.

## <a name="coverage"></a>Cobertura

Actualmente, VBA ofrece una cobertura más completa de Excel características, especialmente las disponibles en el cliente de escritorio. Office scripts cubren casi todos los escenarios para Excel en la Web. Además, a medida que las nuevas características debutan en la web, Office scripts las admitirán tanto para la Grabadora de acciones como para las API de JavaScript.

Office scripts no admiten eventos Excel nivel[.](/office/vba/excel/concepts/events-worksheetfunctions-shapes/using-events-with-excel-objects) Los scripts solo se ejecutan cuando un usuario los inicia manualmente o cuando un flujo de Power Automate llama al script.

## <a name="power-automate"></a>Power Automate

Office scripts se pueden ejecutar a través de Power Automate. El libro se puede actualizar a través de flujos programados o controlados por eventos, lo que le permite automatizar flujos de trabajo sin siquiera abrir Excel. Esto significa que, mientras el libro esté almacenado en OneDrive (y accesible para Power Automate), un flujo puede ejecutar los scripts independientemente de si usted y su organización usan el escritorio, Mac o cliente web de Excel.

VBA no tiene un conector Power Automate de datos. Todos los escenarios de VBA admitidos implican que un usuario asista a la ejecución de la macro.

Pruebe los [scripts de llamada desde un tutorial de flujo Power Automate manual](../tutorials/excel-power-automate-manual.md) para empezar a aprender sobre Power Automate. También puede consultar el ejemplo [De avisos de tareas automatizadas](scenarios/task-reminders.md) para ver Office scripts conectados a Teams a través de Power Automate en un escenario real.

## <a name="see-also"></a>Consulte también

- [Scripts de Office en Excel en la Web](../overview/excel.md)
- [Ejecutar Office scripts con Power Automate](../develop/power-automate-integration.md)
- [Diferencias entre los scripts de Office y los complementos de Office](add-ins-differences.md)
- [Solución de problemas de scripts de Office](../testing/troubleshooting.md)
- [Referencia de VBA para Excel](/office/vba/api/overview/excel)
