---
title: Diferencias entre scripts de Office y macros de VBA
description: El comportamiento y las diferencias de API entre Office Scripts y Excel macros de VBA.
ms.date: 02/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: 60e4fba6e63967302066f544b76fb20a8c8630a6
ms.sourcegitcommit: 34c7740c9bff0e4c7426e01029f967724bfee566
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 05/13/2022
ms.locfileid: "65393617"
---
# <a name="differences-between-office-scripts-and-vba-macros"></a>Diferencias entre scripts de Office y macros de VBA

Office scripts y macros vba tienen mucho en común. Ambos permiten a los usuarios automatizar soluciones a través de una grabadora de acciones fácil de usar y permiten la edición de esas grabaciones. Ambos marcos están diseñados para capacitar a las personas que no se consideren programadores para crear programas pequeños en Excel.

La diferencia fundamental es que las macros de VBA se desarrollan para soluciones de escritorio y Office scripts están diseñados para soluciones seguras basadas en la nube. Actualmente, los scripts de Office solo se admiten en Excel en la Web.

:::image type="content" source="../images/office-programmability-diagram.png" alt-text="Diagrama de cuatro cuadrantes que muestra las áreas de enfoque para diferentes soluciones de extensibilidad Office. Tanto los scripts Office como las macros vba están diseñados para ayudar a los usuarios finales a crear soluciones, pero Office scripts se crean para la web y la colaboración (mientras que VBA es para el escritorio).":::

En este artículo se describen las principales diferencias entre las macros de VBA (así como VBA en general) y los scripts de Office. Dado que Office scripts solo están disponibles para Excel, ese es el único host que se describe aquí.

## <a name="platform-and-ecosystem"></a>Plataforma y ecosistema

VBA es compatible con Excel en Windows y Mac. Office Scripts es compatible con Excel en la Web.

Las dos soluciones se diseñaron para sus respectivas plataformas. VBA puede interactuar con el escritorio de un usuario para conectarse con tecnologías similares, como COM y OLE. Sin embargo, VBA no tiene ninguna manera conveniente de llamar a Internet. Office scripts usan un entorno de ejecución universal para JavaScript. Esto proporciona un comportamiento coherente y accesibilidad, independientemente de la máquina que se use para ejecutar el script. También pueden realizar llamadas a otros servicios web.

### <a name="script-support-for-excel-on-windows"></a>Compatibilidad con scripts para Excel en Windows

[!INCLUDE [Run-from-button support](../includes/run-from-button-desktop-support.md)]

## <a name="security"></a>Seguridad

Las macros de VBA tienen el mismo espacio de seguridad que Excel. Esto les proporciona acceso total al escritorio. Office scripts solo tienen acceso al libro, no a la máquina que hospeda el libro. Además, no se pueden compartir tokens de autenticación de JavaScript con scripts. Esto significa que el script no tiene los tokens del usuario que ha iniciado sesión ni hay ninguna funcionalidad de API para iniciar sesión en un servicio externo, por lo que no puede usar tokens existentes para realizar llamadas externas en nombre del usuario.

Los administradores tienen tres opciones para macros vba: permitir todas las macros en el inquilino, no permitir macros en el inquilino o permitir solo macros con certificados firmados. Esta falta de granularidad hace que sea difícil aislar un solo actor malo. Actualmente, los scripts de Office pueden estar desactivados para todo un inquilino, para todo un inquilino o para un grupo de usuarios de un inquilino. Los administradores también tienen control sobre quién puede compartir scripts con otros usuarios y quién puede usar scripts en Power Automate.

## <a name="coverage"></a>Cobertura

Actualmente, VBA ofrece una cobertura más completa de las características de Excel, especialmente las disponibles en el cliente de escritorio. Office scripts cubren casi todos los escenarios de Excel en la Web. Además, a medida que las nuevas características debutan en la web, Office Scripts los admitirá tanto para la Grabadora de acciones como para las API de JavaScript.

Office scripts no admiten [eventos](/office/vba/excel/concepts/events-worksheetfunctions-shapes/using-events-with-excel-objects) de nivel de Excel. Los scripts solo se ejecutan cuando un usuario los inicia manualmente o cuando un flujo de Power Automate llama al script.

## <a name="power-automate"></a>Power Automate

Office scripts se pueden ejecutar a través de Power Automate. El libro se puede actualizar a través de flujos programados o controlados por eventos, lo que permite automatizar los flujos de trabajo sin siquiera abrir Excel. Esto significa que, siempre y cuando el libro se almacene en OneDrive (y sea accesible para Power Automate), un flujo puede ejecutar los scripts independientemente de si usted y su organización usan el escritorio, Mac o el cliente web de Excel.

VBA no tiene un conector de Power Automate. Todos los escenarios de VBA admitidos implican que un usuario asista a la ejecución de la macro.

Pruebe [los scripts de llamada de un tutorial de flujo de Power Automate manual](../tutorials/excel-power-automate-manual.md) para empezar a aprender sobre Power Automate. También puede consultar el ejemplo [de recordatorios de tareas automatizadas](scenarios/task-reminders.md) para ver Office scripts conectados a Teams a través de Power Automate en un escenario real.

## <a name="see-also"></a>Consulte también

- [scripts de Office en Excel](../overview/excel.md)
- [Ejecución de scripts de Office con Power Automate](../develop/power-automate-integration.md)
- [Diferencias entre los scripts de Office y los complementos de Office](add-ins-differences.md)
- [Solución de problemas de scripts de Office](../testing/troubleshooting.md)
- [Referencia de VBA para Excel](/office/vba/api/overview/excel)
