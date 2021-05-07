---
title: Diferencias entre Office scripts y macros de VBA
description: El comportamiento y las diferencias de API entre Office scripts y Excel macros de VBA.
ms.date: 12/14/2020
localization_priority: Normal
ms.openlocfilehash: ca571e2adad81a87b99696a652a3c49209b870ab
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232847"
---
# <a name="differences-between-office-scripts-and-vba-macros"></a>Diferencias entre Office scripts y macros de VBA

Office Los scripts y las macros de VBA tienen mucho en común. Ambos permiten a los usuarios automatizar soluciones a través de una grabadora de acciones fácil de usar y permitir ediciones de esas grabaciones. Ambos marcos están diseñados para habilitar a las personas que pueden no considerarse programadores para crear pequeños programas en Excel.
La diferencia fundamental es que las macros de VBA se desarrollan para soluciones de escritorio y los scripts de Office están diseñados con la seguridad y la compatibilidad entre plataformas como principios de guía. Actualmente, Office scripts solo se admiten en Excel en la Web.

:::image type="content" source="../images/office-programmability-diagram.png" alt-text="Diagrama de cuatro cuadrantes que muestra las áreas de enfoque para diferentes Office extensibilidad. Tanto Office scripts como macros vba están diseñadas para ayudar a los usuarios finales a crear soluciones, pero Office Scripts se crean para la web y la colaboración (mientras que VBA es para el escritorio)":::

En este artículo se describen las principales diferencias entre macros de VBA (así como VBA en general) y Office scripts. Dado que Office scripts solo están disponibles para Excel, este es el único host que se trata aquí.

## <a name="platform-and-ecosystem"></a>Plataforma y ecosistema

VBA está diseñado para el escritorio y Office scripts están diseñados para la web. VBA puede interactuar con el escritorio de un usuario para conectarse con tecnologías similares, como COM y OLE. Sin embargo, VBA no tiene una forma cómoda de llamar a Internet.

Office Los scripts usan un tiempo de ejecución universal para JavaScript. Esto proporciona un comportamiento y accesibilidad coherentes, independientemente de la máquina que se esté utilizando para ejecutar el script. También pueden realizar llamadas a otros servicios web.

## <a name="security"></a>Seguridad

Las macros de VBA tienen el mismo espacio de seguridad que Excel. Esto les da acceso completo al escritorio. Office Los scripts solo tienen acceso al libro, no al equipo que hospeda el libro. Además, no se pueden compartir tokens de autenticación de JavaScript con scripts. Esto significa que el script no tiene los tokens del usuario que ha iniciado sesión ni hay capacidades de API para iniciar sesión en un servicio externo, por lo que no pueden usar tokens existentes para realizar llamadas externas en nombre del usuario.

Los administradores tienen tres opciones para macros de VBA: permitir todas las macros en el inquilino, no permitir macros en el inquilino o permitir solo macros con certificados firmados. Esta falta de granularidad hace que sea difícil aislar un solo actor malo. Actualmente, Office scripts están en o desactivados para un inquilino. Sin embargo, estamos trabajando para dar a los administradores más control sobre scripts individuales y creadores de scripts.

## <a name="coverage"></a>Cobertura

Actualmente, VBA ofrece una cobertura más completa de Excel características, especialmente las disponibles en el cliente de escritorio. Office Los scripts abarcan casi todos los escenarios para Excel en la Web. Además, a medida que las nuevas características debutan en la web, Office scripts las admitirán tanto para la Grabadora de acciones como para las API de JavaScript.

Office Los scripts no admiten Excel de [nivel de archivo](/office/vba/excel/concepts/events-worksheetfunctions-shapes/using-events-with-excel-objects). Los scripts solo se ejecutan cuando un usuario los inicia manualmente o cuando un flujo Power Automate llama al script.

## <a name="power-automate"></a>Power Automate

Office Los scripts se pueden ejecutar Power Automate. El libro se puede actualizar a través de flujos programados o controlados por eventos, lo que le permite automatizar flujos de trabajo sin siquiera abrir Excel. Esto significa que, mientras el libro esté almacenado en OneDrive (y accesible para Power Automate), un flujo puede ejecutar los scripts independientemente de si usted y su organización usan el escritorio, Mac o cliente web de Excel.

VBA no tiene un conector Power Automate. Todos los escenarios de VBA admitidos implicaban a un usuario que asistía a la ejecución de la macro.

## <a name="see-also"></a>Vea también

- [Scripts de Office en Excel en la Web](../overview/excel.md)
- [Diferencias entre los scripts de Office y los complementos de Office](add-ins-differences.md)
- [Solución de problemas de scripts de Office](../testing/troubleshooting.md)
- [Referencia de VBA para Excel](/office/vba/api/overview/excel)
