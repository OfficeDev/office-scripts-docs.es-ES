---
title: Diferencias entre scripts de Office y macros de VBA
description: El comportamiento y las diferencias de API entre scripts de Office y macros de VBA de Excel.
ms.date: 03/23/2020
localization_priority: Normal
ms.openlocfilehash: 3a0f2c9a2ed7181a10e41d1f45b3af695877a680
ms.sourcegitcommit: d556aaefac80e55f53ac56b7f6ecbc657ebd426f
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 03/26/2020
ms.locfileid: "42978725"
---
# <a name="differences-between-office-scripts-and-vba-macros"></a>Diferencias entre scripts de Office y macros de VBA

Los scripts de Office y las macros de VBA tienen mucho en común. Ambos permiten a los usuarios automatizar soluciones a través de una grabadora de acciones fácil de usar y permitir ediciones de esas grabaciones. Ambos marcos están diseñados para proporcionar a los usuarios que no tienen que considerar a sí mismos programadores que creen pequeños programas en Excel.
La diferencia fundamental es que las macros de VBA que se desarrollan para las soluciones de escritorio y las secuencias de comandos de Office se diseñan con la seguridad y la compatibilidad entre plataformas como principios de guía. Actualmente, los scripts de Office solo se admiten en Excel en la Web.

![Un diagrama de cuatro fases que muestra las áreas de atención para diferentes soluciones de extensibilidad de Office. Tanto los scripts de Office como las macros de VBA están diseñados para ayudar a los usuarios finales a crear soluciones, pero los scripts de Office se crean para la web y la colaboración (mientras que VBA es para el escritorio)).](../images/office-programmability-diagram.png)

En este artículo se describen las principales diferencias entre las macros de VBA (así como VBA en general) y las secuencias de comandos de Office. Como los scripts de Office solo están disponibles para Excel, es el único host que se describe aquí.

## <a name="platform-and-ecosystem"></a>Plataforma y ecosistema

VBA está diseñado para el escritorio y los scripts de Office están diseñados para la Web. VBA puede interactuar con el escritorio de un usuario. Esto permite la integración con tecnologías similares, como COM y OLE. Sin embargo, VBA no tiene ninguna forma cómoda de llamar a Internet.

Los scripts de Office usan un tiempo de ejecución universal o JavaScript. Esto proporciona un comportamiento y una accesibilidad coherentes, independientemente del equipo que se use para ejecutar el script. También pueden realizar llamadas a otros servicios Web.

## <a name="security"></a>Seguridad

Las macros de VBA tienen la misma holgura de seguridad que Excel. Esto les da acceso total a su escritorio. Los scripts de Office solo tienen acceso al libro, no al equipo que hospeda el libro. Además, no se pueden compartir los tokens de autenticación de JavaScript con los scripts, de modo que los scripts no se pueden autenticar nunca con un servicio externo.

Los administradores tienen tres opciones para las macros de VBA: permitir todas las macros en el espacio empresarial, no permitir macros en el espacio empresarial o permitir solo macros con certificados firmados. Esta falta de granularidad hace que sea difícil aislar un solo actor incorrecto. Actualmente, las secuencias de comandos de Office están activadas o desactivadas para un espacio empresarial. Sin embargo, estamos trabajando para dar a los administradores más control sobre los creadores de scripts y scripts individuales.

## <a name="coverage"></a>Infra

Actualmente, VBA ofrece una cobertura más completa de las características de Excel, especialmente las que están disponibles en el cliente de escritorio. Los scripts de Office cubren casi todos los escenarios para Excel en la Web. Además, como nuevas características del lanzamiento en la web, las secuencias de comandos de Office serán compatibles con la grabadora de acciones y las API de JavaScript.

## <a name="power-automate"></a>Automatizar la alimentación

Los scripts de Office se pueden ejecutar a través de la automatización de energía. El libro se puede actualizar mediante flujos programados o controlados por eventos, lo que permite automatizar flujos de trabajo sin necesidad de abrir Excel. Esto significa que, siempre que el libro se almacene en OneDrive (y sea accesible para la automatización automática), un flujo puede ejecutar los scripts independientemente de si usted y su organización usan el cliente de escritorio, Mac o Web de Excel.

VBA no tiene integración con la automatización. Todos los escenarios de VBA admitidos involucrados en un usuario que asiste a la ejecución de la macro.

## <a name="see-also"></a>Consulte también

- [Scripts de Office en Excel en la Web](../overview/excel.md)
- [Diferencias entre scripts de Office y complementos de Office](add-ins-differences.md)
- [Solución de problemas de scripts de Office](../testing/troubleshooting.md)
- [Referencia de VBA para Excel](/office/vba/api/overview/excel)
