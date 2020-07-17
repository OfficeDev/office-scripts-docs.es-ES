---
title: Entorno de editor de código de scripts de Office
description: Los requisitos previos e información del entorno para los scripts de Office en Excel en la Web.
ms.date: 07/10/2020
localization_priority: Normal
ms.openlocfilehash: 643ea2d5bd69adf4311546465ccd65c08dacf4b4
ms.sourcegitcommit: ebd1079c7e2695ac0e7e4c616f2439975e196875
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 07/17/2020
ms.locfileid: "45160498"
---
# <a name="office-scripts-code-editor-environment"></a>Entorno de editor de código de scripts de Office

Los scripts de Office se escriben en [TypeScript o JavaScript](#scripting-language-typescript-or-javascript) , y usan las [API de JavaScript de scripts de Office](#office-scripts-javascript-api) para interactuar con un libro de Excel.

## <a name="scripting-language-typescript-or-javascript"></a>Lenguaje de scripting: TypeScript o JavaScript

Los scripts de Office se escriben en [TypeScript](https://www.typescriptlang.org/docs/home.html) o [JavaScript](https://developer.mozilla.org/docs/Web/JavaScript). La grabadora de acciones genera código en TypeScript (que es un superconjunto de JavaScript). La documentación de scripts de Office usa TypeScript, pero si está más familiarizado con JavaScript, puede usarlo en su lugar.

Los scripts de Office son en gran medida fragmentos de código independientes. Solo se usa una pequeña parte de la funcionalidad de TypeScript. Por lo tanto, puede editar scripts sin tener que aprender las complejidades de TypeScript. El editor de código también controla la instalación, la compilación y la ejecución de código, por lo que no tiene que preocuparse de nada excepto de la propia secuencia de comandos. Es posible aprender el lenguaje y crear scripts sin conocimientos previos de programación. Sin embargo, si no está familiarizado con la programación, le recomendamos que conozca algunos conceptos básicos antes de continuar con los scripts de Office:

- Obtenga información sobre los conceptos básicos de JavaScript. Debe sentirse cómodo con conceptos como variables, flujo de control, funciones y tipos de datos. [Mozilla ofrece un buen y completo tutorial sobre JavaScript](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).
- Obtenga información sobre los tipos de TypeScript. TypeScript se basa en JavaScript al garantizar en tiempo de compilación los tipos correctos se usan para las llamadas de método y las asignaciones. La documentación de TypeScript en [interfaces](https://www.typescriptlang.org/docs/handbook/interfaces.html), [clases](https://www.typescriptlang.org/docs/handbook/classes.html), [inferencia de tipos](https://www.typescriptlang.org/docs/handbook/type-inference.html)y compatibilidad de [tipos](https://www.typescriptlang.org/docs/handbook/type-compatibility.html) será la más útil.

## <a name="office-scripts-javascript-api"></a>API de JavaScript de scripts de Office

Los scripts de Office usan una versión especializada de las API de JavaScript para Office para [Complementos de Office](/office/dev/add-ins/overview/index). Aunque existen similitudes en las dos API, no debe suponer que el código se puede trasladar entre las dos plataformas. Las diferencias entre las dos plataformas se describen en el artículo [diferencias entre scripts de Office y complementos de Office](../resources/add-ins-differences.md#apis) . Puede ver todas las API disponibles para el script en la documentación de referencia de la [API de scripts de Office](/javascript/api/office-scripts/overview).

## <a name="intellisense"></a>Característica

IntelliSense es una característica del editor de código que ayuda a evitar errores tipográficos y sintácticos mientras se edita el script. Muestra los nombres de campo y de objeto posibles mientras se escribe, así como documentación en línea para cada API.

El editor de código de Excel usa el mismo motor de IntelliSense que Visual Studio Code. Para obtener más información acerca de la característica, visite [las características de IntelliSense de Visual Studio Code](https://code.visualstudio.com/docs/editor/intellisense#_intellisense-features).

## <a name="external-library-support"></a>Compatibilidad con bibliotecas externas

Los scripts de Office no admiten el uso de bibliotecas externas de JavaScript de terceros. Actualmente no puede llamar a ninguna biblioteca que no sea las API de scripts de Office desde un script. Todavía tiene acceso a cualquier [objeto de JavaScript integrado](../develop/javascript-objects.md), como [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math).

## <a name="browser-support"></a>Compatibilidad con exploradores

Los scripts de Office funcionan en cualquier explorador que [admita Office para la web](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452). Sin embargo, algunas características de JavaScript no se admiten en Internet Explorer 11 (IE 11). Las características que se incluyen en [ES6 o versiones posteriores](https://www.w3schools.com/Js/js_es6.asp) no funcionarán con IE 11. Si los usuarios de su organización todavía usan ese explorador, asegúrese de probar los scripts en ese entorno al compartirlos.

## <a name="see-also"></a>Vea también

- [Referencia de API de scripts de Office](/javascript/api/office-scripts/overview)
- [Solución de problemas de scripts de Office](../testing/troubleshooting.md)
- [Usar objetos integrados de JavaScript en los scripts de Office](../develop/javascript-objects.md)
