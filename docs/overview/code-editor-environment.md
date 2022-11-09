---
title: Entorno del Editor de código de scripts de Office
description: Los requisitos previos y la información del entorno de los scripts de Office en Excel en la Web.
ms.date: 11/08/2022
ms.localizationpriority: medium
ms.openlocfilehash: a5a7601285553b1da4001a1870b6120f21bf5f2c
ms.sourcegitcommit: 7cadf2b637bf62874e43b6e595286101816662aa
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 11/09/2022
ms.locfileid: "68891256"
---
# <a name="office-scripts-code-editor-environment"></a>Entorno del Editor de código de scripts de Office

Los scripts de Office se escriben en TypeScript o JavaScript y usan las API de JavaScript de Scripts de Office para interactuar con un libro de Excel. El Editor de código se basa en Visual Studio Code, por lo que si ha usado ese entorno antes, se sentirá como en casa.

> [!TIP]
> Si está familiarizado con Visual Studio Code, ahora puede usarlo para escribir scripts. Visite [Visual Studio Code para scripts de Office (versión preliminar)](../develop/vscode-for-scripts.md) para probar esta característica.

## <a name="scripting-language-typescript-or-javascript"></a>Lenguaje de scripting: TypeScript o JavaScript

Los Scripts de Office se escriben en [TypeScript](https://www.typescriptlang.org/docs/home.html), que es un superconjunto de [JavaScript](https://developer.mozilla.org/docs/Web/JavaScript). Action Recorder genera código en TypeScript y la documentación de Scripts de Office usa TypeScript. Dado que TypeScript es un superconjunto de JavaScript, cualquier código de scripting que escriba en JavaScript funcionará perfectamente.

Los scripts de Office son en gran parte fragmentos de código independientes. Solo se usa una pequeña parte de la funcionalidad de TypeScript. Por lo tanto, puede editar scripts sin tener que aprender las complejidades de TypeScript. El Editor de código también controla la instalación, compilación y ejecución de código, por lo que no es necesario preocuparse por nada más que el propio script. Es posible aprender el lenguaje y crear scripts sin conocimientos de programación anteriores. Sin embargo, si no está familiarizado con la programación, se recomienda aprender algunos aspectos básicos antes de continuar con scripts de Office:

[!INCLUDE [Recommended coding resources](../includes/coding-basics-references.md)]

## <a name="office-scripts-javascript-api"></a>Office Scripts JavaScript API

Los scripts de Office usan una versión especializada de las API de JavaScript de Office para [complementos de Office](/office/dev/add-ins/overview/index). Aunque hay similitudes en las dos API, no se debe suponer que el código se puede migrar entre las dos plataformas. Las diferencias entre las dos plataformas se describen en el artículo [Diferencias entre scripts de Office y complementos de Office](../resources/add-ins-differences.md#apis) . Puede ver todas las API disponibles para el script en la documentación de referencia de la [API de Scripts de Office](/javascript/api/office-scripts/overview).

## <a name="external-library-support"></a>Compatibilidad con bibliotecas externas

Office Scripts no admite el uso de bibliotecas de JavaScript externas de terceros. Actualmente, no se puede llamar a ninguna biblioteca distinta de las API de scripts de Office desde un script. Todavía tiene acceso a cualquier [objeto de JavaScript integrado](../develop/javascript-objects.md), como [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math).

## <a name="intellisense"></a>Intellisense

IntelliSense es un conjunto de características del Editor de código que le ayudan a escribir código. Proporciona autocompletar, resaltado de errores de sintaxis y documentación de api insertada.

IntelliSense proporciona sugerencias a medida que escribe, de forma similar al texto sugerido en Excel. Al presionar la tecla Tab o Enter, se inserta el miembro sugerido. Desencadene IntelliSense en la ubicación actual del cursor presionando las teclas Ctrl+Espacio. Estas sugerencias son especialmente útiles al completar un método. La firma del método mostrada por IntelliSense contiene una lista de argumentos que necesita, el tipo de cada argumento, si un argumento determinado es obligatorio u opcional, y el tipo de valor devuelto del método.

Mantenga el cursor sobre un método, una clase u otro objeto de código para ver más información. Mantenga el puntero sobre un error de sintaxis o sugerencia de código, representado por una línea ondulada roja o amarilla, para ver sugerencias sobre cómo solucionar el problema. A menudo, IntelliSense proporciona una opción "Corrección rápida" para cambiar automáticamente el código.

:::image type="content" source="../images/implicit-any-editor-message.png" alt-text="Mensaje de error en el texto del mouse del Editor de código con un botón &quot;Corrección rápida&quot;.":::

El Editor de código de scripts de Office usa el mismo motor de IntelliSense que Visual Studio Code. Para más información sobre la característica, visite [características de IntelliSense de Visual Studio Code](https://code.visualstudio.com/docs/editor/intellisense#_intellisense-features).

## <a name="keyboard-shortcuts"></a>Accesos rápidos de teclado

La mayoría de los métodos abreviados de teclado para Visual Studio Code también funcionan en el Editor de código de scripts de Office. Use los siguientes archivos PDF para obtener información sobre las opciones disponibles y sacar el máximo partido del Editor de código:

- [Métodos abreviados de teclado para macOS](https://code.visualstudio.com/shortcuts/keyboard-shortcuts-macos.pdf).
- [Métodos abreviados de teclado para Windows](https://code.visualstudio.com/shortcuts/keyboard-shortcuts-windows.pdf).

## <a name="see-also"></a>Vea también

- [Referencia de API de scripts de Office](/javascript/api/office-scripts/overview)
- [Solución de problemas de scripts de Office](../testing/troubleshooting.md)
- [Usar objetos integrados de JavaScript en los scripts de Office](../develop/javascript-objects.md)
- [Visual Studio Code para scripts de Office (versión preliminar)](../develop/vscode-for-scripts.md)
