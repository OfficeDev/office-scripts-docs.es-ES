---
title: Office Entorno editor de código de scripts
description: Los requisitos previos y la información del entorno para Office scripts en Excel en la Web.
ms.date: 05/27/2021
localization_priority: Normal
ms.openlocfilehash: 5b2f7afa193dc71e13a3d6763c9e8ff8344ee3e8be18e7e996f8431e03510509
ms.sourcegitcommit: 75f7ed8c2d23a104acc293f8ce29ea580b4fcdc5
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 08/11/2021
ms.locfileid: "57847096"
---
# <a name="office-scripts-code-editor-environment"></a>Office Entorno editor de código de scripts

Office Los scripts se escriben en TypeScript o JavaScript y usan las API de JavaScript de scripts de Office para interactuar con un libro Excel texto. El Editor de código se basa en Visual Studio Code, por lo que si has usado ese entorno antes, te sentirás como en casa.

## <a name="scripting-language-typescript-or-javascript"></a>Lenguaje de scripting: TypeScript o JavaScript

Los Scripts de Office se escriben en [TypeScript](https://www.typescriptlang.org/docs/home.html), que es un superconjunto de [JavaScript](https://developer.mozilla.org/docs/Web/JavaScript). La grabadora de acciones genera código en TypeScript y la documentación Office scripts de texto usa TypeScript. Dado que TypeScript es un superconjunto de JavaScript, cualquier código de scripting que escriba en JavaScript funcionará bien.

Office Los scripts son en gran medida fragmentos de código autocontenido. Solo se usa una pequeña parte de la funcionalidad de TypeScript. Por lo tanto, puede editar scripts sin tener que aprender los entresijos de TypeScript. El Editor de código también controla la instalación, compilación y ejecución de código, por lo que no tiene que preocuparse por nada más que el script en sí. Es posible aprender el idioma y crear scripts sin conocimientos de programación anteriores. Sin embargo, si es nuevo en la programación, se recomienda aprender algunos aspectos básicos antes de continuar con Office scripts:

[!INCLUDE [Recommended coding resources](../includes/coding-basics-references.md)]

## <a name="office-scripts-javascript-api"></a>Office Scripts JavaScript API

Office Los scripts usan una versión especializada de Office API de JavaScript para [Office complementos](/office/dev/add-ins/overview/index). Aunque hay similitudes en las dos API, no debe suponer que el código se puede porte entre las dos plataformas. Las diferencias entre las dos plataformas se describen en el artículo Diferencias entre Office scripts y [Office complementos.](../resources/add-ins-differences.md#apis) Puede ver todas las API disponibles para el script en la documentación de referencia de la API Office [scripts.](/javascript/api/office-scripts/overview)

## <a name="external-library-support"></a>Compatibilidad con bibliotecas externas

Office Los scripts no admiten el uso de bibliotecas de JavaScript externas de terceros. Actualmente, no puede llamar a ninguna biblioteca que no sea Office API de scripts desde un script. Todavía tiene acceso a cualquier objeto [JavaScript integrado,](../develop/javascript-objects.md)como [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math).

## <a name="intellisense"></a>IntelliSense

IntelliSense es un conjunto de características del Editor de código que le ayudan a escribir código. Proporciona autocompletar, resaltado de errores de sintaxis y documentación de API en línea.

IntelliSense sugerencias a medida que escribe, de forma similar al texto sugerido en Excel. Al presionar la tecla Tab o Entrar, se inserta el miembro sugerido. Desencadenador IntelliSense en la ubicación actual del cursor presionando las teclas Ctrl+Espacio. Estas sugerencias son especialmente útiles al completar un método. La firma del método mostrada por IntelliSense contiene una lista de argumentos que necesita, el tipo de cada argumento, si un argumento determinado es obligatorio u opcional, y el tipo devuelto del método.

Mantenga el cursor sobre un método, clase u otro objeto de código para ver más información. Mantenga el puntero sobre un error de sintaxis o una sugerencia de código, representada por una línea roja o amarilla, para ver sugerencias sobre cómo solucionar el problema. A menudo, IntelliSense una opción de "Corrección rápida" para cambiar automáticamente el código.

:::image type="content" source="../images/implicit-any-editor-message.png" alt-text="Un mensaje de error en el texto activa del Editor de código con el botón &quot;Corrección rápida&quot;.":::

El editor Office de código de scripts usa el mismo motor de IntelliSense que Visual Studio Code. Para obtener más información acerca de la característica, visite [Visual Studio Code de IntelliSense características](https://code.visualstudio.com/docs/editor/intellisense#_intellisense-features).

## <a name="keyboard-shortcuts"></a>Accesos rápidos de teclado

La mayoría de los métodos abreviados de teclado Visual Studio Code también funcionan en el Editor de código Office scripts. Use los siguientes ARCHIVOS PDF para obtener información sobre las opciones disponibles y sacar el máximo partido del Editor de código:

- [Métodos abreviados de teclado para macOS](https://code.visualstudio.com/shortcuts/keyboard-shortcuts-macos.pdf).
- [Métodos abreviados de teclado Windows](https://code.visualstudio.com/shortcuts/keyboard-shortcuts-windows.pdf).

## <a name="see-also"></a>Vea también

- [Referencia de API de scripts de Office](/javascript/api/office-scripts/overview)
- [Solución de problemas de scripts de Office](../testing/troubleshooting.md)
- [Usar objetos integrados de JavaScript en los scripts de Office](../develop/javascript-objects.md)
