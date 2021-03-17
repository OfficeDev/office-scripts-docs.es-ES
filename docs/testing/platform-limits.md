---
title: Límites y requisitos de la plataforma con scripts de Office
description: Límites de recursos y compatibilidad con exploradores para scripts de Office cuando se usan con Excel en la web
ms.date: 03/12/2021
localization_priority: Normal
ms.openlocfilehash: 93307b6204f409f26c77b5ead33188205d5c4b4d
ms.sourcegitcommit: 5bde455b06ee2ed007f3e462d8ad485b257774ef
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 03/17/2021
ms.locfileid: "50837268"
---
# <a name="platform-limits-and-requirements-with-office-scripts"></a>Límites y requisitos de la plataforma con scripts de Office

Existen algunas limitaciones de plataforma de las que debe tener en cuenta al desarrollar scripts de Office. En este artículo se detalla la compatibilidad del explorador y los límites de datos para scripts de Office para Excel en la web.

## <a name="browser-support"></a>Compatibilidad con exploradores

Los scripts de Office funcionan en cualquier explorador [que admita Office para la web](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452). Sin embargo, algunas características de JavaScript no son compatibles con Internet Explorer 11 (IE 11). Las características introducidas en [ES6 o versiones posteriores](https://www.w3schools.com/Js/js_es6.asp) no funcionarán con IE 11. Si los usuarios de la organización siguen utilizando ese explorador, asegúrese de probar los scripts en ese entorno al compartirlos.

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

### <a name="third-party-cookies"></a>Cookies de terceros

El explorador necesita cookies de terceros habilitadas para mostrar la pestaña **Automatizar** en Excel en la web. Comprueba la configuración del explorador si no se muestra la pestaña. Si usa una sesión de explorador privado, es posible que deba volver a habilitar esta configuración cada vez.

> [!NOTE]
> Algunos exploradores se refieren a esta configuración como "todas las cookies", en lugar de "cookies de terceros".

#### <a name="instructions-for-adjusting-cookie-settings-in-popular-browsers"></a>Instrucciones para ajustar la configuración de cookies en exploradores populares

- [Chrome](https://support.google.com/chrome/answer/95647)
- [Microsoft Edge](https://support.microsoft.com/microsoft-edge/temporarily-allow-cookies-and-site-data-in-microsoft-edge-597f04f2-c0ce-f08c-7c2b-541086362bd2)
- [Firefox](https://support.mozilla.org/kb/disable-third-party-cookies)
- [Safari](https://support.apple.com/guide/safari/manage-cookies-and-website-data-sfri11471/mac)

## <a name="data-limits"></a>Límites de datos

Hay límites en la cantidad de datos de Excel que se pueden transferir a la vez y cuántas transacciones individuales de Power Automate se pueden llevar a cabo.

### <a name="excel"></a>Excel

Excel para la web tiene las siguientes limitaciones al realizar llamadas al libro a través de un script:

- Las solicitudes y respuestas están limitadas a **5 MB.**
- Un rango está limitado a **cinco millones de celdas.**

Si encuentra errores al tratar con conjuntos de datos grandes, intente usar varios intervalos más pequeños en lugar de intervalos más grandes. También puede api como [Range.getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-) para dirigir celdas específicas en lugar de rangos grandes.

### <a name="power-automate"></a>Power Automate

Al usar scripts de Office con Power Automate, cada usuario está limitado a **200 llamadas al día.** Este límite se restablece a las 12:00 UTC.

La plataforma Power Automate también tiene limitaciones de uso, que se pueden encontrar en los siguientes artículos:

- [Límites y configuración en Power Automate](/power-automate/limits-and-config)
- [Problemas y limitaciones conocidos para el conector de Excel Online (Empresa)](/connectors/excelonlinebusiness/#known-issues-and-limitations)

## <a name="see-also"></a>Consulte también

- [Solución de problemas de scripts de Office](troubleshooting.md)
- [Deshacer los efectos de un script de Office](undo.md)
- [Mejorar el rendimiento de los scripts de Office](../develop/web-client-performance.md)
- [Scripting Fundamentals for Office Scripts in Excel on the web](../develop/scripting-fundamentals.md)
