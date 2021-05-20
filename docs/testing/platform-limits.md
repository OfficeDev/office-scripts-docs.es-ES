---
title: Límites y requisitos de la plataforma con scripts de Office
description: Límites de recursos y compatibilidad con navegadores para scripts de Office cuando se usan con Excel en la Web
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 7e81aaf2f96faeb67c815814fe3b7f1795651318
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545584"
---
# <a name="platform-limits-and-requirements-with-office-scripts"></a>Límites y requisitos de la plataforma con scripts de Office

Hay algunas limitaciones de plataforma de las que debe tener en cuenta al desarrollar scripts de Office. En este artículo se detallan la compatibilidad con el explorador y los límites de datos de Office scripts para Excel en la Web.

## <a name="browser-support"></a>Compatibilidad con exploradores

Office Los scripts funcionan en cualquier navegador que [admita Office para la web.](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452) Sin embargo, algunas características de JavaScript no son compatibles con Internet Explorer 11 (IE 11). Las funciones introducidas en [ES6 o posterior](https://www.w3schools.com/Js/js_es6.asp) no funcionarán con IE 11. Si las personas de su organización siguen usando ese explorador, asegúrese de probar los scripts en ese entorno al compartirlos.

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

### <a name="third-party-cookies"></a>Cookies de terceros

Su navegador necesita cookies de terceros habilitadas para mostrar la pestaña **Automatizar** en Excel en la Web. Compruebe la configuración del navegador si no se muestra la pestaña. Si usa una sesión de explorador privada, es posible que deba volver a habilitar esta configuración cada vez.

> [!NOTE]
> Algunos navegadores se refieren a esta configuración como "todas las cookies", en lugar de "cookies de terceros".

#### <a name="instructions-for-adjusting-cookie-settings-in-popular-browsers"></a>Instrucciones para ajustar la configuración de cookies en navegadores populares

- [Chrome](https://support.google.com/chrome/answer/95647)
- [Microsoft Edge](https://support.microsoft.com/microsoft-edge/temporarily-allow-cookies-and-site-data-in-microsoft-edge-597f04f2-c0ce-f08c-7c2b-541086362bd2)
- [Firefox](https://support.mozilla.org/kb/disable-third-party-cookies)
- [Safari](https://support.apple.com/guide/safari/manage-cookies-and-website-data-sfri11471/mac)

## <a name="data-limits"></a>Límites de datos

Hay límites en la cantidad Excel datos que se pueden transferir a la vez y cuántas transacciones de Power Automate individuales se pueden realizar.

### <a name="excel"></a>Excel

Excel para la web tiene las siguientes limitaciones al realizar llamadas al libro a través de un script:

- Las solicitudes y respuestas están limitadas a **5 MB.**
- Un rango se limita a **cinco millones de celdas.**

Si se producen errores al tratar con conjuntos de datos grandes, intente usar varios rangos más pequeños en lugar de intervalos más grandes. Para obtener un ejemplo, vea el ejemplo [Escribir un conjunto de datos grande.](../resources/samples/write-large-dataset.md) También puede usar API como [Range.getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-) para dirigirse a celdas específicas en lugar de rangos grandes.

### <a name="power-automate"></a>Power Automate

Cuando se utilizan scripts de Office con Power Automate, cada usuario está limitado a **400 llamadas a la acción Ejecutar script por día.** Este límite se restablece a las 12:00 AM UTC.

La plataforma Power Automate también tiene limitaciones de uso, que se pueden encontrar en los siguientes artículos:

- [Límites y configuración en Power Automate](/power-automate/limits-and-config)
- [Problemas y limitaciones conocidos para el conector de Excel En línea (Business)](/connectors/excelonlinebusiness/#known-issues-and-limitations)

## <a name="see-also"></a>Vea también

- [Solución de problemas de scripts Office](troubleshooting.md)
- [Deshacer los efectos de Scripts de Office](undo.md)
- [Mejore el rendimiento de sus scripts de Office](../develop/web-client-performance.md)
- [Fundamentos de scripting para scripts de Office en Excel en la Web](../develop/scripting-fundamentals.md)
