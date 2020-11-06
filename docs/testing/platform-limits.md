---
title: Límites de plataforma y requisitos con scripts de Office
description: Límites de recursos y compatibilidad con exploradores para scripts de Office cuando se usan con Excel en la web
ms.date: 10/23/2020
localization_priority: Normal
ms.openlocfilehash: 61f5c55be278ae056014d3b01e4176354d913f87
ms.sourcegitcommit: d3e7681e262bdccc281fcb7b3c719494202e846b
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 11/06/2020
ms.locfileid: "48930081"
---
# <a name="platform-limits-and-requirements-with-office-scripts"></a>Límites de plataforma y requisitos con scripts de Office

Hay algunas limitaciones de plataforma que debe tener en cuenta al desarrollar scripts de Office. En este artículo se detalla la compatibilidad con exploradores y los límites de datos para los scripts de Office para Excel en la Web.

## <a name="browser-support"></a>Compatibilidad con exploradores

Los scripts de Office funcionan en cualquier explorador que [admita Office para la web](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452). Sin embargo, algunas características de JavaScript no se admiten en Internet Explorer 11 (IE 11). Las características que se incluyen en [ES6 o versiones posteriores](https://www.w3schools.com/Js/js_es6.asp) no funcionarán con IE 11. Si los usuarios de su organización todavía usan ese explorador, asegúrese de probar los scripts en ese entorno al compartirlos.

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

### <a name="third-party-cookies"></a>Cookies de terceros

El explorador necesita las cookies de terceros habilitadas para mostrar la ficha **automatizar** en Excel en la Web. Compruebe la configuración del explorador si no se muestra la pestaña. Si está usando una sesión de explorador privada, es posible que tenga que volver a habilitar esta configuración cada vez.

> [!NOTE]
> Algunos exploradores hacen referencia a esta configuración como "todas las cookies", en lugar de "cookies de terceros".

#### <a name="instructions-for-adjusting-cookie-settings-in-popular-browsers"></a>Instrucciones para ajustar la configuración de cookies en exploradores populares

- [Chrome](https://support.google.com/chrome/answer/95647)
- [Microsoft Edge](https://support.microsoft.com/microsoft-edge/temporarily-allow-cookies-and-site-data-in-microsoft-edge-597f04f2-c0ce-f08c-7c2b-541086362bd2)
- [Firefox](https://support.mozilla.org/kb/disable-third-party-cookies)
- [Safari](https://support.apple.com/guide/safari/manage-cookies-and-website-data-sfri11471/mac)

## <a name="data-limits"></a>Límites de datos

Hay límites en cuanto a la cantidad de datos de Excel que se pueden transferir a la vez y la cantidad de transacciones de automatización individuales que se pueden llevar a cabo.

### <a name="excel"></a>Excel

Excel para la web tiene las siguientes limitaciones cuando se realizan llamadas al libro a través de un script:

- Las solicitudes y respuestas se limitan a **5 MB**.
- Un rango está limitado a **5 millones celdas**.

Si encuentra errores al tratar con conjuntos de valores de gran tamaño, pruebe a usar varios rangos más pequeños en lugar de rangos más grandes. También puede usar API como [Range. getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-) para destinar celdas específicas en lugar de rangos grandes.

### <a name="power-automate"></a>Power Automate

Cuando se usan scripts de Office con la automatización de energía, se limita a **200 llamadas por día**. Este límite se restablece a 12:00 A.M. UTC.

La plataforma de automatización de energía también tiene limitaciones de uso, que se pueden encontrar en los límites de artículo [y en la configuración de la automatización de la energía](/power-automate/limits-and-config).

## <a name="see-also"></a>Vea también

- [Solución de problemas de scripts de Office](troubleshooting.md)
- [Deshacer los efectos de un script de Office](undo.md)
- [Mejorar el rendimiento de los scripts de Office](../develop/web-client-performance.md)
- [Conceptos básicos sobre el scripting de los scripts de Office en Excel en la web](../develop/scripting-fundamentals.md)
