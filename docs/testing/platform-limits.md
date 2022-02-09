---
title: Límites y requisitos de plataforma con Office scripts
description: Límites de recursos y compatibilidad del explorador para Office scripts cuando se usan con Excel en la Web.
ms.date: 01/31/2022
ms.localizationpriority: medium
ms.openlocfilehash: 385248e5c62ed3dbf2827105b3097ef27e5187a7
ms.sourcegitcommit: b84d4c8dd31335e4e39b0da6ad25fd528cb9d8f3
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 02/09/2022
ms.locfileid: "62462505"
---
# <a name="platform-limits-and-requirements-with-office-scripts"></a>Límites y requisitos de plataforma con Office scripts

Hay algunas limitaciones de plataforma de las que debe tener en cuenta al desarrollar scripts Office. En este artículo se detalla la compatibilidad del explorador y los límites de datos para Office scripts para Excel en la Web.

## <a name="browser-support"></a>Compatibilidad con exploradores

Office scripts funcionan en cualquier explorador que [admita Office para la Web](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452). Sin embargo, algunas características de JavaScript no son compatibles con Internet Explorer 11 (IE 11). Las características introducidas en [ES6 o versiones](https://www.w3schools.com/Js/js_es6.asp) posteriores no funcionarán con IE 11. Si los usuarios de la organización siguen utilizando ese explorador, asegúrese de probar los scripts en ese entorno al compartirlos.

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

### <a name="third-party-cookies"></a>Cookies de terceros

El explorador necesita cookies de terceros habilitadas para mostrar la pestaña **Automatizar** en Excel en la Web. Comprueba la configuración del explorador si no se muestra la pestaña. Si usa una sesión de explorador privado, es posible que deba volver a habilitar esta configuración cada vez.

> [!NOTE]
> Algunos exploradores se refieren a esta configuración como "todas las cookies", en lugar de "cookies de terceros".

#### <a name="instructions-for-adjusting-cookie-settings-in-popular-browsers"></a>Instrucciones para ajustar la configuración de cookies en exploradores populares

- [Chrome](https://support.google.com/chrome/answer/95647)
- [Microsoft Edge](https://support.microsoft.com/microsoft-edge/597f04f2-c0ce-f08c-7c2b-541086362bd2)
- [Firefox](https://support.mozilla.org/kb/disable-third-party-cookies)
- [Safari](https://support.apple.com/guide/safari/manage-cookies-and-website-data-sfri11471/mac)

## <a name="data-limits"></a>Límites de datos

Hay límites en la cantidad Excel datos se pueden transferir a la vez y cuántas transacciones individuales Power Automate pueden realizarse.

### <a name="excel"></a>Excel

Excel para la Web tiene las siguientes limitaciones al realizar llamadas al libro a través de un script:

- Las solicitudes y respuestas están limitadas a **5 MB**.
- Un rango está limitado a **cinco millones de celdas**.

Si encuentra errores al tratar con conjuntos de datos grandes, intente usar varios intervalos más pequeños en lugar de intervalos más grandes. Para obtener un ejemplo, vea [el ejemplo Escribir un conjunto de datos](../resources/samples/write-large-dataset.md) grande. También puede usar API como [Range.getSpecialCells para seleccionar celdas](/javascript/api/office-scripts/excelscript/excelscript.range#excelscript-excelscript-range-getspecialcells-member(1)) específicas en lugar de intervalos grandes.

### <a name="power-automate"></a>Power Automate

Al usar Office scripts con Power Automate, cada usuario está limitado a **1.600 llamadas a la acción Ejecutar script por día**. Este límite se restablece a las 12:00 UTC.

La Power Automate también tiene limitaciones de uso, que se pueden encontrar en los siguientes artículos.

- [Límites y configuración en Power Automate](/power-automate/limits-and-config)
- [Problemas y limitaciones conocidos para el conector Excel Online (Empresa)](/connectors/excelonlinebusiness/#known-issues-and-limitations)

> [!NOTE]
> Si tiene un script de larga ejecución, tenga en cuenta el tiempo de espera de [120 segundos para operaciones sincrónicas Power Automate ejecución](/power-automate/limits-and-config#timeout). Tendrás que optimizar el [script o](../develop/web-client-performance.md) dividir la automatización Excel en varios scripts.

## <a name="see-also"></a>Vea también

- [Solucionar problemas Office scripts](troubleshooting.md)
- [Deshacer los efectos de Scripts de Office](undo.md)
- [Mejorar el rendimiento de los scripts Office scripts](../develop/web-client-performance.md)
- [Scripting Fundamentals for Office Scripts in Excel en la Web](../develop/scripting-fundamentals.md)
