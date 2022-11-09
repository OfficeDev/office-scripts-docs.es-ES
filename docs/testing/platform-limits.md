---
title: Límites y requisitos de la plataforma con scripts de Office
description: Límites de recursos y compatibilidad del explorador con scripts de Office cuando se usan con Excel en la Web.
ms.date: 11/07/2022
ms.localizationpriority: medium
ms.openlocfilehash: 764d1eddaf303a941a098ec1d3f3056d63e8693f
ms.sourcegitcommit: 7cadf2b637bf62874e43b6e595286101816662aa
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 11/09/2022
ms.locfileid: "68891249"
---
# <a name="platform-limits-and-requirements-with-office-scripts"></a>Límites y requisitos de la plataforma con scripts de Office

Hay algunas limitaciones de la plataforma que debe tener en cuenta al desarrollar scripts de Office. En este artículo se detallan los límites de datos y compatibilidad del explorador para scripts de Office para Excel en la Web.

## <a name="browser-support"></a>Compatibilidad con exploradores

Los scripts de Office funcionan en cualquier explorador que [admita Office para la Web](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452). Sin embargo, algunas características de JavaScript no se admiten en Internet Explorer 11 (IE 11). Las características introducidas en [ES6 o versiones posteriores](https://www.w3schools.com/Js/js_es6.asp) no funcionarán con IE 11. Si los usuarios de su organización siguen usando ese explorador, asegúrese de probar los scripts en ese entorno al compartirlos.

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

### <a name="third-party-cookies"></a>Cookies de terceros

El explorador necesita cookies de terceros habilitadas para mostrar la pestaña **Automatizar** en Excel en la Web. Compruebe la configuración del explorador si no se muestra la pestaña. Si usa una sesión de explorador privado, es posible que tenga que volver a habilitar esta configuración cada vez.

> [!NOTE]
> Algunos exploradores hacen referencia a esta configuración como "todas las cookies", en lugar de "cookies de terceros".

#### <a name="instructions-for-adjusting-cookie-settings-in-popular-browsers"></a>Instrucciones para ajustar la configuración de cookies en exploradores populares

- [Chrome](https://support.google.com/chrome/answer/95647)
- [Microsoft Edge](https://support.microsoft.com/microsoft-edge/597f04f2-c0ce-f08c-7c2b-541086362bd2)
- [Firefox](https://support.mozilla.org/kb/disable-third-party-cookies)
- [Safari](https://support.apple.com/guide/safari/manage-cookies-and-website-data-sfri11471/mac)

## <a name="data-limits"></a>Límites de datos

Hay límites en cuanto a la cantidad de datos de Excel que se pueden transferir a la vez y cuántas transacciones individuales de Power Automate se pueden realizar.

### <a name="excel"></a>Excel

Excel para la Web tiene las siguientes limitaciones al realizar llamadas al libro a través de un script:

- Las solicitudes y respuestas están limitadas a **5 MB**.
- Un rango está limitado a **cinco millones de celdas**.

Si encuentra errores al tratar con conjuntos de datos grandes, intente usar varios intervalos más pequeños en lugar de intervalos más grandes. Para obtener un ejemplo, consulte el ejemplo [Escribir un conjunto de datos grande](../resources/samples/write-large-dataset.md) . También puede usar API como [Range.getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#excelscript-excelscript-range-getspecialcells-member(1)) para dirigirse a celdas específicas en lugar de rangos grandes.

Los límites de Excel que no son específicos de los scripts de Office se pueden encontrar en el artículo [Especificaciones y límites de Excel](https://support.microsoft.com/office/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3).

### <a name="power-automate"></a>Power Automate

Al usar scripts de Office con Power Automate, cada usuario se limita a **1.600 llamadas a la acción Ejecutar script al día**. Este límite se restablece a las 12:00 UTC.

La plataforma Power Automate también tiene limitaciones de uso, que se pueden encontrar en los artículos siguientes.

- [Límites y configuración en Power Automate](/power-automate/limits-and-config)
- [Problemas conocidos y limitaciones para el conector de Excel Online (empresa)](/connectors/excelonlinebusiness/#known-issues-and-limitations)

> [!NOTE]
> Si tiene un script de ejecución prolongada, tenga en cuenta el [tiempo de espera de 120 segundos para las operaciones sincrónicas de Power Automate](/power-automate/limits-and-config#timeout). Tendrá que [optimizar el script](../develop/web-client-performance.md) o dividir la automatización de Excel en varios scripts.

## <a name="see-also"></a>Vea también

- [Especificaciones y límites de Excel](https://support.microsoft.com/office/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3)
- [Solución de problemas de scripts de Office](troubleshooting.md)
- [Deshacer los efectos de Scripts de Office](undo.md)
- [Mejora del rendimiento de los scripts de Office](../develop/web-client-performance.md)
