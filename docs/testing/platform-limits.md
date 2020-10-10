---
title: Límites de plataforma y requisitos con scripts de Office
description: Límites de recursos y compatibilidad con exploradores para scripts de Office cuando se usan con Excel en la web
ms.date: 10/09/2020
localization_priority: Normal
ms.openlocfilehash: df468192f443b912e26411e46c9f953e046e55ec
ms.sourcegitcommit: 42fa3b629c93930b4e73e9c4c01d0c8bdf6d7487
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 10/09/2020
ms.locfileid: "48411560"
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

## <a name="see-also"></a>Ver también

- [Solución de problemas de scripts de Office](troubleshooting.md)
- [Deshacer los efectos de un script de Office](undo.md)
- [Mejorar el rendimiento de los scripts de Office](../develop/web-client-performance.md)
- [Conceptos básicos sobre el scripting de los scripts de Office en Excel en la web](../develop/scripting-fundamentals.md)
