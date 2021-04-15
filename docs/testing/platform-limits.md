---
title: Límites y requisitos de la plataforma con scripts de Office
description: Límites de recursos y compatibilidad con exploradores para scripts de Office cuando se usan con Excel en la web
ms.date: 03/12/2021
localization_priority: Normal
ms.openlocfilehash: ef733562fb3caa8261fbbd8382923927a46cb7d4
ms.sourcegitcommit: 5ca286615a11d282e3f80023d22d36a039800eed
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 04/13/2021
ms.locfileid: "51689769"
---
# <a name="platform-limits-and-requirements-with-office-scripts"></a><span data-ttu-id="c29f4-103">Límites y requisitos de la plataforma con scripts de Office</span><span class="sxs-lookup"><span data-stu-id="c29f4-103">Platform limits and requirements with Office Scripts</span></span>

<span data-ttu-id="c29f4-104">Existen algunas limitaciones de plataforma de las que debe tener en cuenta al desarrollar scripts de Office.</span><span class="sxs-lookup"><span data-stu-id="c29f4-104">There are some platform limitations of which you should be aware when developing Office Scripts.</span></span> <span data-ttu-id="c29f4-105">En este artículo se detalla la compatibilidad del explorador y los límites de datos para scripts de Office para Excel en la web.</span><span class="sxs-lookup"><span data-stu-id="c29f4-105">This article details the browser support and data limits for Office Scripts for Excel on the web.</span></span>

## <a name="browser-support"></a><span data-ttu-id="c29f4-106">Compatibilidad con exploradores</span><span class="sxs-lookup"><span data-stu-id="c29f4-106">Browser support</span></span>

<span data-ttu-id="c29f4-107">Los scripts de Office funcionan en cualquier explorador [que admita Office para la web](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452).</span><span class="sxs-lookup"><span data-stu-id="c29f4-107">Office Scripts work in any browser that [supports Office for the web](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452).</span></span> <span data-ttu-id="c29f4-108">Sin embargo, algunas características de JavaScript no son compatibles con Internet Explorer 11 (IE 11).</span><span class="sxs-lookup"><span data-stu-id="c29f4-108">However, some JavaScript features aren't supported in Internet Explorer 11 (IE 11).</span></span> <span data-ttu-id="c29f4-109">Las características introducidas en [ES6 o versiones posteriores](https://www.w3schools.com/Js/js_es6.asp) no funcionarán con IE 11.</span><span class="sxs-lookup"><span data-stu-id="c29f4-109">Any features introduced in [ES6 or later](https://www.w3schools.com/Js/js_es6.asp) won't work with IE 11.</span></span> <span data-ttu-id="c29f4-110">Si los usuarios de la organización siguen utilizando ese explorador, asegúrese de probar los scripts en ese entorno al compartirlos.</span><span class="sxs-lookup"><span data-stu-id="c29f4-110">If people in your organization still use that browser, be sure to test your scripts in that environment when sharing them.</span></span>

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

### <a name="third-party-cookies"></a><span data-ttu-id="c29f4-111">Cookies de terceros</span><span class="sxs-lookup"><span data-stu-id="c29f4-111">Third-party cookies</span></span>

<span data-ttu-id="c29f4-112">El explorador necesita cookies de terceros habilitadas para mostrar la pestaña **Automatizar** en Excel en la web.</span><span class="sxs-lookup"><span data-stu-id="c29f4-112">Your browser needs third-party cookies enabled to show the **Automate** tab in Excel on the web.</span></span> <span data-ttu-id="c29f4-113">Comprueba la configuración del explorador si no se muestra la pestaña.</span><span class="sxs-lookup"><span data-stu-id="c29f4-113">Check your browser settings if the tab isn't being displayed.</span></span> <span data-ttu-id="c29f4-114">Si usa una sesión de explorador privado, es posible que deba volver a habilitar esta configuración cada vez.</span><span class="sxs-lookup"><span data-stu-id="c29f4-114">If you're using a private browser session, you may need to re-enable this setting each time.</span></span>

> [!NOTE]
> <span data-ttu-id="c29f4-115">Algunos exploradores se refieren a esta configuración como "todas las cookies", en lugar de "cookies de terceros".</span><span class="sxs-lookup"><span data-stu-id="c29f4-115">Some browsers refer to this setting as "all cookies", instead of "third-party cookies".</span></span>

#### <a name="instructions-for-adjusting-cookie-settings-in-popular-browsers"></a><span data-ttu-id="c29f4-116">Instrucciones para ajustar la configuración de cookies en exploradores populares</span><span class="sxs-lookup"><span data-stu-id="c29f4-116">Instructions for adjusting cookie settings in popular browsers</span></span>

- [<span data-ttu-id="c29f4-117">Chrome</span><span class="sxs-lookup"><span data-stu-id="c29f4-117">Chrome</span></span>](https://support.google.com/chrome/answer/95647)
- [<span data-ttu-id="c29f4-118">Microsoft Edge</span><span class="sxs-lookup"><span data-stu-id="c29f4-118">Edge</span></span>](https://support.microsoft.com/microsoft-edge/temporarily-allow-cookies-and-site-data-in-microsoft-edge-597f04f2-c0ce-f08c-7c2b-541086362bd2)
- [<span data-ttu-id="c29f4-119">Firefox</span><span class="sxs-lookup"><span data-stu-id="c29f4-119">Firefox</span></span>](https://support.mozilla.org/kb/disable-third-party-cookies)
- [<span data-ttu-id="c29f4-120">Safari</span><span class="sxs-lookup"><span data-stu-id="c29f4-120">Safari</span></span>](https://support.apple.com/guide/safari/manage-cookies-and-website-data-sfri11471/mac)

## <a name="data-limits"></a><span data-ttu-id="c29f4-121">Límites de datos</span><span class="sxs-lookup"><span data-stu-id="c29f4-121">Data limits</span></span>

<span data-ttu-id="c29f4-122">Hay límites en la cantidad de datos de Excel que se pueden transferir a la vez y cuántas transacciones individuales de Power Automate se pueden llevar a cabo.</span><span class="sxs-lookup"><span data-stu-id="c29f4-122">There are limits on how much Excel data can be transferred at once and how many individual Power Automate transactions can be conducted.</span></span>

### <a name="excel"></a><span data-ttu-id="c29f4-123">Excel</span><span class="sxs-lookup"><span data-stu-id="c29f4-123">Excel</span></span>

<span data-ttu-id="c29f4-124">Excel para la web tiene las siguientes limitaciones al realizar llamadas al libro a través de un script:</span><span class="sxs-lookup"><span data-stu-id="c29f4-124">Excel for the web has the following limitations when making calls to the workbook through a script:</span></span>

- <span data-ttu-id="c29f4-125">Las solicitudes y respuestas están limitadas a **5 MB.**</span><span class="sxs-lookup"><span data-stu-id="c29f4-125">Requests and responses are limited to **5MB**.</span></span>
- <span data-ttu-id="c29f4-126">Un rango está limitado a **cinco millones de celdas.**</span><span class="sxs-lookup"><span data-stu-id="c29f4-126">A range is limited to **five million cells**.</span></span>

<span data-ttu-id="c29f4-127">Si encuentra errores al tratar con conjuntos de datos grandes, intente usar varios intervalos más pequeños en lugar de intervalos más grandes.</span><span class="sxs-lookup"><span data-stu-id="c29f4-127">If you're encountering errors when dealing with large datasets, try using multiple smaller ranges instead of larger ranges.</span></span> <span data-ttu-id="c29f4-128">También puede api como [Range.getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-) para dirigir celdas específicas en lugar de rangos grandes.</span><span class="sxs-lookup"><span data-stu-id="c29f4-128">You can also APIs like [Range.getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-) to target specific cells instead of large ranges.</span></span>

### <a name="power-automate"></a><span data-ttu-id="c29f4-129">Power Automate</span><span class="sxs-lookup"><span data-stu-id="c29f4-129">Power Automate</span></span>

<span data-ttu-id="c29f4-130">Al usar scripts de Office con Power Automate, cada usuario está limitado a **400** llamadas a la acción Ejecutar script por día.</span><span class="sxs-lookup"><span data-stu-id="c29f4-130">When using Office Scripts with Power Automate, each user is limited to **400 calls to the Run Script action per day**.</span></span> <span data-ttu-id="c29f4-131">Este límite se restablece a las 12:00 UTC.</span><span class="sxs-lookup"><span data-stu-id="c29f4-131">This limit resets at 12:00 AM UTC.</span></span>

<span data-ttu-id="c29f4-132">La plataforma Power Automate también tiene limitaciones de uso, que se pueden encontrar en los siguientes artículos:</span><span class="sxs-lookup"><span data-stu-id="c29f4-132">The Power Automate platform also has usage limitations, which can be found in the following articles:</span></span>

- [<span data-ttu-id="c29f4-133">Límites y configuración en Power Automate</span><span class="sxs-lookup"><span data-stu-id="c29f4-133">Limits and configuration in Power Automate</span></span>](/power-automate/limits-and-config)
- [<span data-ttu-id="c29f4-134">Problemas y limitaciones conocidos para el conector de Excel Online (Empresa)</span><span class="sxs-lookup"><span data-stu-id="c29f4-134">Known issues and limitations for the Excel Online (Business) connector</span></span>](/connectors/excelonlinebusiness/#known-issues-and-limitations)

## <a name="see-also"></a><span data-ttu-id="c29f4-135">Consulte también</span><span class="sxs-lookup"><span data-stu-id="c29f4-135">See also</span></span>

- [<span data-ttu-id="c29f4-136">Solución de problemas de scripts de Office</span><span class="sxs-lookup"><span data-stu-id="c29f4-136">Troubleshooting Office Scripts</span></span>](troubleshooting.md)
- [<span data-ttu-id="c29f4-137">Deshacer los efectos de un script de Office</span><span class="sxs-lookup"><span data-stu-id="c29f4-137">Undo the effects of an Office Script</span></span>](undo.md)
- [<span data-ttu-id="c29f4-138">Mejorar el rendimiento de los scripts de Office</span><span class="sxs-lookup"><span data-stu-id="c29f4-138">Improve the performance of your Office Scripts</span></span>](../develop/web-client-performance.md)
- [<span data-ttu-id="c29f4-139">Scripting Fundamentals for Office Scripts in Excel on the web</span><span class="sxs-lookup"><span data-stu-id="c29f4-139">Scripting Fundamentals for Office Scripts in Excel on the web</span></span>](../develop/scripting-fundamentals.md)
