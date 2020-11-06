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
# <a name="platform-limits-and-requirements-with-office-scripts"></a><span data-ttu-id="881a5-103">Límites de plataforma y requisitos con scripts de Office</span><span class="sxs-lookup"><span data-stu-id="881a5-103">Platform limits and requirements with Office Scripts</span></span>

<span data-ttu-id="881a5-104">Hay algunas limitaciones de plataforma que debe tener en cuenta al desarrollar scripts de Office.</span><span class="sxs-lookup"><span data-stu-id="881a5-104">There are some platform limitations of which you should be aware when developing Office Scripts.</span></span> <span data-ttu-id="881a5-105">En este artículo se detalla la compatibilidad con exploradores y los límites de datos para los scripts de Office para Excel en la Web.</span><span class="sxs-lookup"><span data-stu-id="881a5-105">This article details the browser support and data limits for Office Scripts for Excel on the web.</span></span>

## <a name="browser-support"></a><span data-ttu-id="881a5-106">Compatibilidad con exploradores</span><span class="sxs-lookup"><span data-stu-id="881a5-106">Browser support</span></span>

<span data-ttu-id="881a5-107">Los scripts de Office funcionan en cualquier explorador que [admita Office para la web](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452).</span><span class="sxs-lookup"><span data-stu-id="881a5-107">Office Scripts work in any browser that [supports Office for the web](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452).</span></span> <span data-ttu-id="881a5-108">Sin embargo, algunas características de JavaScript no se admiten en Internet Explorer 11 (IE 11).</span><span class="sxs-lookup"><span data-stu-id="881a5-108">However, some JavaScript features aren't supported in Internet Explorer 11 (IE 11).</span></span> <span data-ttu-id="881a5-109">Las características que se incluyen en [ES6 o versiones posteriores](https://www.w3schools.com/Js/js_es6.asp) no funcionarán con IE 11.</span><span class="sxs-lookup"><span data-stu-id="881a5-109">Any features introduced in [ES6 or later](https://www.w3schools.com/Js/js_es6.asp) won't work with IE 11.</span></span> <span data-ttu-id="881a5-110">Si los usuarios de su organización todavía usan ese explorador, asegúrese de probar los scripts en ese entorno al compartirlos.</span><span class="sxs-lookup"><span data-stu-id="881a5-110">If people in your organization still use that browser, be sure to test your scripts in that environment when sharing them.</span></span>

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

### <a name="third-party-cookies"></a><span data-ttu-id="881a5-111">Cookies de terceros</span><span class="sxs-lookup"><span data-stu-id="881a5-111">Third-party cookies</span></span>

<span data-ttu-id="881a5-112">El explorador necesita las cookies de terceros habilitadas para mostrar la ficha **automatizar** en Excel en la Web.</span><span class="sxs-lookup"><span data-stu-id="881a5-112">Your browser needs third-party cookies enabled to show the **Automate** tab in Excel on the web.</span></span> <span data-ttu-id="881a5-113">Compruebe la configuración del explorador si no se muestra la pestaña.</span><span class="sxs-lookup"><span data-stu-id="881a5-113">Check your browser settings if the tab isn't being displayed.</span></span> <span data-ttu-id="881a5-114">Si está usando una sesión de explorador privada, es posible que tenga que volver a habilitar esta configuración cada vez.</span><span class="sxs-lookup"><span data-stu-id="881a5-114">If you're using a private browser session, you may need to re-enable this setting each time.</span></span>

> [!NOTE]
> <span data-ttu-id="881a5-115">Algunos exploradores hacen referencia a esta configuración como "todas las cookies", en lugar de "cookies de terceros".</span><span class="sxs-lookup"><span data-stu-id="881a5-115">Some browsers refer to this setting as "all cookies", instead of "third-party cookies".</span></span>

#### <a name="instructions-for-adjusting-cookie-settings-in-popular-browsers"></a><span data-ttu-id="881a5-116">Instrucciones para ajustar la configuración de cookies en exploradores populares</span><span class="sxs-lookup"><span data-stu-id="881a5-116">Instructions for adjusting cookie settings in popular browsers</span></span>

- [<span data-ttu-id="881a5-117">Chrome</span><span class="sxs-lookup"><span data-stu-id="881a5-117">Chrome</span></span>](https://support.google.com/chrome/answer/95647)
- [<span data-ttu-id="881a5-118">Microsoft Edge</span><span class="sxs-lookup"><span data-stu-id="881a5-118">Edge</span></span>](https://support.microsoft.com/microsoft-edge/temporarily-allow-cookies-and-site-data-in-microsoft-edge-597f04f2-c0ce-f08c-7c2b-541086362bd2)
- [<span data-ttu-id="881a5-119">Firefox</span><span class="sxs-lookup"><span data-stu-id="881a5-119">Firefox</span></span>](https://support.mozilla.org/kb/disable-third-party-cookies)
- [<span data-ttu-id="881a5-120">Safari</span><span class="sxs-lookup"><span data-stu-id="881a5-120">Safari</span></span>](https://support.apple.com/guide/safari/manage-cookies-and-website-data-sfri11471/mac)

## <a name="data-limits"></a><span data-ttu-id="881a5-121">Límites de datos</span><span class="sxs-lookup"><span data-stu-id="881a5-121">Data limits</span></span>

<span data-ttu-id="881a5-122">Hay límites en cuanto a la cantidad de datos de Excel que se pueden transferir a la vez y la cantidad de transacciones de automatización individuales que se pueden llevar a cabo.</span><span class="sxs-lookup"><span data-stu-id="881a5-122">There are limits on how much Excel data can be transferred at once and how many individual Power Automate transactions can be conducted.</span></span>

### <a name="excel"></a><span data-ttu-id="881a5-123">Excel</span><span class="sxs-lookup"><span data-stu-id="881a5-123">Excel</span></span>

<span data-ttu-id="881a5-124">Excel para la web tiene las siguientes limitaciones cuando se realizan llamadas al libro a través de un script:</span><span class="sxs-lookup"><span data-stu-id="881a5-124">Excel for the web has the following limitations when making calls to the workbook through a script:</span></span>

- <span data-ttu-id="881a5-125">Las solicitudes y respuestas se limitan a **5 MB**.</span><span class="sxs-lookup"><span data-stu-id="881a5-125">Requests and responses are limited to **5MB**.</span></span>
- <span data-ttu-id="881a5-126">Un rango está limitado a **5 millones celdas**.</span><span class="sxs-lookup"><span data-stu-id="881a5-126">A range is limited to **five million cells**.</span></span>

<span data-ttu-id="881a5-127">Si encuentra errores al tratar con conjuntos de valores de gran tamaño, pruebe a usar varios rangos más pequeños en lugar de rangos más grandes.</span><span class="sxs-lookup"><span data-stu-id="881a5-127">If you're encountering errors when dealing with large datasets, try using multiple smaller ranges instead of larger ranges.</span></span> <span data-ttu-id="881a5-128">También puede usar API como [Range. getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-) para destinar celdas específicas en lugar de rangos grandes.</span><span class="sxs-lookup"><span data-stu-id="881a5-128">You can also APIs like [Range.getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-) to target specific cells instead of large ranges.</span></span>

### <a name="power-automate"></a><span data-ttu-id="881a5-129">Power Automate</span><span class="sxs-lookup"><span data-stu-id="881a5-129">Power Automate</span></span>

<span data-ttu-id="881a5-130">Cuando se usan scripts de Office con la automatización de energía, se limita a **200 llamadas por día**.</span><span class="sxs-lookup"><span data-stu-id="881a5-130">When using Office Scripts with Power Automate, you're limited to **200 calls per day**.</span></span> <span data-ttu-id="881a5-131">Este límite se restablece a 12:00 A.M. UTC.</span><span class="sxs-lookup"><span data-stu-id="881a5-131">This limit resets at 12:00 AM UTC.</span></span>

<span data-ttu-id="881a5-132">La plataforma de automatización de energía también tiene limitaciones de uso, que se pueden encontrar en los límites de artículo [y en la configuración de la automatización de la energía](/power-automate/limits-and-config).</span><span class="sxs-lookup"><span data-stu-id="881a5-132">The Power Automate platform also has usage limitations, which can be found in the article [Limits and configuration in Power Automate](/power-automate/limits-and-config).</span></span>

## <a name="see-also"></a><span data-ttu-id="881a5-133">Vea también</span><span class="sxs-lookup"><span data-stu-id="881a5-133">See also</span></span>

- [<span data-ttu-id="881a5-134">Solución de problemas de scripts de Office</span><span class="sxs-lookup"><span data-stu-id="881a5-134">Troubleshooting Office Scripts</span></span>](troubleshooting.md)
- [<span data-ttu-id="881a5-135">Deshacer los efectos de un script de Office</span><span class="sxs-lookup"><span data-stu-id="881a5-135">Undo the effects of an Office Script</span></span>](undo.md)
- [<span data-ttu-id="881a5-136">Mejorar el rendimiento de los scripts de Office</span><span class="sxs-lookup"><span data-stu-id="881a5-136">Improve the performance of your Office Scripts</span></span>](../develop/web-client-performance.md)
- [<span data-ttu-id="881a5-137">Conceptos básicos sobre el scripting de los scripts de Office en Excel en la web</span><span class="sxs-lookup"><span data-stu-id="881a5-137">Scripting Fundamentals for Office Scripts in Excel on the web</span></span>](../develop/scripting-fundamentals.md)
