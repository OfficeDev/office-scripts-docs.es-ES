---
title: Límites y requisitos de plataforma con Office scripts
description: Límites de recursos y compatibilidad del explorador para Office scripts cuando se usan con Excel en la Web
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 7e81aaf2f96faeb67c815814fe3b7f1795651318
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545584"
---
# <a name="platform-limits-and-requirements-with-office-scripts"></a><span data-ttu-id="c3623-103">Límites y requisitos de plataforma con Office scripts</span><span class="sxs-lookup"><span data-stu-id="c3623-103">Platform limits and requirements with Office Scripts</span></span>

<span data-ttu-id="c3623-104">Hay algunas limitaciones de plataforma de las que debe tener en cuenta al desarrollar scripts Office.</span><span class="sxs-lookup"><span data-stu-id="c3623-104">There are some platform limitations of which you should be aware when developing Office Scripts.</span></span> <span data-ttu-id="c3623-105">En este artículo se detalla la compatibilidad del explorador y los límites de datos para Office scripts para Excel en la Web.</span><span class="sxs-lookup"><span data-stu-id="c3623-105">This article details the browser support and data limits for Office Scripts for Excel on the web.</span></span>

## <a name="browser-support"></a><span data-ttu-id="c3623-106">Compatibilidad con exploradores</span><span class="sxs-lookup"><span data-stu-id="c3623-106">Browser support</span></span>

<span data-ttu-id="c3623-107">Office Los scripts funcionan en cualquier explorador que [admita Office para la web](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452).</span><span class="sxs-lookup"><span data-stu-id="c3623-107">Office Scripts work in any browser that [supports Office for the web](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452).</span></span> <span data-ttu-id="c3623-108">Sin embargo, algunas características de JavaScript no son compatibles con Internet Explorer 11 (IE 11).</span><span class="sxs-lookup"><span data-stu-id="c3623-108">However, some JavaScript features aren't supported in Internet Explorer 11 (IE 11).</span></span> <span data-ttu-id="c3623-109">Las características introducidas en [ES6 o versiones posteriores](https://www.w3schools.com/Js/js_es6.asp) no funcionarán con IE 11.</span><span class="sxs-lookup"><span data-stu-id="c3623-109">Any features introduced in [ES6 or later](https://www.w3schools.com/Js/js_es6.asp) won't work with IE 11.</span></span> <span data-ttu-id="c3623-110">Si los usuarios de la organización siguen utilizando ese explorador, asegúrese de probar los scripts en ese entorno al compartirlos.</span><span class="sxs-lookup"><span data-stu-id="c3623-110">If people in your organization still use that browser, be sure to test your scripts in that environment when sharing them.</span></span>

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

### <a name="third-party-cookies"></a><span data-ttu-id="c3623-111">Cookies de terceros</span><span class="sxs-lookup"><span data-stu-id="c3623-111">Third-party cookies</span></span>

<span data-ttu-id="c3623-112">El explorador necesita cookies de terceros habilitadas para mostrar la pestaña **Automatizar** en Excel en la Web.</span><span class="sxs-lookup"><span data-stu-id="c3623-112">Your browser needs third-party cookies enabled to show the **Automate** tab in Excel on the web.</span></span> <span data-ttu-id="c3623-113">Comprueba la configuración del explorador si no se muestra la pestaña.</span><span class="sxs-lookup"><span data-stu-id="c3623-113">Check your browser settings if the tab isn't being displayed.</span></span> <span data-ttu-id="c3623-114">Si usa una sesión de explorador privado, es posible que deba volver a habilitar esta configuración cada vez.</span><span class="sxs-lookup"><span data-stu-id="c3623-114">If you're using a private browser session, you may need to re-enable this setting each time.</span></span>

> [!NOTE]
> <span data-ttu-id="c3623-115">Algunos exploradores se refieren a esta configuración como "todas las cookies", en lugar de "cookies de terceros".</span><span class="sxs-lookup"><span data-stu-id="c3623-115">Some browsers refer to this setting as "all cookies", instead of "third-party cookies".</span></span>

#### <a name="instructions-for-adjusting-cookie-settings-in-popular-browsers"></a><span data-ttu-id="c3623-116">Instrucciones para ajustar la configuración de cookies en exploradores populares</span><span class="sxs-lookup"><span data-stu-id="c3623-116">Instructions for adjusting cookie settings in popular browsers</span></span>

- [<span data-ttu-id="c3623-117">Chrome</span><span class="sxs-lookup"><span data-stu-id="c3623-117">Chrome</span></span>](https://support.google.com/chrome/answer/95647)
- [<span data-ttu-id="c3623-118">Microsoft Edge</span><span class="sxs-lookup"><span data-stu-id="c3623-118">Edge</span></span>](https://support.microsoft.com/microsoft-edge/temporarily-allow-cookies-and-site-data-in-microsoft-edge-597f04f2-c0ce-f08c-7c2b-541086362bd2)
- [<span data-ttu-id="c3623-119">Firefox</span><span class="sxs-lookup"><span data-stu-id="c3623-119">Firefox</span></span>](https://support.mozilla.org/kb/disable-third-party-cookies)
- [<span data-ttu-id="c3623-120">Safari</span><span class="sxs-lookup"><span data-stu-id="c3623-120">Safari</span></span>](https://support.apple.com/guide/safari/manage-cookies-and-website-data-sfri11471/mac)

## <a name="data-limits"></a><span data-ttu-id="c3623-121">Límites de datos</span><span class="sxs-lookup"><span data-stu-id="c3623-121">Data limits</span></span>

<span data-ttu-id="c3623-122">Hay límites en la cantidad Excel datos se pueden transferir a la vez y cuántas transacciones individuales Power Automate pueden realizarse.</span><span class="sxs-lookup"><span data-stu-id="c3623-122">There are limits on how much Excel data can be transferred at once and how many individual Power Automate transactions can be conducted.</span></span>

### <a name="excel"></a><span data-ttu-id="c3623-123">Excel</span><span class="sxs-lookup"><span data-stu-id="c3623-123">Excel</span></span>

<span data-ttu-id="c3623-124">Excel para la web tiene las siguientes limitaciones al realizar llamadas al libro a través de un script:</span><span class="sxs-lookup"><span data-stu-id="c3623-124">Excel for the web has the following limitations when making calls to the workbook through a script:</span></span>

- <span data-ttu-id="c3623-125">Las solicitudes y respuestas están limitadas a **5 MB.**</span><span class="sxs-lookup"><span data-stu-id="c3623-125">Requests and responses are limited to **5MB**.</span></span>
- <span data-ttu-id="c3623-126">Un rango está limitado a **cinco millones de celdas.**</span><span class="sxs-lookup"><span data-stu-id="c3623-126">A range is limited to **five million cells**.</span></span>

<span data-ttu-id="c3623-127">Si encuentra errores al tratar con conjuntos de datos grandes, intente usar varios intervalos más pequeños en lugar de intervalos más grandes.</span><span class="sxs-lookup"><span data-stu-id="c3623-127">If you're encountering errors when dealing with large datasets, try using multiple smaller ranges instead of larger ranges.</span></span> <span data-ttu-id="c3623-128">Para obtener un ejemplo, vea [el ejemplo Escribir un conjunto de datos](../resources/samples/write-large-dataset.md) grande.</span><span class="sxs-lookup"><span data-stu-id="c3623-128">For an example, see the [Write a large dataset](../resources/samples/write-large-dataset.md) sample.</span></span> <span data-ttu-id="c3623-129">También puede usar API como [Range.getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-) para seleccionar celdas específicas en lugar de intervalos grandes.</span><span class="sxs-lookup"><span data-stu-id="c3623-129">You can also use APIs like [Range.getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-) to target specific cells instead of large ranges.</span></span>

### <a name="power-automate"></a><span data-ttu-id="c3623-130">Power Automate</span><span class="sxs-lookup"><span data-stu-id="c3623-130">Power Automate</span></span>

<span data-ttu-id="c3623-131">Al usar Office scripts con Power Automate, cada usuario está limitado a **400** llamadas a la acción Ejecutar script por día .</span><span class="sxs-lookup"><span data-stu-id="c3623-131">When using Office Scripts with Power Automate, each user is limited to **400 calls to the Run Script action per day**.</span></span> <span data-ttu-id="c3623-132">Este límite se restablece a las 12:00 UTC.</span><span class="sxs-lookup"><span data-stu-id="c3623-132">This limit resets at 12:00 AM UTC.</span></span>

<span data-ttu-id="c3623-133">La Power Automate también tiene limitaciones de uso, que se pueden encontrar en los siguientes artículos:</span><span class="sxs-lookup"><span data-stu-id="c3623-133">The Power Automate platform also has usage limitations, which can be found in the following articles:</span></span>

- [<span data-ttu-id="c3623-134">Límites y configuración en Power Automate</span><span class="sxs-lookup"><span data-stu-id="c3623-134">Limits and configuration in Power Automate</span></span>](/power-automate/limits-and-config)
- [<span data-ttu-id="c3623-135">Problemas y limitaciones conocidos para el conector Excel Online (Empresa)</span><span class="sxs-lookup"><span data-stu-id="c3623-135">Known issues and limitations for the Excel Online (Business) connector</span></span>](/connectors/excelonlinebusiness/#known-issues-and-limitations)

## <a name="see-also"></a><span data-ttu-id="c3623-136">Consulte también</span><span class="sxs-lookup"><span data-stu-id="c3623-136">See also</span></span>

- [<span data-ttu-id="c3623-137">Solucionar problemas Office scripts</span><span class="sxs-lookup"><span data-stu-id="c3623-137">Troubleshoot Office Scripts</span></span>](troubleshooting.md)
- [<span data-ttu-id="c3623-138">Deshacer los efectos de Scripts de Office</span><span class="sxs-lookup"><span data-stu-id="c3623-138">Undo the effects of Office Scripts</span></span>](undo.md)
- [<span data-ttu-id="c3623-139">Mejorar el rendimiento de los scripts Office scripts</span><span class="sxs-lookup"><span data-stu-id="c3623-139">Improve the performance of your Office Scripts</span></span>](../develop/web-client-performance.md)
- [<span data-ttu-id="c3623-140">Scripting Fundamentals for Office Scripts in Excel en la Web</span><span class="sxs-lookup"><span data-stu-id="c3623-140">Scripting Fundamentals for Office Scripts in Excel on the web</span></span>](../develop/scripting-fundamentals.md)
