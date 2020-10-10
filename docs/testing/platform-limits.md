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
# <a name="platform-limits-and-requirements-with-office-scripts"></a><span data-ttu-id="8b9de-103">Límites de plataforma y requisitos con scripts de Office</span><span class="sxs-lookup"><span data-stu-id="8b9de-103">Platform limits and requirements with Office Scripts</span></span>

<span data-ttu-id="8b9de-104">Hay algunas limitaciones de plataforma que debe tener en cuenta al desarrollar scripts de Office.</span><span class="sxs-lookup"><span data-stu-id="8b9de-104">There are some platform limitations of which you should be aware when developing Office Scripts.</span></span> <span data-ttu-id="8b9de-105">En este artículo se detalla la compatibilidad con exploradores y los límites de datos para los scripts de Office para Excel en la Web.</span><span class="sxs-lookup"><span data-stu-id="8b9de-105">This article details the browser support and data limits for Office Scripts for Excel on the web.</span></span>

## <a name="browser-support"></a><span data-ttu-id="8b9de-106">Compatibilidad con exploradores</span><span class="sxs-lookup"><span data-stu-id="8b9de-106">Browser support</span></span>

<span data-ttu-id="8b9de-107">Los scripts de Office funcionan en cualquier explorador que [admita Office para la web](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452).</span><span class="sxs-lookup"><span data-stu-id="8b9de-107">Office Scripts work in any browser that [supports Office for the web](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452).</span></span> <span data-ttu-id="8b9de-108">Sin embargo, algunas características de JavaScript no se admiten en Internet Explorer 11 (IE 11).</span><span class="sxs-lookup"><span data-stu-id="8b9de-108">However, some JavaScript features aren't supported in Internet Explorer 11 (IE 11).</span></span> <span data-ttu-id="8b9de-109">Las características que se incluyen en [ES6 o versiones posteriores](https://www.w3schools.com/Js/js_es6.asp) no funcionarán con IE 11.</span><span class="sxs-lookup"><span data-stu-id="8b9de-109">Any features introduced in [ES6 or later](https://www.w3schools.com/Js/js_es6.asp) won't work with IE 11.</span></span> <span data-ttu-id="8b9de-110">Si los usuarios de su organización todavía usan ese explorador, asegúrese de probar los scripts en ese entorno al compartirlos.</span><span class="sxs-lookup"><span data-stu-id="8b9de-110">If people in your organization still use that browser, be sure to test your scripts in that environment when sharing them.</span></span>

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

### <a name="third-party-cookies"></a><span data-ttu-id="8b9de-111">Cookies de terceros</span><span class="sxs-lookup"><span data-stu-id="8b9de-111">Third-party cookies</span></span>

<span data-ttu-id="8b9de-112">El explorador necesita las cookies de terceros habilitadas para mostrar la ficha **automatizar** en Excel en la Web.</span><span class="sxs-lookup"><span data-stu-id="8b9de-112">Your browser needs third-party cookies enabled to show the **Automate** tab in Excel on the web.</span></span> <span data-ttu-id="8b9de-113">Compruebe la configuración del explorador si no se muestra la pestaña.</span><span class="sxs-lookup"><span data-stu-id="8b9de-113">Check your browser settings if the tab isn't being displayed.</span></span> <span data-ttu-id="8b9de-114">Si está usando una sesión de explorador privada, es posible que tenga que volver a habilitar esta configuración cada vez.</span><span class="sxs-lookup"><span data-stu-id="8b9de-114">If you're using a private browser session, you may need to re-enable this setting each time.</span></span>

> [!NOTE]
> <span data-ttu-id="8b9de-115">Algunos exploradores hacen referencia a esta configuración como "todas las cookies", en lugar de "cookies de terceros".</span><span class="sxs-lookup"><span data-stu-id="8b9de-115">Some browsers refer to this setting as "all cookies", instead of "third-party cookies".</span></span>

## <a name="data-limits"></a><span data-ttu-id="8b9de-116">Límites de datos</span><span class="sxs-lookup"><span data-stu-id="8b9de-116">Data limits</span></span>

<span data-ttu-id="8b9de-117">Hay límites en cuanto a la cantidad de datos de Excel que se pueden transferir a la vez y la cantidad de transacciones de automatización individuales que se pueden llevar a cabo.</span><span class="sxs-lookup"><span data-stu-id="8b9de-117">There are limits on how much Excel data can be transferred at once and how many individual Power Automate transactions can be conducted.</span></span>

### <a name="excel"></a><span data-ttu-id="8b9de-118">Excel</span><span class="sxs-lookup"><span data-stu-id="8b9de-118">Excel</span></span>

<span data-ttu-id="8b9de-119">Excel para la web tiene las siguientes limitaciones cuando se realizan llamadas al libro a través de un script:</span><span class="sxs-lookup"><span data-stu-id="8b9de-119">Excel for the web has the following limitations when making calls to the workbook through a script:</span></span>

- <span data-ttu-id="8b9de-120">Las solicitudes y respuestas se limitan a **5 MB**.</span><span class="sxs-lookup"><span data-stu-id="8b9de-120">Requests and responses are limited to **5MB**.</span></span>
- <span data-ttu-id="8b9de-121">Un rango está limitado a **5 millones celdas**.</span><span class="sxs-lookup"><span data-stu-id="8b9de-121">A range is limited to **five million cells**.</span></span>

<span data-ttu-id="8b9de-122">Si encuentra errores al tratar con conjuntos de valores de gran tamaño, pruebe a usar varios rangos más pequeños en lugar de rangos más grandes.</span><span class="sxs-lookup"><span data-stu-id="8b9de-122">If you're encountering errors when dealing with large datasets, try using multiple smaller ranges instead of larger ranges.</span></span> <span data-ttu-id="8b9de-123">También puede usar API como [Range. getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-) para destinar celdas específicas en lugar de rangos grandes.</span><span class="sxs-lookup"><span data-stu-id="8b9de-123">You can also APIs like [Range.getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-) to target specific cells instead of large ranges.</span></span>

### <a name="power-automate"></a><span data-ttu-id="8b9de-124">Power Automate</span><span class="sxs-lookup"><span data-stu-id="8b9de-124">Power Automate</span></span>

<span data-ttu-id="8b9de-125">Cuando se usan scripts de Office con la automatización de energía, se limita a **200 llamadas por día**.</span><span class="sxs-lookup"><span data-stu-id="8b9de-125">When using Office Scripts with Power Automate, you're limited to **200 calls per day**.</span></span> <span data-ttu-id="8b9de-126">Este límite se restablece a 12:00 A.M. UTC.</span><span class="sxs-lookup"><span data-stu-id="8b9de-126">This limit resets at 12:00 AM UTC.</span></span>

<span data-ttu-id="8b9de-127">La plataforma de automatización de energía también tiene limitaciones de uso, que se pueden encontrar en los límites de artículo [y en la configuración de la automatización de la energía](/power-automate/limits-and-config).</span><span class="sxs-lookup"><span data-stu-id="8b9de-127">The Power Automate platform also has usage limitations, which can be found in the article [Limits and configuration in Power Automate](/power-automate/limits-and-config).</span></span>

## <a name="see-also"></a><span data-ttu-id="8b9de-128">Ver también</span><span class="sxs-lookup"><span data-stu-id="8b9de-128">See also</span></span>

- [<span data-ttu-id="8b9de-129">Solución de problemas de scripts de Office</span><span class="sxs-lookup"><span data-stu-id="8b9de-129">Troubleshooting Office Scripts</span></span>](troubleshooting.md)
- [<span data-ttu-id="8b9de-130">Deshacer los efectos de un script de Office</span><span class="sxs-lookup"><span data-stu-id="8b9de-130">Undo the effects of an Office Script</span></span>](undo.md)
- [<span data-ttu-id="8b9de-131">Mejorar el rendimiento de los scripts de Office</span><span class="sxs-lookup"><span data-stu-id="8b9de-131">Improve the performance of your Office Scripts</span></span>](../develop/web-client-performance.md)
- [<span data-ttu-id="8b9de-132">Conceptos básicos sobre el scripting de los scripts de Office en Excel en la web</span><span class="sxs-lookup"><span data-stu-id="8b9de-132">Scripting Fundamentals for Office Scripts in Excel on the web</span></span>](../develop/scripting-fundamentals.md)
