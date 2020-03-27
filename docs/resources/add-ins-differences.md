---
title: Diferencias entre scripts de Office y complementos de Office
description: El comportamiento y las diferencias de API entre scripts de Office y complementos de Office.
ms.date: 03/23/2020
localization_priority: Normal
ms.openlocfilehash: 2290d4e34b7a7286d67443de9e9c64bad4fcd4b7
ms.sourcegitcommit: d556aaefac80e55f53ac56b7f6ecbc657ebd426f
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 03/26/2020
ms.locfileid: "42978730"
---
# <a name="differences-between-office-scripts-and-office-add-ins"></a><span data-ttu-id="46cd1-103">Diferencias entre scripts de Office y complementos de Office</span><span class="sxs-lookup"><span data-stu-id="46cd1-103">Differences between Office Scripts and Office Add-ins</span></span>

<span data-ttu-id="46cd1-104">Los complementos de Office y los scripts de Office tienen mucho en común.</span><span class="sxs-lookup"><span data-stu-id="46cd1-104">Office Add-ins and Office Scripts have a lot in common.</span></span> <span data-ttu-id="46cd1-105">Ambos ofrecen el control automatizado de un libro de Excel `Excel` a través del espacio de nombres de la API de JavaScript de Office.</span><span class="sxs-lookup"><span data-stu-id="46cd1-105">They both offer automated control of an Excel workbook through the `Excel` namespace of the Office JavaScript API.</span></span> <span data-ttu-id="46cd1-106">Sin embargo, las secuencias de comandos de Office están más limitadas en su ámbito.</span><span class="sxs-lookup"><span data-stu-id="46cd1-106">However, Office Scripts are more limited in their scope.</span></span>

![Un diagrama de cuatro fases que muestra las áreas de enfoque para diferentes soluciones de extensibilidad de Office.](../images/office-programmability-diagram.png)

<span data-ttu-id="46cd1-109">Los scripts de Office se ejecutan hasta el final con una pulsación de botón manual o como un paso de la [automatización de energía](https://flow.microsoft.com/), mientras que los complementos de Office se conservan mientras los paneles de tareas están abiertos.</span><span class="sxs-lookup"><span data-stu-id="46cd1-109">Office Scripts run to completion with a manual button press or as a step in [Power Automate](https://flow.microsoft.com/), whereas Office Add-ins persist while their task panes are open.</span></span> <span data-ttu-id="46cd1-110">Esto significa que los complementos pueden mantener el estado durante una sesión, mientras que los scripts de Office no mantienen un estado interno entre ejecuciones.</span><span class="sxs-lookup"><span data-stu-id="46cd1-110">This means the add-ins can maintain state during a session, whereas Office Scripts do not maintain an internal state between runs.</span></span> <span data-ttu-id="46cd1-111">Si observa que su extensión de Excel debe superar las capacidades de la plataforma de scripting, visite la documentación de los complementos de [Office](/office/dev/add-ins) para obtener más información sobre los complementos de Office.</span><span class="sxs-lookup"><span data-stu-id="46cd1-111">If you find that your Excel extension needs to exceed the scripting platform's capabilities, visit the [Office Add-ins documentation](/office/dev/add-ins) to learn more about Office Add-ins.</span></span>

<span data-ttu-id="46cd1-112">En el resto de este artículo se describen las principales diferencias entre los complementos de Office y los scripts de Office.</span><span class="sxs-lookup"><span data-stu-id="46cd1-112">The rest of this article describes on the main differences between Office Add-ins and Office Scripts.</span></span>

## <a name="platform-support"></a><span data-ttu-id="46cd1-113">Compatibilidad con plataformas</span><span class="sxs-lookup"><span data-stu-id="46cd1-113">Platform Support</span></span>

<span data-ttu-id="46cd1-114">Los complementos de Office son para varias plataformas.</span><span class="sxs-lookup"><span data-stu-id="46cd1-114">Office Add-ins are cross-platform.</span></span> <span data-ttu-id="46cd1-115">Funcionan en plataformas de escritorio de Windows, Mac, iOS y Web y proporcionan la misma experiencia en cada uno de ellos.</span><span class="sxs-lookup"><span data-stu-id="46cd1-115">They work across Windows desktop, Mac, iOS, and web platforms and provide the same experience on each.</span></span> <span data-ttu-id="46cd1-116">Cualquier excepción a esto se indica en la documentación de la API individual.</span><span class="sxs-lookup"><span data-stu-id="46cd1-116">Any exception to this is noted in the documentation of the individual API.</span></span>

<span data-ttu-id="46cd1-117">Los scripts de Office solo están actualmente admitidos por para Excel en la Web.</span><span class="sxs-lookup"><span data-stu-id="46cd1-117">Office Scripts are currently only supported by for Excel on the web.</span></span> <span data-ttu-id="46cd1-118">Todas las operaciones de grabación, edición y ejecución se realizan en la plataforma Web.</span><span class="sxs-lookup"><span data-stu-id="46cd1-118">All recording, editing, and running is done on the web platform.</span></span>

## <a name="apis"></a><span data-ttu-id="46cd1-119">API</span><span class="sxs-lookup"><span data-stu-id="46cd1-119">APIs</span></span>

<span data-ttu-id="46cd1-120">Los scripts de Office admiten la mayoría de las API de JavaScript de Excel, lo que significa que hay mucha funcionalidad superpuesta entre las dos plataformas.</span><span class="sxs-lookup"><span data-stu-id="46cd1-120">Office Scripts support most of the Excel JavaScript APIs, which means there's  a lot of functionality overlap between the two platforms.</span></span> <span data-ttu-id="46cd1-121">Hay dos excepciones: Events y Common API.</span><span class="sxs-lookup"><span data-stu-id="46cd1-121">There are two exceptions: events and Common APIs.</span></span>

### <a name="events"></a><span data-ttu-id="46cd1-122">Eventos</span><span class="sxs-lookup"><span data-stu-id="46cd1-122">Events</span></span>

<span data-ttu-id="46cd1-123">Los scripts de Office no admiten [eventos](/office/dev/add-ins/excel/excel-add-ins-events).</span><span class="sxs-lookup"><span data-stu-id="46cd1-123">Office Scripts do not support [events](/office/dev/add-ins/excel/excel-add-ins-events).</span></span> <span data-ttu-id="46cd1-124">Cada secuencia de comandos ejecuta el código en `main` un solo método y, a continuación, finaliza.</span><span class="sxs-lookup"><span data-stu-id="46cd1-124">Every script runs the code in a single `main` method, then ends.</span></span> <span data-ttu-id="46cd1-125">No se reactiva cuando se desencadenan eventos y, por lo tanto, no pueden registrar los eventos.</span><span class="sxs-lookup"><span data-stu-id="46cd1-125">It does not reactivate when events are triggered, and thus, cannot register events.</span></span>

### <a name="common-apis"></a><span data-ttu-id="46cd1-126">API comunes</span><span class="sxs-lookup"><span data-stu-id="46cd1-126">Common APIs</span></span>

<span data-ttu-id="46cd1-127">Los scripts de Office no pueden usar [API comunes](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="46cd1-127">Office Scripts cannot use [Common APIs](/javascript/api/office).</span></span> <span data-ttu-id="46cd1-128">Si necesita la autenticación, ventanas de cuadro de diálogo u otras características que solo se admiten en las API comunes, es probable que deba crear un complemento de Office en lugar de un script de Office.</span><span class="sxs-lookup"><span data-stu-id="46cd1-128">If you need authentication, dialog windows, or other features that are only supported by Common APIs, you'll likely need to create an Office Add-in instead of an Office Script.</span></span>

## <a name="see-also"></a><span data-ttu-id="46cd1-129">Consulte también</span><span class="sxs-lookup"><span data-stu-id="46cd1-129">See also</span></span>

- [<span data-ttu-id="46cd1-130">Scripts de Office en Excel en la Web</span><span class="sxs-lookup"><span data-stu-id="46cd1-130">Office Scripts in Excel on the web</span></span>](../overview/excel.md)
- [<span data-ttu-id="46cd1-131">Diferencias entre scripts de Office y macros de VBA</span><span class="sxs-lookup"><span data-stu-id="46cd1-131">Differences between Office Scripts and VBA macros</span></span>](vba-differences.md)
- [<span data-ttu-id="46cd1-132">Solución de problemas de scripts de Office</span><span class="sxs-lookup"><span data-stu-id="46cd1-132">Troubleshooting Office Scripts</span></span>](../testing/troubleshooting.md)
- [<span data-ttu-id="46cd1-133">Crear un complemento de panel de tareas de Excel</span><span class="sxs-lookup"><span data-stu-id="46cd1-133">Build an Excel task pane add-in</span></span>](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)
