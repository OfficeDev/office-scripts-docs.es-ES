---
title: Diferencias entre scripts de Office y complementos de Office
description: El comportamiento y las diferencias de API entre scripts de Office y complementos de Office.
ms.date: 12/12/2019
localization_priority: Normal
ms.openlocfilehash: 4626afb66b54c94a72f29b039c601435c089d64d
ms.sourcegitcommit: b075eed5a6f275274fbbf6d62633219eac416f26
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 03/10/2020
ms.locfileid: "42700397"
---
# <a name="differences-between-office-scripts-and-office-add-ins"></a><span data-ttu-id="283ed-103">Diferencias entre scripts de Office y complementos de Office</span><span class="sxs-lookup"><span data-stu-id="283ed-103">Differences between Office Scripts and Office Add-ins</span></span>

<span data-ttu-id="283ed-104">Los complementos de Office y los scripts de Office tienen mucho en común.</span><span class="sxs-lookup"><span data-stu-id="283ed-104">Office Add-ins and Office Scripts have a lot in common.</span></span> <span data-ttu-id="283ed-105">Ambos ofrecen el control automatizado de un libro de Excel `Excel` a través del espacio de nombres de la API de JavaScript de Office.</span><span class="sxs-lookup"><span data-stu-id="283ed-105">They both offer automated control of an Excel workbook through the `Excel` namespace of the Office JavaScript API.</span></span> <span data-ttu-id="283ed-106">Sin embargo, las secuencias de comandos de Office están más limitadas en su ámbito.</span><span class="sxs-lookup"><span data-stu-id="283ed-106">However, Office Scripts are more limited in their scope.</span></span>

<span data-ttu-id="283ed-107">Los scripts de Office se ejecutan hasta el final con una pulsación de botón manual, mientras que los complementos de Office se basan en la interacción del usuario y se conservan mientras el libro está en uso.</span><span class="sxs-lookup"><span data-stu-id="283ed-107">Office Scripts run to completion with a manual button press, whereas Office Add-ins rely on user interaction and persist while the workbook is in use.</span></span> <span data-ttu-id="283ed-108">Si observa que su extensión de Excel debe superar las capacidades de la plataforma de scripting, visite la documentación de los complementos de [Office](/office/dev/add-ins) para obtener más información sobre los complementos de Office.</span><span class="sxs-lookup"><span data-stu-id="283ed-108">If you find that your Excel extension needs to exceed the scripting platform's capabilities, visit the [Office Add-ins documentation](/office/dev/add-ins) to learn more about Office Add-ins.</span></span>

<span data-ttu-id="283ed-109">En el resto de este artículo se describen las principales diferencias entre los complementos de Office y los scripts de Office.</span><span class="sxs-lookup"><span data-stu-id="283ed-109">The rest of this article describes on the main differences between Office Add-ins and Office Scripts.</span></span>

## <a name="platform-support"></a><span data-ttu-id="283ed-110">Compatibilidad con plataformas</span><span class="sxs-lookup"><span data-stu-id="283ed-110">Platform Support</span></span>

<span data-ttu-id="283ed-111">Los complementos de Office son para varias plataformas.</span><span class="sxs-lookup"><span data-stu-id="283ed-111">Office Add-ins are cross-platform.</span></span> <span data-ttu-id="283ed-112">Funcionan en plataformas de escritorio de Windows, Mac, iOS y Web y proporcionan la misma experiencia en cada uno de ellos.</span><span class="sxs-lookup"><span data-stu-id="283ed-112">They work across Windows desktop, Mac, iOS, and web platforms and provide the same experience on each.</span></span> <span data-ttu-id="283ed-113">Cualquier excepción a esto se indica en la documentación de la API individual.</span><span class="sxs-lookup"><span data-stu-id="283ed-113">Any exception to this is noted in the documentation of the individual API.</span></span>

<span data-ttu-id="283ed-114">Los scripts de Office solo están actualmente admitidos por para Excel en la Web.</span><span class="sxs-lookup"><span data-stu-id="283ed-114">Office Scripts are currently only supported by for Excel on the web.</span></span> <span data-ttu-id="283ed-115">Todas las operaciones de grabación, edición y ejecución se realizan en la plataforma Web.</span><span class="sxs-lookup"><span data-stu-id="283ed-115">All recording, editing, and running is done on the web platform.</span></span>

## <a name="apis"></a><span data-ttu-id="283ed-116">API</span><span class="sxs-lookup"><span data-stu-id="283ed-116">APIs</span></span>

<span data-ttu-id="283ed-117">Los scripts de Office admiten la mayoría de las API de JavaScript de Excel, lo que significa que hay mucha funcionalidad superpuesta entre las dos plataformas.</span><span class="sxs-lookup"><span data-stu-id="283ed-117">Office Scripts support most of the Excel JavaScript APIs, which means there's  a lot of functionality overlap between the two platforms.</span></span> <span data-ttu-id="283ed-118">Hay dos excepciones: Events y Common API.</span><span class="sxs-lookup"><span data-stu-id="283ed-118">There are two exceptions: events and Common APIs.</span></span>

### <a name="events"></a><span data-ttu-id="283ed-119">Eventos</span><span class="sxs-lookup"><span data-stu-id="283ed-119">Events</span></span>

<span data-ttu-id="283ed-120">Los scripts de Office no admiten [eventos](/office/dev/add-ins/excel/excel-add-ins-events).</span><span class="sxs-lookup"><span data-stu-id="283ed-120">Office Scripts do not support [events](/office/dev/add-ins/excel/excel-add-ins-events).</span></span> <span data-ttu-id="283ed-121">Cada secuencia de comandos ejecuta el código en `main` un solo método y, a continuación, finaliza.</span><span class="sxs-lookup"><span data-stu-id="283ed-121">Every script runs the code in a single `main` method, then ends.</span></span> <span data-ttu-id="283ed-122">No se reactiva cuando se desencadenan eventos y, por lo tanto, no pueden registrar los eventos.</span><span class="sxs-lookup"><span data-stu-id="283ed-122">It does not reactivate when events are triggered, and thus, cannot register events.</span></span>

### <a name="common-apis"></a><span data-ttu-id="283ed-123">API comunes</span><span class="sxs-lookup"><span data-stu-id="283ed-123">Common APIs</span></span>

<span data-ttu-id="283ed-124">Los scripts de Office no pueden usar [API comunes](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="283ed-124">Office Scripts cannot use [Common APIs](/javascript/api/office).</span></span> <span data-ttu-id="283ed-125">Si necesita la autenticación, ventanas de cuadro de diálogo u otras características que solo se admiten en las API comunes, es probable que deba crear un complemento de Office en lugar de un script de Office.</span><span class="sxs-lookup"><span data-stu-id="283ed-125">If you need authentication, dialog windows, or other features that are only supported by Common APIs, you'll likely need to create an Office Add-in instead of an Office Script.</span></span>

## <a name="see-also"></a><span data-ttu-id="283ed-126">Vea también</span><span class="sxs-lookup"><span data-stu-id="283ed-126">See also</span></span>

- [<span data-ttu-id="283ed-127">Scripts de Office en Excel en la web</span><span class="sxs-lookup"><span data-stu-id="283ed-127">Office Scripts in Excel on the web</span></span>](../overview/excel.md)
- [<span data-ttu-id="283ed-128">Solución de problemas de scripts de Office</span><span class="sxs-lookup"><span data-stu-id="283ed-128">Troubleshooting Office Scripts</span></span>](../testing/troubleshooting.md)
- [<span data-ttu-id="283ed-129">Crear un complemento de panel de tareas de Excel</span><span class="sxs-lookup"><span data-stu-id="283ed-129">Build an Excel task pane add-in</span></span>](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)