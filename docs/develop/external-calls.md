---
title: Compatibilidad con llamadas a API externas en scripts de Office
description: Soporte técnico y guía para realizar llamadas a API externas en un script de Office.
ms.date: 06/25/2020
localization_priority: Normal
ms.openlocfilehash: ec8281551cbe7c500eee40ec86067e5efbfcfc31
ms.sourcegitcommit: aec3c971c6640429f89b6bb99d2c95ea06725599
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 06/25/2020
ms.locfileid: "44878821"
---
# <a name="external-api-call-support-in-office-scripts"></a><span data-ttu-id="014fb-103">Compatibilidad con llamadas a API externas en scripts de Office</span><span class="sxs-lookup"><span data-stu-id="014fb-103">External API call support in Office Scripts</span></span>

<span data-ttu-id="014fb-104">La plataforma de scripts de Office no admite llamadas a [API externas](https://developer.mozilla.org/docs/Web/API).</span><span class="sxs-lookup"><span data-stu-id="014fb-104">The Office Scripts platform doesn't support calls to [external APIs](https://developer.mozilla.org/docs/Web/API).</span></span> <span data-ttu-id="014fb-105">Sin embargo, estas llamadas pueden ejecutarse en las circunstancias adecuadas.</span><span class="sxs-lookup"><span data-stu-id="014fb-105">However, these calls can be run under the right circumstances.</span></span> <span data-ttu-id="014fb-106">Las llamadas externas solo pueden realizarse a través del cliente de Excel, no a través de la automatización de la energía [en circunstancias normales](#external-calls-from-power-automate).</span><span class="sxs-lookup"><span data-stu-id="014fb-106">External calls can be only be made through the Excel client, not through Power Automate [under normal circumstances](#external-calls-from-power-automate).</span></span>

<span data-ttu-id="014fb-107">Los autores de scripts no deberían esperar un comportamiento coherente al usar las API externas durante la fase de vista previa de la plataforma.</span><span class="sxs-lookup"><span data-stu-id="014fb-107">Script authors shouldn't expect consistent behavior when using external APIs during the platform's preview phase.</span></span> <span data-ttu-id="014fb-108">Esto se debe a la forma en que el tiempo de ejecución de JavaScript administra la interacción con el libro.</span><span class="sxs-lookup"><span data-stu-id="014fb-108">This is due how the JavaScript runtime manages interacting with the workbook.</span></span> <span data-ttu-id="014fb-109">El script puede finalizar antes de que finalice la llamada a la API (o de que `Promise` se haya resuelto completamente).</span><span class="sxs-lookup"><span data-stu-id="014fb-109">The script may end before the API call completes (or its `Promise` is fully resolved).</span></span> <span data-ttu-id="014fb-110">Por lo tanto, no confíe en las API externas en escenarios de script críticos.</span><span class="sxs-lookup"><span data-stu-id="014fb-110">As such, do not rely on external APIs for critical script scenarios.</span></span>

> [!CAUTION]
> <span data-ttu-id="014fb-111">Las llamadas externas pueden dar lugar a que los datos confidenciales se expongan a extremos no deseados.</span><span class="sxs-lookup"><span data-stu-id="014fb-111">External calls may result in sensitive data being exposed to undesirable endpoints.</span></span> <span data-ttu-id="014fb-112">El administrador puede establecer la protección del firewall en dichas llamadas.</span><span class="sxs-lookup"><span data-stu-id="014fb-112">Your admin can establish firewall protection against such calls.</span></span>

## <a name="definition-files-for-external-apis"></a><span data-ttu-id="014fb-113">Archivos de definición para API externas</span><span class="sxs-lookup"><span data-stu-id="014fb-113">Definition files for external APIs</span></span>

<span data-ttu-id="014fb-114">Los archivos de definición para las API externas no se incluyen con los scripts de Office.</span><span class="sxs-lookup"><span data-stu-id="014fb-114">The definition files for external APIs aren't included with Office Scripts.</span></span> <span data-ttu-id="014fb-115">El uso de estas API genera errores en tiempo de compilación para las definiciones que faltan.</span><span class="sxs-lookup"><span data-stu-id="014fb-115">The use of such APIs generates compile-time errors for missing definitions.</span></span> <span data-ttu-id="014fb-116">Las API se siguen ejecutando (aunque solo cuando se ejecutan mediante el cliente de Excel), como se muestra en el siguiente script:</span><span class="sxs-lookup"><span data-stu-id="014fb-116">The APIs still run (though only when run through the Excel client), as shown in the following script:</span></span>

```typescript
async function main(workbook: ExcelScript.Workbook): Promise <void> {
  /* The following line of code generates the error:
   * "Cannot find name 'fetch'".
   * It will still run and return the JSON from the testing service.
   */
  let fetchResult = await fetch('https://jsonplaceholder.typicode.com/todos/1');
  let json = await fetchResult.json();

  // Displays the content from https://jsonplaceholder.typicode.com/todos/1
  console.log(JSON.stringify(json));
}
```

## <a name="external-calls-from-power-automate"></a><span data-ttu-id="014fb-117">Llamadas externas de la automatización de la alimentación</span><span class="sxs-lookup"><span data-stu-id="014fb-117">External calls from Power Automate</span></span>

<span data-ttu-id="014fb-118">Se produce un error en cualquier llamada de API externa cuando se ejecuta un script con la automatización de energía.</span><span class="sxs-lookup"><span data-stu-id="014fb-118">Any external API calls fail when a script is run with Power Automate.</span></span> <span data-ttu-id="014fb-119">Esta es una diferencia de comportamiento entre la ejecución de un script a través del cliente de Excel y a través de la automatización de la energía.</span><span class="sxs-lookup"><span data-stu-id="014fb-119">This is a behavioral difference between running a script through the Excel client and through Power Automate.</span></span> <span data-ttu-id="014fb-120">Asegúrese de comprobar las secuencias de comandos para las referencias antes de crearlas en un flujo.</span><span class="sxs-lookup"><span data-stu-id="014fb-120">Be sure to check your scripts for such references before building them into a flow.</span></span>

> [!WARNING]
> <span data-ttu-id="014fb-121">El error de las llamadas externas [Excel online Connector](/connectors/excelonlinebusiness) en Power Automatically para ayudar a preservar las directivas de prevención de pérdida de datos existentes.</span><span class="sxs-lookup"><span data-stu-id="014fb-121">The failure of external calls [Excel Online connector](/connectors/excelonlinebusiness) in Power Automate is there to help uphold existing data loss prevention policies.</span></span> <span data-ttu-id="014fb-122">Sin embargo, los scripts que se ejecutan mediante la automatización de la alimentación se realizan fuera de la organización y fuera de los firewalls de la organización.</span><span class="sxs-lookup"><span data-stu-id="014fb-122">However, the scripts run through Power Automate are done so outside of your organization, and outside of your organization's firewalls.</span></span> <span data-ttu-id="014fb-123">Para obtener protección adicional de usuarios malintencionados en este entorno externo, el administrador puede controlar el uso de scripts de Office.</span><span class="sxs-lookup"><span data-stu-id="014fb-123">For additional protection from malicious users in this external environment, your admin can control the use of Office Scripts.</span></span> <span data-ttu-id="014fb-124">El administrador puede deshabilitar el conector de Excel online con Power automatice o desactivar los scripts de Office para Excel en la web a través de los [controles de administrador de scripts de Office](https://support.microsoft.com/office/19d3c51a-6ca2-40ab-978d-60fa49554dcf).</span><span class="sxs-lookup"><span data-stu-id="014fb-124">Your admin can either disable the Excel Online connector in Power Automate or turn off Office Scripts for Excel on the web through the [Office Scripts administrator controls](https://support.microsoft.com/office/19d3c51a-6ca2-40ab-978d-60fa49554dcf).</span></span>

## <a name="see-also"></a><span data-ttu-id="014fb-125">Vea también</span><span class="sxs-lookup"><span data-stu-id="014fb-125">See also</span></span>

- [<span data-ttu-id="014fb-126">Usar objetos integrados de JavaScript en los scripts de Office</span><span class="sxs-lookup"><span data-stu-id="014fb-126">Using built-in JavaScript objects in Office Scripts</span></span>](javascript-objects.md)