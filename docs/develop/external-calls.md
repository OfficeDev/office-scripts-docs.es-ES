---
title: Soporte de llamadas de API externas en Scripts de Office
description: Soporte técnico e instrucciones para realizar llamadas a API externas en un script de Office.
ms.date: 01/05/2021
localization_priority: Normal
ms.openlocfilehash: 74b8750f609370370759ca4a4a1daa998363ac2e
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 04/02/2021
ms.locfileid: "51570314"
---
# <a name="external-api-call-support-in-office-scripts"></a><span data-ttu-id="d50ca-103">Soporte de llamadas de API externas en Scripts de Office</span><span class="sxs-lookup"><span data-stu-id="d50ca-103">External API call support in Office Scripts</span></span>

<span data-ttu-id="d50ca-104">Los autores de scripts no deben esperar un comportamiento coherente al usar [API externas](https://developer.mozilla.org/docs/Web/API) durante la fase de vista previa de la plataforma.</span><span class="sxs-lookup"><span data-stu-id="d50ca-104">Script authors shouldn't expect consistent behavior when using [external APIs](https://developer.mozilla.org/docs/Web/API) during the platform's preview phase.</span></span> <span data-ttu-id="d50ca-105">Por lo tanto, no confíe en api externas para escenarios de script críticos.</span><span class="sxs-lookup"><span data-stu-id="d50ca-105">As such, do not rely on external APIs for critical script scenarios.</span></span>

<span data-ttu-id="d50ca-106">Las llamadas a API externas solo se pueden realizar a través de la aplicación de Excel, no a través de Power Automate [en circunstancias normales.](#external-calls-from-power-automate)</span><span class="sxs-lookup"><span data-stu-id="d50ca-106">Calls to external APIs can be only be made through the Excel application, not through Power Automate [under normal circumstances](#external-calls-from-power-automate).</span></span>

> [!CAUTION]
> <span data-ttu-id="d50ca-107">Las llamadas externas pueden provocar que los datos confidenciales se exponán a extremos no deseados.</span><span class="sxs-lookup"><span data-stu-id="d50ca-107">External calls may result in sensitive data being exposed to undesirable endpoints.</span></span> <span data-ttu-id="d50ca-108">El administrador puede establecer la protección del firewall frente a estas llamadas.</span><span class="sxs-lookup"><span data-stu-id="d50ca-108">Your admin can establish firewall protection against such calls.</span></span>

## <a name="working-with-fetch"></a><span data-ttu-id="d50ca-109">Trabajar con `fetch`</span><span class="sxs-lookup"><span data-stu-id="d50ca-109">Working with `fetch`</span></span>

<span data-ttu-id="d50ca-110">La [API de captura](https://developer.mozilla.org/docs/Web/API/Fetch_API) recupera información de servicios externos.</span><span class="sxs-lookup"><span data-stu-id="d50ca-110">The [fetch API](https://developer.mozilla.org/docs/Web/API/Fetch_API) retrieves information from external services.</span></span> <span data-ttu-id="d50ca-111">Es una `async` API, por lo que tendrá que ajustar la `main` firma del script.</span><span class="sxs-lookup"><span data-stu-id="d50ca-111">It is an `async` API, so you will need to adjust the `main` signature of your script.</span></span> <span data-ttu-id="d50ca-112">Haga que `main` la función y haga que devuelva un `async` `Promise<void>` .</span><span class="sxs-lookup"><span data-stu-id="d50ca-112">Make the `main` function `async` and have it return a `Promise<void>`.</span></span> <span data-ttu-id="d50ca-113">También debe asegurarse de la llamada `await` `fetch` y la `json` recuperación.</span><span class="sxs-lookup"><span data-stu-id="d50ca-113">You should also be sure to `await` the `fetch` call and `json` retrieval.</span></span> <span data-ttu-id="d50ca-114">Esto garantiza que las operaciones se completen antes de que finalice el script.</span><span class="sxs-lookup"><span data-stu-id="d50ca-114">This ensures those operations complete before the script ends.</span></span>

<span data-ttu-id="d50ca-115">El siguiente script usa `fetch` para recuperar datos JSON del servidor de prueba en la dirección URL determinada.</span><span class="sxs-lookup"><span data-stu-id="d50ca-115">The following script uses `fetch` to retrieve JSON data from the test server in the given URL.</span></span>

```TypeScript
async function main(workbook: ExcelScript.Workbook): Promise <void> {
  /* 
   * Retrieve JSON data from a test server.
   */
  let fetchResult = await fetch('https://jsonplaceholder.typicode.com/todos/1');
  let json = await fetchResult.json();

  // Displays the content from https://jsonplaceholder.typicode.com/todos/1
  console.log(JSON.stringify(json));
}
```

<span data-ttu-id="d50ca-116">El escenario de ejemplo scripts de Office: Los datos de nivel de agua de gráfico de [NOAA](../resources/scenarios/noaa-data-fetch.md) muestran el comando de captura que se usa para recuperar registros de la base de datos de corrientes y mareos de la Administración nacional oceánica y atmosférico.</span><span class="sxs-lookup"><span data-stu-id="d50ca-116">The [Office Scripts sample scenario: Graph water-level data from NOAA](../resources/scenarios/noaa-data-fetch.md) demonstrates the fetch command being used to retrieve records from the National Oceanic and Atmospheric Administration's Tides and Currents database.</span></span>

## <a name="external-calls-from-power-automate"></a><span data-ttu-id="d50ca-117">Llamadas externas desde Power Automate</span><span class="sxs-lookup"><span data-stu-id="d50ca-117">External calls from Power Automate</span></span>

<span data-ttu-id="d50ca-118">Las llamadas a API externas fallan cuando se ejecuta un script con Power Automate.</span><span class="sxs-lookup"><span data-stu-id="d50ca-118">Any external API calls fail when a script is run with Power Automate.</span></span> <span data-ttu-id="d50ca-119">Esta es una diferencia de comportamiento entre ejecutar un script a través del cliente de Excel y a través de Power Automate.</span><span class="sxs-lookup"><span data-stu-id="d50ca-119">This is a behavioral difference between running a script through the Excel client and through Power Automate.</span></span> <span data-ttu-id="d50ca-120">Asegúrese de comprobar las referencias de los scripts antes de crearlas en un flujo.</span><span class="sxs-lookup"><span data-stu-id="d50ca-120">Be sure to check your scripts for such references before building them into a flow.</span></span>

> [!WARNING]
> <span data-ttu-id="d50ca-121">Las llamadas externas realizadas a través del conector de Power Automate [Excel Online](/connectors/excelonlinebusiness) fallan para ayudar a mantener las directivas de prevención de pérdida de datos existentes.</span><span class="sxs-lookup"><span data-stu-id="d50ca-121">External calls made through the Power Automate [Excel Online connector](/connectors/excelonlinebusiness) fail in order to help uphold existing data loss prevention policies.</span></span> <span data-ttu-id="d50ca-122">Sin embargo, los scripts que se ejecutan a través de Power Automate se realizan fuera de la organización y fuera de los firewalls de la organización.</span><span class="sxs-lookup"><span data-stu-id="d50ca-122">However, scripts that are run through Power Automate are done so outside of your organization, and outside of your organization's firewalls.</span></span> <span data-ttu-id="d50ca-123">Para obtener protección adicional contra usuarios malintencionados en este entorno externo, el administrador puede controlar el uso de scripts de Office.</span><span class="sxs-lookup"><span data-stu-id="d50ca-123">For additional protection from malicious users in this external environment, your admin can control the use of Office Scripts.</span></span> <span data-ttu-id="d50ca-124">El administrador puede deshabilitar el conector de Excel Online en Power Automate o desactivar scripts de Office para Excel en la web a través de los controles de administrador [de Scripts de Office.](/microsoft-365/admin/manage/manage-office-scripts-settings)</span><span class="sxs-lookup"><span data-stu-id="d50ca-124">Your admin can either disable the Excel Online connector in Power Automate or turn off Office Scripts for Excel on the web through the [Office Scripts administrator controls](/microsoft-365/admin/manage/manage-office-scripts-settings).</span></span>

## <a name="see-also"></a><span data-ttu-id="d50ca-125">Vea también</span><span class="sxs-lookup"><span data-stu-id="d50ca-125">See also</span></span>

- [<span data-ttu-id="d50ca-126">Usar objetos integrados de JavaScript en los scripts de Office</span><span class="sxs-lookup"><span data-stu-id="d50ca-126">Using built-in JavaScript objects in Office Scripts</span></span>](javascript-objects.md)
- [<span data-ttu-id="d50ca-127">Escenario de ejemplo de scripts de Office: gráfico de datos de nivel de agua de NOAA</span><span class="sxs-lookup"><span data-stu-id="d50ca-127">Office Scripts sample scenario: Graph water-level data from NOAA</span></span>](../resources/scenarios/noaa-data-fetch.md)
