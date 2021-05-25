---
title: Soporte de llamadas de API externas en Scripts de Office
description: Soporte técnico e instrucciones para realizar llamadas API externas en un script Office script.
ms.date: 05/21/2021
localization_priority: Normal
ms.openlocfilehash: 5d768b53112473c1774f8fe8257b197ffead4a63
ms.sourcegitcommit: 09d8859d5269ada8f1d0e141f6b5a4f96d95a739
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 05/24/2021
ms.locfileid: "52631646"
---
# <a name="external-api-call-support-in-office-scripts"></a><span data-ttu-id="9fcfd-103">Soporte de llamadas de API externas en Scripts de Office</span><span class="sxs-lookup"><span data-stu-id="9fcfd-103">External API call support in Office Scripts</span></span>

<span data-ttu-id="9fcfd-104">Los scripts admiten llamadas a servicios externos.</span><span class="sxs-lookup"><span data-stu-id="9fcfd-104">Scripts support calls to external services.</span></span> <span data-ttu-id="9fcfd-105">Use estos servicios para proporcionar datos y otra información al libro.</span><span class="sxs-lookup"><span data-stu-id="9fcfd-105">Use these services to supply data and other information to your workbook.</span></span>

> [!CAUTION]
> <span data-ttu-id="9fcfd-106">Las llamadas externas pueden provocar que los datos confidenciales se exponán a extremos no deseados.</span><span class="sxs-lookup"><span data-stu-id="9fcfd-106">External calls may result in sensitive data being exposed to undesirable endpoints.</span></span> <span data-ttu-id="9fcfd-107">El administrador puede establecer la protección del firewall frente a estas llamadas.</span><span class="sxs-lookup"><span data-stu-id="9fcfd-107">Your admin can establish firewall protection against such calls.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="9fcfd-108">Las llamadas a API externas solo se pueden realizar a través de la aplicación Excel, no a través de Power Automate [en circunstancias normales](#external-calls-from-power-automate).</span><span class="sxs-lookup"><span data-stu-id="9fcfd-108">Calls to external APIs can be only be made through the Excel application, not through Power Automate [under normal circumstances](#external-calls-from-power-automate).</span></span>

## <a name="configure-your-script-for-external-calls"></a><span data-ttu-id="9fcfd-109">Configurar el script para llamadas externas</span><span class="sxs-lookup"><span data-stu-id="9fcfd-109">Configure your script for external calls</span></span>

<span data-ttu-id="9fcfd-110">Las llamadas externas [son asincrónicas](https://developer.mozilla.org/docs/Learn/JavaScript/Asynchronous/Async_await) y requieren que el script esté marcado como `async` .</span><span class="sxs-lookup"><span data-stu-id="9fcfd-110">External calls are [asynchronous](https://developer.mozilla.org/docs/Learn/JavaScript/Asynchronous/Async_await) and require that your script is marked as `async`.</span></span> <span data-ttu-id="9fcfd-111">Agregue el `async` prefijo a la función y haga que devuelva un , como se muestra `main` `Promise` aquí:</span><span class="sxs-lookup"><span data-stu-id="9fcfd-111">Add the `async` prefix to your `main` function and have it return a `Promise`, as shown here:</span></span>

```typescript
async function main(workbook: ExcelScript.Workbook) : Promise <void>
```

> [!NOTE]
> <span data-ttu-id="9fcfd-112">Los scripts que devuelven otra información pueden devolver un `Promise` de ese tipo.</span><span class="sxs-lookup"><span data-stu-id="9fcfd-112">Scripts that return other information can return a `Promise` of that type.</span></span> <span data-ttu-id="9fcfd-113">Por ejemplo, si el script necesita devolver un `Employee` objeto, la firma de devolución sería `: Promise <Employee>`</span><span class="sxs-lookup"><span data-stu-id="9fcfd-113">For example, if your script needs to return an `Employee` object, the return signature would be `: Promise <Employee>`</span></span>

<span data-ttu-id="9fcfd-114">Tendrás que aprender las interfaces del servicio externo para realizar llamadas a ese servicio.</span><span class="sxs-lookup"><span data-stu-id="9fcfd-114">You'll need to learn the external service's interfaces to make calls to that service.</span></span> <span data-ttu-id="9fcfd-115">Si está usando o `fetch` [LAS API de REST,](https://wikipedia.org/wiki/Representational_state_transfer)debe determinar la estructura JSON de los datos devueltos.</span><span class="sxs-lookup"><span data-stu-id="9fcfd-115">If you are using `fetch` or [REST APIs](https://wikipedia.org/wiki/Representational_state_transfer), you need to determine the JSON structure of the returned data.</span></span> <span data-ttu-id="9fcfd-116">Para la entrada y salida desde el script, considere la posibilidad de realizar una para que `interface` coincida con las estructuras JSON necesarias.</span><span class="sxs-lookup"><span data-stu-id="9fcfd-116">For both input to and output from your script, consider making an `interface` to match the needed JSON structures.</span></span> <span data-ttu-id="9fcfd-117">Esto proporciona al script más seguridad de tipo.</span><span class="sxs-lookup"><span data-stu-id="9fcfd-117">This gives the script more type safety.</span></span> <span data-ttu-id="9fcfd-118">Puede ver un ejemplo de esto en [Using fetch from Office Scripts](../resources/samples/external-fetch-calls.md).</span><span class="sxs-lookup"><span data-stu-id="9fcfd-118">You can see an example of this in [Using fetch from Office Scripts](../resources/samples/external-fetch-calls.md).</span></span>

### <a name="limitations-with-external-calls-from-office-scripts"></a><span data-ttu-id="9fcfd-119">Limitaciones con llamadas externas desde Office scripts</span><span class="sxs-lookup"><span data-stu-id="9fcfd-119">Limitations with external calls from Office Scripts</span></span>

* <span data-ttu-id="9fcfd-120">No hay forma de iniciar sesión o usar el tipo OAuth2 de flujos de autenticación.</span><span class="sxs-lookup"><span data-stu-id="9fcfd-120">There is no way to sign in or use OAuth2 type of authentication flows.</span></span> <span data-ttu-id="9fcfd-121">Todas las claves y credenciales deben codificarse de forma rígida (o leerse desde otro origen).</span><span class="sxs-lookup"><span data-stu-id="9fcfd-121">All keys and credentials have to be hardcoded (or read from another source).</span></span>
* <span data-ttu-id="9fcfd-122">No hay ninguna infraestructura para almacenar credenciales y claves de API.</span><span class="sxs-lookup"><span data-stu-id="9fcfd-122">There is no infrastructure to store API credentials and keys.</span></span> <span data-ttu-id="9fcfd-123">El usuario tendrá que administrarlo.</span><span class="sxs-lookup"><span data-stu-id="9fcfd-123">This will have to be managed by the user.</span></span>
* <span data-ttu-id="9fcfd-124">No se admiten las cookies de documento `localStorage` `sessionStorage` ni los objetos.</span><span class="sxs-lookup"><span data-stu-id="9fcfd-124">Document cookies, `localStorage`, and `sessionStorage` objects are not supported.</span></span>
* <span data-ttu-id="9fcfd-125">Las llamadas externas pueden provocar que los datos confidenciales se exponán a extremos no deseados o que los datos externos se puedan incluir en libros internos.</span><span class="sxs-lookup"><span data-stu-id="9fcfd-125">External calls may result in sensitive data being exposed to undesirable endpoints, or external data to be brought into internal workbooks.</span></span> <span data-ttu-id="9fcfd-126">El administrador puede establecer la protección del firewall frente a estas llamadas.</span><span class="sxs-lookup"><span data-stu-id="9fcfd-126">Your admin can establish firewall protection against such calls.</span></span> <span data-ttu-id="9fcfd-127">Asegúrese de comprobar con las directivas locales antes de confiar en llamadas externas.</span><span class="sxs-lookup"><span data-stu-id="9fcfd-127">Be sure to check with local policies prior to relying on external calls.</span></span>
* <span data-ttu-id="9fcfd-128">Asegúrese de comprobar la cantidad de rendimiento de datos antes de tomar una dependencia.</span><span class="sxs-lookup"><span data-stu-id="9fcfd-128">Be sure to check the amount of data throughput prior to taking a dependency.</span></span> <span data-ttu-id="9fcfd-129">Por ejemplo, extraer todo el conjunto de datos externo puede no ser la mejor opción y, en su lugar, se debe usar la paginación para obtener datos en fragmentos.</span><span class="sxs-lookup"><span data-stu-id="9fcfd-129">For instance, pulling down the entire external dataset may not be the best option and instead pagination should be used to get data in chunks.</span></span>

## <a name="retrieve-information-with-fetch"></a><span data-ttu-id="9fcfd-130">Recuperar información con `fetch`</span><span class="sxs-lookup"><span data-stu-id="9fcfd-130">Retrieve information with `fetch`</span></span>

<span data-ttu-id="9fcfd-131">La [API de captura](https://developer.mozilla.org/docs/Web/API/Fetch_API) recupera información de servicios externos.</span><span class="sxs-lookup"><span data-stu-id="9fcfd-131">The [fetch API](https://developer.mozilla.org/docs/Web/API/Fetch_API) retrieves information from external services.</span></span> <span data-ttu-id="9fcfd-132">Es una `async` API, por lo que debe ajustar la `main` firma del script.</span><span class="sxs-lookup"><span data-stu-id="9fcfd-132">It is an `async` API, so you need to adjust the `main` signature of your script.</span></span> <span data-ttu-id="9fcfd-133">Haga que `main` la función y haga que devuelva un `async` `Promise<void>` .</span><span class="sxs-lookup"><span data-stu-id="9fcfd-133">Make the `main` function `async` and have it return a `Promise<void>`.</span></span> <span data-ttu-id="9fcfd-134">También debe asegurarse de la llamada `await` `fetch` y la `json` recuperación.</span><span class="sxs-lookup"><span data-stu-id="9fcfd-134">You should also be sure to `await` the `fetch` call and `json` retrieval.</span></span> <span data-ttu-id="9fcfd-135">Esto garantiza que las operaciones se completen antes de que finalice el script.</span><span class="sxs-lookup"><span data-stu-id="9fcfd-135">This ensures those operations complete before the script ends.</span></span>

<span data-ttu-id="9fcfd-136">Los datos JSON recuperados por `fetch` deben coincidir con una interfaz definida en el script.</span><span class="sxs-lookup"><span data-stu-id="9fcfd-136">Any JSON data retrieved by `fetch` must match an interface defined in the script.</span></span> <span data-ttu-id="9fcfd-137">El valor devuelto debe asignarse a un tipo específico porque Office [scripts no admiten el `any` tipo](typescript-restrictions.md#no-any-type-in-office-scripts).</span><span class="sxs-lookup"><span data-stu-id="9fcfd-137">The returned value must be assigned to a specific type because [Office Scripts do not support the `any` type](typescript-restrictions.md#no-any-type-in-office-scripts).</span></span> <span data-ttu-id="9fcfd-138">Debe consultar la documentación del servicio para ver cuáles son los nombres y tipos de las propiedades devueltas.</span><span class="sxs-lookup"><span data-stu-id="9fcfd-138">You should refer to the documentation for your service to see what the names and types of the returned properties are.</span></span> <span data-ttu-id="9fcfd-139">A continuación, agregue la interfaz o las interfaces correspondientes al script.</span><span class="sxs-lookup"><span data-stu-id="9fcfd-139">Then, add the matching interface or interfaces to your script.</span></span>

<span data-ttu-id="9fcfd-140">El siguiente script usa `fetch` para recuperar datos JSON del servidor de prueba en la dirección URL determinada.</span><span class="sxs-lookup"><span data-stu-id="9fcfd-140">The following script uses `fetch` to retrieve JSON data from the test server in the given URL.</span></span> <span data-ttu-id="9fcfd-141">Tenga en `JSONData` cuenta la interfaz para almacenar los datos como un tipo de coincidencia.</span><span class="sxs-lookup"><span data-stu-id="9fcfd-141">Note the `JSONData` interface to store the data as a matching type.</span></span>

```TypeScript
async function main(workbook: ExcelScript.Workbook): Promise<void> {
  // Retrieve sample JSON data from a test server.
  let fetchResult = await fetch('https://jsonplaceholder.typicode.com/todos/1');

  // Convert the returned data to the expected JSON structure.
  let json : JSONData = await fetchResult.json();

  // Display the content in a readable format.
  console.log(JSON.stringify(json));
}

/**
 * An interface that matches the returned JSON structure.
 * The property names match exactly.
 */
interface JSONData {
  userId: number;
  id: number;
  title: string;
  completed: boolean;
}
```

### <a name="other-fetch-samples"></a><span data-ttu-id="9fcfd-142">Otras `fetch` muestras</span><span class="sxs-lookup"><span data-stu-id="9fcfd-142">Other `fetch` samples</span></span>

* <span data-ttu-id="9fcfd-143">El [ejemplo Use external fetch calls in Office Scripts](../resources/samples/external-fetch-calls.md) muestra cómo obtener información básica sobre los repositorios de GitHub usuario.</span><span class="sxs-lookup"><span data-stu-id="9fcfd-143">The [Use external fetch calls in Office Scripts](../resources/samples/external-fetch-calls.md) sample shows how to get basic information about a user's GitHub repositories.</span></span>
* <span data-ttu-id="9fcfd-144">Escenario de ejemplo de scripts de Office: Graph datos de nivel de agua de [NOAA](../resources/scenarios/noaa-data-fetch.md) muestra el comando de captura que se usa para recuperar registros de la base de datos de corrientes y mareos de la Administración oceánica y atmosférica nacional.</span><span class="sxs-lookup"><span data-stu-id="9fcfd-144">The [Office Scripts sample scenario: Graph water-level data from NOAA](../resources/scenarios/noaa-data-fetch.md) demonstrates the fetch command being used to retrieve records from the National Oceanic and Atmospheric Administration's Tides and Currents database.</span></span>

## <a name="external-calls-from-power-automate"></a><span data-ttu-id="9fcfd-145">Llamadas externas desde Power Automate</span><span class="sxs-lookup"><span data-stu-id="9fcfd-145">External calls from Power Automate</span></span>

<span data-ttu-id="9fcfd-146">Cualquier llamada de API externa produce un error cuando se ejecuta un script con Power Automate.</span><span class="sxs-lookup"><span data-stu-id="9fcfd-146">Any external API call fails when a script is run with Power Automate.</span></span> <span data-ttu-id="9fcfd-147">Esta es una diferencia de comportamiento entre ejecutar un script a través de la aplicación Excel y a través de Power Automate.</span><span class="sxs-lookup"><span data-stu-id="9fcfd-147">This is a behavioral difference between running a script through the Excel application and through Power Automate.</span></span> <span data-ttu-id="9fcfd-148">Asegúrese de comprobar las referencias de los scripts antes de crearlas en un flujo.</span><span class="sxs-lookup"><span data-stu-id="9fcfd-148">Be sure to check your scripts for such references before building them into a flow.</span></span>

<span data-ttu-id="9fcfd-149">Tendrás que usar [HTTP](/connectors/webcontents/) con Azure AD u otras acciones equivalentes para extraer datos de un servicio externo o insertarlo en él.</span><span class="sxs-lookup"><span data-stu-id="9fcfd-149">You'll have to use [HTTP with Azure AD](/connectors/webcontents/) or other equivalent actions to pull data from or push it to an external service.</span></span>

> [!WARNING]
> <span data-ttu-id="9fcfd-150">Las llamadas externas realizadas a través del conector [Power Automate Excel Online fallan](/connectors/excelonlinebusiness) para ayudar a mantener las directivas de prevención de pérdida de datos existentes.</span><span class="sxs-lookup"><span data-stu-id="9fcfd-150">External calls made through the Power Automate [Excel Online connector](/connectors/excelonlinebusiness) fail in order to help uphold existing data loss prevention policies.</span></span> <span data-ttu-id="9fcfd-151">Sin embargo, los scripts que se ejecutan Power Automate se realizan fuera de la organización y fuera de los firewalls de la organización.</span><span class="sxs-lookup"><span data-stu-id="9fcfd-151">However, scripts that are run through Power Automate are done so outside of your organization, and outside of your organization's firewalls.</span></span> <span data-ttu-id="9fcfd-152">Para obtener protección adicional contra usuarios malintencionados en este entorno externo, el administrador puede controlar el uso de Office scripts.</span><span class="sxs-lookup"><span data-stu-id="9fcfd-152">For additional protection from malicious users in this external environment, your admin can control the use of Office Scripts.</span></span> <span data-ttu-id="9fcfd-153">El administrador puede deshabilitar el conector de Excel Online en Power Automate o desactivar scripts de Office para Excel en la Web a través de los controles de administrador [de scripts](/microsoft-365/admin/manage/manage-office-scripts-settings)Office.</span><span class="sxs-lookup"><span data-stu-id="9fcfd-153">Your admin can either disable the Excel Online connector in Power Automate or turn off Office Scripts for Excel on the web through the [Office Scripts administrator controls](/microsoft-365/admin/manage/manage-office-scripts-settings).</span></span>

## <a name="see-also"></a><span data-ttu-id="9fcfd-154">Vea también</span><span class="sxs-lookup"><span data-stu-id="9fcfd-154">See also</span></span>

* [<span data-ttu-id="9fcfd-155">Usar objetos integrados de JavaScript en los scripts de Office</span><span class="sxs-lookup"><span data-stu-id="9fcfd-155">Using built-in JavaScript objects in Office Scripts</span></span>](javascript-objects.md)
* [<span data-ttu-id="9fcfd-156">Usar llamadas de captura externa en Scripts de Office</span><span class="sxs-lookup"><span data-stu-id="9fcfd-156">Use external fetch calls in Office Scripts</span></span>](../resources/samples/external-fetch-calls.md)
* [<span data-ttu-id="9fcfd-157">Office Escenario de ejemplo de scripts: Graph datos de nivel de agua de NOAA</span><span class="sxs-lookup"><span data-stu-id="9fcfd-157">Office Scripts sample scenario: Graph water-level data from NOAA</span></span>](../resources/scenarios/noaa-data-fetch.md)
