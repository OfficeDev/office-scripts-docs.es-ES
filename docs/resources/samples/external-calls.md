---
title: Realizar llamadas api externas en scripts de Office
description: Obtenga información sobre cómo realizar llamadas a API externas en scripts de Office.
ms.date: 03/30/2021
localization_priority: Normal
ms.openlocfilehash: d0abfa0bb1adedc7535059ed359b8053d9f1c84d
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571608"
---
# <a name="external-api-calls-from-office-scripts"></a><span data-ttu-id="d4cff-103">Llamadas de API externas desde scripts de Office</span><span class="sxs-lookup"><span data-stu-id="d4cff-103">External API calls from Office Scripts</span></span>

<span data-ttu-id="d4cff-104">Los scripts de Office [permiten una compatibilidad limitada con llamadas de API externas.](../../develop/external-calls.md)</span><span class="sxs-lookup"><span data-stu-id="d4cff-104">Office Scripts allows [limited external API call support](../../develop/external-calls.md).</span></span>

> [!IMPORTANT]
>
> * <span data-ttu-id="d4cff-105">No hay forma de iniciar sesión o usar el tipo OAuth2 de flujos de autenticación.</span><span class="sxs-lookup"><span data-stu-id="d4cff-105">There is no way to sign in or use OAuth2 type of authentication flows.</span></span> <span data-ttu-id="d4cff-106">Todas las claves y credenciales deben codificarse de forma rígida (o leerse desde otro origen).</span><span class="sxs-lookup"><span data-stu-id="d4cff-106">All keys and credentials have to be hardcoded (or read from another source).</span></span>
> * <span data-ttu-id="d4cff-107">No hay ninguna infraestructura para almacenar credenciales y claves de API.</span><span class="sxs-lookup"><span data-stu-id="d4cff-107">There is no infrastructure to store API credentials and keys.</span></span> <span data-ttu-id="d4cff-108">El usuario tendrá que administrarlo.</span><span class="sxs-lookup"><span data-stu-id="d4cff-108">This will have to be managed by the user.</span></span>
> * <span data-ttu-id="d4cff-109">Las llamadas externas pueden provocar que los datos confidenciales se exponán a extremos no deseados o que los datos externos se puedan incluir en libros internos.</span><span class="sxs-lookup"><span data-stu-id="d4cff-109">External calls may result in sensitive data being exposed to undesirable endpoints, or external data to be brought into internal workbooks.</span></span> <span data-ttu-id="d4cff-110">El administrador puede establecer la protección del firewall frente a estas llamadas.</span><span class="sxs-lookup"><span data-stu-id="d4cff-110">Your admin can establish firewall protection against such calls.</span></span> <span data-ttu-id="d4cff-111">Asegúrese de comprobar con las directivas locales antes de confiar en llamadas externas.</span><span class="sxs-lookup"><span data-stu-id="d4cff-111">Be sure to check with local policies prior to relying on external calls.</span></span>
> * <span data-ttu-id="d4cff-112">Si un script usa una llamada a la API, no funcionará en un escenario de Power Automate.</span><span class="sxs-lookup"><span data-stu-id="d4cff-112">If a script uses an API call, it will not function in a Power Automate scenario.</span></span> <span data-ttu-id="d4cff-113">Tendrás que usar la acción HTTP de Power Automate o acciones equivalentes para extraer datos de un servicio externo o insertarlo en él.</span><span class="sxs-lookup"><span data-stu-id="d4cff-113">You'll have to use Power Automate's HTTP action or equivalent actions to pull data from or push it to an external service.</span></span>
> * <span data-ttu-id="d4cff-114">Una llamada de API externa implica una sintaxis asincrónica de API y requiere conocimientos ligeramente avanzados de cómo funciona la comunicación asincrónica.</span><span class="sxs-lookup"><span data-stu-id="d4cff-114">An external API call involves asynchronous API syntax and requires slightly advanced knowledge of the way async communication works.</span></span>
> * <span data-ttu-id="d4cff-115">Asegúrese de comprobar la cantidad de rendimiento de datos antes de tomar una dependencia.</span><span class="sxs-lookup"><span data-stu-id="d4cff-115">Be sure to check the amount of data throughput prior to taking a dependency.</span></span> <span data-ttu-id="d4cff-116">Por ejemplo, extraer todo el conjunto de datos externo puede no ser la mejor opción y, en su lugar, se debe usar la paginación para obtener datos en fragmentos.</span><span class="sxs-lookup"><span data-stu-id="d4cff-116">For instance, pulling down the entire external dataset may not be the best option and instead pagination should be used to get data in chunks.</span></span>

## <a name="useful-knowledge-and-resources"></a><span data-ttu-id="d4cff-117">Conocimientos y recursos útiles</span><span class="sxs-lookup"><span data-stu-id="d4cff-117">Useful knowledge and resources</span></span>

* <span data-ttu-id="d4cff-118">[API de REST:](https://en.wikipedia.org/wiki/Representational_state_transfer)es muy probable que use la llamada a la API.</span><span class="sxs-lookup"><span data-stu-id="d4cff-118">[REST API](https://en.wikipedia.org/wiki/Representational_state_transfer): Most likely way you'll use the API call.</span></span>
* <span data-ttu-id="d4cff-119">[ `async` : Comprenda cómo funciona esto. `await` ](https://developer.mozilla.org/docs/Learn/JavaScript/Asynchronous/Async_await)</span><span class="sxs-lookup"><span data-stu-id="d4cff-119">[`async` `await`](https://developer.mozilla.org/docs/Learn/JavaScript/Asynchronous/Async_await): Understand how this works.</span></span>
* <span data-ttu-id="d4cff-120">[`fetch`](https://developer.mozilla.org/docs/Web/API/Fetch_API/Using_Fetch): Comprenda cómo funciona esto.</span><span class="sxs-lookup"><span data-stu-id="d4cff-120">[`fetch`](https://developer.mozilla.org/docs/Web/API/Fetch_API/Using_Fetch): Understand how this works.</span></span>

## <a name="steps"></a><span data-ttu-id="d4cff-121">Pasos</span><span class="sxs-lookup"><span data-stu-id="d4cff-121">Steps</span></span>

1. <span data-ttu-id="d4cff-122">Marca la `main` función como una función asincrónica agregando `async` prefijo.</span><span class="sxs-lookup"><span data-stu-id="d4cff-122">Mark your `main` function as an asynchronous function by adding `async` prefix.</span></span> <span data-ttu-id="d4cff-123">Por ejemplo, `async function main(workbook: ExcelScript.Workbook)`.</span><span class="sxs-lookup"><span data-stu-id="d4cff-123">For example, `async function main(workbook: ExcelScript.Workbook)`.</span></span>
1. <span data-ttu-id="d4cff-124">¿Qué tipo de llamada api realiza?</span><span class="sxs-lookup"><span data-stu-id="d4cff-124">Which type of API call are you making?</span></span> <span data-ttu-id="d4cff-125">`GET`, `POST`, `PUT`, `DELETE`, `PATCH`?</span><span class="sxs-lookup"><span data-stu-id="d4cff-125">`GET`, `POST`, `PUT`, `DELETE`, `PATCH`?</span></span> <span data-ttu-id="d4cff-126">Consulte el material de la API de REST para obtener más información.</span><span class="sxs-lookup"><span data-stu-id="d4cff-126">Refer to REST API material for details.</span></span>
1. <span data-ttu-id="d4cff-127">Obtenga el punto de conexión de la API de servicio, los requisitos de autenticación, los encabezados, etc.</span><span class="sxs-lookup"><span data-stu-id="d4cff-127">Obtain the service API endpoint, authentication requirements, headers, etc.</span></span>
1. <span data-ttu-id="d4cff-128">Defina la entrada o salida para `interface` ayudar con la finalización del código y la comprobación del tiempo de desarrollo.</span><span class="sxs-lookup"><span data-stu-id="d4cff-128">Define the input or output `interface` to help with code completion and development time verification.</span></span> <span data-ttu-id="d4cff-129">Vea [el vídeo](#training-video-how-to-make-external-api-calls) para obtener más información.</span><span class="sxs-lookup"><span data-stu-id="d4cff-129">See [video](#training-video-how-to-make-external-api-calls) for details.</span></span>
1. <span data-ttu-id="d4cff-130">Código, prueba, optimiza.</span><span class="sxs-lookup"><span data-stu-id="d4cff-130">Code, test, optimize.</span></span> <span data-ttu-id="d4cff-131">Puede crear una función para la rutina de llamada de la API para que sea reutilizable desde otras partes del script o para su reutilización en un script diferente (copiar y pegar se vuelve mucho más fácil de esta manera).</span><span class="sxs-lookup"><span data-stu-id="d4cff-131">You can create a function for your API call routine to make it reusable from other parts of your script or for reuse in a different script (copy-paste becomes much easier this way).</span></span>

## <a name="scenario"></a><span data-ttu-id="d4cff-132">Escenario</span><span class="sxs-lookup"><span data-stu-id="d4cff-132">Scenario</span></span>

<span data-ttu-id="d4cff-133">Este script obtiene información básica sobre los repositorios de GitHub del usuario.</span><span class="sxs-lookup"><span data-stu-id="d4cff-133">This script gets basic information about the user's GitHub repositories.</span></span>

![Ejemplo de obtener información de repositorios](../../images/git.png)

## <a name="resources-used-in-the-sample"></a><span data-ttu-id="d4cff-135">Recursos usados en el ejemplo</span><span class="sxs-lookup"><span data-stu-id="d4cff-135">Resources used in the sample</span></span>

1. [<span data-ttu-id="d4cff-136">Obtener la referencia de la API de Github de repositorios.</span><span class="sxs-lookup"><span data-stu-id="d4cff-136">Get repositories Github API reference.</span></span>](https://docs.github.com/rest/reference/repos#list-repositories-for-a-user)
1. <span data-ttu-id="d4cff-137">Salida de llamada api: vaya a un explorador web o a cualquier interfaz HTTP y escriba , reemplazando el marcador `https://api.github.com/users/{USERNAME}/repos` de posición {USERNAME} por su id. de Github.</span><span class="sxs-lookup"><span data-stu-id="d4cff-137">API call output: Go to a web browser or any HTTP interface and type in `https://api.github.com/users/{USERNAME}/repos`, replacing the {USERNAME} placeholder with your Github ID.</span></span>
1. <span data-ttu-id="d4cff-138">Información recuperada: repo.name, repo.size, repo.owner.id, repo.license?. nombre</span><span class="sxs-lookup"><span data-stu-id="d4cff-138">Information fetched: repo.name, repo.size, repo.owner.id, repo.license?.name</span></span>

## <a name="sample-code-get-basic-information-about-users-github-repositories"></a><span data-ttu-id="d4cff-139">Código de ejemplo: obtener información básica sobre los repositorios de GitHub del usuario</span><span class="sxs-lookup"><span data-stu-id="d4cff-139">Sample code: Get basic information about user's GitHub repositories</span></span>

```TypeScript
async function main(workbook: ExcelScript.Workbook) {

  // Replace the {USERNAME} placeholder with your GitHub username.
  const response = await fetch('https://api.github.com/users/{USERNAME}/repos');
  const repos: Repository[] = await response.json();
  
  const rows: (string | boolean | number)[][] = [];
  for (let repo of repos){ 
    rows.push([repo.id, repo.name, repo.license?.name, repo.license?.url])
  }
  const sheet = workbook.getActiveWorksheet();
  const range = sheet.getRange('A2').getResizedRange(rows.length - 1, rows[0].length - 1);
  range.setValues(rows);
  return;
}

interface Repository {
  name: string,
  id: string,
  license?: License 
}

interface License {
  name: string,
  url: string
}
```

## <a name="training-video-how-to-make-external-api-calls"></a><span data-ttu-id="d4cff-140">Vídeo de aprendizaje: Cómo realizar llamadas a API externas</span><span class="sxs-lookup"><span data-stu-id="d4cff-140">Training video: How to make external API calls</span></span>

<span data-ttu-id="d4cff-141">[![Ver vídeo sobre cómo realizar llamadas a API externas](../../images/api-vid.png)](https://youtu.be/fulP29J418E "Vídeo sobre cómo realizar llamadas a API externas")</span><span class="sxs-lookup"><span data-stu-id="d4cff-141">[![Watch video on how to make external API calls](../../images/api-vid.png)](https://youtu.be/fulP29J418E "Video on how to make external API calls")</span></span>
