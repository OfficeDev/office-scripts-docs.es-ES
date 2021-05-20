---
title: Usar llamadas de captura externa en Scripts de Office
description: Obtenga información sobre cómo realizar llamadas a API externas en scripts de Office.
ms.date: 05/14/2021
localization_priority: Normal
ms.openlocfilehash: df8814cbab16969a1140aecfe526fd68e609d43c
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545755"
---
# <a name="use-external-fetch-calls-in-office-scripts"></a><span data-ttu-id="db0a5-103">Usar llamadas de captura externa en Scripts de Office</span><span class="sxs-lookup"><span data-stu-id="db0a5-103">Use external fetch calls in Office Scripts</span></span>

<span data-ttu-id="db0a5-104">Este script obtiene información básica sobre los repositorios de GitHub de un usuario.</span><span class="sxs-lookup"><span data-stu-id="db0a5-104">This script gets basic information about a user's GitHub repositories.</span></span> <span data-ttu-id="db0a5-105">Muestra cómo usar `fetch` en un escenario simple.</span><span class="sxs-lookup"><span data-stu-id="db0a5-105">It shows how to use `fetch` in a simple scenario.</span></span> <span data-ttu-id="db0a5-106">Para obtener más información sobre el uso `fetch` u otras llamadas externas, lea Compatibilidad con [llamadas a la API externa en scripts de Office](../../develop/external-calls.md)</span><span class="sxs-lookup"><span data-stu-id="db0a5-106">For more information about using `fetch` or other external calls, read [External API call support in Office Scripts](../../develop/external-calls.md)</span></span>

<span data-ttu-id="db0a5-107">Puede obtener más información acerca de las API de GItHub que se usan en la [referencia de la API de GitHub](https://docs.github.com/rest/reference/repos#list-repositories-for-a-user).</span><span class="sxs-lookup"><span data-stu-id="db0a5-107">You can learn more about the GItHub APIs being used in the [GitHub API reference](https://docs.github.com/rest/reference/repos#list-repositories-for-a-user).</span></span> <span data-ttu-id="db0a5-108">También puede ver la salida de llamada a la API sin procesar visitando `https://api.github.com/users/{USERNAME}/repos` en un explorador web (asegúrese de reemplazar el marcador de posición {USERNAME} por el ID de GitHub).</span><span class="sxs-lookup"><span data-stu-id="db0a5-108">You can also see the raw API call output by visiting `https://api.github.com/users/{USERNAME}/repos` in a web browser (be sure to replace the {USERNAME} placeholder with your GitHub ID).</span></span>

![Obtener ejemplos de información de repositorios](../../images/git.png)

## <a name="sample-code-get-basic-information-about-users-github-repositories"></a><span data-ttu-id="db0a5-110">Código de ejemplo: obtenga información básica sobre los repositorios de GitHub del usuario</span><span class="sxs-lookup"><span data-stu-id="db0a5-110">Sample code: Get basic information about user's GitHub repositories</span></span>

```TypeScript
async function main(workbook: ExcelScript.Workbook) {
  // Call the GitHub REST API.
  // Replace the {USERNAME} placeholder with your GitHub username.
  const response = await fetch('https://api.github.com/users/{USERNAME}/repos');
  const repos: Repository[] = await response.json();
  
  // Create an array to hold the returned values.
  const rows: (string | boolean | number)[][] = [];

  // Convert each repository block into a row.
  for (let repo of repos){ 
    rows.push([repo.id, repo.name, repo.license?.name, repo.license?.url])
  }

  // Add the data to the current worksheet, starting at "A2".
  const sheet = workbook.getActiveWorksheet();
  const range = sheet.getRange('A2').getResizedRange(rows.length - 1, rows[0].length - 1);
  range.setValues(rows);
}

// An interface matching the returned JSON for a GitHub repository.
interface Repository {
  name: string,
  id: string,
  license?: License 
}

// An interface matching the returned JSON for a GitHub repo license.
interface License {
  name: string,
  url: string
}
```

## <a name="training-video-how-to-make-external-api-calls"></a><span data-ttu-id="db0a5-111">Vídeo de formación: cómo realizar llamadas a la API externas</span><span class="sxs-lookup"><span data-stu-id="db0a5-111">Training video: How to make external API calls</span></span>

<span data-ttu-id="db0a5-112">[Mira a Sudhi Ramamurthy caminar a través de esta muestra en YouTube.](https://youtu.be/fulP29J418E)</span><span class="sxs-lookup"><span data-stu-id="db0a5-112">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/fulP29J418E).</span></span>
