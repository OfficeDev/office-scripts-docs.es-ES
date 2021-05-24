---
title: Usar llamadas de captura externa en Scripts de Office
description: Obtenga información sobre cómo realizar llamadas API externas en Office scripts.
ms.date: 05/14/2021
localization_priority: Normal
ms.openlocfilehash: df8814cbab16969a1140aecfe526fd68e609d43c
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545755"
---
# <a name="use-external-fetch-calls-in-office-scripts"></a><span data-ttu-id="74f36-103">Usar llamadas de captura externa en Scripts de Office</span><span class="sxs-lookup"><span data-stu-id="74f36-103">Use external fetch calls in Office Scripts</span></span>

<span data-ttu-id="74f36-104">Este script obtiene información básica sobre los repositorios de GitHub usuario.</span><span class="sxs-lookup"><span data-stu-id="74f36-104">This script gets basic information about a user's GitHub repositories.</span></span> <span data-ttu-id="74f36-105">Muestra cómo usar en `fetch` un escenario simple.</span><span class="sxs-lookup"><span data-stu-id="74f36-105">It shows how to use `fetch` in a simple scenario.</span></span> <span data-ttu-id="74f36-106">Para obtener más información sobre el uso u otras llamadas externas, lea Compatibilidad con llamadas `fetch` de API externa en scripts Office [externos](../../develop/external-calls.md)</span><span class="sxs-lookup"><span data-stu-id="74f36-106">For more information about using `fetch` or other external calls, read [External API call support in Office Scripts](../../develop/external-calls.md)</span></span>

<span data-ttu-id="74f36-107">Puede obtener más información sobre las API de GItHub que se usan en la referencia GitHub [API](https://docs.github.com/rest/reference/repos#list-repositories-for-a-user).</span><span class="sxs-lookup"><span data-stu-id="74f36-107">You can learn more about the GItHub APIs being used in the [GitHub API reference](https://docs.github.com/rest/reference/repos#list-repositories-for-a-user).</span></span> <span data-ttu-id="74f36-108">También puede ver el resultado de la llamada api sin procesar visitando en un explorador web (asegúrese de reemplazar el marcador de posición `https://api.github.com/users/{USERNAME}/repos` {USERNAME} por su identificador GitHub usuario).</span><span class="sxs-lookup"><span data-stu-id="74f36-108">You can also see the raw API call output by visiting `https://api.github.com/users/{USERNAME}/repos` in a web browser (be sure to replace the {USERNAME} placeholder with your GitHub ID).</span></span>

![Ejemplo de obtener información de repositorios](../../images/git.png)

## <a name="sample-code-get-basic-information-about-users-github-repositories"></a><span data-ttu-id="74f36-110">Código de ejemplo: obtener información básica sobre los repositorios de GitHub usuario</span><span class="sxs-lookup"><span data-stu-id="74f36-110">Sample code: Get basic information about user's GitHub repositories</span></span>

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

## <a name="training-video-how-to-make-external-api-calls"></a><span data-ttu-id="74f36-111">Vídeo de aprendizaje: Cómo realizar llamadas a API externas</span><span class="sxs-lookup"><span data-stu-id="74f36-111">Training video: How to make external API calls</span></span>

<span data-ttu-id="74f36-112">[Vea el recorrido de Sudhi Ramamurthy por este ejemplo en YouTube](https://youtu.be/fulP29J418E).</span><span class="sxs-lookup"><span data-stu-id="74f36-112">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/fulP29J418E).</span></span>
