---
title: Usar llamadas de captura externa en Office scripts
description: Obtenga información sobre cómo realizar llamadas API externas en Office scripts.
ms.date: 04/05/2021
localization_priority: Normal
ms.openlocfilehash: a77ceb61c2ff46a7b6226b798462b7be2c8e1c54
ms.sourcegitcommit: 1f003c9924e651600c913d84094506125f1055ab
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 04/26/2021
ms.locfileid: "52026997"
---
# <a name="use-external-fetch-calls-in-office-scripts"></a><span data-ttu-id="ff2fa-103">Usar llamadas de captura externa en Office scripts</span><span class="sxs-lookup"><span data-stu-id="ff2fa-103">Use external fetch calls in Office Scripts</span></span>

<span data-ttu-id="ff2fa-104">Este script obtiene información básica sobre los repositorios de GitHub usuario.</span><span class="sxs-lookup"><span data-stu-id="ff2fa-104">This script gets basic information about a user's GitHub repositories.</span></span> <span data-ttu-id="ff2fa-105">Muestra cómo usar en `fetch` un escenario simple.</span><span class="sxs-lookup"><span data-stu-id="ff2fa-105">It shows how to use `fetch` in a simple scenario.</span></span>

<span data-ttu-id="ff2fa-106">Puede obtener más información sobre las API de GItHub que se usan en la referencia GitHub [API](https://docs.github.com/rest/reference/repos#list-repositories-for-a-user).</span><span class="sxs-lookup"><span data-stu-id="ff2fa-106">You can learn more about the GItHub APIs being used in the [GitHub API reference](https://docs.github.com/rest/reference/repos#list-repositories-for-a-user).</span></span> <span data-ttu-id="ff2fa-107">También puede ver el resultado de la llamada api sin procesar visitando en un explorador web (asegúrese de reemplazar el marcador `https://api.github.com/users/{USERNAME}/repos` de posición {USERNAME} por su id. de Github).</span><span class="sxs-lookup"><span data-stu-id="ff2fa-107">You can also see the raw API call output by visiting `https://api.github.com/users/{USERNAME}/repos` in a web browser (be sure to replace the {USERNAME} placeholder with your Github ID).</span></span>

![Ejemplo de obtener información de repositorios](../../images/git.png)

## <a name="sample-code-get-basic-information-about-users-github-repositories"></a><span data-ttu-id="ff2fa-109">Código de ejemplo: obtener información básica sobre los repositorios de GitHub usuario</span><span class="sxs-lookup"><span data-stu-id="ff2fa-109">Sample code: Get basic information about user's GitHub repositories</span></span>

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

## <a name="training-video-how-to-make-external-api-calls"></a><span data-ttu-id="ff2fa-110">Vídeo de aprendizaje: Cómo realizar llamadas a API externas</span><span class="sxs-lookup"><span data-stu-id="ff2fa-110">Training video: How to make external API calls</span></span>

<span data-ttu-id="ff2fa-111">[![Ver vídeo sobre cómo realizar llamadas a API externas](../../images/api-vid.png)](https://youtu.be/fulP29J418E "Vídeo sobre cómo realizar llamadas a API externas")</span><span class="sxs-lookup"><span data-stu-id="ff2fa-111">[![Watch video on how to make external API calls](../../images/api-vid.png)](https://youtu.be/fulP29J418E "Video on how to make external API calls")</span></span>
