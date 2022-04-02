---
title: Usar llamadas de captura externa en Scripts de Office
description: Obtenga información sobre cómo realizar llamadas API externas en Office scripts.
ms.date: 05/14/2021
ms.localizationpriority: medium
ms.openlocfilehash: feff9d49f9f50f14fd83b1864568df8dab02d417
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585530"
---
# <a name="use-external-fetch-calls-in-office-scripts"></a>Usar llamadas de captura externa en Scripts de Office

Este script obtiene información básica sobre los repositorios de GitHub usuario. Muestra cómo usar en `fetch` un escenario simple. Para obtener más información sobre el uso u `fetch` otras llamadas externas, lea [Soporte de llamadas de API externa en Office scripts](../../develop/external-calls.md)

Puede obtener más información sobre las API de GItHub que se usan en la referencia GitHub [API.](https://docs.github.com/rest/reference/repos#list-repositories-for-a-user) También puede ver el resultado de la llamada api `https://api.github.com/users/{USERNAME}/repos` sin procesar visitando en un explorador web (asegúrese de reemplazar el marcador de posición {USERNAME} por su GitHub id.).

![Ejemplo de obtener información de repositorios](../../images/git.png)

## <a name="sample-code-get-basic-information-about-users-github-repositories"></a>Código de ejemplo: obtener información básica sobre los repositorios de GitHub usuario

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

## <a name="training-video-how-to-make-external-api-calls"></a>Vídeo de aprendizaje: Cómo realizar llamadas a API externas

[Vea el recorrido de Sudhi Ramamurthy a través de esta muestra en YouTube](https://youtu.be/fulP29J418E).
