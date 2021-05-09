---
title: Usar llamadas de captura externa en Scripts de Office
description: Obtenga información sobre cómo realizar llamadas API externas en Office scripts.
ms.date: 05/06/2021
localization_priority: Normal
ms.openlocfilehash: e8f46f552dee2c1ea43a321c968b00f02ffba49a
ms.sourcegitcommit: 763d341857bcb209b2f2c278a82fdb63d0e18f0a
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 05/08/2021
ms.locfileid: "52285825"
---
# <a name="use-external-fetch-calls-in-office-scripts"></a>Usar llamadas de captura externa en Scripts de Office

Este script obtiene información básica sobre los repositorios de GitHub usuario. Muestra cómo usar en `fetch` un escenario simple.

Puede obtener más información sobre las API de GItHub que se usan en la referencia GitHub [API](https://docs.github.com/rest/reference/repos#list-repositories-for-a-user). También puede ver el resultado de la llamada api sin procesar visitando en un explorador web (asegúrese de reemplazar el marcador `https://api.github.com/users/{USERNAME}/repos` de posición {USERNAME} por su id. de Github).

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

[Vea el recorrido de Sudhi Ramamurthy por este ejemplo en YouTube](https://youtu.be/fulP29J418E).
