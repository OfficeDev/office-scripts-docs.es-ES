---
title: Usar llamadas de captura externa en Scripts de Office
description: Obtenga información sobre cómo realizar llamadas API externas en scripts de Office.
ms.date: 06/10/2022
ms.localizationpriority: medium
ms.openlocfilehash: 569d74f1ca8996cd8fe8a4ba3163445d57676d27
ms.sourcegitcommit: dd01979d34b3499360d2f79a56f8a8f24f480eed
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 06/15/2022
ms.locfileid: "66088095"
---
# <a name="use-external-fetch-calls-in-office-scripts"></a>Usar llamadas de captura externa en Scripts de Office

Este script obtiene información básica sobre los repositorios de GitHub de un usuario. Muestra cómo usar `fetch` en un escenario sencillo. Para obtener más información sobre el uso `fetch` de u otras llamadas externas, consulte [Compatibilidad con llamadas API externas en scripts de Office](../../develop/external-calls.md). Para obtener información sobre cómo trabajar con objetos [JSON]](https://www.w3schools.com/whatis/whatis_json.asp), como lo que devuelven las API de GitHub, lea [Uso de JSON para pasar datos hacia y desde scripts de Office](../../develop/use-json.md).

Obtenga más información sobre las API de GItHub que se usan en la [referencia de GitHub API](https://docs.github.com/rest/reference/repos#list-repositories-for-a-user). También puede ver la salida de la llamada API sin procesar visitando `https://api.github.com/users/{USERNAME}/repos` en un explorador web (asegúrese de reemplazar el marcador de posición {USERNAME} por el identificador de GitHub).

![Ejemplo de información de obtención de repositorios](../../images/git.png)

## <a name="sample-code-get-basic-information-about-users-github-repositories"></a>Código de ejemplo: obtener información básica sobre los repositorios de GitHub del usuario

```TypeScript
async function main(workbook: ExcelScript.Workbook) {
  // Call the GitHub REST API.
  // Replace the {USERNAME} placeholder with your GitHub username.
  const response = await fetch('https://api.github.com/users/{USERNAME}/repos');
  const repos: Repository[] = await response.json();

  // Create an array to hold the returned values.
  const rows: (string | boolean | number)[][] = [];

  // Convert each repository block into a row.
  for (let repo of repos) {
    rows.push([repo.id, repo.name, repo.license?.name, repo.license?.url]);
  }
  // Create a header row.
  const sheet = workbook.getActiveWorksheet();
  sheet.getRange('A1:D1').setValues([["ID", "Name", "License Name", "License URL"]]);

  // Add the data to the current worksheet, starting at "A2".
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

## <a name="training-video-how-to-make-external-api-calls"></a>Vídeo de entrenamiento: Cómo realizar llamadas API externas

[Vea cómo Sudhi Ramamurthy recorre este ejemplo en YouTube](https://youtu.be/fulP29J418E).
