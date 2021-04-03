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
# <a name="external-api-calls-from-office-scripts"></a>Llamadas de API externas desde scripts de Office

Los scripts de Office [permiten una compatibilidad limitada con llamadas de API externas.](../../develop/external-calls.md)

> [!IMPORTANT]
>
> * No hay forma de iniciar sesión o usar el tipo OAuth2 de flujos de autenticación. Todas las claves y credenciales deben codificarse de forma rígida (o leerse desde otro origen).
> * No hay ninguna infraestructura para almacenar credenciales y claves de API. El usuario tendrá que administrarlo.
> * Las llamadas externas pueden provocar que los datos confidenciales se exponán a extremos no deseados o que los datos externos se puedan incluir en libros internos. El administrador puede establecer la protección del firewall frente a estas llamadas. Asegúrese de comprobar con las directivas locales antes de confiar en llamadas externas.
> * Si un script usa una llamada a la API, no funcionará en un escenario de Power Automate. Tendrás que usar la acción HTTP de Power Automate o acciones equivalentes para extraer datos de un servicio externo o insertarlo en él.
> * Una llamada de API externa implica una sintaxis asincrónica de API y requiere conocimientos ligeramente avanzados de cómo funciona la comunicación asincrónica.
> * Asegúrese de comprobar la cantidad de rendimiento de datos antes de tomar una dependencia. Por ejemplo, extraer todo el conjunto de datos externo puede no ser la mejor opción y, en su lugar, se debe usar la paginación para obtener datos en fragmentos.

## <a name="useful-knowledge-and-resources"></a>Conocimientos y recursos útiles

* [API de REST:](https://en.wikipedia.org/wiki/Representational_state_transfer)es muy probable que use la llamada a la API.
* [ `async` : Comprenda cómo funciona esto. `await` ](https://developer.mozilla.org/docs/Learn/JavaScript/Asynchronous/Async_await)
* [`fetch`](https://developer.mozilla.org/docs/Web/API/Fetch_API/Using_Fetch): Comprenda cómo funciona esto.

## <a name="steps"></a>Pasos

1. Marca la `main` función como una función asincrónica agregando `async` prefijo. Por ejemplo, `async function main(workbook: ExcelScript.Workbook)`.
1. ¿Qué tipo de llamada api realiza? `GET`, `POST`, `PUT`, `DELETE`, `PATCH`? Consulte el material de la API de REST para obtener más información.
1. Obtenga el punto de conexión de la API de servicio, los requisitos de autenticación, los encabezados, etc.
1. Defina la entrada o salida para `interface` ayudar con la finalización del código y la comprobación del tiempo de desarrollo. Vea [el vídeo](#training-video-how-to-make-external-api-calls) para obtener más información.
1. Código, prueba, optimiza. Puede crear una función para la rutina de llamada de la API para que sea reutilizable desde otras partes del script o para su reutilización en un script diferente (copiar y pegar se vuelve mucho más fácil de esta manera).

## <a name="scenario"></a>Escenario

Este script obtiene información básica sobre los repositorios de GitHub del usuario.

![Ejemplo de obtener información de repositorios](../../images/git.png)

## <a name="resources-used-in-the-sample"></a>Recursos usados en el ejemplo

1. [Obtener la referencia de la API de Github de repositorios.](https://docs.github.com/rest/reference/repos#list-repositories-for-a-user)
1. Salida de llamada api: vaya a un explorador web o a cualquier interfaz HTTP y escriba , reemplazando el marcador `https://api.github.com/users/{USERNAME}/repos` de posición {USERNAME} por su id. de Github.
1. Información recuperada: repo.name, repo.size, repo.owner.id, repo.license?. nombre

## <a name="sample-code-get-basic-information-about-users-github-repositories"></a>Código de ejemplo: obtener información básica sobre los repositorios de GitHub del usuario

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

## <a name="training-video-how-to-make-external-api-calls"></a>Vídeo de aprendizaje: Cómo realizar llamadas a API externas

[![Ver vídeo sobre cómo realizar llamadas a API externas](../../images/api-vid.png)](https://youtu.be/fulP29J418E "Vídeo sobre cómo realizar llamadas a API externas")
