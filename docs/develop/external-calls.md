---
title: Soporte de llamadas de API externas en Scripts de Office
description: Soporte técnico e instrucciones para realizar llamadas API externas en un script Office script.
ms.date: 05/21/2021
ms.localizationpriority: medium
ms.openlocfilehash: 14b98e49907ff989684eceb9509edf56a1a72d9e
ms.sourcegitcommit: d3ed4bdeeba805d97c930394e172e8306a0cf484
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 09/15/2021
ms.locfileid: "59332055"
---
# <a name="external-api-call-support-in-office-scripts"></a>Soporte de llamadas de API externas en Scripts de Office

Los scripts admiten llamadas a servicios externos. Use estos servicios para proporcionar datos y otra información al libro.

> [!CAUTION]
> Las llamadas externas pueden provocar que los datos confidenciales se exponán a extremos no deseados. El administrador puede establecer la protección del firewall frente a estas llamadas.

> [!IMPORTANT]
> Las llamadas a API externas solo se pueden realizar a través de la aplicación Excel, no a través de Power Automate [en circunstancias normales](#external-calls-from-power-automate).

## <a name="configure-your-script-for-external-calls"></a>Configurar el script para llamadas externas

Las llamadas externas [son asincrónicas](https://developer.mozilla.org/docs/Learn/JavaScript/Asynchronous/Async_await) y requieren que el script esté marcado como `async` . Agregue el `async` prefijo a la función y haga que devuelva un , como se muestra `main` `Promise` aquí:

```typescript
async function main(workbook: ExcelScript.Workbook) : Promise <void>
```

> [!NOTE]
> Los scripts que devuelven otra información pueden devolver un `Promise` de ese tipo. Por ejemplo, si el script necesita devolver un `Employee` objeto, la firma de devolución sería `: Promise <Employee>`

Tendrás que aprender las interfaces del servicio externo para realizar llamadas a ese servicio. Si está usando o `fetch` [LAS API de REST,](https://wikipedia.org/wiki/Representational_state_transfer)debe determinar la estructura JSON de los datos devueltos. Para la entrada y salida desde el script, considere la posibilidad de realizar una para que `interface` coincida con las estructuras JSON necesarias. Esto proporciona al script más seguridad de tipo. Puede ver un ejemplo de esto en [Using fetch from Office Scripts](../resources/samples/external-fetch-calls.md).

### <a name="limitations-with-external-calls-from-office-scripts"></a>Limitaciones con llamadas externas desde Office scripts

* No hay forma de iniciar sesión o usar el tipo OAuth2 de flujos de autenticación. Todas las claves y credenciales deben codificarse de forma rígida (o leerse desde otro origen).
* No hay ninguna infraestructura para almacenar credenciales y claves de API. El usuario tendrá que administrarlo.
* No se admiten las cookies de documento `localStorage` `sessionStorage` ni los objetos.
* Las llamadas externas pueden provocar que los datos confidenciales se exponán a extremos no deseados o que los datos externos se puedan incluir en libros internos. El administrador puede establecer la protección del firewall frente a estas llamadas. Asegúrese de comprobar con las directivas locales antes de confiar en llamadas externas.
* Asegúrese de comprobar la cantidad de rendimiento de datos antes de tomar una dependencia. Por ejemplo, extraer todo el conjunto de datos externo puede no ser la mejor opción y, en su lugar, se debe usar la paginación para obtener datos en fragmentos.

## <a name="retrieve-information-with-fetch"></a>Recuperar información con `fetch`

La [API de captura](https://developer.mozilla.org/docs/Web/API/Fetch_API) recupera información de servicios externos. Es una `async` API, por lo que debe ajustar la `main` firma del script. Haga que `main` la función y haga que devuelva un `async` `Promise<void>` . También debe asegurarse de la llamada `await` `fetch` y la `json` recuperación. Esto garantiza que las operaciones se completen antes de que finalice el script.

Los datos JSON recuperados por `fetch` deben coincidir con una interfaz definida en el script. El valor devuelto debe asignarse a un tipo específico porque Office [scripts no admiten el `any` tipo](typescript-restrictions.md#no-any-type-in-office-scripts). Debe consultar la documentación del servicio para ver cuáles son los nombres y tipos de las propiedades devueltas. A continuación, agregue la interfaz o las interfaces correspondientes al script.

El siguiente script usa `fetch` para recuperar datos JSON del servidor de prueba en la dirección URL determinada. Tenga en `JSONData` cuenta la interfaz para almacenar los datos como un tipo de coincidencia.

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

### <a name="other-fetch-samples"></a>Otras `fetch` muestras

* El [ejemplo Use external fetch calls in Office Scripts](../resources/samples/external-fetch-calls.md) muestra cómo obtener información básica sobre los repositorios de GitHub usuario.
* Escenario de ejemplo de scripts de Office: Graph datos de nivel de agua de [NOAA](../resources/scenarios/noaa-data-fetch.md) muestra el comando de captura que se usa para recuperar registros de la base de datos de corrientes y mareos de la Administración oceánica y atmosférica nacional.

## <a name="external-calls-from-power-automate"></a>Llamadas externas desde Power Automate

Cualquier llamada de API externa produce un error cuando se ejecuta un script con Power Automate. Esta es una diferencia de comportamiento entre ejecutar un script a través de la aplicación Excel y a través de Power Automate. Asegúrese de comprobar las referencias de los scripts antes de crearlas en un flujo.

Tendrás que usar [HTTP](/connectors/webcontents/) con Azure AD u otras acciones equivalentes para extraer datos de un servicio externo o insertarlo en él.

> [!WARNING]
> Las llamadas externas realizadas a través del conector [Power Automate Excel Online fallan](/connectors/excelonlinebusiness) para ayudar a mantener las directivas de prevención de pérdida de datos existentes. Sin embargo, los scripts que se ejecutan Power Automate se realizan fuera de la organización y fuera de los firewalls de la organización. Para obtener protección adicional contra usuarios malintencionados en este entorno externo, el administrador puede controlar el uso de Office scripts. El administrador puede deshabilitar el conector de Excel Online en Power Automate o desactivar scripts de Office para Excel en la Web a través de los controles de administrador [de scripts](/microsoft-365/admin/manage/manage-office-scripts-settings)Office.

## <a name="see-also"></a>Ver también

* [Usar objetos integrados de JavaScript en los scripts de Office](javascript-objects.md)
* [Usar llamadas de captura externa en Scripts de Office](../resources/samples/external-fetch-calls.md)
* [Office Escenario de ejemplo de scripts: Graph datos de nivel de agua de NOAA](../resources/scenarios/noaa-data-fetch.md)
