---
title: Soporte de llamadas de API externas en Scripts de Office
description: Compatibilidad e instrucciones para realizar llamadas API externas en un script de Office.
ms.date: 06/10/2022
ms.localizationpriority: medium
ms.openlocfilehash: b847400893184533c250ab99b640563ff0cbdb3e
ms.sourcegitcommit: dd01979d34b3499360d2f79a56f8a8f24f480eed
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 06/15/2022
ms.locfileid: "66088046"
---
# <a name="external-api-call-support-in-office-scripts"></a>Soporte de llamadas de API externas en Scripts de Office

Los scripts admiten llamadas a servicios externos. Use estos servicios para proporcionar datos y otra información al libro.

> [!CAUTION]
> Las llamadas externas pueden dar lugar a la exposición de datos confidenciales a puntos de conexión no deseados. El administrador puede establecer la protección del firewall frente a estas llamadas.

> [!IMPORTANT]
> Las llamadas a API externas solo se pueden realizar a través de la aplicación Excel, no a través de Power Automate [en circunstancias normales](#external-calls-from-power-automate).

## <a name="configure-your-script-for-external-calls"></a>Configuración del script para llamadas externas

Las llamadas externas son [asincrónicas](https://developer.mozilla.org/docs/Learn/JavaScript/Asynchronous/Async_await) y requieren que el script esté marcado como `async`. Agregue el prefijo a `main` la `async` función y haga que devuelva un `Promise`, como se muestra aquí:

```typescript
async function main(workbook: ExcelScript.Workbook) : Promise <void>
```

> [!NOTE]
> Los scripts que devuelven otra información pueden devolver un `Promise` de ese tipo. Por ejemplo, si el script necesita devolver un `Employee` objeto, la firma de devolución sería `: Promise <Employee>`

Tendrá que aprender las interfaces del servicio externo para realizar llamadas a ese servicio. Si usa `fetch` o [api rest](https://wikipedia.org/wiki/Representational_state_transfer), debe determinar la estructura JSON de los datos devueltos. Para la entrada y la salida del script, considere la posibilidad de crear una `interface` para que coincida con las estructuras JSON necesarias. Esto proporciona al script más seguridad de tipos. Puede ver un ejemplo de esto en [Uso de la captura de scripts de Office](../resources/samples/external-fetch-calls.md).

### <a name="limitations-with-external-calls-from-office-scripts"></a>Limitaciones con llamadas externas desde scripts de Office

* No hay ninguna manera de iniciar sesión o usar flujos de autenticación de tipo OAuth2. Todas las claves y credenciales deben codificarse de forma rígida (o leerlas desde otro origen).
* No hay ninguna infraestructura para almacenar las claves y las credenciales de API. El usuario tendrá que administrar esto.
* No se admiten las cookies, `localStorage`y los `sessionStorage` objetos de documento.
* Las llamadas externas pueden dar lugar a la exposición de datos confidenciales a puntos de conexión no deseados o a datos externos que se van a incluir en libros internos. El administrador puede establecer la protección del firewall frente a estas llamadas. Asegúrese de comprobar con las directivas locales antes de confiar en llamadas externas.
* Asegúrese de comprobar la cantidad de rendimiento de datos antes de tomar una dependencia. Por ejemplo, extraer todo el conjunto de datos externo puede no ser la mejor opción y, en su lugar, se debe usar la paginación para obtener datos en fragmentos.

## <a name="retrieve-information-with-fetch"></a>Recuperar información con `fetch`

La [API de captura](https://developer.mozilla.org/docs/Web/API/Fetch_API) recupera información de servicios externos. Es una `async` API, por lo que debe ajustar la `main` firma del script. Haga que la `main` función `async`. También debe asegurarse de la `await` `fetch` llamada y `json` recuperación. Esto garantiza que esas operaciones se completen antes de que finalice el script.

Los datos JSON recuperados por `fetch` deben coincidir con una interfaz definida en el script. El valor devuelto debe asignarse a un tipo específico porque [Office scripts no admiten el `any` tipo](typescript-restrictions.md#no-any-type-in-office-scripts). Debe consultar la documentación del servicio para ver cuáles son los nombres y tipos de las propiedades devueltas. A continuación, agregue la interfaz o las interfaces coincidentes al script.

El siguiente script usa `fetch` para recuperar datos JSON del servidor de prueba en la dirección URL especificada. Tenga en cuenta la `JSONData` interfaz para almacenar los datos como un tipo coincidente.

```TypeScript
async function main(workbook: ExcelScript.Workbook) {
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

### <a name="other-fetch-samples"></a>Otros `fetch` ejemplos

* El ejemplo [Uso de llamadas de captura externas en scripts de Office](../resources/samples/external-fetch-calls.md) muestra cómo obtener información básica sobre los repositorios de GitHub de un usuario.
* El [escenario de ejemplo Office Scripts: Graph datos de nivel de agua de NOAA](../resources/scenarios/noaa-data-fetch.md) muestra el comando fetch que se usa para recuperar registros de la base de datos Mareas y corrientes de la Administración Nacional Oceánica y Atmosférica.

## <a name="external-calls-from-power-automate"></a>Llamadas externas desde Power Automate

Se produce un error en cualquier llamada API externa cuando se ejecuta un script con Power Automate. Se trata de una diferencia de comportamiento entre ejecutar un script a través de la aplicación Excel y a través de Power Automate. Asegúrese de comprobar las referencias de los scripts antes de compilarlas en un flujo.

Tendrá que usar [HTTP con Azure AD](/connectors/webcontents/) u otras acciones equivalentes para extraer datos o insertarlos en un servicio externo.

> [!WARNING]
> Las llamadas externas realizadas a través de Power Automate [Excel conector en línea](/connectors/excelonlinebusiness) producen un error para ayudar a mantener las directivas de prevención de pérdida de datos existentes. Sin embargo, los scripts que se ejecutan a través de Power Automate se realizan fuera de la organización y fuera de los firewalls de la organización. Para una protección adicional frente a usuarios malintencionados en este entorno externo, el administrador puede controlar el uso de scripts de Office. El administrador puede deshabilitar el conector de Excel Online en Power Automate o desactivar Office scripts para Excel en la Web a través de [los controles de administrador de scripts de Office](/microsoft-365/admin/manage/manage-office-scripts-settings).

## <a name="see-also"></a>Consulte también

* [Uso de JSON para pasar datos hacia y desde scripts de Office](use-json.md)
* [Usar objetos integrados de JavaScript en los scripts de Office](javascript-objects.md)
* [Usar llamadas de captura externa en Scripts de Office](../resources/samples/external-fetch-calls.md)
* [Escenario de ejemplo de scripts de Office: Graph datos de nivel de agua de NOAA](../resources/scenarios/noaa-data-fetch.md)
