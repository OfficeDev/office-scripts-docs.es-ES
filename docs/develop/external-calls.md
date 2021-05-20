---
title: Soporte de llamadas de API externas en Scripts de Office
description: Soporte y orientación para realizar llamadas a la API externas en un script de Office.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: fd6ba0c57bf4cabb2d07421355cacff373f6706c
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545085"
---
# <a name="external-api-call-support-in-office-scripts"></a>Soporte de llamadas de API externas en Scripts de Office

Los autores de scripts no deben esperar un comportamiento coherente al usar [API externas](https://developer.mozilla.org/docs/Web/API) durante la fase de vista previa de la plataforma. Como tal, no confíe en apis externas para escenarios de script críticos.

Las llamadas a API externas solo se pueden realizar a través de la aplicación Excel, no a través de Power Automate [en circunstancias normales.](#external-calls-from-power-automate)

> [!CAUTION]
> Las llamadas externas pueden dar lugar a que los datos confidenciales se expongan a puntos de conexión indeseables. El administrador puede establecer protección de firewall contra dichas llamadas.

## <a name="configure-your-script-for-external-calls"></a>Configure el script para llamadas externas

Las llamadas externas son [asincrónicas](https://developer.mozilla.org/docs/Learn/JavaScript/Asynchronous/Async_await) y requieren que el script esté marcado como `async` . Agregue el `async` prefijo a su `main` función y haga que devuelva un , tal y como se muestra `Promise` aquí:

```typescript
async function main(workbook: ExcelScript.Workbook) : Promise <void>
```

> [!NOTE]
> Los scripts que devuelven otra información pueden devolver una `Promise` de ese tipo. Por ejemplo, si el script necesita devolver un `Employee` objeto, la firma de devolución sería `: Promise <Employee>`

Deberá aprender las interfaces del servicio externo para realizar llamadas a ese servicio. Si está utilizando `fetch` o las API de [REST,](https://wikipedia.org/wiki/Representational_state_transfer)debe determinar la estructura JSON de los datos devueltos. Para obtener la entrada y la salida del script, considere la posibilidad de hacer una `interface` coincidencia con las estructuras JSON necesarias. Esto le da al script más seguridad de tipo. Puede ver un ejemplo de esto en [Uso de captura de scripts de Office](../resources/samples/external-fetch-calls.md).

### <a name="limitations-with-external-calls-from-office-scripts"></a>Limitaciones con llamadas externas de scripts de Office

* No hay forma de iniciar sesión ni usar el tipo de flujos de autenticación OAuth2. Todas las claves y credenciales deben codificarse de forma rígida (o leerse desde otro origen).
* No hay infraestructura para almacenar credenciales y claves de API. Esto tendrá que ser administrado por el usuario.
* No se admiten cookies de documentos `localStorage` ni `sessionStorage` objetos. 
* Las llamadas externas pueden dar lugar a que los datos confidenciales se expongan a puntos de conexión indeseables o a que los datos externos se insistan en libros internos. El administrador puede establecer protección de firewall contra dichas llamadas. Asegúrese de comprobar con directivas locales antes de confiar en llamadas externas.
* Asegúrese de comprobar la cantidad de rendimiento de los datos antes de tomar una dependencia. Por ejemplo, tirar hacia abajo todo el conjunto de datos externo puede no ser la mejor opción y en su lugar la paginación debe usarse para obtener datos en fragmentos.

## <a name="retrieve-information-with-fetch"></a>Recuperar información con `fetch`

La [API fetch](https://developer.mozilla.org/docs/Web/API/Fetch_API) recupera información de servicios externos. Es una `async` API, por lo que debe ajustar la `main` firma del script. Haga la `main` función y haga que devuelva un archivo `async` `Promise<void>` . También debe estar seguro de `await` la llamada y la `fetch` `json` recuperación. Esto garantiza que esas operaciones se completen antes de que finalice el script.

Los datos JSON recuperados por `fetch` deben coincidir con una interfaz definida en el script. El valor devuelto debe asignarse a un tipo específico porque [Office Scripts no admiten el `any` tipo](typescript-restrictions.md#no-any-type-in-office-scripts). Debe hacer referencia a la documentación del servicio para ver cuáles son los nombres y tipos de las propiedades devueltas. A continuación, agregue la interfaz o las interfaces coincidentes al script.

El siguiente script se utiliza `fetch` para recuperar datos JSON del servidor de prueba en la dirección URL especificada. Anote la `JSONData` interfaz para almacenar los datos como un tipo coincidente.

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

* El [ejemplo Usar llamadas de captura externas en scripts de Office](../resources/samples/external-fetch-calls.md) muestra cómo obtener información básica sobre los repositorios de GitHub de un usuario.
* El [escenario de ejemplo Office Scripts: Graph datos a nivel de agua de la NOAA](../resources/scenarios/noaa-data-fetch.md) muestra el comando fetch que se utiliza para recuperar registros de la base de datos Tides and Currents de la Administración Nacional Oceánica y Atmosférica.

## <a name="external-calls-from-power-automate"></a>Llamadas externas desde Power Automate

Cualquier llamada a la API externa falla cuando se ejecuta un script con Power Automate. Esta es una diferencia de comportamiento entre ejecutar un script a través de la aplicación Excel y a través de Power Automate. Asegúrese de comprobar los scripts en busca de dichas referencias antes de crearlas en un flujo.

Tendrá que usar [HTTP con Azure AD](/connectors/webcontents/) u otras acciones equivalentes para extraer datos o insertarlo en un servicio externo.

> [!WARNING]
> Las llamadas externas realizadas a través de la Power Automate [Excel conector en línea](/connectors/excelonlinebusiness) fallan con el fin de ayudar a mantener las políticas de prevención de pérdida de datos existentes. Sin embargo, los scripts que se ejecutan a través de Power Automate se realizan fuera de su organización y fuera de los firewalls de su organización. Para obtener protección adicional contra usuarios malintencionados en este entorno externo, el administrador puede controlar el uso de scripts de Office. El administrador puede deshabilitar el conector en línea Excel en Power Automate o desactivar scripts Office para Excel en la Web a través de los [controles de administrador de scripts de Office.](/microsoft-365/admin/manage/manage-office-scripts-settings)

## <a name="see-also"></a>Vea también

* [Usar objetos integrados de JavaScript en los scripts de Office](javascript-objects.md)
* [Usar llamadas de captura externa en Scripts de Office](../resources/samples/external-fetch-calls.md)
* [Office Escenario de ejemplo de scripts: Graph datos a nivel del agua de la NOAA](../resources/scenarios/noaa-data-fetch.md)
