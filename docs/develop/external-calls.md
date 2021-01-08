---
title: Soporte de llamadas de API externas en Scripts de Office
description: Soporte técnico e instrucciones para realizar llamadas a API externas en un script de Office.
ms.date: 01/05/2021
localization_priority: Normal
ms.openlocfilehash: 1091031bc2e12f3e1e79b177c69874ee4ce61dd8
ms.sourcegitcommit: 30c4b731dc8d18fca5aa74ce59e18a4a63eb4ffc
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 01/08/2021
ms.locfileid: "49784147"
---
# <a name="external-api-call-support-in-office-scripts"></a>Soporte de llamadas de API externas en Scripts de Office

Los autores de scripts no deben esperar un comportamiento coherente al usar API externas [durante](https://developer.mozilla.org/docs/Web/API) la fase de vista previa de la plataforma. Por lo tanto, no confíe en api externas para escenarios de script críticos.

Las llamadas a API externas solo se pueden realizar a través de la aplicación de Excel, no a través de Power Automate [en circunstancias normales.](#external-calls-from-power-automate)

> [!CAUTION]
> Las llamadas externas pueden dar lugar a la exposición de datos confidenciales a extremos no deseados. El administrador puede establecer una protección de firewall frente a dichas llamadas.

## <a name="working-with-fetch"></a>Trabajar con `fetch`

La [API de recuperación](https://developer.mozilla.org/docs/Web/API/Fetch_API) recupera información de servicios externos. Es una `async` API, por lo que tendrá que ajustar la `main` firma del script. Hacer que `main` la función y hacer que devuelva un `async` `Promise<void>` . También debe asegurarse de que la `await` llamada y la `fetch` `json` recuperación. Esto garantiza que esas operaciones se completen antes de que finalice el script.

El siguiente script usa `fetch` para recuperar datos JSON del servidor de prueba en la dirección URL determinada.

```typescript
async function main(workbook: ExcelScript.Workbook): Promise <void> {
  /* 
   * Retrieve JSON data from a test server.
   */
  let fetchResult = await fetch('https://jsonplaceholder.typicode.com/todos/1');
  let json = await fetchResult.json();

  // Displays the content from https://jsonplaceholder.typicode.com/todos/1
  console.log(JSON.stringify(json));
}
```

El escenario de ejemplo de scripts de Office: los datos de nivel de agua de Graph de [NOAA](../resources/scenarios/noaa-data-fetch.md) muestran el comando de captura que se usa para recuperar registros de la base de datos National Oceanic and Desplomeys and Currents.

## <a name="external-calls-from-power-automate"></a>Llamadas externas de Power Automate

Si se ejecuta un script con Power Automate, se producirá un error en las llamadas a API externas. Esta es una diferencia de comportamiento entre ejecutar un script a través del cliente de Excel y a través de Power Automate. Asegúrese de comprobar las referencias de los scripts antes de crearlas en un flujo.

> [!WARNING]
> Se producirá un error en las llamadas externas realizadas a través del conector de Power Automate [Excel Online](/connectors/excelonlinebusiness) para ayudar a evitar las directivas de prevención de pérdida de datos existentes. Sin embargo, los scripts que se ejecutan a través de Power Automate se realizan fuera de la organización y fuera de los firewalls de la organización. Para obtener protección adicional contra usuarios malintencionados en este entorno externo, el administrador puede controlar el uso de scripts de Office. El administrador puede deshabilitar el conector de Excel Online en Power Automate o desactivar scripts de Office para Excel en la web a través de los controles de administrador [de scripts de Office.](/microsoft-365/admin/manage/manage-office-scripts-settings)

## <a name="see-also"></a>Vea también

- [Usar objetos integrados de JavaScript en los scripts de Office](javascript-objects.md)
- [Escenario de ejemplo de scripts de Office: datos de nivel de agua de Gráfico de NOAA](../resources/scenarios/noaa-data-fetch.md)
