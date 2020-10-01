---
title: Soporte de llamadas de API externas en Scripts de Office
description: Soporte técnico y guía para realizar llamadas a API externas en un script de Office.
ms.date: 09/24/2020
localization_priority: Normal
ms.openlocfilehash: fa77e606e2b3ab90144507660d71561b278e82e5
ms.sourcegitcommit: ce72354381561dc167ea0092efd915642a9161b3
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 09/30/2020
ms.locfileid: "48319633"
---
# <a name="external-api-call-support-in-office-scripts"></a>Soporte de llamadas de API externas en Scripts de Office

La plataforma de scripts de Office no admite llamadas a [API externas](https://developer.mozilla.org/docs/Web/API). Sin embargo, estas llamadas pueden ejecutarse en las circunstancias adecuadas. Las llamadas externas solo pueden realizarse a través del cliente de Excel, no a través de la automatización de la energía [en circunstancias normales](#external-calls-from-power-automate).

Los autores de scripts no deberían esperar un comportamiento coherente al usar las API externas durante la fase de vista previa de la plataforma. Esto se debe a la forma en que el tiempo de ejecución de JavaScript administra la interacción con el libro. El script puede finalizar antes de que finalice la llamada a la API (o de que `Promise` se haya resuelto completamente). Por lo tanto, no confíe en las API externas en escenarios de script críticos.

> [!CAUTION]
> Las llamadas externas pueden dar lugar a que los datos confidenciales se expongan a extremos no deseados. El administrador puede establecer la protección del firewall en dichas llamadas.

## <a name="definition-files-for-external-apis"></a>Archivos de definición para API externas

Los archivos de definición para las API externas no se incluyen con los scripts de Office. El uso de estas API genera errores en tiempo de compilación para las definiciones que faltan. Las API se siguen ejecutando (aunque solo cuando se ejecutan mediante el cliente de Excel), como se muestra en el siguiente script:

```typescript
async function main(workbook: ExcelScript.Workbook): Promise <void> {
  /* The following line of code generates the error:
   * "Cannot find name 'fetch'".
   * It will still run and return the JSON from the testing service.
   */
  let fetchResult = await fetch('https://jsonplaceholder.typicode.com/todos/1');
  let json = await fetchResult.json();

  // Displays the content from https://jsonplaceholder.typicode.com/todos/1
  console.log(JSON.stringify(json));
}
```

## <a name="external-calls-from-power-automate"></a>Llamadas externas de la automatización de la alimentación

Se produce un error en cualquier llamada de API externa cuando se ejecuta un script con la automatización de energía. Esta es una diferencia de comportamiento entre la ejecución de un script a través del cliente de Excel y a través de la automatización de la energía. Asegúrese de comprobar las secuencias de comandos para las referencias antes de crearlas en un flujo.

> [!WARNING]
> El error de las llamadas externas [Excel online Connector](/connectors/excelonlinebusiness) en Power Automatically para ayudar a preservar las directivas de prevención de pérdida de datos existentes. Sin embargo, los scripts que se ejecutan mediante la automatización de la alimentación se realizan fuera de la organización y fuera de los firewalls de la organización. Para obtener protección adicional de usuarios malintencionados en este entorno externo, el administrador puede controlar el uso de scripts de Office. El administrador puede deshabilitar el conector de Excel online con Power automatice o desactivar los scripts de Office para Excel en la web a través de los [controles de administrador de scripts de Office](/microsoft-365/admin/manage/manage-office-scripts-settings).

## <a name="see-also"></a>Vea también

- [Usar objetos integrados de JavaScript en los scripts de Office](javascript-objects.md)