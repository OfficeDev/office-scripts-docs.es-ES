---
title: Admitir scripts de Office antiguos que usan las API asincrónicas
description: Una introducción a las API asincrónicas de Scripts de Office y cómo usar el patrón de carga y sincronización para scripts más antiguos.
ms.date: 02/08/2021
localization_priority: Normal
ms.openlocfilehash: 143f52a7ffefb4f19ee36ba4343fd7c2f1cbdffe
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 04/14/2021
ms.locfileid: "51755080"
---
# <a name="support-older-office-scripts-that-use-the-async-apis"></a>Admitir scripts de Office antiguos que usan las API asincrónicas

Este artículo le enseñará a mantener y actualizar scripts que usan las API asincrónicas del modelo anterior. Estas API tienen la misma funcionalidad principal que las API de scripts de Office sincrónicas estándar, pero requieren que el script controle la sincronización de datos entre el script y el libro.

> [!IMPORTANT]
> El modelo asincrónico solo se puede usar con scripts creados antes de la implementación del modelo [de API actual.](scripting-fundamentals.md) Los scripts se bloquean permanentemente en el modelo de API que tienen al crearse. Esto también significa que si desea convertir un script antiguo al nuevo modelo, debe crear un script nuevo. Se recomienda actualizar los scripts antiguos al nuevo modelo al realizar cambios, ya que el modelo actual es más fácil de usar. La [sección Convertir scripts asincrónicos en el modelo](#converting-async-scripts-to-the-current-model) actual ofrece consejos sobre cómo realizar esta transición.

## <a name="main-function"></a>`main` Función

Los scripts que usan las API asincrónicas tienen una función `main` diferente. Es una función `async` que tiene un como primer `Excel.RequestContext` parámetro.

```TypeScript
async function main(context: Excel.RequestContext) {
    // Your async Office Script
}
```

## <a name="context"></a>Context

La función `main` acepta un parámetro de `Excel.RequestContext`, denominado `context`. Considere `context` como el puente entre el script y el libro. El script obtiene acceso al libro con el objeto `context` y usa ese `context` para enviar datos hacia adelante y hacia atrás.

El objeto `context` es necesario porque el script y Excel se ejecutan en diferentes procesos y ubicaciones. El script tendrá que realizar cambios o consultar datos en el libro en la nube. El objeto `context` administra estas transacciones.

## <a name="sync-and-load"></a>Sync y Load

Como el script y el libro se ejecutan en distintas ubicaciones, cualquier transferencia de datos entre ambos necesita tiempo. En la API asincrónica, los comandos se ponen en cola hasta que el script llama explícitamente a la operación `sync` para sincronizar el script y el libro. El script puede funcionar de forma independiente hasta que necesite realizar cualquiera de las siguientes acciones:

- Lea los datos del libro (después de una operación `load` o método que devuelve un [ClientResult](/javascript/api/office/officeextension.clientresult?view=excel-js-online&preserve-view=true)).
- Escribir datos en el libro (por lo general, porque el script ha terminado).

En la imagen siguiente se muestra un ejemplo de flujo de control entre el script y el libro:

:::image type="content" source="../images/load-sync.png" alt-text="Diagrama en el que se muestran las operaciones de lectura y escritura en el libro desde el script.":::

### <a name="sync"></a>Sincronizar

Siempre que el script asincrónico necesite leer o escribir datos en el libro, llame al método como `RequestContext.sync` se muestra aquí:

```TypeScript
await context.sync();
```

> [!NOTE]
> Se llama de forma implícita a `context.sync()` cuando finaliza un script.

Una vez completada la operación `sync`, el libro se actualiza para reflejar las operaciones de escritura que haya especificado el script. Una operación de escritura consiste en establecer cualquier propiedad en un objeto de Excel (por ejemplo, ) o llamar a un método que cambie una propiedad `range.format.fill.color = "red"` (por ejemplo, `range.format.autoFitColumns()` ). La operación `sync` también lee cualquier valor del libro solicitado por el script mediante una operación `load` o un método que devuelve un `ClientResult`(como se describe en la sección siguiente).

Sincronizar el script con el libro puede tardar un tiempo, según la red. Minimice el número de `sync` llamadas para ayudar a que el script se ejecute rápidamente. De lo contrario, las API asincrónicas no son más rápidas que las API sincrónicas estándar.

### <a name="load"></a>Load

Un script asincrónico debe cargar datos del libro antes de leerlo. Sin embargo, la carga de datos de todo el libro reduciría considerablemente la velocidad del script. El `load` método permite al script especificar específicamente qué datos se deben recuperar del libro.

El método `load` está disponible en cada objeto de Excel. El script debe cargar las propiedades de un objeto antes de poder leerlas. Si no lo hace, se producirá un error.

Los ejemplos siguientes usan un objeto `Range` para mostrar las tres formas en que se puede usar el método `load` para cargar datos.

|Objetivo |Comando de ejemplo | Efecto |
|:--|:--|:--|
|Cargar una propiedad |`myRange.load("values");` | Carga una única propiedad, en este caso la matriz bidimensional de valores en este rango. |
|Cargar varias propiedades |`myRange.load("values, rowCount, columnCount");`| Carga todas las propiedades de una lista delimitada por comas, en este ejemplo, los valores, el número de filas y el número de columnas. |
|Cargar todo | `myRange.load();`|Carga todas las propiedades en el rango. Esta no es una solución recomendada, ya que ralentizará el script obteniendo datos innecesarios. Solo úse esto mientras prueba el script o si necesita todas las propiedades del objeto. |

El script debe llamar a `context.sync()` antes de leer cualquier valor cargado.

```TypeScript
/**
 * This script uses the async API to get the row count for a range.
 * It shows how to load a property in the async model.
 */
async function main(context: Excel.RequestContext) {
    let selectedSheet = context.workbook.worksheets.getActiveWorksheet();
    let range = selectedSheet.getRange("A1:B3");

    // Load the property.
    range.load("rowCount");

    // Synchronize with the workbook to get the property.
    await context.sync();

    // Read and log the property value (3).
    console.log(range.rowCount);
}
```

También puede cargar propiedades de toda la colección. Cada objeto de colección de la API asincrónica tiene una propiedad `items` que es una matriz que contiene los objetos de esa colección. El uso de `items` como inicio de una llamada jerárquica (`items\myProperty`) a `load` carga las propiedades especificadas en cada uno de esos elementos. El ejemplo siguiente carga la propiedad `resolved` en cada objeto `Comment` del objeto `CommentCollection` de una hoja de cálculo.

```TypeScript
/**
 * This script uses the async API to get resolved property on every comment in the worksheet.
 * It shows how to load a property from every object in a collection.
 */
async function main(context: Excel.RequestContext){
    let selectedSheet = context.workbook.worksheets.getActiveWorksheet();
    let comments = selectedSheet.comments;

    // Load the `resolved` property from every comment in this collection.
    comments.load("items/resolved");

    // Synchronize with the workbook to get the properties.
    await context.sync();
}
```

### <a name="clientresult"></a>ClientResult

Los métodos de la API asincrónica que devuelven información del libro tienen un patrón similar al `load` / `sync` paradigma. Por ejemplo, `TableCollection.getCount` obtiene el número de tablas de la colección. `getCount` devuelve un `ClientResult<number>`, lo que significa que la propiedad `value` en el [`ClientResult`](/javascript/api/office/officeextension.clientresult?view=excel-js-online&preserve-view=true) de retorno es un número. El script no puede acceder a ese valor hasta que se llama a `context.sync()`. De forma muy similar a la carga de una propiedad, el `value` es un valor local "vacío" hasta esa llamada `sync`.

El siguiente script obtiene el número total de tablas en el libro y registra ese número en la consola.

```TypeScript
/**
 * This script uses the async API to get the table count of the workbook.
 * It shows how ClientResult objects return workbook information.
 */
async function main(context: Excel.RequestContext) {
    let tableCount = context.workbook.tables.getCount();

    // This sync call implicitly loads tableCount.value.
    // Any other ClientResult values are loaded too.
    await context.sync();

    // Trying to log the value before calling sync would throw an error.
    console.log(tableCount.value);
}
```

## <a name="converting-async-scripts-to-the-current-model"></a>Convertir scripts asincrónicos al modelo actual

El modelo de API actual no usa `load` , `sync` o un `RequestContext` . Esto hace que los scripts sean mucho más fáciles de escribir y mantener. El mejor recurso para convertir scripts antiguos es [Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts). Allí, puede solicitar ayuda a la comunidad con escenarios específicos. La siguiente guía debe ayudar a describir los pasos generales que debe seguir.

1. Cree un nuevo script y copie el código asincrónico antiguo en él. Asegúrese de no incluir la firma del método `main` antiguo, usando la actual en `function main(workbook: ExcelScript.Workbook)` su lugar.

2. Quite todas las `load` llamadas `sync` y. Ya no son necesarias.

3. Se han quitado todas las propiedades. Ahora tiene acceso a esos objetos a través de y métodos, por lo que tendrá que cambiar esas referencias `get` de propiedad a llamadas a `set` métodos. Por ejemplo, en lugar de establecer el color de relleno de una celda a través del acceso a la propiedad de este modo: , ahora usará `mySheet.getRange("A2:C2").format.fill.color = "blue";` métodos como este: `mySheet.getRange("A2:C2").getFormat().getFill().setColor("blue");`

4. Las clases de colección se han reemplazado por matrices. Los métodos y de esas clases de colección se movieron al objeto que era propietario de la colección, por lo que las referencias `add` `get` deben actualizarse en consecuencia. Por ejemplo, para obtener un gráfico denominado "MyChart" de la primera hoja de cálculo del libro, use el siguiente código: `workbook.getWorksheets()[0].getChart("MyChart");` . Tenga en `[0]` cuenta el acceso al primer valor del devuelto por `Worksheet[]` `getWorksheets()` .

5. Algunos métodos se cambiaron por claridad y se agregaron para mayor comodidad. Consulte la referencia de la [API de scripts de Office](/javascript/api/office-scripts/overview) para obtener más información.

## <a name="office-scripts-async-api-reference-documentation"></a>Documentación de referencia de api asincrónica de Scripts de Office

Las API asincrónicas son equivalentes a las que se usan en los complementos de Office. La documentación de referencia se encuentra en la sección Excel de la referencia de la API de [JavaScript de](/javascript/api/excel?view=excel-js-online&preserve-view=true)complementos de Office.
