---
title: Compatibilidad con scripts de Office anteriores que usan las API asincrónicas
description: Un manual sobre las API asincrónicas de scripts de Office y cómo usar el patrón Load/Sync para scripts más antiguos.
ms.date: 07/08/2020
localization_priority: Normal
ms.openlocfilehash: e7ca5b276cff0e3a38bffc2af1541c0051cf5490
ms.sourcegitcommit: ebd1079c7e2695ac0e7e4c616f2439975e196875
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 07/17/2020
ms.locfileid: "45160463"
---
# <a name="support-older-office-scripts-that-use-the-async-apis"></a><span data-ttu-id="f3b6f-103">Compatibilidad con scripts de Office anteriores que usan las API asincrónicas</span><span class="sxs-lookup"><span data-stu-id="f3b6f-103">Support older Office Scripts that use the async APIs</span></span>

<span data-ttu-id="f3b6f-104">Este artículo le enseñará a mantener y actualizar scripts que usan las API asincrónicas de los modelos anteriores.</span><span class="sxs-lookup"><span data-stu-id="f3b6f-104">This article will teach you how to maintain and update scripts that use the older model's async APIs.</span></span> <span data-ttu-id="f3b6f-105">Estas API tienen la misma funcionalidad principal que las API de scripts de Office sincrónicas ahora, pero requieren el script para controlar la sincronización de datos entre el script y el libro.</span><span class="sxs-lookup"><span data-stu-id="f3b6f-105">These APIs have the same core functionality as the now-standard, synchronous Office Scripts APIs, but they require your script to control the data synchronization between the script and the workbook.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="f3b6f-106">El modelo Async solo se puede usar con scripts creados antes de la implementación del [modelo de API](scripting-fundamentals.md?view=office-scripts)actual.</span><span class="sxs-lookup"><span data-stu-id="f3b6f-106">The async model can only be used with scripts created before the implementation of the current [API model](scripting-fundamentals.md?view=office-scripts).</span></span> <span data-ttu-id="f3b6f-107">Los scripts se bloquean permanentemente en el modelo de API que tienen tras la creación.</span><span class="sxs-lookup"><span data-stu-id="f3b6f-107">Scripts are permanently locked to the API model they have upon creation.</span></span> <span data-ttu-id="f3b6f-108">Esto también significa que si desea convertir un script antiguo en el nuevo modelo, debe crear un nuevo script de marca.</span><span class="sxs-lookup"><span data-stu-id="f3b6f-108">This also means that if you want to convert an old script to the new model, you must create a brand new script.</span></span> <span data-ttu-id="f3b6f-109">Le recomendamos que actualice los scripts antiguos al nuevo modelo cuando realice cambios, ya que el modelo actual es más fácil de usar.</span><span class="sxs-lookup"><span data-stu-id="f3b6f-109">We recommend you update your old scripts to the new model when making changes, since the current model is easier to use.</span></span> <span data-ttu-id="f3b6f-110">La sección [convertir scripts asíncronos en el modelo actual](#converting-async-scripts-to-the-current-model) tiene consejos sobre cómo realizar esta transición.</span><span class="sxs-lookup"><span data-stu-id="f3b6f-110">The [Converting async scripts to the current model](#converting-async-scripts-to-the-current-model) section has advice on how to make this transition.</span></span>

## <a name="main-function"></a><span data-ttu-id="f3b6f-111">`main` Función</span><span class="sxs-lookup"><span data-stu-id="f3b6f-111">`main` function</span></span>

<span data-ttu-id="f3b6f-112">Los scripts que usan las API asincrónicas tienen una `main` función diferente.</span><span class="sxs-lookup"><span data-stu-id="f3b6f-112">Scripts that use the async APIs have a different `main` function.</span></span> <span data-ttu-id="f3b6f-113">Es una `async` función que tiene un `Excel.RequestContext` como el primer parámetro.</span><span class="sxs-lookup"><span data-stu-id="f3b6f-113">It's an `async` function that has an `Excel.RequestContext` as the first parameter.</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
    // Your async Office Script
}
```

## <a name="context"></a><span data-ttu-id="f3b6f-114">Context</span><span class="sxs-lookup"><span data-stu-id="f3b6f-114">Context</span></span>

<span data-ttu-id="f3b6f-115">La función `main` acepta un parámetro de `Excel.RequestContext`, denominado `context`.</span><span class="sxs-lookup"><span data-stu-id="f3b6f-115">The `main` function accepts an `Excel.RequestContext` parameter, named `context`.</span></span> <span data-ttu-id="f3b6f-116">Considere `context` como el puente entre el script y el libro.</span><span class="sxs-lookup"><span data-stu-id="f3b6f-116">Think of `context` as the bridge between your script and the workbook.</span></span> <span data-ttu-id="f3b6f-117">El script obtiene acceso al libro con el objeto `context` y usa ese `context` para enviar datos hacia adelante y hacia atrás.</span><span class="sxs-lookup"><span data-stu-id="f3b6f-117">Your script accesses the workbook with the `context` object and uses that `context` to send data back and forth.</span></span>

<span data-ttu-id="f3b6f-118">El objeto `context` es necesario porque el script y Excel se ejecutan en diferentes procesos y ubicaciones.</span><span class="sxs-lookup"><span data-stu-id="f3b6f-118">The `context` object is necessary because the script and Excel are running in different processes and locations.</span></span> <span data-ttu-id="f3b6f-119">El script tendrá que realizar cambios o consultar datos en el libro en la nube.</span><span class="sxs-lookup"><span data-stu-id="f3b6f-119">The script will need to make changes to or query data from the workbook in the cloud.</span></span> <span data-ttu-id="f3b6f-120">El objeto `context` administra estas transacciones.</span><span class="sxs-lookup"><span data-stu-id="f3b6f-120">The `context` object manages those transactions.</span></span>

## <a name="sync-and-load"></a><span data-ttu-id="f3b6f-121">Sync y Load</span><span class="sxs-lookup"><span data-stu-id="f3b6f-121">Sync and Load</span></span>

<span data-ttu-id="f3b6f-122">Como el script y el libro se ejecutan en distintas ubicaciones, cualquier transferencia de datos entre ambos necesita tiempo.</span><span class="sxs-lookup"><span data-stu-id="f3b6f-122">Because your script and workbook run in different locations, any data transfer between the two takes time.</span></span> <span data-ttu-id="f3b6f-123">En la API asincrónica, los comandos se ponen en cola hasta que el script llame explícitamente `sync` a la operación para sincronizar el script y el libro.</span><span class="sxs-lookup"><span data-stu-id="f3b6f-123">In the async API, commands are queued up until the script explicitly calls the `sync` operation to synchronize the script and workbook.</span></span> <span data-ttu-id="f3b6f-124">El script puede funcionar de forma independiente hasta que necesite realizar cualquiera de las siguientes acciones:</span><span class="sxs-lookup"><span data-stu-id="f3b6f-124">Your script can work independently until it needs to do either of the following:</span></span>

- <span data-ttu-id="f3b6f-125">Lea los datos del libro (después de una operación `load` o método que devuelve un [ClientResult](/javascript/api/office-scripts/excelscript/excelscript.clientresult?view=office-scripts-async)).</span><span class="sxs-lookup"><span data-stu-id="f3b6f-125">Read data from the workbook (following a `load` operation or method that returns a [ClientResult](/javascript/api/office-scripts/excelscript/excelscript.clientresult?view=office-scripts-async)).</span></span>
- <span data-ttu-id="f3b6f-126">Escribir datos en el libro (por lo general, porque el script ha terminado).</span><span class="sxs-lookup"><span data-stu-id="f3b6f-126">Write data to the workbook (usually because the script has finished).</span></span>

<span data-ttu-id="f3b6f-127">En la imagen siguiente se muestra un ejemplo de flujo de control entre el script y el libro:</span><span class="sxs-lookup"><span data-stu-id="f3b6f-127">The following image shows an example control flow between the script and workbook:</span></span>

![Diagrama en el que se muestran las operaciones de lectura y escritura en el libro desde el script.](../images/load-sync.png)

### <a name="sync"></a><span data-ttu-id="f3b6f-129">Sync</span><span class="sxs-lookup"><span data-stu-id="f3b6f-129">Sync</span></span>

<span data-ttu-id="f3b6f-130">Siempre que el script asincrónico necesite leer datos de un libro o escribir datos en él, llame al `RequestContext.sync` método como se muestra a continuación:</span><span class="sxs-lookup"><span data-stu-id="f3b6f-130">Whenever your async script needs to read data from or write data to the workbook, call the `RequestContext.sync` method as shown here:</span></span>

```TypeScript
await context.sync();
```

> [!NOTE]
> <span data-ttu-id="f3b6f-131">Se llama de forma implícita a `context.sync()` cuando finaliza un script.</span><span class="sxs-lookup"><span data-stu-id="f3b6f-131">`context.sync()` is implicitly called when a script ends.</span></span>

<span data-ttu-id="f3b6f-132">Una vez completada la operación `sync`, el libro se actualiza para reflejar las operaciones de escritura que haya especificado el script.</span><span class="sxs-lookup"><span data-stu-id="f3b6f-132">After the `sync` operation completes, the workbook updates to reflect any write operations that script has specified.</span></span> <span data-ttu-id="f3b6f-133">Una operación de escritura está estableciendo cualquier propiedad en un objeto de Excel (por ejemplo, `range.format.fill.color = "red"` ) o llamando a un método que cambia una propiedad (por ejemplo, `range.format.autoFitColumns()` ).</span><span class="sxs-lookup"><span data-stu-id="f3b6f-133">A write operation is setting any property on a Excel object (e.g., `range.format.fill.color = "red"`) or calling a method that changes a property (e.g., `range.format.autoFitColumns()`).</span></span> <span data-ttu-id="f3b6f-134">La operación `sync` también lee cualquier valor del libro solicitado por el script mediante una operación `load` o un método que devuelve un `ClientResult`(como se describe en la sección siguiente).</span><span class="sxs-lookup"><span data-stu-id="f3b6f-134">The `sync` operation also reads any values from the workbook that the script requested by using a `load` operation or a method that returns a `ClientResult` (as discussed in the next sections).</span></span>

<span data-ttu-id="f3b6f-135">Sincronizar el script con el libro puede tardar un tiempo, según la red.</span><span class="sxs-lookup"><span data-stu-id="f3b6f-135">Synchronizing your script with the workbook can take time, depending on your network.</span></span> <span data-ttu-id="f3b6f-136">Minimice el número de `sync` llamadas para ayudar a que el script se ejecute rápidamente.</span><span class="sxs-lookup"><span data-stu-id="f3b6f-136">Minimize the number of `sync` calls to help your script run fast.</span></span> <span data-ttu-id="f3b6f-137">De lo contrario, las API asincrónicas no son más rápidas las API estándar y sincrónicas.</span><span class="sxs-lookup"><span data-stu-id="f3b6f-137">Otherwise, the async APIs are not faster the standard, synchronous APIs.</span></span>

### <a name="load"></a><span data-ttu-id="f3b6f-138">Load</span><span class="sxs-lookup"><span data-stu-id="f3b6f-138">Load</span></span>

<span data-ttu-id="f3b6f-139">Un script asincrónico debe cargar datos del libro antes de leerlo.</span><span class="sxs-lookup"><span data-stu-id="f3b6f-139">An async script must load data from the workbook before reading it.</span></span> <span data-ttu-id="f3b6f-140">Sin embargo, si se cargan datos de todo el libro, se reduce en gran medida la velocidad del script.</span><span class="sxs-lookup"><span data-stu-id="f3b6f-140">However, loading data from the entire workbook would greatly reduce the script's speed.</span></span> <span data-ttu-id="f3b6f-141">El `load` método permite que su script indique específicamente qué datos deben recuperarse del libro.</span><span class="sxs-lookup"><span data-stu-id="f3b6f-141">The `load` method lets your script specifically state what data should be retrieved from the workbook.</span></span>

<span data-ttu-id="f3b6f-142">El método `load` está disponible en cada objeto de Excel.</span><span class="sxs-lookup"><span data-stu-id="f3b6f-142">The `load` method is available on every Excel object.</span></span> <span data-ttu-id="f3b6f-143">El script debe cargar las propiedades de un objeto antes de poder leerlas.</span><span class="sxs-lookup"><span data-stu-id="f3b6f-143">Your script must load an object's properties before it can read them.</span></span> <span data-ttu-id="f3b6f-144">Si no lo hace, se producirá un error.</span><span class="sxs-lookup"><span data-stu-id="f3b6f-144">Not doing so results in an error.</span></span>

<span data-ttu-id="f3b6f-145">Los ejemplos siguientes usan un objeto `Range` para mostrar las tres formas en que se puede usar el método `load` para cargar datos.</span><span class="sxs-lookup"><span data-stu-id="f3b6f-145">The following examples use a `Range` object to show the three ways the `load` method can be used to load data.</span></span>

|<span data-ttu-id="f3b6f-146">Objetivo</span><span class="sxs-lookup"><span data-stu-id="f3b6f-146">Intent</span></span> |<span data-ttu-id="f3b6f-147">Comando de ejemplo</span><span class="sxs-lookup"><span data-stu-id="f3b6f-147">Example Command</span></span> | <span data-ttu-id="f3b6f-148">Efecto</span><span class="sxs-lookup"><span data-stu-id="f3b6f-148">Effect</span></span> |
|:--|:--|:--|
|<span data-ttu-id="f3b6f-149">Cargar una propiedad</span><span class="sxs-lookup"><span data-stu-id="f3b6f-149">Load one property</span></span> |`myRange.load("values");` | <span data-ttu-id="f3b6f-150">Carga una única propiedad, en este caso la matriz bidimensional de valores en este rango.</span><span class="sxs-lookup"><span data-stu-id="f3b6f-150">Loads a single property, in this case the two-dimensional array of values in this range.</span></span> |
|<span data-ttu-id="f3b6f-151">Cargar varias propiedades</span><span class="sxs-lookup"><span data-stu-id="f3b6f-151">Load multiple properties</span></span> |`myRange.load("values, rowCount, columnCount");`| <span data-ttu-id="f3b6f-152">Carga todas las propiedades de una lista delimitada por comas, en este ejemplo, los valores, el número de filas y el número de columnas.</span><span class="sxs-lookup"><span data-stu-id="f3b6f-152">Loads all the properties from a comma-delimited list, in this example the values, row count, and column count.</span></span> |
|<span data-ttu-id="f3b6f-153">Cargar todo</span><span class="sxs-lookup"><span data-stu-id="f3b6f-153">Load everything</span></span> | `myRange.load();`|<span data-ttu-id="f3b6f-154">Carga todas las propiedades en el rango.</span><span class="sxs-lookup"><span data-stu-id="f3b6f-154">Loads all the properties on the range.</span></span> <span data-ttu-id="f3b6f-155">Esta solución no se recomienda, ya que ralentizará el script al obtener datos innecesarios.</span><span class="sxs-lookup"><span data-stu-id="f3b6f-155">This isn't a recommended solution, since it will slow down your script by getting unnecessary data.</span></span> <span data-ttu-id="f3b6f-156">Úsela solamente mientras prueba el script o si necesita todas las propiedades del objeto.</span><span class="sxs-lookup"><span data-stu-id="f3b6f-156">Only use this while testing your script or if you need every property from the object.</span></span> |

<span data-ttu-id="f3b6f-157">El script debe llamar a `context.sync()` antes de leer cualquier valor cargado.</span><span class="sxs-lookup"><span data-stu-id="f3b6f-157">Your script must call `context.sync()` before reading any loaded values.</span></span>

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

<span data-ttu-id="f3b6f-158">También puede cargar propiedades de toda la colección.</span><span class="sxs-lookup"><span data-stu-id="f3b6f-158">You can also load properties across an entire collection.</span></span> <span data-ttu-id="f3b6f-159">Cada objeto de colección de la API asincrónica tiene una `items` propiedad que es una matriz que contiene los objetos de esa colección.</span><span class="sxs-lookup"><span data-stu-id="f3b6f-159">Every collection object in the async API has an `items` property that is an array containing the objects in that collection.</span></span> <span data-ttu-id="f3b6f-160">El uso de `items` como inicio de una llamada jerárquica (`items\myProperty`) a `load` carga las propiedades especificadas en cada uno de esos elementos.</span><span class="sxs-lookup"><span data-stu-id="f3b6f-160">Using `items` as the start of a hierarchical call (`items\myProperty`) to `load` loads the specified properties on each of those items.</span></span> <span data-ttu-id="f3b6f-161">El ejemplo siguiente carga la propiedad `resolved` en cada objeto `Comment` del objeto `CommentCollection` de una hoja de cálculo.</span><span class="sxs-lookup"><span data-stu-id="f3b6f-161">The following example loads the `resolved` property on every `Comment` object in the `CommentCollection` object of a worksheet.</span></span>

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

### <a name="clientresult"></a><span data-ttu-id="f3b6f-162">ClientResult</span><span class="sxs-lookup"><span data-stu-id="f3b6f-162">ClientResult</span></span>

<span data-ttu-id="f3b6f-163">Los métodos de la API asíncrona que devuelven información del libro tienen un patrón similar al `load` / `sync` paradigma.</span><span class="sxs-lookup"><span data-stu-id="f3b6f-163">Methods in the async API that return information from the workbook have a similar pattern to the `load`/`sync` paradigm.</span></span> <span data-ttu-id="f3b6f-164">Por ejemplo, `TableCollection.getCount` obtiene el número de tablas de la colección.</span><span class="sxs-lookup"><span data-stu-id="f3b6f-164">As an example, `TableCollection.getCount` gets the number of tables in the collection.</span></span> <span data-ttu-id="f3b6f-165">`getCount`devuelve a `ClientResult<number>` , lo que significa que la `value` propiedad en el devuelto [`ClientResult`](/javascript/api/office-scripts/excelscript/excelscript.clientresult?view=office-scripts-async) es un número.</span><span class="sxs-lookup"><span data-stu-id="f3b6f-165">`getCount` returns a `ClientResult<number>`, meaning the `value` property in the returned [`ClientResult`](/javascript/api/office-scripts/excelscript/excelscript.clientresult?view=office-scripts-async) is a number.</span></span> <span data-ttu-id="f3b6f-166">El script no puede acceder a ese valor hasta que se llama a `context.sync()`.</span><span class="sxs-lookup"><span data-stu-id="f3b6f-166">Your script can't access that value until `context.sync()` is called.</span></span> <span data-ttu-id="f3b6f-167">De forma muy similar a la carga de una propiedad, el `value` es un valor local "vacío" hasta esa llamada `sync`.</span><span class="sxs-lookup"><span data-stu-id="f3b6f-167">Much like loading a property, the `value` is a local "empty" value until that `sync` call.</span></span>

<span data-ttu-id="f3b6f-168">El siguiente script obtiene el número total de tablas en el libro y registra ese número en la consola.</span><span class="sxs-lookup"><span data-stu-id="f3b6f-168">The following script gets the total number of tables in the workbook and logs that number to the console.</span></span>

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

## <a name="converting-async-scripts-to-the-current-model"></a><span data-ttu-id="f3b6f-169">Conversión de scripts asincrónicos al modelo actual</span><span class="sxs-lookup"><span data-stu-id="f3b6f-169">Converting async scripts to the current model</span></span>

<span data-ttu-id="f3b6f-170">El modelo de API actual no usa `load` , `sync` o a `RequestContext` .</span><span class="sxs-lookup"><span data-stu-id="f3b6f-170">The current API model doesn't use `load`, `sync`, or a `RequestContext`.</span></span> <span data-ttu-id="f3b6f-171">Esto hace que los scripts sean mucho más fáciles de escribir y mantener.</span><span class="sxs-lookup"><span data-stu-id="f3b6f-171">This makes the scripts much easier to write and maintain.</span></span> <span data-ttu-id="f3b6f-172">El mejor recurso para convertir secuencias de comandos antiguas es [desbordamiento de pila](https://stackoverflow.com/questions/tagged/office-scripts).</span><span class="sxs-lookup"><span data-stu-id="f3b6f-172">Your best resource for converting old scripts is [Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts).</span></span> <span data-ttu-id="f3b6f-173">Allí puede solicitar ayuda a la comunidad para escenarios específicos.</span><span class="sxs-lookup"><span data-stu-id="f3b6f-173">There, you can ask the community for help with specific scenarios.</span></span> <span data-ttu-id="f3b6f-174">Las siguientes instrucciones le ayudarán a describir los pasos generales que tendrá que realizar.</span><span class="sxs-lookup"><span data-stu-id="f3b6f-174">The following guidance should help outline the general steps you'll need to take.</span></span>

1. <span data-ttu-id="f3b6f-175">Cree un nuevo script y copie en él el código asincrónico anterior.</span><span class="sxs-lookup"><span data-stu-id="f3b6f-175">Create a new script and copy the old async code into it.</span></span> <span data-ttu-id="f3b6f-176">Asegúrese de no incluir la firma del `main` método anterior, utilizando la actual `function main(workbook: ExcelScript.Workbook)` en su lugar.</span><span class="sxs-lookup"><span data-stu-id="f3b6f-176">Be sure not to include the old `main` method signature, using the current `function main(workbook: ExcelScript.Workbook)` instead.</span></span>

2. <span data-ttu-id="f3b6f-177">Quite todas las `load` `sync` llamadas y.</span><span class="sxs-lookup"><span data-stu-id="f3b6f-177">Remove all the `load` and `sync` calls.</span></span> <span data-ttu-id="f3b6f-178">Ya no son necesarios.</span><span class="sxs-lookup"><span data-stu-id="f3b6f-178">They are no longer necessary.</span></span>

3. <span data-ttu-id="f3b6f-179">Se han quitado todas las propiedades.</span><span class="sxs-lookup"><span data-stu-id="f3b6f-179">All properties have been removed.</span></span> <span data-ttu-id="f3b6f-180">Ahora tiene acceso a los objetos mediante `get` y `set` métodos, por lo que tendrá que cambiar las referencias de propiedades a llamadas a métodos.</span><span class="sxs-lookup"><span data-stu-id="f3b6f-180">You now access those objects through `get` and `set` methods, so you'll need to switch those property references to method calls.</span></span> <span data-ttu-id="f3b6f-181">Por ejemplo, en lugar de establecer el color de relleno de una celda mediante el acceso a propiedades como este: `mySheet.getRange("A2:C2").format.fill.color = "blue";` , ahora usará métodos como este:`mySheet.getRange("A2:C2").getFormat().getFill().setColor("blue");`</span><span class="sxs-lookup"><span data-stu-id="f3b6f-181">For example, instead of setting a cell's fill color through property access like this: `mySheet.getRange("A2:C2").format.fill.color = "blue";`, you'll now use methods like this: `mySheet.getRange("A2:C2").getFormat().getFill().setColor("blue");`</span></span>

4. <span data-ttu-id="f3b6f-182">Las clases de colección se han reemplazado por matrices.</span><span class="sxs-lookup"><span data-stu-id="f3b6f-182">Collection classes have been replaced by arrays.</span></span> <span data-ttu-id="f3b6f-183">Los `add` `get` métodos y de las clases de colección se movieron al objeto que poseía la colección, por lo que las referencias deben actualizarse en consecuencia.</span><span class="sxs-lookup"><span data-stu-id="f3b6f-183">The `add` and `get` methods of those collection classes were moved to the object that owned the collection, so your references must be updated accordingly.</span></span> <span data-ttu-id="f3b6f-184">Por ejemplo, para obtener un gráfico denominado "MyChart" de la primera hoja de cálculo del libro, use el siguiente código: `workbook.getWorksheets()[0].getChart("MyChart");` .</span><span class="sxs-lookup"><span data-stu-id="f3b6f-184">For example, to get a chart named "MyChart" from the first worksheet in the workbook, use the following code: `workbook.getWorksheets()[0].getChart("MyChart");`.</span></span> <span data-ttu-id="f3b6f-185">Tenga en cuenta el `[0]` para obtener acceso al primer valor de la `Worksheet[]` devuelta por `getWorksheets()` .</span><span class="sxs-lookup"><span data-stu-id="f3b6f-185">Note the `[0]` to access the first value of the `Worksheet[]` returned by `getWorksheets()`.</span></span>

5. <span data-ttu-id="f3b6f-186">Se ha cambiado el nombre de algunos métodos para mayor claridad y se ha agregado para mayor comodidad.</span><span class="sxs-lookup"><span data-stu-id="f3b6f-186">Some methods have been renamed for clarity and added for convenience.</span></span> <span data-ttu-id="f3b6f-187">Consulte la referencia de la [API de scripts de Office](/javascript/api/office-scripts/overview?view=office-scripts) para obtener más información.</span><span class="sxs-lookup"><span data-stu-id="f3b6f-187">Please consult the [Office Scripts API reference](/javascript/api/office-scripts/overview?view=office-scripts) for more details.</span></span>

## <a name="office-scripts-async-api-reference-documentation"></a><span data-ttu-id="f3b6f-188">Documentación de referencia de API asincrónica de scripts de Office</span><span class="sxs-lookup"><span data-stu-id="f3b6f-188">Office Scripts Async API reference documentation</span></span>

[!INCLUDE [Async reference documentation](../includes/async-reference-documentation-link.md)]
