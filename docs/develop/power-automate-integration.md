---
title: Ejecutar scripts de Office con Power automatization
description: Cómo obtener scripts de Office para Excel en la web trabajar con un flujo de trabajo de Power automatization.
ms.date: 07/01/2020
localization_priority: Normal
ms.openlocfilehash: 40a67f3d0e8f049a8ec5516c0af54c5fc6fb9319
ms.sourcegitcommit: edf58aed3cd38f57e5e7227465a1ef5515e15703
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 07/08/2020
ms.locfileid: "45081596"
---
# <a name="run-office-scripts-with-power-automate"></a><span data-ttu-id="294bf-103">Ejecutar scripts de Office con Power automatization</span><span class="sxs-lookup"><span data-stu-id="294bf-103">Run Office Scripts with Power Automate</span></span>

<span data-ttu-id="294bf-104">La [automatización de energía](https://flow.microsoft.com) permite agregar scripts de Office a un flujo de trabajo más grande y automatizado.</span><span class="sxs-lookup"><span data-stu-id="294bf-104">[Power Automate](https://flow.microsoft.com) lets you add Office Scripts to a larger, automated workflow.</span></span> <span data-ttu-id="294bf-105">Puede usar la función automatizar acciones, como agregar el contenido de un correo electrónico a una tabla de una hoja de cálculo o crear acciones en las herramientas de administración de proyectos en función de los comentarios del libro.</span><span class="sxs-lookup"><span data-stu-id="294bf-105">You can use Power Automate do things like add the contents of an email to a worksheet's table or create actions in your project management tools based on workbook comments.</span></span> <span data-ttu-id="294bf-106">Si es la novedad de la automatización de energía, le recomendamos que visite [Introducción a Power Automate](/power-automate/getting-started).</span><span class="sxs-lookup"><span data-stu-id="294bf-106">If you are new to Power Automate, we recommend visiting [Get started with Power Automate](/power-automate/getting-started).</span></span> <span data-ttu-id="294bf-107">Aquí puede obtener más información sobre cómo automatizar los flujos de trabajo en varios servicios.</span><span class="sxs-lookup"><span data-stu-id="294bf-107">There, you can learn more about automating your workflows across multiple services.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="294bf-108">Actualmente, no se pueden ejecutar scripts de Office desde un [flujo compartido](/power-automate/share-buttons).</span><span class="sxs-lookup"><span data-stu-id="294bf-108">Currently, you can't run Office Scripts from a [shared flow](/power-automate/share-buttons).</span></span> <span data-ttu-id="294bf-109">Solo el usuario que creó un script puede ejecutarlo, incluso a través de la automatización de la energía.</span><span class="sxs-lookup"><span data-stu-id="294bf-109">Only the user who created a script can run it, even through Power Automate.</span></span>

## <a name="getting-started"></a><span data-ttu-id="294bf-110">Introducción</span><span class="sxs-lookup"><span data-stu-id="294bf-110">Getting started</span></span>

<span data-ttu-id="294bf-111">Para empezar a combinar la automatización de la alimentación y los scripts de Office, siga el tutorial [comenzar a usar scripts con Power automatization](../tutorials/excel-power-automate-manual.md).</span><span class="sxs-lookup"><span data-stu-id="294bf-111">To begin combining Power Automate and Office Scripts, follow the tutorial [Start using scripts with Power Automate](../tutorials/excel-power-automate-manual.md).</span></span> <span data-ttu-id="294bf-112">Esto le enseñará a crear un flujo que llame a un script sencillo.</span><span class="sxs-lookup"><span data-stu-id="294bf-112">This will teach you how to create a flow that calls a simple script.</span></span> <span data-ttu-id="294bf-113">Una vez que haya completado ese tutorial y el tutorial de [ejecutar scripts automáticamente con flujos automáticos de alimentación automatizada](../tutorials/excel-power-automate-trigger.md) , vuelva aquí para obtener información detallada sobre la conexión de scripts de Office para automatizar los flujos de alimentación.</span><span class="sxs-lookup"><span data-stu-id="294bf-113">After you've completed that tutorial and the [Automatically run scripts with automated Power Automate flows](../tutorials/excel-power-automate-trigger.md) tutorial, return here for detailed information about connecting Office Scripts to Power Automate flows.</span></span>

## <a name="excel-online-business-connector"></a><span data-ttu-id="294bf-114">Conector de Excel online (Business)</span><span class="sxs-lookup"><span data-stu-id="294bf-114">Excel Online (Business) connector</span></span>

<span data-ttu-id="294bf-115">Los [conectores](/connectors/connectors) son los puentes entre las aplicaciones y la automatización de la alimentación.</span><span class="sxs-lookup"><span data-stu-id="294bf-115">[Connectors](/connectors/connectors) are the bridges between Power Automate and applications.</span></span> <span data-ttu-id="294bf-116">El [conector de Excel online (Business)](/connectors/excelonlinebusiness) proporciona a los flujos acceso a los libros de Excel.</span><span class="sxs-lookup"><span data-stu-id="294bf-116">The [Excel Online (Business) connector](/connectors/excelonlinebusiness) gives your flows access to Excel workbooks.</span></span> <span data-ttu-id="294bf-117">La acción "ejecutar script" permite llamar a cualquier script de Office accesible a través del libro seleccionado.</span><span class="sxs-lookup"><span data-stu-id="294bf-117">The "Run script" action lets you call any Office Script accessible through the selected workbook.</span></span> <span data-ttu-id="294bf-118">No solo puede ejecutar scripts mediante un flujo, sino que puede pasar datos del libro y del flujo de trabajo a través de los scripts.</span><span class="sxs-lookup"><span data-stu-id="294bf-118">Not only can you run scripts through a flow, you can pass data to and from the workbook with the flow through the scripts.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="294bf-119">La acción "ejecutar script" da a los usuarios que usan el conector de Excel acceso significativo al libro y a sus datos.</span><span class="sxs-lookup"><span data-stu-id="294bf-119">The "Run script" action gives people who use the Excel connector significant access to your workbook and its data.</span></span> <span data-ttu-id="294bf-120">Además, existen riesgos de seguridad con los scripts que realizan llamadas externas a la API, como se explica en [llamadas externas de la automatización de la alimentación](external-calls.md).</span><span class="sxs-lookup"><span data-stu-id="294bf-120">Additionally, there are security risks with scripts that make external API calls, as explained in [External calls from Power Automate](external-calls.md).</span></span> <span data-ttu-id="294bf-121">Si su administrador está preocupado por la exposición de datos extremadamente confidenciales, puede desactivar el conector de Excel online o restringir el acceso a los scripts de Office a través de los [controles de administrador de scripts de Office](https://support.microsoft.com/office/19d3c51a-6ca2-40ab-978d-60fa49554dcf).</span><span class="sxs-lookup"><span data-stu-id="294bf-121">If your admin is concerned with the exposure of highly sensitive data, they can either turn off the Excel Online connector or restrict access to Office Scripts through the [Office Scripts administrator controls](https://support.microsoft.com/office/19d3c51a-6ca2-40ab-978d-60fa49554dcf).</span></span>

## <a name="data-transfer-in-flows-for-scripts"></a><span data-ttu-id="294bf-122">Transferencia de datos en flujos para scripts</span><span class="sxs-lookup"><span data-stu-id="294bf-122">Data transfer in flows for scripts</span></span>

<span data-ttu-id="294bf-123">La automatización de energía permite pasar datos entre los pasos de su flujo.</span><span class="sxs-lookup"><span data-stu-id="294bf-123">Power Automate lets you pass pieces of data between steps of your flow.</span></span> <span data-ttu-id="294bf-124">Los scripts se pueden configurar para que acepten los tipos de información que necesite y devuelvan cualquier elemento del libro que desee en su flujo.</span><span class="sxs-lookup"><span data-stu-id="294bf-124">Scripts can be configured to accept whatever types of information you need and return anything from your workbook that you want in your flow.</span></span> <span data-ttu-id="294bf-125">La entrada para el script se especifica agregando parámetros a la `main` función (además de `workbook: ExcelScript.Workbook` ).</span><span class="sxs-lookup"><span data-stu-id="294bf-125">Input for your script is specified by adding parameters to the `main` function (in addition to `workbook: ExcelScript.Workbook`).</span></span> <span data-ttu-id="294bf-126">El resultado del script se declara agregando un tipo de valor devuelto a `main` .</span><span class="sxs-lookup"><span data-stu-id="294bf-126">Output from the script is declared by adding a return type to `main`.</span></span>

> [!NOTE]
> <span data-ttu-id="294bf-127">Cuando se crea un bloque "ejecutar secuencia de comandos" en el flujo, se rellenan los parámetros y los tipos devueltos aceptados.</span><span class="sxs-lookup"><span data-stu-id="294bf-127">When you create a "Run Script" block in you flow, the accepted parameters and returned types are populated.</span></span> <span data-ttu-id="294bf-128">Si cambia los parámetros o tipos de valores devueltos del script, deberá rehacer el bloque "ejecutar script" del flujo.</span><span class="sxs-lookup"><span data-stu-id="294bf-128">If you change the parameters or return types of your script, you'll need to redo the "Run script" block of your flow.</span></span> <span data-ttu-id="294bf-129">Esto garantiza que los datos se analizan correctamente.</span><span class="sxs-lookup"><span data-stu-id="294bf-129">This ensure the data is being parsed correctly.</span></span>

<span data-ttu-id="294bf-130">En las secciones siguientes se describen los detalles de entrada y salida de las secuencias de comandos que se usan en la automatización de la energía.</span><span class="sxs-lookup"><span data-stu-id="294bf-130">The following sections cover the details of input and output for scripts used in Power Automate.</span></span> <span data-ttu-id="294bf-131">Si desea obtener un enfoque práctico para aprender este tema, pruebe el tutorial de [ejecución automática de secuencias de comandos con flujos de energía automatizada](../tutorials/excel-power-automate-trigger.md) automatizada o explorar el escenario de ejemplo de [avisos de tareas automatizadas](../resources/scenarios/task-reminders.md) .</span><span class="sxs-lookup"><span data-stu-id="294bf-131">If you'd like a hands-on approach to learning this topic, try out the [Automatically run scripts with automated Power Automate flows](../tutorials/excel-power-automate-trigger.md) tutorial or explore the [Automated task reminders](../resources/scenarios/task-reminders.md) sample scenario.</span></span>

### <a name="main-parameters-passing-data-to-a-script"></a><span data-ttu-id="294bf-132">`main`Parámetros: pasar datos a un script</span><span class="sxs-lookup"><span data-stu-id="294bf-132">`main` Parameters: Passing data to a script</span></span>

<span data-ttu-id="294bf-133">Todas las entradas de script se especifican como parámetros adicionales para la `main` función.</span><span class="sxs-lookup"><span data-stu-id="294bf-133">All script input is specified as additional parameters for the `main` function.</span></span> <span data-ttu-id="294bf-134">Por ejemplo, si desea que un script acepte un `string` que represente un nombre como entrada, cambiaría la `main` firma a `function main(workbook: ExcelScript.Workbook, name: string)` .</span><span class="sxs-lookup"><span data-stu-id="294bf-134">For example, if you wanted a script to accept a `string` that represents a name as input, you would change the `main` signature to `function main(workbook: ExcelScript.Workbook, name: string)`.</span></span>

<span data-ttu-id="294bf-135">Si está configurando un flujo con la automatización de la alimentación, puede especificar la entrada del script como valores estáticos, [expresiones](/power-automate/use-expressions-in-conditions)o contenido dinámico.</span><span class="sxs-lookup"><span data-stu-id="294bf-135">When you're configuring a flow in Power Automate, you can specify script input as static values, [expressions](/power-automate/use-expressions-in-conditions), or dynamic content.</span></span> <span data-ttu-id="294bf-136">Para obtener información sobre el conector de un servicio individual, vaya a la [documentación del conector Power Automated](/connectors/).</span><span class="sxs-lookup"><span data-stu-id="294bf-136">Details on an individual service's connector can be found in the [Power Automate Connector documentation](/connectors/).</span></span>

<span data-ttu-id="294bf-137">Al agregar parámetros de entrada a la función de una secuencia de comandos `main` , tenga en cuenta las siguientes restricciones y concesiones.</span><span class="sxs-lookup"><span data-stu-id="294bf-137">When adding input parameters to a script's `main` function, consider the following allowances and restrictions.</span></span>

1. <span data-ttu-id="294bf-138">El primer parámetro debe ser de tipo `ExcelScript.Workbook` .</span><span class="sxs-lookup"><span data-stu-id="294bf-138">The first parameter must be of type `ExcelScript.Workbook`.</span></span> <span data-ttu-id="294bf-139">El nombre del parámetro no importa.</span><span class="sxs-lookup"><span data-stu-id="294bf-139">Its parameter name does not matter.</span></span>

2. <span data-ttu-id="294bf-140">Cada parámetro debe tener un tipo.</span><span class="sxs-lookup"><span data-stu-id="294bf-140">Every parameter must have a type.</span></span>

3. <span data-ttu-id="294bf-141">`string` `number` `boolean` `any` `unknown` `object` `undefined` Se admiten los tipos básicos,,,,, y.</span><span class="sxs-lookup"><span data-stu-id="294bf-141">The basic types `string`, `number`, `boolean`, `any`, `unknown`, `object`, and `undefined` are supported.</span></span>

4. <span data-ttu-id="294bf-142">Se admiten las matrices de los tipos básicos enumerados anteriormente.</span><span class="sxs-lookup"><span data-stu-id="294bf-142">Arrays of the previously listed basic types are supported.</span></span>

5. <span data-ttu-id="294bf-143">Las matrices anidadas se admiten como parámetros (pero no como tipos devueltos).</span><span class="sxs-lookup"><span data-stu-id="294bf-143">Nested arrays are supported as parameters (but not as return types).</span></span>

6. <span data-ttu-id="294bf-144">Los tipos de Unión están permitidos si son una Unión de literales que pertenecen a un tipo único ( `string` , `number` o `boolean` ).</span><span class="sxs-lookup"><span data-stu-id="294bf-144">Union types are allowed if they are a union of literals belonging to a single type (`string`, `number`, or `boolean`).</span></span> <span data-ttu-id="294bf-145">También se admiten las uniones de un tipo compatible con undefined.</span><span class="sxs-lookup"><span data-stu-id="294bf-145">Unions of a supported type with undefined are also supported.</span></span>

7. <span data-ttu-id="294bf-146">Los tipos de objeto están permitidos si contienen propiedades de tipo `string` ,, `number` `boolean` , matrices admitidas u otros objetos admitidos.</span><span class="sxs-lookup"><span data-stu-id="294bf-146">Object types are allowed if they contain properties of type `string`, `number`, `boolean`, supported arrays, or other supported objects.</span></span> <span data-ttu-id="294bf-147">En el ejemplo siguiente se muestran los objetos anidados que se admiten como tipos de parámetro:</span><span class="sxs-lookup"><span data-stu-id="294bf-147">The following example shows nested objects that are supported as parameter types:</span></span>

    ```TypeScript
    // Office Scripts can return an Employee object because Position only contains strings and numbers.
    interface Employee {
        name: string;
        job: Position;
    }

    interface Position {
        id: number;
        title: string;
    }
    ```

8. <span data-ttu-id="294bf-148">Los objetos deben tener su definición de interfaz o clase definida en el script.</span><span class="sxs-lookup"><span data-stu-id="294bf-148">Objects must have their interface or class definition defined in the script.</span></span> <span data-ttu-id="294bf-149">Un objeto también puede definirse de forma anónima en línea, como en el ejemplo siguiente:</span><span class="sxs-lookup"><span data-stu-id="294bf-149">An object can also be defined anonymously inline, as in the following example:</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook): {name: string, email: string}
    ```

9. <span data-ttu-id="294bf-150">Los parámetros opcionales están permitidos y se pueden marcar como tales mediante el modificador Optional `?` (por ejemplo, `function main(workbook: ExcelScript.Workbook, Name?: string)` ).</span><span class="sxs-lookup"><span data-stu-id="294bf-150">Optional parameters are allowed and can be denoted as such by using the optional modifier `?` (for example, `function main(workbook: ExcelScript.Workbook, Name?: string)`).</span></span>

10. <span data-ttu-id="294bf-151">Se permiten los valores predeterminados de parámetro (por ejemplo,) `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')` .</span><span class="sxs-lookup"><span data-stu-id="294bf-151">Default parameter values are allowed (for example `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')`.</span></span>

## <a name="returning-data-from-a-script"></a><span data-ttu-id="294bf-152">Devolución de datos de un script</span><span class="sxs-lookup"><span data-stu-id="294bf-152">Returning data from a script</span></span>

<span data-ttu-id="294bf-153">Los scripts pueden devolver datos del libro que se van a usar como contenido dinámico en un flujo de automatización energética.</span><span class="sxs-lookup"><span data-stu-id="294bf-153">Scripts can return data from the workbook to be used as dynamic content in a Power Automate flow.</span></span> <span data-ttu-id="294bf-154">Al igual que con los parámetros de entrada, la automatización de energía coloca algunas restricciones en el tipo de valor devuelto.</span><span class="sxs-lookup"><span data-stu-id="294bf-154">As with input parameters, Power Automate places some restrictions on the return type.</span></span>

1. <span data-ttu-id="294bf-155">Se admiten los tipos básicos `string` , `number` ,, `boolean` `void` y `undefined` .</span><span class="sxs-lookup"><span data-stu-id="294bf-155">The basic types `string`, `number`, `boolean`, `void`, and `undefined` are supported.</span></span>

2. <span data-ttu-id="294bf-156">Los tipos de unión usados como tipos de valor devuelto siguen las mismas restricciones que los que se usan cuando se usan como parámetros de script.</span><span class="sxs-lookup"><span data-stu-id="294bf-156">Union types used as return types follow the same restrictions as they do when used as script parameters.</span></span>

3. <span data-ttu-id="294bf-157">Los tipos de matriz están permitidos si son del tipo `string` , `number` o `boolean` .</span><span class="sxs-lookup"><span data-stu-id="294bf-157">Array types are allowed if they are of type `string`, `number`, or `boolean`.</span></span> <span data-ttu-id="294bf-158">También se permiten si el tipo es una Unión compatible o un tipo literal admitido.</span><span class="sxs-lookup"><span data-stu-id="294bf-158">They are also allowed if the type is a supported union or supported literal type.</span></span>

4. <span data-ttu-id="294bf-159">Los tipos de objeto que se usan como tipos de valor devuelto siguen las mismas restricciones que cuando se usan como parámetros de script.</span><span class="sxs-lookup"><span data-stu-id="294bf-159">Object types used as return types follow the same restrictions as they do when used as script parameters.</span></span>

5. <span data-ttu-id="294bf-160">Se admite la escritura implícita, aunque debe seguir las mismas reglas que un tipo definido.</span><span class="sxs-lookup"><span data-stu-id="294bf-160">Implicit typing is supported, though it must follow the same rules as a defined type.</span></span>

## <a name="avoid-using-relative-references"></a><span data-ttu-id="294bf-161">Evitar el uso de referencias relativas</span><span class="sxs-lookup"><span data-stu-id="294bf-161">Avoid using relative references</span></span>

<span data-ttu-id="294bf-162">Power automaticing ejecuta el script en el libro de Excel elegido en su nombre.</span><span class="sxs-lookup"><span data-stu-id="294bf-162">Power Automate runs your script in the chosen Excel workbook on your behalf.</span></span> <span data-ttu-id="294bf-163">Es posible que el libro se cierre cuando esto suceda.</span><span class="sxs-lookup"><span data-stu-id="294bf-163">The workbook might be closed when this happens.</span></span> <span data-ttu-id="294bf-164">Cualquier API que se base en el estado actual del usuario, como `Workbook.getActiveWorksheet` , se producirá un error al ejecutarse a través de la automatización de la energía.</span><span class="sxs-lookup"><span data-stu-id="294bf-164">Any API that relies on the user's current state, such as `Workbook.getActiveWorksheet`, will fail when run through Power Automate.</span></span> <span data-ttu-id="294bf-165">Al diseñar los scripts, asegúrese de usar referencias absolutas para las hojas de cálculo y los rangos.</span><span class="sxs-lookup"><span data-stu-id="294bf-165">When designing your scripts, be sure to use absolute references for worksheets and ranges.</span></span>

<span data-ttu-id="294bf-166">Las siguientes funciones producirán un error y se producirá un error cuando se llame desde un script en un flujo de automatización de energía.</span><span class="sxs-lookup"><span data-stu-id="294bf-166">The following functions will throw an error and fail when called from a script in a Power Automate flow.</span></span>

- `Chart.activate`
- `Range.select`
- `Workbook.getActiveCell`
- `Workbook.getActiveChart`
- `Workbook.getActiveChartOrNullObject`
- `Workbook.getActiveSlicer`
- `Workbook.getActiveSlicerOrNullObject`
- `Workbook.getActiveWorksheet`
- `Workbook.getSelectedRange`
- `Workbook.getSelectedRanges`
- `Worksheet.activate`

## <a name="example"></a><span data-ttu-id="294bf-167">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="294bf-167">Example</span></span>

<span data-ttu-id="294bf-168">En la siguiente captura de pantalla se muestra un flujo de automatización de energía que se desencadena cuando se le asigna un problema de [GitHub](https://github.com/) .</span><span class="sxs-lookup"><span data-stu-id="294bf-168">The following screenshot shows a Power Automate flow that's triggered whenever a [GitHub](https://github.com/) issue is assigned to you.</span></span> <span data-ttu-id="294bf-169">El flujo ejecuta un script que agrega el problema a una tabla de un libro de Excel.</span><span class="sxs-lookup"><span data-stu-id="294bf-169">The flow runs a script that adds the issue to a table in an Excel workbook.</span></span> <span data-ttu-id="294bf-170">Si la tabla tiene cinco o más problemas, el flujo envía un aviso de correo electrónico.</span><span class="sxs-lookup"><span data-stu-id="294bf-170">If there are five or more issues in that table, the flow sends an email reminder.</span></span>

![El flujo de ejemplo, tal como se muestra en el editor de flujo de Power Automate.](../images/power-automate-parameter-return-sample.png)

<span data-ttu-id="294bf-172">La `main` función del script especifica el identificador del problema y el título del problema como parámetros de entrada, y el script devuelve el número de filas de la tabla Issue.</span><span class="sxs-lookup"><span data-stu-id="294bf-172">The `main` function of the script specifies the issue ID and issue title as input parameters, and the script returns the number of rows in the issue table.</span></span>

```TypeScript
function main(
  workbook: ExcelScript.Workbook,
  issueId: string,
  issueTitle: string): number {
  // Get the "GitHub" worksheet.
  let worksheet = workbook.getWorksheet("GitHub");

  // Get the first table in this worksheet, which contains the table of GitHub issues.
  let issueTable = worksheet.getTables()[0];

  // Add the issue ID and issue title as a row.
  issueTable.addRow(-1, [issueId, issueTitle]);

  // Return the number of rows in the table, which represents how many issues are assigned to this user.
  return issueTable.getRangeBetweenHeaderAndTotal().getRowCount();
}
```

## <a name="see-also"></a><span data-ttu-id="294bf-173">Vea también</span><span class="sxs-lookup"><span data-stu-id="294bf-173">See also</span></span>

- [<span data-ttu-id="294bf-174">Ejecutar scripts de Office en Excel en la web con la automatización de energía</span><span class="sxs-lookup"><span data-stu-id="294bf-174">Run Office Scripts in Excel on the web with Power Automate</span></span>](../tutorials/excel-power-automate-manual.md)
- [<span data-ttu-id="294bf-175">Ejecutar scripts automáticamente con flujos automatizar la alimentación automatizada</span><span class="sxs-lookup"><span data-stu-id="294bf-175">Automatically run scripts with automated Power Automate flows</span></span>](../tutorials/excel-power-automate-trigger.md)
- [<span data-ttu-id="294bf-176">Conceptos básicos de los Scripts de Office en Excel en la web</span><span class="sxs-lookup"><span data-stu-id="294bf-176">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](scripting-fundamentals.md)
- [<span data-ttu-id="294bf-177">Introducción a Power Automate</span><span class="sxs-lookup"><span data-stu-id="294bf-177">Get started with Power Automate</span></span>](/power-automate/getting-started)
- [<span data-ttu-id="294bf-178">Documentación de referencia de Excel online (Business) Connector</span><span class="sxs-lookup"><span data-stu-id="294bf-178">Excel Online (Business) connector reference documentation</span></span>](/connectors/excelonlinebusiness/)
