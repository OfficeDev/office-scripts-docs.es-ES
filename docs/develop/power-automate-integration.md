---
title: Ejecutar scripts de Office con Power automatization
description: Cómo obtener scripts de Office para Excel en la web trabajar con un flujo de trabajo de Power automatization.
ms.date: 06/29/2020
localization_priority: Normal
ms.openlocfilehash: 0ea58324998d23020e04cb37dfeea065791757f5
ms.sourcegitcommit: bf9f33c37c6f7805d6b408aa648bb9785a7cd133
ms.contentlocale: es-ES
ms.lasthandoff: 07/06/2020
ms.locfileid: "45043387"
---
# <a name="run-office-scripts-with-power-automate"></a><span data-ttu-id="0b101-103">Ejecutar scripts de Office con Power automatization</span><span class="sxs-lookup"><span data-stu-id="0b101-103">Run Office Scripts with Power Automate</span></span>

<span data-ttu-id="0b101-104">La [automatización de energía](https://flow.microsoft.com) permite agregar scripts de Office a un flujo de trabajo más grande y automatizado.</span><span class="sxs-lookup"><span data-stu-id="0b101-104">[Power Automate](https://flow.microsoft.com) lets you add Office Scripts to a larger, automated workflow.</span></span> <span data-ttu-id="0b101-105">Puede usar la función automatizar acciones, como agregar el contenido de un correo electrónico a una tabla de una hoja de cálculo o crear acciones en las herramientas de administración de proyectos en función de los comentarios del libro.</span><span class="sxs-lookup"><span data-stu-id="0b101-105">You can use Power Automate do things like add the contents of an email to a worksheet's table or create actions in your project management tools based on workbook comments.</span></span> <span data-ttu-id="0b101-106">Si es la novedad de la automatización de energía, le recomendamos que visite [Introducción a Power Automate](/power-automate/getting-started).</span><span class="sxs-lookup"><span data-stu-id="0b101-106">If you are new to Power Automate, we recommend visiting [Get started with Power Automate](/power-automate/getting-started).</span></span> <span data-ttu-id="0b101-107">Aquí puede obtener más información sobre cómo automatizar los flujos de trabajo en varios servicios.</span><span class="sxs-lookup"><span data-stu-id="0b101-107">There, you can learn more about automating your workflows across multiple services.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="0b101-108">Actualmente, no se pueden ejecutar scripts de Office desde un [flujo compartido](/power-automate/share-buttons).</span><span class="sxs-lookup"><span data-stu-id="0b101-108">Currently, you can't run Office Scripts from a [shared flow](/power-automate/share-buttons).</span></span> <span data-ttu-id="0b101-109">Solo el usuario que creó un script puede ejecutarlo, incluso a través de la automatización de la energía.</span><span class="sxs-lookup"><span data-stu-id="0b101-109">Only the user who created a script can run it, even through Power Automate.</span></span>

## <a name="getting-started"></a><span data-ttu-id="0b101-110">Introducción</span><span class="sxs-lookup"><span data-stu-id="0b101-110">Getting started</span></span>

<span data-ttu-id="0b101-111">Para empezar a combinar la automatización de la alimentación y los scripts de Office, siga el tutorial [comenzar a usar scripts con Power automatization](../tutorials/excel-power-automate-manual.md).</span><span class="sxs-lookup"><span data-stu-id="0b101-111">To begin combining Power Automate and Office Scripts, follow the tutorial [Start using scripts with Power Automate](../tutorials/excel-power-automate-manual.md).</span></span> <span data-ttu-id="0b101-112">Esto le enseñará a crear un flujo que llame a un script sencillo.</span><span class="sxs-lookup"><span data-stu-id="0b101-112">This will teach you how to create a flow that calls a simple script.</span></span> <span data-ttu-id="0b101-113">Después de completar ese tutorial y [ejecutar automáticamente scripts con el tutorial de Power automatization](../tutorials/excel-power-automate-trigger.md) , vuelva aquí para obtener información detallada sobre la conexión de scripts de Office para automatizar los flujos de alimentación.</span><span class="sxs-lookup"><span data-stu-id="0b101-113">After you've completed that tutorial and the [Automatically run scripts with Power Automate](../tutorials/excel-power-automate-trigger.md) tutorial, return here for detailed information about connecting Office Scripts to Power Automate flows.</span></span>

## <a name="excel-online-business-connector"></a><span data-ttu-id="0b101-114">Conector de Excel online (Business)</span><span class="sxs-lookup"><span data-stu-id="0b101-114">Excel Online (Business) connector</span></span>

<span data-ttu-id="0b101-115">Los [conectores](/connectors/connectors) son los puentes entre las aplicaciones y la automatización de la alimentación.</span><span class="sxs-lookup"><span data-stu-id="0b101-115">[Connectors](/connectors/connectors) are the bridges between Power Automate and applications.</span></span> <span data-ttu-id="0b101-116">El [conector de Excel online (Business)](/connectors/excelonlinebusiness) proporciona a los flujos acceso a los libros de Excel.</span><span class="sxs-lookup"><span data-stu-id="0b101-116">The [Excel Online (Business) connector](/connectors/excelonlinebusiness) gives your flows access to Excel workbooks.</span></span> <span data-ttu-id="0b101-117">La acción "ejecutar script" permite llamar a cualquier script de Office accesible a través del libro seleccionado.</span><span class="sxs-lookup"><span data-stu-id="0b101-117">The "Run script" action lets you call any Office Script accessible through the selected workbook.</span></span> <span data-ttu-id="0b101-118">No solo puede ejecutar scripts mediante un flujo, sino que puede pasar datos del libro y del flujo de trabajo a través de los scripts.</span><span class="sxs-lookup"><span data-stu-id="0b101-118">Not only can you run scripts through a flow, you can pass data to and from the workbook with the flow through the scripts.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="0b101-119">La acción "ejecutar script" da a los usuarios que usan el conector de Excel acceso significativo al libro y a sus datos.</span><span class="sxs-lookup"><span data-stu-id="0b101-119">The "Run script" action gives people who use the Excel connector significant access to your workbook and its data.</span></span> <span data-ttu-id="0b101-120">Además, existen riesgos de seguridad con los scripts que realizan llamadas externas a la API, como se explica en [llamadas externas de la automatización de la alimentación](external-calls.md).</span><span class="sxs-lookup"><span data-stu-id="0b101-120">Additionally, there are security risks with scripts that make external API calls, as explained in [External calls from Power Automate](external-calls.md).</span></span> <span data-ttu-id="0b101-121">Si su administrador está preocupado por la exposición de datos extremadamente confidenciales, puede desactivar el conector de Excel online o restringir el acceso a los scripts de Office a través de los [controles de administrador de scripts de Office](https://support.microsoft.com/office/19d3c51a-6ca2-40ab-978d-60fa49554dcf).</span><span class="sxs-lookup"><span data-stu-id="0b101-121">If your admin is concerned with the exposure of highly sensitive data, they can either turn off the Excel Online connector or restrict access to Office Scripts through the [Office Scripts administrator controls](https://support.microsoft.com/office/19d3c51a-6ca2-40ab-978d-60fa49554dcf).</span></span>

## <a name="passing-data-from-power-automate-into-a-script"></a><span data-ttu-id="0b101-122">Pasar datos de Automatic Power a un script</span><span class="sxs-lookup"><span data-stu-id="0b101-122">Passing data from Power Automate into a script</span></span>

<span data-ttu-id="0b101-123">Todas las entradas de script se especifican como parámetros adicionales para la `main` función.</span><span class="sxs-lookup"><span data-stu-id="0b101-123">All script input is specified as additional parameters for the `main` function.</span></span> <span data-ttu-id="0b101-124">Por ejemplo, si desea que un script acepte un `string` que represente un nombre como entrada, cambiaría la `main` firma a `function main(workbook: ExcelScript.Workbook, name: string)` .</span><span class="sxs-lookup"><span data-stu-id="0b101-124">For example, if you wanted a script to accept a `string` that represents a name as input, you would change the `main` signature to `function main(workbook: ExcelScript.Workbook, name: string)`.</span></span>

<span data-ttu-id="0b101-125">Si está configurando un flujo con la automatización de la alimentación, puede especificar la entrada del script como valores estáticos, [expresiones](/power-automate/use-expressions-in-conditions)o contenido dinámico.</span><span class="sxs-lookup"><span data-stu-id="0b101-125">When you're configuring a flow in Power Automate, you can specify script input as static values, [expressions](/power-automate/use-expressions-in-conditions), or dynamic content.</span></span> <span data-ttu-id="0b101-126">Para obtener información sobre el conector de un servicio individual, vaya a la [documentación del conector Power Automated](/connectors/).</span><span class="sxs-lookup"><span data-stu-id="0b101-126">Details on an individual service's connector can be found in the [Power Automate Connector documentation](/connectors/).</span></span>

<span data-ttu-id="0b101-127">Al agregar parámetros de entrada a la función de una secuencia de comandos `main` , tenga en cuenta las siguientes restricciones y concesiones.</span><span class="sxs-lookup"><span data-stu-id="0b101-127">When adding input parameters to a script's `main` function, consider the following allowances and restrictions.</span></span>

1. <span data-ttu-id="0b101-128">El primer parámetro debe ser de tipo `ExcelScript.Workbook` .</span><span class="sxs-lookup"><span data-stu-id="0b101-128">The first parameter must be of type `ExcelScript.Workbook`.</span></span> <span data-ttu-id="0b101-129">El nombre del parámetro no importa.</span><span class="sxs-lookup"><span data-stu-id="0b101-129">Its parameter name does not matter.</span></span>

2. <span data-ttu-id="0b101-130">Cada parámetro debe tener un tipo.</span><span class="sxs-lookup"><span data-stu-id="0b101-130">Every parameter must have a type.</span></span>

3. <span data-ttu-id="0b101-131">`string` `number` `boolean` `any` `unknown` `object` `undefined` Se admiten los tipos básicos,,,,, y.</span><span class="sxs-lookup"><span data-stu-id="0b101-131">The basic types `string`, `number`, `boolean`, `any`, `unknown`, `object`, and `undefined` are supported.</span></span>

4. <span data-ttu-id="0b101-132">Se admiten las matrices de los tipos básicos enumerados anteriormente.</span><span class="sxs-lookup"><span data-stu-id="0b101-132">Arrays of the previously listed basic types are supported.</span></span>

5. <span data-ttu-id="0b101-133">Las matrices anidadas se admiten como parámetros (pero no como tipos devueltos).</span><span class="sxs-lookup"><span data-stu-id="0b101-133">Nested arrays are supported as parameters (but not as return types).</span></span>

6. <span data-ttu-id="0b101-134">Los tipos de Unión están permitidos si son una Unión de literales que pertenecen a un tipo único ( `string` , `number` o `boolean` ).</span><span class="sxs-lookup"><span data-stu-id="0b101-134">Union types are allowed if they are a union of literals belonging to a single type (`string`, `number`, or `boolean`).</span></span> <span data-ttu-id="0b101-135">También se admiten las uniones de un tipo compatible con undefined.</span><span class="sxs-lookup"><span data-stu-id="0b101-135">Unions of a supported type with undefined are also supported.</span></span>

7. <span data-ttu-id="0b101-136">Los tipos de objeto están permitidos si contienen propiedades de tipo `string` ,, `number` `boolean` , matrices admitidas u otros objetos admitidos.</span><span class="sxs-lookup"><span data-stu-id="0b101-136">Object types are allowed if they contain properties of type `string`, `number`, `boolean`, supported arrays, or other supported objects.</span></span> <span data-ttu-id="0b101-137">En el ejemplo siguiente se muestran los objetos anidados que se admiten como tipos de parámetro:</span><span class="sxs-lookup"><span data-stu-id="0b101-137">The following example shows nested objects that are supported as parameter types:</span></span>

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

8. <span data-ttu-id="0b101-138">Los objetos deben tener su definición de interfaz o clase definida en el script.</span><span class="sxs-lookup"><span data-stu-id="0b101-138">Objects must have their interface or class definition defined in the script.</span></span> <span data-ttu-id="0b101-139">Un objeto también puede definirse de forma anónima en línea, como en el ejemplo siguiente:</span><span class="sxs-lookup"><span data-stu-id="0b101-139">An object can also be defined anonymously inline, as in the following example:</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook): {name: string, email: string}
    ```

9. <span data-ttu-id="0b101-140">Los parámetros opcionales están permitidos y se pueden marcar como tales mediante el modificador Optional `?` (por ejemplo, `function main(workbook: ExcelScript.Workbook, Name?: string)` ).</span><span class="sxs-lookup"><span data-stu-id="0b101-140">Optional parameters are allowed and can be denoted as such by using the optional modifier `?` (for example, `function main(workbook: ExcelScript.Workbook, Name?: string)`).</span></span>

10. <span data-ttu-id="0b101-141">Se permiten los valores predeterminados de parámetro (por ejemplo,) `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')` .</span><span class="sxs-lookup"><span data-stu-id="0b101-141">Default parameter values are allowed (for example `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')`.</span></span>

## <a name="returning-data-from-a-script-back-to-power-automate"></a><span data-ttu-id="0b101-142">Devolución de datos de un script a la automatización de la energía</span><span class="sxs-lookup"><span data-stu-id="0b101-142">Returning data from a script back to Power Automate</span></span>

<span data-ttu-id="0b101-143">Los scripts pueden devolver datos del libro que se van a usar como contenido dinámico en un flujo de automatización energética.</span><span class="sxs-lookup"><span data-stu-id="0b101-143">Scripts can return data from the workbook to be used as dynamic content in a Power Automate flow.</span></span> <span data-ttu-id="0b101-144">Al igual que con los parámetros de entrada, la automatización de energía coloca algunas restricciones en el tipo de valor devuelto.</span><span class="sxs-lookup"><span data-stu-id="0b101-144">As with input parameters, Power Automate places some restrictions on the return type.</span></span>

1. <span data-ttu-id="0b101-145">Se admiten los tipos básicos `string` , `number` ,, `boolean` `void` y `undefined` .</span><span class="sxs-lookup"><span data-stu-id="0b101-145">The basic types `string`, `number`, `boolean`, `void`, and `undefined` are supported.</span></span>

2. <span data-ttu-id="0b101-146">Los tipos de unión usados como tipos de valor devuelto siguen las mismas restricciones que los que se usan cuando se usan como parámetros de script.</span><span class="sxs-lookup"><span data-stu-id="0b101-146">Union types used as return types follow the same restrictions as they do when used as script parameters.</span></span>

3. <span data-ttu-id="0b101-147">Los tipos de matriz están permitidos si son del tipo `string` , `number` o `boolean` .</span><span class="sxs-lookup"><span data-stu-id="0b101-147">Array types are allowed if they are of type `string`, `number`, or `boolean`.</span></span> <span data-ttu-id="0b101-148">También se permiten si el tipo es una Unión compatible o un tipo literal admitido.</span><span class="sxs-lookup"><span data-stu-id="0b101-148">They are also allowed if the type is a supported union or supported literal type.</span></span>

4. <span data-ttu-id="0b101-149">Los tipos de objeto que se usan como tipos de valor devuelto siguen las mismas restricciones que cuando se usan como parámetros de script.</span><span class="sxs-lookup"><span data-stu-id="0b101-149">Object types used as return types follow the same restrictions as they do when used as script parameters.</span></span>

5. <span data-ttu-id="0b101-150">Se admite la escritura implícita, aunque debe seguir las mismas reglas que un tipo definido.</span><span class="sxs-lookup"><span data-stu-id="0b101-150">Implicit typing is supported, though it must follow the same rules as a defined type.</span></span>

## <a name="avoid-using-relative-references"></a><span data-ttu-id="0b101-151">Evitar el uso de referencias relativas</span><span class="sxs-lookup"><span data-stu-id="0b101-151">Avoid using relative references</span></span>

<span data-ttu-id="0b101-152">Power automaticing ejecuta el script en el libro de Excel elegido en su nombre.</span><span class="sxs-lookup"><span data-stu-id="0b101-152">Power Automate runs your script in the chosen Excel workbook on your behalf.</span></span> <span data-ttu-id="0b101-153">Es posible que el libro se cierre cuando esto suceda.</span><span class="sxs-lookup"><span data-stu-id="0b101-153">The workbook might be closed when this happens.</span></span> <span data-ttu-id="0b101-154">Cualquier API que se base en el estado actual del usuario, como `Workbook.getActiveWorksheet` , se producirá un error al ejecutarse a través de la automatización de la energía.</span><span class="sxs-lookup"><span data-stu-id="0b101-154">Any API that relies on the user's current state, such as `Workbook.getActiveWorksheet`, will fail when run through Power Automate.</span></span> <span data-ttu-id="0b101-155">Al diseñar los scripts, asegúrese de usar referencias absolutas para las hojas de cálculo y los rangos.</span><span class="sxs-lookup"><span data-stu-id="0b101-155">When designing your scripts, be sure to use absolute references for worksheets and ranges.</span></span>

<span data-ttu-id="0b101-156">Las siguientes funciones producirán un error y se producirá un error cuando se llame desde un script en un flujo de automatización de energía.</span><span class="sxs-lookup"><span data-stu-id="0b101-156">The following functions will throw an error and fail when called from a script in a Power Automate flow.</span></span>

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

## <a name="example"></a><span data-ttu-id="0b101-157">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="0b101-157">Example</span></span>

<span data-ttu-id="0b101-158">En la siguiente captura de pantalla se muestra un flujo de automatización de energía que se desencadena cuando se le asigna un problema de [GitHub](https://github.com/) .</span><span class="sxs-lookup"><span data-stu-id="0b101-158">The following screenshot shows a Power Automate flow that's triggered whenever a [GitHub](https://github.com/) issue is assigned to you.</span></span> <span data-ttu-id="0b101-159">El flujo ejecuta un script que agrega el problema a una tabla de un libro de Excel.</span><span class="sxs-lookup"><span data-stu-id="0b101-159">The flow runs a script that adds the issue to a table in an Excel workbook.</span></span> <span data-ttu-id="0b101-160">Si la tabla tiene cinco o más problemas, el flujo envía un aviso de correo electrónico.</span><span class="sxs-lookup"><span data-stu-id="0b101-160">If there are five or more issues in that table, the flow sends an email reminder.</span></span>

![El flujo de ejemplo, tal como se muestra en el editor de flujo de Power Automate.](../images/power-automate-parameter-return-sample.png)

<span data-ttu-id="0b101-162">La `main` función del script especifica el identificador del problema y el título del problema como parámetros de entrada, y el script devuelve el número de filas de la tabla Issue.</span><span class="sxs-lookup"><span data-stu-id="0b101-162">The `main` function of the script specifies the issue ID and issue title as input parameters, and the script returns the number of rows in the issue table.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="0b101-163">Vea también</span><span class="sxs-lookup"><span data-stu-id="0b101-163">See also</span></span>

- [<span data-ttu-id="0b101-164">Ejecutar scripts de Office en Excel en la web con la automatización de energía</span><span class="sxs-lookup"><span data-stu-id="0b101-164">Run Office Scripts in Excel on the web with Power Automate</span></span>](../tutorials/excel-power-automate-manual.md)
- [<span data-ttu-id="0b101-165">Ejecutar automáticamente scripts con Power Automate</span><span class="sxs-lookup"><span data-stu-id="0b101-165">Automatically run scripts with Power Automate</span></span>](../tutorials/excel-power-automate-trigger.md)
- [<span data-ttu-id="0b101-166">Conceptos básicos de los scripts de Office en Excel en la Web</span><span class="sxs-lookup"><span data-stu-id="0b101-166">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](scripting-fundamentals.md)
- [<span data-ttu-id="0b101-167">Introducción a Power Automate</span><span class="sxs-lookup"><span data-stu-id="0b101-167">Get started with Power Automate</span></span>](/power-automate/getting-started)
- [<span data-ttu-id="0b101-168">Documentación de referencia de Excel online (Business) Connector</span><span class="sxs-lookup"><span data-stu-id="0b101-168">Excel Online (Business) connector reference documentation</span></span>](/connectors/excelonlinebusiness/)
