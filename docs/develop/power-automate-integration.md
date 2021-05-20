---
title: Ejecute scripts de Office con Power Automate
description: Cómo obtener scripts de Office para Excel en la Web trabajando con un flujo de trabajo Power Automate.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 7562a2b2359cde67a9a47e0640515018fe23ac35
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545043"
---
# <a name="run-office-scripts-with-power-automate"></a><span data-ttu-id="4cffc-103">Ejecute scripts de Office con Power Automate</span><span class="sxs-lookup"><span data-stu-id="4cffc-103">Run Office Scripts with Power Automate</span></span>

<span data-ttu-id="4cffc-104">[Power Automate](https://flow.microsoft.com) le permite agregar scripts de Office a un flujo de trabajo automatizado más grande.</span><span class="sxs-lookup"><span data-stu-id="4cffc-104">[Power Automate](https://flow.microsoft.com) lets you add Office Scripts to a larger, automated workflow.</span></span> <span data-ttu-id="4cffc-105">Puede usar Power Automate hacer cosas como agregar el contenido de un correo electrónico a la tabla de una hoja de trabajo o crear acciones en las herramientas de administración de proyectos en función de los comentarios del libro.</span><span class="sxs-lookup"><span data-stu-id="4cffc-105">You can use Power Automate do things like add the contents of an email to a worksheet's table or create actions in your project management tools based on workbook comments.</span></span>

## <a name="get-started"></a><span data-ttu-id="4cffc-106">Comenzar</span><span class="sxs-lookup"><span data-stu-id="4cffc-106">Get started</span></span>

<span data-ttu-id="4cffc-107">Si eres nuevo en Power Automate, te recomendamos que te [desamiendes Comenzar con Power Automate.](/power-automate/getting-started)</span><span class="sxs-lookup"><span data-stu-id="4cffc-107">If you are new to Power Automate, we recommend visiting [Get started with Power Automate](/power-automate/getting-started).</span></span> <span data-ttu-id="4cffc-108">Allí, puede obtener más información sobre todas las posibilidades de automatización disponibles para usted.</span><span class="sxs-lookup"><span data-stu-id="4cffc-108">There, you can learn more about all the automation possibilities available to you.</span></span> <span data-ttu-id="4cffc-109">Los documentos aquí se centran en cómo funcionan Office Scripts con Power Automate y cómo eso puede ayudar a mejorar su experiencia Excel.</span><span class="sxs-lookup"><span data-stu-id="4cffc-109">The documents here focus on how Office Scripts work with Power Automate and how that can help improve your Excel experience.</span></span>

<span data-ttu-id="4cffc-110">Para comenzar a combinar scripts de Power Automate y Office, siga el tutorial [Iniciar uso de scripts con Power Automate](../tutorials/excel-power-automate-manual.md).</span><span class="sxs-lookup"><span data-stu-id="4cffc-110">To begin combining Power Automate and Office Scripts, follow the tutorial [Start using scripts with Power Automate](../tutorials/excel-power-automate-manual.md).</span></span> <span data-ttu-id="4cffc-111">Esto le enseñará cómo crear un flujo que llame a un script simple.</span><span class="sxs-lookup"><span data-stu-id="4cffc-111">This will teach you how to create a flow that calls a simple script.</span></span> <span data-ttu-id="4cffc-112">Después de completar ese tutorial y pasar datos a scripts en un tutorial [de flujo de Power Automate de ejecución automática,](../tutorials/excel-power-automate-trigger.md) vuelva aquí para obtener información detallada sobre cómo conectar scripts de Office a flujos de Power Automate.</span><span class="sxs-lookup"><span data-stu-id="4cffc-112">After you've completed that tutorial and the [Pass data to scripts in an automatically-run Power Automate flow](../tutorials/excel-power-automate-trigger.md) tutorial, return here for detailed information about connecting Office Scripts to Power Automate flows.</span></span>

## <a name="excel-online-business-connector"></a><span data-ttu-id="4cffc-113">Excel Conector en línea (business)</span><span class="sxs-lookup"><span data-stu-id="4cffc-113">Excel Online (Business) connector</span></span>

<span data-ttu-id="4cffc-114">[Los conectores](/connectors/connectors) son los puentes entre Power Automate y aplicaciones.</span><span class="sxs-lookup"><span data-stu-id="4cffc-114">[Connectors](/connectors/connectors) are the bridges between Power Automate and applications.</span></span> <span data-ttu-id="4cffc-115">El [conector Excel online (business)](/connectors/excelonlinebusiness) proporciona a los flujos acceso a Excel libros de trabajo.</span><span class="sxs-lookup"><span data-stu-id="4cffc-115">The [Excel Online (Business) connector](/connectors/excelonlinebusiness) gives your flows access to Excel workbooks.</span></span> <span data-ttu-id="4cffc-116">La acción "Ejecutar script" le permite llamar a cualquier script Office accesible a través del libro seleccionado.</span><span class="sxs-lookup"><span data-stu-id="4cffc-116">The "Run script" action lets you call any Office Script accessible through the selected workbook.</span></span> <span data-ttu-id="4cffc-117">También puede proporcionar los parámetros de entrada de scripts para que el flujo pueda proporcionar datos o hacer que el script devuelva información para pasos posteriores en el flujo.</span><span class="sxs-lookup"><span data-stu-id="4cffc-117">You can also give your scripts input parameters so data can be provided by the flow, or have your script return information for later steps in the flow.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="4cffc-118">La acción "Ejecutar script" proporciona a las personas que usan el conector de Excel acceso significativo a su libro de trabajo y sus datos.</span><span class="sxs-lookup"><span data-stu-id="4cffc-118">The "Run script" action gives people who use the Excel connector significant access to your workbook and its data.</span></span> <span data-ttu-id="4cffc-119">Además, existen riesgos de seguridad con los scripts que realizan llamadas a la API externa, como se explica en [Llamadas externas desde Power Automate](external-calls.md).</span><span class="sxs-lookup"><span data-stu-id="4cffc-119">Additionally, there are security risks with scripts that make external API calls, as explained in [External calls from Power Automate](external-calls.md).</span></span> <span data-ttu-id="4cffc-120">Si el administrador está preocupado por la exposición de datos altamente confidenciales, puede desactivar el conector Excel online o restringir el acceso a scripts de Office a través de los [controles de administrador de scripts de Office.](/microsoft-365/admin/manage/manage-office-scripts-settings)</span><span class="sxs-lookup"><span data-stu-id="4cffc-120">If your admin is concerned with the exposure of highly sensitive data, they can either turn off the Excel Online connector or restrict access to Office Scripts through the [Office Scripts administrator controls](/microsoft-365/admin/manage/manage-office-scripts-settings).</span></span>

## <a name="data-transfer-in-flows-for-scripts"></a><span data-ttu-id="4cffc-121">Transferencia de datos en flujos para scripts</span><span class="sxs-lookup"><span data-stu-id="4cffc-121">Data transfer in flows for scripts</span></span>

<span data-ttu-id="4cffc-122">Power Automate le permite pasar fragmentos de datos entre los pasos del flujo.</span><span class="sxs-lookup"><span data-stu-id="4cffc-122">Power Automate lets you pass pieces of data between steps of your flow.</span></span> <span data-ttu-id="4cffc-123">Los scripts se pueden configurar para aceptar los tipos de información que necesite y devolver cualquier cosa de su libro de trabajo que desee en el flujo.</span><span class="sxs-lookup"><span data-stu-id="4cffc-123">Scripts can be configured to accept whatever types of information you need and return anything from your workbook that you want in your flow.</span></span> <span data-ttu-id="4cffc-124">La entrada del script se especifica agregando parámetros a la `main` función (además de `workbook: ExcelScript.Workbook` ).</span><span class="sxs-lookup"><span data-stu-id="4cffc-124">Input for your script is specified by adding parameters to the `main` function (in addition to `workbook: ExcelScript.Workbook`).</span></span> <span data-ttu-id="4cffc-125">La salida del script se declara agregando un tipo de valor devuelto a `main` .</span><span class="sxs-lookup"><span data-stu-id="4cffc-125">Output from the script is declared by adding a return type to `main`.</span></span>

> [!NOTE]
> <span data-ttu-id="4cffc-126">Al crear un bloque "Ejecutar script" en el flujo, se rellenan los parámetros aceptados y los tipos devueltos.</span><span class="sxs-lookup"><span data-stu-id="4cffc-126">When you create a "Run Script" block in your flow, the accepted parameters and returned types are populated.</span></span> <span data-ttu-id="4cffc-127">Si cambia los parámetros o los tipos de valor devuelto del script, deberá volver a hacer el bloque "Ejecutar script" del flujo.</span><span class="sxs-lookup"><span data-stu-id="4cffc-127">If you change the parameters or return types of your script, you'll need to redo the "Run script" block of your flow.</span></span> <span data-ttu-id="4cffc-128">Esto garantiza que los datos se estén analizando correctamente.</span><span class="sxs-lookup"><span data-stu-id="4cffc-128">This ensures the data is being parsed correctly.</span></span>

<span data-ttu-id="4cffc-129">En las secciones siguientes se tratan los detalles de entrada y salida de los scripts utilizados en Power Automate.</span><span class="sxs-lookup"><span data-stu-id="4cffc-129">The following sections cover the details of input and output for scripts used in Power Automate.</span></span> <span data-ttu-id="4cffc-130">Si desea un enfoque práctico para aprender este tema, pruebe los [datos de paso a scripts en un](../tutorials/excel-power-automate-trigger.md) tutorial de flujo de Power Automate de ejecución automática o explore el escenario de ejemplo [Recordatorios de tareas automatizadas.](../resources/scenarios/task-reminders.md)</span><span class="sxs-lookup"><span data-stu-id="4cffc-130">If you'd like a hands-on approach to learning this topic, try out the [Pass data to scripts in an automatically-run Power Automate flow](../tutorials/excel-power-automate-trigger.md) tutorial or explore the [Automated task reminders](../resources/scenarios/task-reminders.md) sample scenario.</span></span>

### <a name="main-parameters-pass-data-to-a-script"></a><span data-ttu-id="4cffc-131">`main` Parámetros: pasar datos a un script</span><span class="sxs-lookup"><span data-stu-id="4cffc-131">`main` Parameters: Pass data to a script</span></span>

<span data-ttu-id="4cffc-132">Toda la entrada de script se especifica como parámetros adicionales para la `main` función.</span><span class="sxs-lookup"><span data-stu-id="4cffc-132">All script input is specified as additional parameters for the `main` function.</span></span> <span data-ttu-id="4cffc-133">Por ejemplo, si desea que un script acepte un `string` que represente un nombre como entrada, cambiaría la `main` firma a `function main(workbook: ExcelScript.Workbook, name: string)` .</span><span class="sxs-lookup"><span data-stu-id="4cffc-133">For example, if you wanted a script to accept a `string` that represents a name as input, you would change the `main` signature to `function main(workbook: ExcelScript.Workbook, name: string)`.</span></span>

<span data-ttu-id="4cffc-134">Al configurar un flujo en Power Automate, puede especificar la entrada de script como valores [estáticos, expresiones](/power-automate/use-expressions-in-conditions)o contenido dinámico.</span><span class="sxs-lookup"><span data-stu-id="4cffc-134">When you're configuring a flow in Power Automate, you can specify script input as static values, [expressions](/power-automate/use-expressions-in-conditions), or dynamic content.</span></span> <span data-ttu-id="4cffc-135">Los detalles del conector de un servicio individual se pueden encontrar en la documentación del [conector Power Automate.](/connectors/)</span><span class="sxs-lookup"><span data-stu-id="4cffc-135">Details on an individual service's connector can be found in the [Power Automate Connector documentation](/connectors/).</span></span>

<span data-ttu-id="4cffc-136">Al agregar parámetros de entrada a la función de un `main` script, tenga en cuenta las siguientes asignaciones y restricciones.</span><span class="sxs-lookup"><span data-stu-id="4cffc-136">When adding input parameters to a script's `main` function, consider the following allowances and restrictions.</span></span>

1. <span data-ttu-id="4cffc-137">El primer parámetro debe ser de tipo `ExcelScript.Workbook` .</span><span class="sxs-lookup"><span data-stu-id="4cffc-137">The first parameter must be of type `ExcelScript.Workbook`.</span></span> <span data-ttu-id="4cffc-138">Su nombre de parámetro no importa.</span><span class="sxs-lookup"><span data-stu-id="4cffc-138">Its parameter name does not matter.</span></span>

2. <span data-ttu-id="4cffc-139">Cada parámetro debe tener un tipo (por `string` ejemplo, o `number` ).</span><span class="sxs-lookup"><span data-stu-id="4cffc-139">Every parameter must have a type (such as `string` or `number`).</span></span>

3. <span data-ttu-id="4cffc-140">Los tipos básicos `string` , , , , y se `number` `boolean` `unknown` `object` `undefined` admiten.</span><span class="sxs-lookup"><span data-stu-id="4cffc-140">The basic types `string`, `number`, `boolean`, `unknown`, `object`, and `undefined` are supported.</span></span>

4. <span data-ttu-id="4cffc-141">Se admiten matrices de los tipos básicos enumerados anteriormente.</span><span class="sxs-lookup"><span data-stu-id="4cffc-141">Arrays of the previously listed basic types are supported.</span></span>

5. <span data-ttu-id="4cffc-142">Las matrices anidadas se admiten como parámetros (pero no como tipos de valor devuelto).</span><span class="sxs-lookup"><span data-stu-id="4cffc-142">Nested arrays are supported as parameters (but not as return types).</span></span>

6. <span data-ttu-id="4cffc-143">Los tipos de unión se permiten si son una unión de literales que pertenecen a un único tipo (por `"Left" | "Right"` ejemplo).</span><span class="sxs-lookup"><span data-stu-id="4cffc-143">Union types are allowed if they are a union of literals belonging to a single type (such as `"Left" | "Right"`).</span></span> <span data-ttu-id="4cffc-144">También se admiten uniones de tipo admitido con indefinidos (por `string | undefined` ejemplo).</span><span class="sxs-lookup"><span data-stu-id="4cffc-144">Unions of a supported type with undefined are also supported (such as `string | undefined`).</span></span>

7. <span data-ttu-id="4cffc-145">Los tipos de objeto se permiten si contienen propiedades de tipo `string` `number` , , `boolean` matrices admitidas u otros objetos admitidos.</span><span class="sxs-lookup"><span data-stu-id="4cffc-145">Object types are allowed if they contain properties of type `string`, `number`, `boolean`, supported arrays, or other supported objects.</span></span> <span data-ttu-id="4cffc-146">En el ejemplo siguiente se muestran los objetos anidados que se admiten como tipos de parámetro:</span><span class="sxs-lookup"><span data-stu-id="4cffc-146">The following example shows nested objects that are supported as parameter types:</span></span>

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

8. <span data-ttu-id="4cffc-147">Los objetos deben tener definida su interfaz o definición de clase en el script.</span><span class="sxs-lookup"><span data-stu-id="4cffc-147">Objects must have their interface or class definition defined in the script.</span></span> <span data-ttu-id="4cffc-148">Un objeto también se puede definir de forma anónima en línea, como en el ejemplo siguiente:</span><span class="sxs-lookup"><span data-stu-id="4cffc-148">An object can also be defined anonymously inline, as in the following example:</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook): {name: string, email: string}
    ```

9. <span data-ttu-id="4cffc-149">Los parámetros opcionales están permitidos y se pueden denota como tales mediante el modificador opcional `?` (por ejemplo, `function main(workbook: ExcelScript.Workbook, Name?: string)` ).</span><span class="sxs-lookup"><span data-stu-id="4cffc-149">Optional parameters are allowed and can be denoted as such by using the optional modifier `?` (for example, `function main(workbook: ExcelScript.Workbook, Name?: string)`).</span></span>

10. <span data-ttu-id="4cffc-150">Se permiten valores de parámetro predeterminados (por `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')` ejemplo.</span><span class="sxs-lookup"><span data-stu-id="4cffc-150">Default parameter values are allowed (for example `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')`.</span></span>

### <a name="return-data-from-a-script"></a><span data-ttu-id="4cffc-151">Devolver datos de un script</span><span class="sxs-lookup"><span data-stu-id="4cffc-151">Return data from a script</span></span>

<span data-ttu-id="4cffc-152">Los scripts pueden devolver datos del libro que se usarán como contenido dinámico en un flujo de Power Automate.</span><span class="sxs-lookup"><span data-stu-id="4cffc-152">Scripts can return data from the workbook to be used as dynamic content in a Power Automate flow.</span></span> <span data-ttu-id="4cffc-153">Al igual que con los parámetros de entrada, Power Automate impone algunas restricciones al tipo de valor devuelto.</span><span class="sxs-lookup"><span data-stu-id="4cffc-153">As with input parameters, Power Automate places some restrictions on the return type.</span></span>

1. <span data-ttu-id="4cffc-154">Los tipos básicos `string` , , , y se `number` `boolean` `void` `undefined` admiten.</span><span class="sxs-lookup"><span data-stu-id="4cffc-154">The basic types `string`, `number`, `boolean`, `void`, and `undefined` are supported.</span></span>

2. <span data-ttu-id="4cffc-155">Los tipos de unión utilizados como tipos de valor devuelto siguen las mismas restricciones que cuando se usan como parámetros de script.</span><span class="sxs-lookup"><span data-stu-id="4cffc-155">Union types used as return types follow the same restrictions as they do when used as script parameters.</span></span>

3. <span data-ttu-id="4cffc-156">Los tipos de matriz se permiten si son de tipo `string` `number` , o `boolean` .</span><span class="sxs-lookup"><span data-stu-id="4cffc-156">Array types are allowed if they are of type `string`, `number`, or `boolean`.</span></span> <span data-ttu-id="4cffc-157">También se permiten si el tipo es una unión compatible o un tipo literal admitido.</span><span class="sxs-lookup"><span data-stu-id="4cffc-157">They are also allowed if the type is a supported union or supported literal type.</span></span>

4. <span data-ttu-id="4cffc-158">Los tipos de objeto utilizados como tipos de valor devuelto siguen las mismas restricciones que cuando se usan como parámetros de script.</span><span class="sxs-lookup"><span data-stu-id="4cffc-158">Object types used as return types follow the same restrictions as they do when used as script parameters.</span></span>

5. <span data-ttu-id="4cffc-159">Se admite la escritura implícita, aunque debe seguir las mismas reglas que un tipo definido.</span><span class="sxs-lookup"><span data-stu-id="4cffc-159">Implicit typing is supported, though it must follow the same rules as a defined type.</span></span>

## <a name="example"></a><span data-ttu-id="4cffc-160">Ejemplo</span><span class="sxs-lookup"><span data-stu-id="4cffc-160">Example</span></span>

<span data-ttu-id="4cffc-161">La siguiente captura de pantalla muestra un flujo de Power Automate que se desencadena cada vez que se le asigna un problema [de GitHub.](https://github.com/)</span><span class="sxs-lookup"><span data-stu-id="4cffc-161">The following screenshot shows a Power Automate flow that's triggered whenever a [GitHub](https://github.com/) issue is assigned to you.</span></span> <span data-ttu-id="4cffc-162">El flujo ejecuta un script que agrega el problema a una tabla de un libro de Excel.</span><span class="sxs-lookup"><span data-stu-id="4cffc-162">The flow runs a script that adds the issue to a table in an Excel workbook.</span></span> <span data-ttu-id="4cffc-163">Si hay cinco o más problemas en esa tabla, el flujo envía un recordatorio por correo electrónico.</span><span class="sxs-lookup"><span data-stu-id="4cffc-163">If there are five or more issues in that table, the flow sends an email reminder.</span></span>

:::image type="content" source="../images/power-automate-parameter-return-sample.png" alt-text="El editor de flujo Power Automate que muestra el flujo de ejemplo":::

<span data-ttu-id="4cffc-165">La `main` función del script especifica el id.</span><span class="sxs-lookup"><span data-stu-id="4cffc-165">The `main` function of the script specifies the issue ID and issue title as input parameters, and the script returns the number of rows in the issue table.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="4cffc-166">Vea también</span><span class="sxs-lookup"><span data-stu-id="4cffc-166">See also</span></span>

- [<span data-ttu-id="4cffc-167">Ejecute scripts de Office en Excel en la Web con Power Automate</span><span class="sxs-lookup"><span data-stu-id="4cffc-167">Run Office Scripts in Excel on the web with Power Automate</span></span>](../tutorials/excel-power-automate-manual.md)
- [<span data-ttu-id="4cffc-168">Pasar datos a scripts en un flujo de Power Automate ejecutado automáticamente</span><span class="sxs-lookup"><span data-stu-id="4cffc-168">Pass data to scripts in an automatically-run Power Automate flow</span></span>](../tutorials/excel-power-automate-trigger.md)
- [<span data-ttu-id="4cffc-169">Devolver datos de un script a un flujo de Power Automate ejecutado automáticamente</span><span class="sxs-lookup"><span data-stu-id="4cffc-169">Return data from a script to an automatically-run Power Automate flow</span></span>](../tutorials/excel-power-automate-returns.md)
- [<span data-ttu-id="4cffc-170">Información de solución de problemas para Power Automate con scripts de Office</span><span class="sxs-lookup"><span data-stu-id="4cffc-170">Troubleshooting information for Power Automate with Office Scripts</span></span>](../testing/power-automate-troubleshooting.md)
- [<span data-ttu-id="4cffc-171">Introducción a Power Automate</span><span class="sxs-lookup"><span data-stu-id="4cffc-171">Get started with Power Automate</span></span>](/power-automate/getting-started)
- [<span data-ttu-id="4cffc-172">Excel Documentación de referencia del conector en línea (business)</span><span class="sxs-lookup"><span data-stu-id="4cffc-172">Excel Online (Business) connector reference documentation</span></span>](/connectors/excelonlinebusiness/)
