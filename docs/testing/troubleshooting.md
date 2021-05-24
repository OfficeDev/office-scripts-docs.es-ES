---
title: Solucionar problemas Office scripts
description: Sugerencias y técnicas de depuración para Office scripts, así como recursos de ayuda.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: ff0ac1e63084c7c541d2a4925f1f011d16fa4992
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545562"
---
# <a name="troubleshoot-office-scripts"></a><span data-ttu-id="7fe9f-103">Solucionar problemas Office scripts</span><span class="sxs-lookup"><span data-stu-id="7fe9f-103">Troubleshoot Office Scripts</span></span>

<span data-ttu-id="7fe9f-104">A medida que desarrolla Office scripts, puede cometer errores.</span><span class="sxs-lookup"><span data-stu-id="7fe9f-104">As you develop Office Scripts, you may make mistakes.</span></span> <span data-ttu-id="7fe9f-105">Está bien.</span><span class="sxs-lookup"><span data-stu-id="7fe9f-105">It's okay.</span></span> <span data-ttu-id="7fe9f-106">Tiene las herramientas para ayudar a encontrar los problemas y hacer que los scripts funcionen perfectamente.</span><span class="sxs-lookup"><span data-stu-id="7fe9f-106">You have the tools to help find the problems and get your scripts working perfectly.</span></span>

## <a name="types-of-errors"></a><span data-ttu-id="7fe9f-107">Tipos de errores</span><span class="sxs-lookup"><span data-stu-id="7fe9f-107">Types of errors</span></span>

<span data-ttu-id="7fe9f-108">Office Los errores de scripts se ensoyen en una de dos categorías:</span><span class="sxs-lookup"><span data-stu-id="7fe9f-108">Office Scripts errors fall into one of two categories:</span></span>

* <span data-ttu-id="7fe9f-109">Errores o advertencias en tiempo de compilación</span><span class="sxs-lookup"><span data-stu-id="7fe9f-109">Compile-time errors or warnings</span></span>
* <span data-ttu-id="7fe9f-110">Errores en tiempo de ejecución</span><span class="sxs-lookup"><span data-stu-id="7fe9f-110">Runtime errors</span></span>

### <a name="compile-time-errors"></a><span data-ttu-id="7fe9f-111">Errores en tiempo de compilación</span><span class="sxs-lookup"><span data-stu-id="7fe9f-111">Compile-time errors</span></span>

<span data-ttu-id="7fe9f-112">Los errores y advertencias en tiempo de compilación se muestran inicialmente en el Editor de código.</span><span class="sxs-lookup"><span data-stu-id="7fe9f-112">Compile-time errors and warnings are initially shown in the Code Editor.</span></span> <span data-ttu-id="7fe9f-113">Estos se muestran con los subrayados rojos ondulados del editor.</span><span class="sxs-lookup"><span data-stu-id="7fe9f-113">These are shown by the wavy red underlines in the editor.</span></span> <span data-ttu-id="7fe9f-114">También se muestran en la pestaña **Problemas** en la parte inferior del panel de tareas Editor de código.</span><span class="sxs-lookup"><span data-stu-id="7fe9f-114">They are also displayed under the **Problems** tab at the bottom of the Code Editor task pane.</span></span> <span data-ttu-id="7fe9f-115">Al seleccionar el error, se darán más detalles sobre el problema y se sugerirán soluciones.</span><span class="sxs-lookup"><span data-stu-id="7fe9f-115">Selecting the error will give more details about the problem and suggest solutions.</span></span> <span data-ttu-id="7fe9f-116">Los errores en tiempo de compilación deben solucionarse antes de ejecutar el script.</span><span class="sxs-lookup"><span data-stu-id="7fe9f-116">Compile-time errors should be addressed before running the script.</span></span>

:::image type="content" source="../images/explicit-any-editor-message.png" alt-text="Error del compilador que se muestra en el texto activado del Editor de código":::

<span data-ttu-id="7fe9f-118">También puede ver subrayados de advertencia naranja y mensajes informativos grises.</span><span class="sxs-lookup"><span data-stu-id="7fe9f-118">You may also see orange warning underlines and grey informational messages.</span></span> <span data-ttu-id="7fe9f-119">Estas indican sugerencias de rendimiento u otras posibilidades en las que el script puede tener efectos involuntarias.</span><span class="sxs-lookup"><span data-stu-id="7fe9f-119">These indicate performance suggestions or other possibilities where the script may have unintentional effects.</span></span> <span data-ttu-id="7fe9f-120">Estas advertencias deben examinarse estrechamente antes de descartarlas.</span><span class="sxs-lookup"><span data-stu-id="7fe9f-120">Such warnings should be examined closely before dismissing them.</span></span>

### <a name="runtime-errors"></a><span data-ttu-id="7fe9f-121">Errores en tiempo de ejecución</span><span class="sxs-lookup"><span data-stu-id="7fe9f-121">Runtime errors</span></span>

<span data-ttu-id="7fe9f-122">Los errores en tiempo de ejecución se producen debido a problemas de lógica en el script.</span><span class="sxs-lookup"><span data-stu-id="7fe9f-122">Runtime errors happen because of logic issues in the script.</span></span> <span data-ttu-id="7fe9f-123">Esto podría deberse a que un objeto usado en el script no está en el libro, una tabla tiene un formato diferente al previsto o alguna otra discrepancia leve entre los requisitos del script y el libro actual.</span><span class="sxs-lookup"><span data-stu-id="7fe9f-123">This could be because an object used in the script isn't in the workbook, a table is formatted differently than anticipated, or some other slight discrepancy between the script's requirements and the current workbook.</span></span> <span data-ttu-id="7fe9f-124">El siguiente script genera un error cuando no está presente una hoja de cálculo denominada "TestSheet".</span><span class="sxs-lookup"><span data-stu-id="7fe9f-124">The following script generates an error when a worksheet named "TestSheet" is not present.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  let mySheet = workbook.getWorksheet('TestSheet');

  // This will throw an error if there is no "TestSheet".
  mySheet.getRange("A1");
}
```

### <a name="console-messages"></a><span data-ttu-id="7fe9f-125">Mensajes de consola</span><span class="sxs-lookup"><span data-stu-id="7fe9f-125">Console messages</span></span>

<span data-ttu-id="7fe9f-126">Tanto los errores en tiempo de compilación como en tiempo de ejecución muestran mensajes de error en la consola cuando se ejecuta un script.</span><span class="sxs-lookup"><span data-stu-id="7fe9f-126">Both compile-time and runtime errors display error messages in the console when a script runs.</span></span> <span data-ttu-id="7fe9f-127">Dan un número de línea donde se encontró el problema.</span><span class="sxs-lookup"><span data-stu-id="7fe9f-127">They give a line number where the problem was encountered.</span></span> <span data-ttu-id="7fe9f-128">Tenga en cuenta que la causa raíz de cualquier problema puede ser una línea de código diferente a la que se indica en la consola.</span><span class="sxs-lookup"><span data-stu-id="7fe9f-128">Keep in mind that the root cause of any issue may be a different line of code than what is indicated in the console.</span></span>

<span data-ttu-id="7fe9f-129">En la imagen siguiente se muestra el resultado de la consola del [error explícito `any` ](../develop/typescript-restrictions.md) del compilador.</span><span class="sxs-lookup"><span data-stu-id="7fe9f-129">The following image shows the console output for the [explicit `any`](../develop/typescript-restrictions.md) compiler error.</span></span> <span data-ttu-id="7fe9f-130">Tenga en cuenta `[5, 16]` el texto al principio de la cadena de error.</span><span class="sxs-lookup"><span data-stu-id="7fe9f-130">Note the text `[5, 16]` at the beginning of the error string.</span></span> <span data-ttu-id="7fe9f-131">Esto indica que el error está en la línea 5, empezando por el carácter 16.</span><span class="sxs-lookup"><span data-stu-id="7fe9f-131">This indicates the error is on line 5, starting at character 16.</span></span>
:::image type="content" source="../images/explicit-any-error-message.png" alt-text="La consola del Editor de código que muestra un mensaje de error explícito de &quot;cualquiera&quot;":::

<span data-ttu-id="7fe9f-133">La imagen siguiente muestra el resultado de la consola de un error en tiempo de ejecución.</span><span class="sxs-lookup"><span data-stu-id="7fe9f-133">The follow image shows the console output for a runtime error.</span></span> <span data-ttu-id="7fe9f-134">Aquí, el script intenta agregar una hoja de cálculo con el nombre de una hoja de cálculo existente.</span><span class="sxs-lookup"><span data-stu-id="7fe9f-134">Here, the script tries to add a worksheet with a the name of an existing worksheet.</span></span> <span data-ttu-id="7fe9f-135">De nuevo, anote la "Línea 2" anterior al error para mostrar la línea que se debe investigar.</span><span class="sxs-lookup"><span data-stu-id="7fe9f-135">Again, note the "Line 2" preceding the error to show which line to investigate.</span></span>
:::image type="content" source="../images/runtime-error-console.png" alt-text="La consola del Editor de código que muestra un error de la llamada &quot;addWorksheet&quot;":::

## <a name="console-logs"></a><span data-ttu-id="7fe9f-137">Registros de consola</span><span class="sxs-lookup"><span data-stu-id="7fe9f-137">Console logs</span></span>

<span data-ttu-id="7fe9f-138">Imprimir mensajes en la pantalla con la `console.log` instrucción.</span><span class="sxs-lookup"><span data-stu-id="7fe9f-138">Print messages to the screen with the `console.log` statement.</span></span> <span data-ttu-id="7fe9f-139">Estos registros pueden mostrar el valor actual de las variables o qué rutas de código se están desencadenando.</span><span class="sxs-lookup"><span data-stu-id="7fe9f-139">These logs can show you the current value of variables or which code paths are being triggered.</span></span> <span data-ttu-id="7fe9f-140">Para ello, llame `console.log` con cualquier objeto como parámetro.</span><span class="sxs-lookup"><span data-stu-id="7fe9f-140">To do this, call `console.log` with any object as a parameter.</span></span> <span data-ttu-id="7fe9f-141">Por lo general, a `string` es el tipo más fácil de leer en la consola.</span><span class="sxs-lookup"><span data-stu-id="7fe9f-141">Usually, a `string` is the easiest type to read in the console.</span></span>

```TypeScript
console.log("Logging myRange's address.");
console.log(myRange.getAddress());
```

<span data-ttu-id="7fe9f-142">Las cadenas pasadas se muestran en la consola de registro del Editor de `console.log` código, en la parte inferior del panel de tareas.</span><span class="sxs-lookup"><span data-stu-id="7fe9f-142">Strings passed to `console.log` are displayed in the Code Editor's logging console, at the bottom of the task pane.</span></span> <span data-ttu-id="7fe9f-143">Los registros se encuentran en la **pestaña Salida,** aunque la pestaña aumenta automáticamente el foco cuando se escribe un registro.</span><span class="sxs-lookup"><span data-stu-id="7fe9f-143">Logs are found on the **Output** tab, though the tab automatically gains focus when a log is written.</span></span>

<span data-ttu-id="7fe9f-144">Los registros no afectan al libro.</span><span class="sxs-lookup"><span data-stu-id="7fe9f-144">Logs do not affect the workbook.</span></span>

## <a name="automate-tab-not-appearing-or-office-scripts-unavailable"></a><span data-ttu-id="7fe9f-145">La pestaña Automatizar no aparece ni Office scripts no están disponibles</span><span class="sxs-lookup"><span data-stu-id="7fe9f-145">Automate tab not appearing or Office Scripts unavailable</span></span>

<span data-ttu-id="7fe9f-146">Los siguientes pasos deben ayudar a solucionar los problemas relacionados con la pestaña **Automatizar** que no aparezcan en Excel en la Web.</span><span class="sxs-lookup"><span data-stu-id="7fe9f-146">The following steps should help troubleshoot any problems related to the **Automate** tab not appearing in Excel on the web.</span></span>

1. <span data-ttu-id="7fe9f-147">[Asegúrese de que su Microsoft 365 incluye Office scripts](../overview/excel.md#requirements).</span><span class="sxs-lookup"><span data-stu-id="7fe9f-147">[Make sure your Microsoft 365 license includes Office Scripts](../overview/excel.md#requirements).</span></span>
1. <span data-ttu-id="7fe9f-148">[Compruebe que el explorador es compatible.](platform-limits.md#browser-support)</span><span class="sxs-lookup"><span data-stu-id="7fe9f-148">[Check that your browser is supported](platform-limits.md#browser-support).</span></span>
1. <span data-ttu-id="7fe9f-149">[Asegúrese de que las cookies de terceros están habilitadas](platform-limits.md#third-party-cookies).</span><span class="sxs-lookup"><span data-stu-id="7fe9f-149">[Ensure third-party cookies are enabled](platform-limits.md#third-party-cookies).</span></span>
1. <span data-ttu-id="7fe9f-150">[Asegúrese de que el administrador no ha deshabilitado Office scripts en el centro Microsoft 365 de administración](/microsoft-365/admin/manage/manage-office-scripts-settings).</span><span class="sxs-lookup"><span data-stu-id="7fe9f-150">[Ensure that your admin has not disabled Office Scripts in the Microsoft 365 admin center](/microsoft-365/admin/manage/manage-office-scripts-settings).</span></span>

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

## <a name="troubleshoot-scripts-in-power-automate"></a><span data-ttu-id="7fe9f-151">Solucionar problemas de scripts en Power Automate</span><span class="sxs-lookup"><span data-stu-id="7fe9f-151">Troubleshoot scripts in Power Automate</span></span>

<span data-ttu-id="7fe9f-152">Para obtener información específica sobre la ejecución de scripts Power Automate, vea [Troubleshoot Office Scripts running in Power Automate](power-automate-troubleshooting.md).</span><span class="sxs-lookup"><span data-stu-id="7fe9f-152">For information specific to running scripts through Power Automate, see [Troubleshoot Office Scripts running in Power Automate](power-automate-troubleshooting.md).</span></span>

## <a name="help-resources"></a><span data-ttu-id="7fe9f-153">Recursos de ayuda</span><span class="sxs-lookup"><span data-stu-id="7fe9f-153">Help resources</span></span>

<span data-ttu-id="7fe9f-154">[Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts) es una comunidad de desarrolladores dispuestos a ayudar con problemas de codificación.</span><span class="sxs-lookup"><span data-stu-id="7fe9f-154">[Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts) is a community of developers willing to help with coding problems.</span></span> <span data-ttu-id="7fe9f-155">A menudo, podrás encontrar la solución al problema mediante una búsqueda rápida de desbordamiento de pila.</span><span class="sxs-lookup"><span data-stu-id="7fe9f-155">Often, you'll be able to find the solution to your problem through a quick Stack Overflow search.</span></span> <span data-ttu-id="7fe9f-156">Si no es así, haga su pregunta y etiquete con la etiqueta "office-scripts".</span><span class="sxs-lookup"><span data-stu-id="7fe9f-156">If not, ask your question and tag it with the "office-scripts" tag.</span></span> <span data-ttu-id="7fe9f-157">Asegúrese de mencionar que está creando un script Office *,* no un Office *complemento*.</span><span class="sxs-lookup"><span data-stu-id="7fe9f-157">Be sure to mention you're creating an Office *Script*, not an Office *Add-in*.</span></span>

<span data-ttu-id="7fe9f-158">Si encuentra un problema con la API de JavaScript Office, cree un problema en el repositorio de GitHub [OfficeDev/office-js.](https://github.com/OfficeDev/office-js)</span><span class="sxs-lookup"><span data-stu-id="7fe9f-158">If you encounter a problem with the Office JavaScript API, create an issue in the [OfficeDev/office-js](https://github.com/OfficeDev/office-js) GitHub repository.</span></span> <span data-ttu-id="7fe9f-159">Los miembros del equipo del producto responderán a los problemas y proporcionarán más asistencia.</span><span class="sxs-lookup"><span data-stu-id="7fe9f-159">Members of the product team will respond to issues and provide further assistance.</span></span> <span data-ttu-id="7fe9f-160">La creación de un problema en el repositorio de **OfficeDev/office-js** indica que ha encontrado un error en la biblioteca de la API de JavaScript de Office que el equipo de producto debe solucionar.</span><span class="sxs-lookup"><span data-stu-id="7fe9f-160">Creating an issue in the **OfficeDev/office-js** repository indicates you have found a flaw in the Office JavaScript API library that the product team should address.</span></span>

<span data-ttu-id="7fe9f-161">Si hay un problema con la Grabadora de acciones o el Editor, envía comentarios a través del botón Ayuda **> comentarios** en Excel.</span><span class="sxs-lookup"><span data-stu-id="7fe9f-161">If there is a problem with the Action Recorder or Editor, send feedback through the **Help > Feedback** button in Excel.</span></span>

## <a name="see-also"></a><span data-ttu-id="7fe9f-162">Consulte también</span><span class="sxs-lookup"><span data-stu-id="7fe9f-162">See also</span></span>

- [<span data-ttu-id="7fe9f-163">Procedimientos recomendados para Scripts de Office</span><span class="sxs-lookup"><span data-stu-id="7fe9f-163">Best practices in Office Scripts</span></span>](../develop/best-practices.md)
- [<span data-ttu-id="7fe9f-164">Límites de plataforma con Office scripts</span><span class="sxs-lookup"><span data-stu-id="7fe9f-164">Platform limits with Office Scripts</span></span>](platform-limits.md)
- [<span data-ttu-id="7fe9f-165">Mejorar el rendimiento de los scripts Office scripts</span><span class="sxs-lookup"><span data-stu-id="7fe9f-165">Improve the performance of your Office Scripts</span></span>](../develop/web-client-performance.md)
- [<span data-ttu-id="7fe9f-166">Solucionar Office scripts que se ejecutan en PowerAutomate</span><span class="sxs-lookup"><span data-stu-id="7fe9f-166">Troubleshoot Office Scripts running in PowerAutomate</span></span>](power-automate-troubleshooting.md)
- [<span data-ttu-id="7fe9f-167">Deshacer los efectos de Scripts de Office</span><span class="sxs-lookup"><span data-stu-id="7fe9f-167">Undo the effects of Office Scripts</span></span>](undo.md)
