---
title: Solución de problemas de scripts Office
description: Sugerencias y técnicas de depuración para scripts de Office, así como recursos de ayuda.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: ff0ac1e63084c7c541d2a4925f1f011d16fa4992
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545562"
---
# <a name="troubleshoot-office-scripts"></a><span data-ttu-id="40e30-103">Solución de problemas de scripts Office</span><span class="sxs-lookup"><span data-stu-id="40e30-103">Troubleshoot Office Scripts</span></span>

<span data-ttu-id="40e30-104">A medida que desarrolla Office scripts, puede cometer errores.</span><span class="sxs-lookup"><span data-stu-id="40e30-104">As you develop Office Scripts, you may make mistakes.</span></span> <span data-ttu-id="40e30-105">Está bien.</span><span class="sxs-lookup"><span data-stu-id="40e30-105">It's okay.</span></span> <span data-ttu-id="40e30-106">Usted tiene las herramientas para ayudar a encontrar los problemas y hacer que sus scripts funcionen perfectamente.</span><span class="sxs-lookup"><span data-stu-id="40e30-106">You have the tools to help find the problems and get your scripts working perfectly.</span></span>

## <a name="types-of-errors"></a><span data-ttu-id="40e30-107">Tipos de errores</span><span class="sxs-lookup"><span data-stu-id="40e30-107">Types of errors</span></span>

<span data-ttu-id="40e30-108">Office Los errores de scripts se dividen en una de las dos categorías:</span><span class="sxs-lookup"><span data-stu-id="40e30-108">Office Scripts errors fall into one of two categories:</span></span>

* <span data-ttu-id="40e30-109">Compilar errores o advertencias en tiempo de compilación</span><span class="sxs-lookup"><span data-stu-id="40e30-109">Compile-time errors or warnings</span></span>
* <span data-ttu-id="40e30-110">Errores en tiempo de ejecución</span><span class="sxs-lookup"><span data-stu-id="40e30-110">Runtime errors</span></span>

### <a name="compile-time-errors"></a><span data-ttu-id="40e30-111">Errores de tiempo de compilación</span><span class="sxs-lookup"><span data-stu-id="40e30-111">Compile-time errors</span></span>

<span data-ttu-id="40e30-112">Los errores y advertencias en tiempo de compilación se muestran inicialmente en el Editor de código.</span><span class="sxs-lookup"><span data-stu-id="40e30-112">Compile-time errors and warnings are initially shown in the Code Editor.</span></span> <span data-ttu-id="40e30-113">Estos se muestran por los subrayados rojos ondulados en el editor.</span><span class="sxs-lookup"><span data-stu-id="40e30-113">These are shown by the wavy red underlines in the editor.</span></span> <span data-ttu-id="40e30-114">También se muestran en la pestaña **Problemas** en la parte inferior del panel de tareas Editor de código.</span><span class="sxs-lookup"><span data-stu-id="40e30-114">They are also displayed under the **Problems** tab at the bottom of the Code Editor task pane.</span></span> <span data-ttu-id="40e30-115">La selección del error dará más detalles sobre el problema y sugerirá soluciones.</span><span class="sxs-lookup"><span data-stu-id="40e30-115">Selecting the error will give more details about the problem and suggest solutions.</span></span> <span data-ttu-id="40e30-116">Los errores en tiempo de compilación deben solucionarse antes de ejecutar el script.</span><span class="sxs-lookup"><span data-stu-id="40e30-116">Compile-time errors should be addressed before running the script.</span></span>

:::image type="content" source="../images/explicit-any-editor-message.png" alt-text="Error del compilador que se muestra en el texto flotante del Editor de código":::

<span data-ttu-id="40e30-118">También puede ver subrayados de advertencia naranja y mensajes informativos grises.</span><span class="sxs-lookup"><span data-stu-id="40e30-118">You may also see orange warning underlines and grey informational messages.</span></span> <span data-ttu-id="40e30-119">Estos indican sugerencias de rendimiento u otras posibilidades donde el script puede tener efectos involuntarios.</span><span class="sxs-lookup"><span data-stu-id="40e30-119">These indicate performance suggestions or other possibilities where the script may have unintentional effects.</span></span> <span data-ttu-id="40e30-120">Estas advertencias deben examinarse detenidamente antes de desestimarlas.</span><span class="sxs-lookup"><span data-stu-id="40e30-120">Such warnings should be examined closely before dismissing them.</span></span>

### <a name="runtime-errors"></a><span data-ttu-id="40e30-121">Errores en tiempo de ejecución</span><span class="sxs-lookup"><span data-stu-id="40e30-121">Runtime errors</span></span>

<span data-ttu-id="40e30-122">Los errores en tiempo de ejecución se producen debido a problemas lógicos en el script.</span><span class="sxs-lookup"><span data-stu-id="40e30-122">Runtime errors happen because of logic issues in the script.</span></span> <span data-ttu-id="40e30-123">Esto podría deberse a que un objeto utilizado en el script no está en el libro, una tabla tiene un formato diferente al previsto o alguna otra discrepancia leve entre los requisitos del script y el libro de trabajo actual.</span><span class="sxs-lookup"><span data-stu-id="40e30-123">This could be because an object used in the script isn't in the workbook, a table is formatted differently than anticipated, or some other slight discrepancy between the script's requirements and the current workbook.</span></span> <span data-ttu-id="40e30-124">El siguiente script genera un error cuando una hoja de cálculo denominada "TestSheet" no está presente.</span><span class="sxs-lookup"><span data-stu-id="40e30-124">The following script generates an error when a worksheet named "TestSheet" is not present.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  let mySheet = workbook.getWorksheet('TestSheet');

  // This will throw an error if there is no "TestSheet".
  mySheet.getRange("A1");
}
```

### <a name="console-messages"></a><span data-ttu-id="40e30-125">Mensajes de consola</span><span class="sxs-lookup"><span data-stu-id="40e30-125">Console messages</span></span>

<span data-ttu-id="40e30-126">Los errores en tiempo de compilación y tiempo de ejecución muestran mensajes de error en la consola cuando se ejecuta un script.</span><span class="sxs-lookup"><span data-stu-id="40e30-126">Both compile-time and runtime errors display error messages in the console when a script runs.</span></span> <span data-ttu-id="40e30-127">Dan un número de línea donde se encontró el problema.</span><span class="sxs-lookup"><span data-stu-id="40e30-127">They give a line number where the problem was encountered.</span></span> <span data-ttu-id="40e30-128">Tenga en cuenta que la causa raíz de cualquier problema puede ser una línea de código diferente a la indicada en la consola.</span><span class="sxs-lookup"><span data-stu-id="40e30-128">Keep in mind that the root cause of any issue may be a different line of code than what is indicated in the console.</span></span>

<span data-ttu-id="40e30-129">La siguiente imagen muestra la salida de la consola para el error [explícito `any` ](../develop/typescript-restrictions.md) del compilador.</span><span class="sxs-lookup"><span data-stu-id="40e30-129">The following image shows the console output for the [explicit `any`](../develop/typescript-restrictions.md) compiler error.</span></span> <span data-ttu-id="40e30-130">Anote el texto `[5, 16]` al principio de la cadena de error.</span><span class="sxs-lookup"><span data-stu-id="40e30-130">Note the text `[5, 16]` at the beginning of the error string.</span></span> <span data-ttu-id="40e30-131">Esto indica que el error está en la línea 5, comenzando en el carácter 16.</span><span class="sxs-lookup"><span data-stu-id="40e30-131">This indicates the error is on line 5, starting at character 16.</span></span>
:::image type="content" source="../images/explicit-any-error-message.png" alt-text="La consola del Editor de código que muestra un mensaje de error explícito de &quot;cualquier&quot;":::

<span data-ttu-id="40e30-133">La siguiente imagen muestra la salida de la consola para un error en tiempo de ejecución.</span><span class="sxs-lookup"><span data-stu-id="40e30-133">The follow image shows the console output for a runtime error.</span></span> <span data-ttu-id="40e30-134">Aquí, el script intenta agregar una hoja de cálculo con un nombre de una hoja de cálculo existente.</span><span class="sxs-lookup"><span data-stu-id="40e30-134">Here, the script tries to add a worksheet with a the name of an existing worksheet.</span></span> <span data-ttu-id="40e30-135">Una vez más, observe la "Línea 2" anterior al error para mostrar qué línea investigar.</span><span class="sxs-lookup"><span data-stu-id="40e30-135">Again, note the "Line 2" preceding the error to show which line to investigate.</span></span>
:::image type="content" source="../images/runtime-error-console.png" alt-text="La consola del Editor de código que muestra un error de la llamada a 'addWorksheet'":::

## <a name="console-logs"></a><span data-ttu-id="40e30-137">Registros de consola</span><span class="sxs-lookup"><span data-stu-id="40e30-137">Console logs</span></span>

<span data-ttu-id="40e30-138">Imprima mensajes en la pantalla con la `console.log` instrucción.</span><span class="sxs-lookup"><span data-stu-id="40e30-138">Print messages to the screen with the `console.log` statement.</span></span> <span data-ttu-id="40e30-139">Estos registros pueden mostrar el valor actual de las variables o qué rutas de código se están desencadenando.</span><span class="sxs-lookup"><span data-stu-id="40e30-139">These logs can show you the current value of variables or which code paths are being triggered.</span></span> <span data-ttu-id="40e30-140">Para ello, llame `console.log` con cualquier objeto como parámetro.</span><span class="sxs-lookup"><span data-stu-id="40e30-140">To do this, call `console.log` with any object as a parameter.</span></span> <span data-ttu-id="40e30-141">Por lo general, a `string` es el tipo más fácil de leer en la consola.</span><span class="sxs-lookup"><span data-stu-id="40e30-141">Usually, a `string` is the easiest type to read in the console.</span></span>

```TypeScript
console.log("Logging myRange's address.");
console.log(myRange.getAddress());
```

<span data-ttu-id="40e30-142">Las cadenas a las que `console.log` se pasa se muestran en la consola de registro del Editor de código, en la parte inferior del panel de tareas.</span><span class="sxs-lookup"><span data-stu-id="40e30-142">Strings passed to `console.log` are displayed in the Code Editor's logging console, at the bottom of the task pane.</span></span> <span data-ttu-id="40e30-143">Los registros se encuentran en la pestaña **Salida,** aunque la pestaña gana automáticamente el foco cuando se escribe un registro.</span><span class="sxs-lookup"><span data-stu-id="40e30-143">Logs are found on the **Output** tab, though the tab automatically gains focus when a log is written.</span></span>

<span data-ttu-id="40e30-144">Los registros no afectan al libro.</span><span class="sxs-lookup"><span data-stu-id="40e30-144">Logs do not affect the workbook.</span></span>

## <a name="automate-tab-not-appearing-or-office-scripts-unavailable"></a><span data-ttu-id="40e30-145">Automatice la pestaña que no aparece o Office scripts no disponibles</span><span class="sxs-lookup"><span data-stu-id="40e30-145">Automate tab not appearing or Office Scripts unavailable</span></span>

<span data-ttu-id="40e30-146">Los pasos siguientes deben ayudar a solucionar cualquier problema relacionado con la pestaña **Automatizar** que no aparezca en Excel en la Web.</span><span class="sxs-lookup"><span data-stu-id="40e30-146">The following steps should help troubleshoot any problems related to the **Automate** tab not appearing in Excel on the web.</span></span>

1. <span data-ttu-id="40e30-147">[Asegúrese de que la licencia de Microsoft 365 incluya scripts Office](../overview/excel.md#requirements).</span><span class="sxs-lookup"><span data-stu-id="40e30-147">[Make sure your Microsoft 365 license includes Office Scripts](../overview/excel.md#requirements).</span></span>
1. <span data-ttu-id="40e30-148">[Compruebe que su navegador es compatible.](platform-limits.md#browser-support)</span><span class="sxs-lookup"><span data-stu-id="40e30-148">[Check that your browser is supported](platform-limits.md#browser-support).</span></span>
1. <span data-ttu-id="40e30-149">[Asegúrese de que las cookies de terceros estén habilitadas.](platform-limits.md#third-party-cookies)</span><span class="sxs-lookup"><span data-stu-id="40e30-149">[Ensure third-party cookies are enabled](platform-limits.md#third-party-cookies).</span></span>
1. <span data-ttu-id="40e30-150">[Asegúrese de que el administrador no ha deshabilitado Office scripts en el Centro de administración de Microsoft 365.](/microsoft-365/admin/manage/manage-office-scripts-settings)</span><span class="sxs-lookup"><span data-stu-id="40e30-150">[Ensure that your admin has not disabled Office Scripts in the Microsoft 365 admin center](/microsoft-365/admin/manage/manage-office-scripts-settings).</span></span>

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

## <a name="troubleshoot-scripts-in-power-automate"></a><span data-ttu-id="40e30-151">Solucionar problemas de scripts en Power Automate</span><span class="sxs-lookup"><span data-stu-id="40e30-151">Troubleshoot scripts in Power Automate</span></span>

<span data-ttu-id="40e30-152">Para obtener información específica para ejecutar scripts a través de Power Automate, consulte [Solución de problemas Office scripts que se ejecutan en Power Automate](power-automate-troubleshooting.md).</span><span class="sxs-lookup"><span data-stu-id="40e30-152">For information specific to running scripts through Power Automate, see [Troubleshoot Office Scripts running in Power Automate](power-automate-troubleshooting.md).</span></span>

## <a name="help-resources"></a><span data-ttu-id="40e30-153">Ayudar a los recursos</span><span class="sxs-lookup"><span data-stu-id="40e30-153">Help resources</span></span>

<span data-ttu-id="40e30-154">[Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts) es una comunidad de desarrolladores dispuestos a ayudar con los problemas de codificación.</span><span class="sxs-lookup"><span data-stu-id="40e30-154">[Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts) is a community of developers willing to help with coding problems.</span></span> <span data-ttu-id="40e30-155">A menudo, podrás encontrar la solución a tu problema a través de una búsqueda rápida de Stack Overflow.</span><span class="sxs-lookup"><span data-stu-id="40e30-155">Often, you'll be able to find the solution to your problem through a quick Stack Overflow search.</span></span> <span data-ttu-id="40e30-156">Si no es así, haga su pregunta y etiquete con la etiqueta "office-scripts".</span><span class="sxs-lookup"><span data-stu-id="40e30-156">If not, ask your question and tag it with the "office-scripts" tag.</span></span> <span data-ttu-id="40e30-157">Asegúrese de mencionar que está creando un *script* Office, no un *complemento Office*.</span><span class="sxs-lookup"><span data-stu-id="40e30-157">Be sure to mention you're creating an Office *Script*, not an Office *Add-in*.</span></span>

<span data-ttu-id="40e30-158">Si tiene un problema con la API de JavaScript Office, cree un problema en el repositorio de [officedev/office-js](https://github.com/OfficeDev/office-js) GitHub.</span><span class="sxs-lookup"><span data-stu-id="40e30-158">If you encounter a problem with the Office JavaScript API, create an issue in the [OfficeDev/office-js](https://github.com/OfficeDev/office-js) GitHub repository.</span></span> <span data-ttu-id="40e30-159">Los miembros del equipo del producto responderán a las cuestiones y proporcionarán más asistencia.</span><span class="sxs-lookup"><span data-stu-id="40e30-159">Members of the product team will respond to issues and provide further assistance.</span></span> <span data-ttu-id="40e30-160">La creación de un problema en el repositorio **OfficeDev/office-js** indica que ha encontrado un defecto en la biblioteca de API de JavaScript Office que el equipo del producto debe abordar.</span><span class="sxs-lookup"><span data-stu-id="40e30-160">Creating an issue in the **OfficeDev/office-js** repository indicates you have found a flaw in the Office JavaScript API library that the product team should address.</span></span>

<span data-ttu-id="40e30-161">Si hay un problema con el Grabador de acciones o editor, envíe comentarios a través del botón **Ayuda > Comentarios** en Excel.</span><span class="sxs-lookup"><span data-stu-id="40e30-161">If there is a problem with the Action Recorder or Editor, send feedback through the **Help > Feedback** button in Excel.</span></span>

## <a name="see-also"></a><span data-ttu-id="40e30-162">Vea también</span><span class="sxs-lookup"><span data-stu-id="40e30-162">See also</span></span>

- [<span data-ttu-id="40e30-163">Procedimientos recomendados para Scripts de Office</span><span class="sxs-lookup"><span data-stu-id="40e30-163">Best practices in Office Scripts</span></span>](../develop/best-practices.md)
- [<span data-ttu-id="40e30-164">Límites de plataforma con scripts de Office</span><span class="sxs-lookup"><span data-stu-id="40e30-164">Platform limits with Office Scripts</span></span>](platform-limits.md)
- [<span data-ttu-id="40e30-165">Mejore el rendimiento de sus scripts de Office</span><span class="sxs-lookup"><span data-stu-id="40e30-165">Improve the performance of your Office Scripts</span></span>](../develop/web-client-performance.md)
- [<span data-ttu-id="40e30-166">Solucionar problemas Office scripts que se ejecutan en PowerAutomate</span><span class="sxs-lookup"><span data-stu-id="40e30-166">Troubleshoot Office Scripts running in PowerAutomate</span></span>](power-automate-troubleshooting.md)
- [<span data-ttu-id="40e30-167">Deshacer los efectos de Scripts de Office</span><span class="sxs-lookup"><span data-stu-id="40e30-167">Undo the effects of Office Scripts</span></span>](undo.md)
